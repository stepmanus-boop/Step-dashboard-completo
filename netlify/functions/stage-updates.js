
const { jsonResponse, requireSession, normalizeSectorValue } = require('./_auth');
const { readJson, writeJson } = require('./_githubStore');
const { isSupabaseConfigured, listStageUpdates, createStageUpdate, updateStageUpdate } = require('./_supabase');
const { findProjectAndSpool, loadProjectPayload } = require('./_projectLookup');

const DATA_PATH = 'data/stage-updates.json';
const SUPPORTED_SECTORS = ['pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const PROGRESS_OPTIONS = [25, 50, 75, 100];
const PENDING_STATUSES = ['pending', 'pending_advance', 'pending_review'];

const SMARTSHEET_API_BASE = process.env.SMARTSHEET_API_BASE || 'https://api.smartsheet.com/2.0';
const SMARTSHEET_TOKEN = process.env.SMARTSHEET_TOKEN || '5pP36OjBaD1W2HWyxf6aoGxXasPvEl8gbqOmQ';

const TRACKING_UPDATE_COLUMN_BY_SECTOR = {
  pintura: 'Surface preparation and/or coating',
  solda: 'Full welding execution',
  producao: 'Spool Assemble and tack weld',
  calderaria: 'Spool Assemble and tack weld',
  inspecao: 'Final Inspection',
  pendente_envio: 'Package and Delivered',
};

const TRACKING_FIELDS_BY_SECTOR = {
  pintura: ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)'],
  inspecao: ['Final Inspection', 'Hydro Test Pressure (QC)', 'Non Destructive Examination (QC)', 'Final Dimensional Inpection/3D (QC)', 'Initial Dimensional Inspection/3D'],
  pendente_envio: ['Package and Delivered', 'Final Inspection'],
  producao: ['Spool Assemble and tack weld', 'Welding Preparation'],
  calderaria: ['Spool Assemble and tack weld', 'Welding Preparation', 'Material Separation', 'Material Release to Fabrication'],
  solda: ['Full welding execution'],
};

function normalizeColumnTitle(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '')
    .toLowerCase();
}

async function smartsheetFetch(path, options = {}) {
  if (!SMARTSHEET_TOKEN) {
    const err = new Error('SMARTSHEET_TOKEN não configurado.');
    err.statusCode = 500;
    throw err;
  }

  const response = await fetch(`${SMARTSHEET_API_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${SMARTSHEET_TOKEN}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const message = await response.text().catch(() => '');
    const err = new Error(`Smartsheet ${response.status}: ${message || 'Falha ao atualizar tracking.'}`);
    err.statusCode = response.status;
    throw err;
  }

  return response.json().catch(() => ({}));
}

async function getSheetColumns(sheetId) {
  const sheet = await smartsheetFetch(`/sheets/${sheetId}?pageSize=1`);
  return Array.isArray(sheet?.columns) ? sheet.columns : [];
}

function getTrackingUpdateColumnForSector(sector) {
  return TRACKING_UPDATE_COLUMN_BY_SECTOR[normalizeSectorValue(sector)] || '';
}

function findColumnId(columns, columnTitle) {
  const target = normalizeColumnTitle(columnTitle);
  const found = (Array.isArray(columns) ? columns : []).find((column) => normalizeColumnTitle(column?.title) === target);
  return found?.id || null;
}

async function updateSmartsheetCellWithPercent(sheetId, rowId, columnId, progress) {
  const value = Number(progress || 0);

  if (![25, 50, 75, 100].includes(value)) {
    const err = new Error('Valor de avanço inválido. Use apenas 25%, 50%, 75% ou 100%.');
    err.statusCode = 400;
    throw err;
  }

  const percentValue = `${value}%`;

  await smartsheetFetch(`/sheets/${sheetId}/rows`, {
    method: 'PUT',
    body: JSON.stringify([
      {
        id: Number(rowId),
        cells: [
          {
            columnId: Number(columnId),
            value: percentValue,
            strict: false,
          },
        ],
      },
    ]),
  });

  return { ok: true, value: percentValue };
}

async function updateTrackingCellForStageUpdate(update, session) {
  const { project, spool, payload } = await findProjectAndSpool(update.projectRowId, update.spoolIso);
  if (!project || !spool) {
    const err = new Error('BSP ou spool não localizado no Tracking.');
    err.statusCode = 404;
    throw err;
  }

  const columnTitle = getTrackingUpdateColumnForSector(update.sector);
  if (!columnTitle) {
    const err = new Error(`Não existe coluna configurada para atualizar o setor ${update.sector || 'informado'}.`);
    err.statusCode = 400;
    throw err;
  }

  const progress = Number(update.progress || 0);
  const currentProgress = getTrackingProgressForSector(spool, update.sector);
  if (currentProgress != null && currentProgress >= progress) {
    return {
      update: applyTrackingVerification({
        ...update,
        trackingUpdatedBy: session.username || '',
        trackingUpdatedByName: session.name || session.username || 'PCP',
        trackingUpdatedAt: new Date().toISOString(),
        trackingUpdatedColumn: columnTitle,
      }, project, spool),
      applied: false,
      alreadyUpdated: true,
      columnTitle,
      currentProgress,
    };
  }

  const sheetId = payload?.meta?.sheetId;
  if (!sheetId) {
    const err = new Error('Sheet ID do Tracking não localizado.');
    err.statusCode = 500;
    throw err;
  }

  const rowId = Number(spool.rowId || 0);
  if (!rowId) {
    const err = new Error('Linha da spool no Tracking não localizada.');
    err.statusCode = 404;
    throw err;
  }

  const columns = await getSheetColumns(sheetId);
  const columnId = findColumnId(columns, columnTitle);
  if (!columnId) {
    const err = new Error(`Coluna "${columnTitle}" não encontrada no Tracking.`);
    err.statusCode = 404;
    throw err;
  }

  await updateSmartsheetCellWithPercent(sheetId, rowId, columnId, progress);

  const now = new Date().toISOString();
  return {
    update: {
      ...update,
      trackingCheckedAt: now,
      trackingProgress: progress,
      trackingMatched: true,
      trackingStatus: 'matched',
      trackingUpdatedBy: session.username || '',
      trackingUpdatedByName: session.name || session.username || 'PCP',
      trackingUpdatedAt: now,
      trackingUpdatedColumn: columnTitle,
    },
    applied: true,
    alreadyUpdated: false,
    columnTitle,
    currentProgress: progress,
  };
}

function parseTrackingPercent(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return null;
    return value >= 0 && value <= 1 ? value * 100 : value;
  }
  let raw = String(value || '').trim();
  if (!raw) return null;
  raw = raw.replace('%', '').replace(/\s/g, '').replace(',', '.');
  const parsed = Number(raw.replace(/[^\d.-]/g, ''));
  if (!Number.isFinite(parsed)) return null;
  return parsed >= 0 && parsed <= 1 ? parsed * 100 : parsed;
}

function getTrackingProgressForSector(spool, sector) {
  const normalizedSector = normalizeSectorValue(sector);
  const stageValues = spool?.stageValues || {};
  const fields = TRACKING_FIELDS_BY_SECTOR[normalizedSector] || [];
  const values = [];

  for (const field of fields) {
    const parsed = parseTrackingPercent(stageValues[field]);
    if (parsed != null) values.push(parsed);
  }

  const currentSector = normalizeSectorValue(spool?.currentSector || spool?.operationalSector || spool?.flow?.sector);
  if (currentSector === normalizedSector) {
    const stagePercent = parseTrackingPercent(spool?.stagePercent ?? spool?.flow?.percent);
    if (stagePercent != null) values.push(stagePercent);
  }

  if (!values.length) return null;
  return Math.max(...values.map((value) => Math.max(0, Math.min(100, Number(value)))));
}

function findProjectInPayload(projects, projectRowId) {
  const normalizedProjectId = String(projectRowId ?? '').trim();
  return (Array.isArray(projects) ? projects : []).find((item) => {
    const rowId = String(item?.rowId ?? '').trim();
    const rowNumber = String(item?.rowNumber ?? '').trim();
    return (rowId && rowId === normalizedProjectId) || (rowNumber && rowNumber === normalizedProjectId);
  }) || null;
}

function findSpoolInProject(project, spoolIso) {
  const normalizedSpoolIso = String(spoolIso || '').trim().toLowerCase();
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  return spools.find((item) => String(item?.iso || '').trim().toLowerCase() === normalizedSpoolIso) || null;
}

function applyTrackingVerification(update, project, spool) {
  const progress = Number(update?.progress || 0);
  const trackingProgress = getTrackingProgressForSector(spool, update?.sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;

  return {
    ...update,
    trackingCheckedAt: new Date().toISOString(),
    trackingProgress: trackingProgress == null ? null : Number(trackingProgress.toFixed(2)),
    trackingMatched,
    trackingStatus: trackingProgress == null
      ? 'not_found'
      : (trackingMatched ? 'matched' : 'waiting'),
  };
}

async function enrichUpdatesWithTracking(updates) {
  const list = Array.isArray(updates) ? updates : [];
  if (!list.length) return [];

  let payload = null;
  try {
    payload = await loadProjectPayload();
  } catch (_) {
    payload = { projects: [] };
  }
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];

  return list.map((item) => {
    const project = findProjectInPayload(projects, item?.projectRowId);
    const spool = project ? findSpoolInProject(project, item?.spoolIso) : null;
    if (!project || !spool) {
      return {
        ...item,
        trackingCheckedAt: new Date().toISOString(),
        trackingProgress: null,
        trackingMatched: false,
        trackingStatus: 'not_found',
      };
    }
    return applyTrackingVerification(item, project, spool);
  });
}

function canValidate(session) {
  const sector = normalizeSectorValue(session?.sector);
  return session?.role === 'admin' || sector === 'pcp';
}

function canCreate(session) {
  const sector = normalizeSectorValue(session?.sector);
  return SUPPORTED_SECTORS.includes(sector);
}

function getActorSector(session) {
  return normalizeSectorValue(session?.sector);
}

async function listUpdates() {
  if (isSupabaseConfigured()) {
    return listStageUpdates();
  }
  const rows = await readJson(DATA_PATH, []);
  return Array.isArray(rows) ? rows : [];
}

async function saveUpdates(rows) {
  return writeJson(DATA_PATH, rows, 'chore: atualiza apontamentos setoriais');
}

async function createSingleUpdate(payload, session, existingUpdates = null) {
  const projectRowId = Number(payload.projectRowId || 0);
  const spoolIso = String(payload.spoolIso || '').trim();
  const progress = Number(payload.progress || 0);
  const completionDate = String(payload.completionDate || '').trim();
  const note = String(payload.note || '').trim();
  const actionType = String(payload.actionType || 'advance').trim().toLowerCase() === 'review' ? 'review' : 'advance';
  const sector = session.role === 'admin'
    ? normalizeSectorValue(payload.sector || session.sector)
    : getActorSector(session);

  if (!projectRowId || !spoolIso || !SUPPORTED_SECTORS.includes(sector)) {
    throw new Error('Informe BSP, spool e uma etapa válida.');
  }
  if (!PROGRESS_OPTIONS.includes(progress)) {
    throw new Error('Selecione um avanço válido: 25%, 50%, 75% ou 100%.');
  }
  const { project, spool } = await findProjectAndSpool(projectRowId, spoolIso);
  if (!project || !spool) {
    const err = new Error('BSP ou spool não localizado para este apontamento.');
    err.statusCode = 404;
    throw err;
  }
  const trackingProgress = getTrackingProgressForSector(spool, sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;
  const updates = existingUpdates || await listUpdates();
  const pendingExists = updates.find((item) =>
    PENDING_STATUSES.includes(String(item.status || 'pending').trim().toLowerCase())
    && Number(item.projectRowId || 0) === projectRowId
    && String(item.spoolIso || '').trim().toLowerCase() === spoolIso.toLowerCase()
    && normalizeSectorValue(item.sector) === sector
  );
  if (pendingExists) {
    const err = new Error('Já existe um apontamento pendente desta etapa para este spool.');
    err.statusCode = 409;
    throw err;
  }
  const now = new Date().toISOString();
  const record = {
    id: `stg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}` ,
    projectRowId,
    projectNumber: project.projectNumber || project.projectDisplay || `Projeto ${projectRowId}`,
    projectDisplay: project.projectDisplay || project.projectNumber || `Projeto ${projectRowId}`,
    client: project.client || '',
    spoolIso,
    spoolDescription: spool.description || '',
    sector,
    progress,
    completionDate: completionDate || (progress === 100 ? now.slice(0, 10) : ''),
    note,
    status: actionType === 'review' ? 'pending_review' : 'pending_advance',
    trackingCheckedAt: now,
    trackingProgress: trackingProgress == null ? null : Number(trackingProgress.toFixed(2)),
    trackingMatched,
    trackingStatus: trackingProgress == null ? 'not_found' : (trackingMatched ? 'matched' : 'waiting'),
    createdBy: session.username || '',
    createdByName: session.name || session.username || 'Usuário',
    createdAt: now,
    resolvedBy: '',
    resolvedByName: '',
    resolvedAt: '',
    resolutionNote: '',
  };
  if (isSupabaseConfigured()) {
    const saved = await createStageUpdate(record);
    return saved || record;
  }
  updates.unshift(record);
  return record;
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return jsonResponse(200, { ok: true });
  }
  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  const session = auth.session;
  try {
    if (event.httpMethod === 'GET') {
      const updates = await enrichUpdatesWithTracking(await listUpdates());
      return jsonResponse(200, {
        ok: true,
        updates,
        permissions: {
          canCreate: canCreate(session),
          canValidate: canValidate(session),
          sector: getActorSector(session),
        },
        progressOptions: PROGRESS_OPTIONS,
      });
    }

    if (event.httpMethod === 'POST') {
      if (!canCreate(session) && session.role !== 'admin') {
        return jsonResponse(403, { ok: false, error: 'Seu perfil não pode lançar apontamentos setoriais.' });
      }
      const body = JSON.parse(event.body || '{}');
      const items = Array.isArray(body.items) ? body.items : null;
      if (items && items.length) {
        const baseUpdates = isSupabaseConfigured() ? [] : await listUpdates();
        const created = [];
        const errors = [];
        for (const item of items) {
          try {
            const saved = await createSingleUpdate(item, session, baseUpdates);
            created.push(saved);
          } catch (error) {
            errors.push({
              projectRowId: item?.projectRowId || 0,
              spoolIso: item?.spoolIso || '',
              error: error.message || 'Falha ao enviar item.',
            });
          }
        }
        if (!isSupabaseConfigured() && created.length) {
          await saveUpdates(baseUpdates);
        }
        return jsonResponse(created.length ? 200 : 400, {
          ok: created.length > 0,
          updates: created,
          errors,
          storage: isSupabaseConfigured() ? 'supabase' : 'json',
        });
      }
      try {
        const updates = isSupabaseConfigured() ? null : await listUpdates();
        const saved = await createSingleUpdate(body, session, updates);
        if (!isSupabaseConfigured()) {
          await saveUpdates(updates);
        }
        return jsonResponse(200, { ok: true, update: saved, storage: isSupabaseConfigured() ? 'supabase' : 'json' });
      } catch (error) {
        return jsonResponse(error.statusCode || 400, { ok: false, error: error.message || 'Falha ao enviar apontamento.' });
      }
    }

    if (event.httpMethod === 'PATCH') {
      if (!canValidate(session)) {
        return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode concluir apontamentos.' });
      }
      const body = JSON.parse(event.body || '{}');

      if (String(body.action || '').trim().toLowerCase() === 'update_tracking') {
        const requestedIds = Array.isArray(body.ids)
          ? body.ids.map((item) => String(item || '').trim()).filter(Boolean)
          : [String(body.id || '').trim()].filter(Boolean);

        const idsToUpdate = Array.from(new Set(requestedIds));
        if (!idsToUpdate.length) {
          return jsonResponse(400, { ok: false, error: 'Selecione ao menos um apontamento para atualizar o Tracking.' });
        }

        const updates = await listUpdates();
        const updated = [];
        const errors = [];
        const results = [];

        for (const id of idsToUpdate) {
          const index = updates.findIndex((item) => String(item.id) === id);
          if (index < 0) {
            errors.push({ id, error: 'Apontamento não encontrado.' });
            continue;
          }

          const current = updates[index];
          const currentStatus = String(current.status || 'pending').trim().toLowerCase();

          if (!PENDING_STATUSES.includes(currentStatus)) {
            errors.push({ id, error: 'Somente apontamentos pendentes podem atualizar o Tracking.' });
            continue;
          }

          if (currentStatus === 'pending_review') {
            errors.push({ id, error: 'Apontamento de revisão não atualiza percentual do Tracking.' });
            continue;
          }

          try {
            const result = await updateTrackingCellForStageUpdate(current, session);
            const updatedRecord = {
              ...current,
              ...result.update,
            };

            if (!isSupabaseConfigured()) {
              updates[index] = updatedRecord;
            }

            updated.push(updatedRecord);
            results.push({
              id,
              applied: result.applied,
              alreadyUpdated: result.alreadyUpdated,
              columnTitle: result.columnTitle,
            });
          } catch (error) {
            errors.push({ id, error: error.message || 'Falha ao atualizar Tracking.' });
          }
        }

        if (!isSupabaseConfigured() && updated.length) {
          await saveUpdates(updates);
        }

        return jsonResponse(updated.length ? 200 : 400, {
          ok: updated.length > 0,
          error: updated.length ? '' : (errors[0]?.error || 'Falha ao atualizar Tracking em lote.'),
          update: updated[0] || null,
          updates: updated,
          results,
          errors,
          storage: isSupabaseConfigured() ? 'supabase' : 'json',
        });
      }

      const ids = Array.isArray(body.ids) ? body.ids.map((id) => String(id || '').trim()).filter(Boolean) : [];
      const resolutionNote = String(body.resolutionNote || '').trim();
      if (ids.length) {
        const updates = await listUpdates();
        const updated = [];
        for (const id of ids) {
          const index = updates.findIndex((item) => String(item.id) === id);
          if (index < 0) continue;
          const currentStatus = String(updates[index]?.status || 'pending').trim().toLowerCase();
          const nextStatus = currentStatus === 'pending_review' ? 'resolved_review' : 'resolved_advance';
          const updatedRecord = {
            ...updates[index],
            status: nextStatus,
            resolvedBy: session.username || '',
            resolvedByName: session.name || session.username || 'Usuário',
            resolvedAt: new Date().toISOString(),
            resolutionNote,
          };
          if (isSupabaseConfigured()) {
            const saved = await updateStageUpdate(id, {
              status: nextStatus,
              resolvedBy: updatedRecord.resolvedBy,
              resolvedByName: updatedRecord.resolvedByName,
              resolvedAt: updatedRecord.resolvedAt,
              resolutionNote: updatedRecord.resolutionNote,
            });
            updated.push(saved || updatedRecord);
          } else {
            updates[index] = updatedRecord;
            updated.push(updatedRecord);
          }
        }
        if (!isSupabaseConfigured()) {
          await saveUpdates(updates);
        }
        return jsonResponse(200, { ok: true, updates: await enrichUpdatesWithTracking(updated), storage: isSupabaseConfigured() ? 'supabase' : 'json' });
      }
      const id = String(body.id || '').trim();
      if (!id) return jsonResponse(400, { ok: false, error: 'Informe o apontamento para concluir.' });
      const updates = await listUpdates();
      const index = updates.findIndex((item) => String(item.id) === id);
      if (index < 0) return jsonResponse(404, { ok: false, error: 'Apontamento não encontrado.' });
      const currentStatus = String(updates[index]?.status || 'pending').trim().toLowerCase();
      const nextStatus = currentStatus === 'pending_review' ? 'resolved_review' : 'resolved_advance';
      const updatedRecord = {
        ...updates[index],
        status: nextStatus,
        resolvedBy: session.username || '',
        resolvedByName: session.name || session.username || 'Usuário',
        resolvedAt: new Date().toISOString(),
        resolutionNote,
      };
      if (isSupabaseConfigured()) {
        const saved = await updateStageUpdate(id, {
          status: nextStatus,
          resolvedBy: updatedRecord.resolvedBy,
          resolvedByName: updatedRecord.resolvedByName,
          resolvedAt: updatedRecord.resolvedAt,
          resolutionNote: updatedRecord.resolutionNote,
        });
        const enriched = await enrichUpdatesWithTracking([saved || updatedRecord]);
        return jsonResponse(200, { ok: true, update: enriched[0] || saved || updatedRecord, storage: 'supabase' });
      }
      updates[index] = updatedRecord;
      await saveUpdates(updates);
      const enriched = await enrichUpdatesWithTracking([updates[index]]);
      return jsonResponse(200, { ok: true, update: enriched[0] || updates[index], storage: 'json' });
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao processar apontamentos setoriais.' });
  }
};
