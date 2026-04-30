const { jsonResponse, requireSession, normalizeSectorValue } = require('./_auth');
const { readJson, writeJson } = require('./_githubStore');
const { isSupabaseConfigured, listStageUpdates, createStageUpdate, updateStageUpdate, deleteStageUpdates } = require('./_supabase');
const { findProjectAndSpool, loadProjectPayload } = require('./_projectLookup');
const { applyStageUpdatesToTracking, listHistoryDatePendencies } = require('./_smartsheetTracking');

const DATA_PATH = 'data/stage-updates.json';
const SUPPORTED_SECTORS = ['pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const PROGRESS_OPTIONS = [25, 50, 75, 100];
const PENDING_STATUSES = ['pending', 'pending_advance', 'pending_review'];
const RESOLVED_STATUSES = ['resolved', 'resolved_advance', 'resolved_review'];

const TRACKING_FIELDS_BY_SECTOR = {
  pintura: ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)'],
  inspecao: ['Final Inspection', 'Hydro Test Pressure (QC)', 'Non Destructive Examination (QC)', 'Final Dimensional Inpection/3D (QC)', 'Initial Dimensional Inspection/3D'],
  pendente_envio: ['Package and Delivered', 'Final Inspection'],
  producao: ['Spool Assemble and tack weld', 'Welding Preparation'],
  calderaria: ['Spool Assemble and tack weld', 'Welding Preparation', 'Material Separation', 'Material Release to Fabrication'],
  solda: ['Full welding execution'],
};

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

function isPendingStatus(status) {
  return PENDING_STATUSES.includes(String(status || 'pending').trim().toLowerCase());
}

function isResolvedStatus(status) {
  return RESOLVED_STATUSES.includes(String(status || '').trim().toLowerCase());
}

function isReviewStatus(status) {
  return String(status || '').trim().toLowerCase().includes('review');
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
    trackingStatus: trackingProgress == null ? 'not_found' : (trackingMatched ? 'matched' : 'waiting'),
  };
}

async function enrichUpdatesWithTracking(updates) {
  const list = Array.isArray(updates) ? updates : [];
  if (!list.length) return [];
  let payload = null;
  try {
    payload = await loadProjectPayload({ allowFallback: false });
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
  if (isSupabaseConfigured()) return listStageUpdates();
  const rows = await readJson(DATA_PATH, []);
  return Array.isArray(rows) ? rows : [];
}

async function saveUpdates(rows) {
  return writeJson(DATA_PATH, rows, 'chore: atualiza apontamentos setoriais');
}

function normalizeUpdateForJson(record) {
  return { ...record, updatedAt: new Date().toISOString() };
}

async function resolveStageUpdateRecord(id, updates, session, resolutionNote = '') {
  const index = updates.findIndex((item) => String(item.id) === String(id));
  if (index < 0) return null;
  const current = updates[index];
  const currentStatus = String(current?.status || 'pending').trim().toLowerCase();
  const nextStatus = currentStatus === 'pending_review' ? 'resolved_review' : 'resolved_advance';
  const resolvedAt = new Date().toISOString();
  const updatedRecord = normalizeUpdateForJson({
    ...current,
    status: nextStatus,
    resolvedBy: session.username || '',
    resolvedByName: session.name || session.username || 'Usuário',
    resolvedAt,
    resolutionNote,
  });
  if (isSupabaseConfigured()) {
    const saved = await updateStageUpdate(id, {
      status: nextStatus,
      resolvedBy: updatedRecord.resolvedBy,
      resolvedByName: updatedRecord.resolvedByName,
      resolvedAt: updatedRecord.resolvedAt,
      resolutionNote: updatedRecord.resolutionNote,
    });
    return saved || updatedRecord;
  }
  updates[index] = updatedRecord;
  return updatedRecord;
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
  const { project, spool } = await findProjectAndSpool(projectRowId, spoolIso, { allowFallback: false });
  if (!project || !spool) {
    const err = new Error('BSP ou spool não localizado para este apontamento.');
    err.statusCode = 404;
    throw err;
  }
  const trackingProgress = getTrackingProgressForSector(spool, sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;
  const updates = existingUpdates || await listUpdates();
  const pendingExists = updates.find((item) =>
    isPendingStatus(item.status)
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
    id: `stg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    projectRowId,
    projectNumber: project.projectNumber || project.projectDisplay || `Projeto ${projectRowId}`,
    projectDisplay: project.projectDisplay || project.projectNumber || `Projeto ${projectRowId}`,
    client: project.client || '',
    spoolIso,
    spoolDescription: spool.description || spool.drawing || '',
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

function getUpdatesByIds(updates, ids) {
  const cleanIds = new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean));
  return updates.filter((item) => cleanIds.has(String(item.id || '')));
}

async function updateTrackingAndResolve(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode atualizar o Tracking.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe os apontamentos para atualizar.' });

  const forceRewrite = Boolean(body.forceRewrite || body.rewrite);
  const dateOnly = Boolean(body.dateOnly);
  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => {
    if (dateOnly) return isResolvedStatus(item.status) && Number(item.progress || 0) === 100 && !isReviewStatus(item.status);
    return isPendingStatus(item.status) && !isReviewStatus(item.status);
  });
  if (!selected.length) {
    return jsonResponse(404, { ok: false, error: 'Nenhum apontamento elegível encontrado para atualizar o Tracking.' });
  }

  const trackingResult = await applyStageUpdatesToTracking(selected, { forceRewrite, dateOnly });
  const successResultIds = new Set((trackingResult.results || []).filter((item) => item.success).map((item) => String(item.id || '')));
  const resolutionNoteBase = dateOnly
    ? 'Pendência de data do histórico corrigida no Smartsheet/Tracking.'
    : (forceRewrite ? 'Tracking regravado e apontamento validado automaticamente.' : 'Tracking atualizado e apontamento validado automaticamente.');

  const resolved = [];
  if (!dateOnly) {
    for (const item of selected) {
      if (!successResultIds.has(String(item.id))) continue;
      const saved = await resolveStageUpdateRecord(item.id, updates, session, resolutionNoteBase);
      if (saved) resolved.push(saved);
    }
    if (!isSupabaseConfigured() && resolved.length) await saveUpdates(updates);
  }

  const hasSuccess = (trackingResult.results || []).some((item) => item.success);
  const hasErrors = Array.isArray(trackingResult.errors) && trackingResult.errors.length > 0;
  return jsonResponse(hasSuccess ? 200 : 400, {
    ok: hasSuccess,
    partial: hasSuccess && hasErrors,
    tracking: trackingResult,
    updates: resolved,
    errors: trackingResult.errors || [],
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}


async function deleteStageUpdateRecords(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode remover apontamentos.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe os apontamentos para remover.' });

  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => isPendingStatus(item.status));
  if (!selected.length) return jsonResponse(404, { ok: false, error: 'Nenhum apontamento pendente encontrado para remover.' });
  const selectedIds = selected.map((item) => String(item.id));

  if (isSupabaseConfigured()) {
    await deleteStageUpdates(selectedIds);
  } else {
    const selectedSet = new Set(selectedIds);
    const remaining = updates.filter((item) => !selectedSet.has(String(item.id || '')));
    await saveUpdates(remaining);
  }

  return jsonResponse(200, {
    ok: true,
    removed: selected,
    removedCount: selected.length,
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}

async function concludeTrackingOkOnly(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode concluir apontamentos.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  const resolutionNote = String(body.resolutionNote || '').trim();
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe o apontamento para concluir.' });

  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => isPendingStatus(item.status));
  if (!selected.length) return jsonResponse(404, { ok: false, error: 'Apontamento não encontrado.' });

  const reviewItems = selected.filter((item) => isReviewStatus(item.status));
  const advanceItems = selected.filter((item) => !isReviewStatus(item.status));
  let eligibleAdvance = [];
  let blockedErrors = [];

  if (advanceItems.length) {
    const dryRun = await applyStageUpdatesToTracking(advanceItems, { dryRun: true });
    eligibleAdvance = advanceItems.filter((item) => {
      const result = (dryRun.results || []).find((entry) => String(entry.id) === String(item.id));
      return result?.trackingOk === true;
    });
    blockedErrors = (dryRun.results || [])
      .filter((entry) => entry?.trackingOk !== true)
      .map((entry) => ({ id: entry.id, error: entry.message || 'Tracking ainda pendente de atualização.' }));
  }

  const toResolve = [...reviewItems, ...eligibleAdvance];
  if (!toResolve.length) {
    return jsonResponse(409, {
      ok: false,
      error: 'Nenhum item está com Tracking OK para concluir. Atualize ou regrave o Tracking primeiro.',
      errors: blockedErrors,
    });
  }

  const resolved = [];
  const note = resolutionNote || 'Concluído pelo PCP após conferência de Tracking OK.';
  for (const item of toResolve) {
    const saved = await resolveStageUpdateRecord(item.id, updates, session, note);
    if (saved) resolved.push(saved);
  }
  if (!isSupabaseConfigured() && resolved.length) await saveUpdates(updates);

  return jsonResponse(200, {
    ok: blockedErrors.length === 0,
    partial: blockedErrors.length > 0,
    updates: resolved,
    errors: blockedErrors,
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return jsonResponse(200, { ok: true });
  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  const session = auth.session;

  try {
    if (event.httpMethod === 'GET') {
      const mode = String(event.queryStringParameters?.mode || '').trim();
      const updates = await listUpdates();
      if (mode === 'history-date-pending') {
        if (!canValidate(session)) return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode consultar pendências.' });
        const pendencies = await listHistoryDatePendencies(updates);
        return jsonResponse(200, { ok: true, pendencies });
      }
      const enriched = await enrichUpdatesWithTracking(updates);
      return jsonResponse(200, {
        ok: true,
        updates: enriched,
        autoResolvedCount: 0,
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
        if (!isSupabaseConfigured() && created.length) await saveUpdates(baseUpdates);
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
        if (!isSupabaseConfigured()) await saveUpdates(updates);
        return jsonResponse(200, { ok: true, update: saved, storage: isSupabaseConfigured() ? 'supabase' : 'json' });
      } catch (error) {
        return jsonResponse(error.statusCode || 400, { ok: false, error: error.message || 'Falha ao enviar apontamento.' });
      }
    }

    if (event.httpMethod === 'PUT') {
      const body = JSON.parse(event.body || '{}');
      const action = String(body.action || '').trim().toLowerCase();
      if (action === 'update-tracking') return updateTrackingAndResolve(body, session);
      if (action === 'fix-history-dates') return updateTrackingAndResolve({ ...body, dateOnly: true, forceRewrite: true }, session);
      return jsonResponse(400, { ok: false, error: 'Ação de atualização não reconhecida.' });
    }

    if (event.httpMethod === 'DELETE') {
      const body = JSON.parse(event.body || '{}');
      return deleteStageUpdateRecords(body, session);
    }

    if (event.httpMethod === 'PATCH') {
      const body = JSON.parse(event.body || '{}');
      return concludeTrackingOkOnly(body, session);
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao processar apontamentos setoriais.' });
  }
};
