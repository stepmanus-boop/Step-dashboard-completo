
const { jsonResponse, requireSession, normalizeSectorValue } = require('./_auth');
const { readJson, writeJson } = require('./_githubStore');
const { isSupabaseConfigured, listStageUpdates, createStageUpdate, updateStageUpdate } = require('./_supabase');
const { findProjectAndSpool } = require('./_projectLookup');

const DATA_PATH = 'data/stage-updates.json';
const SUPPORTED_SECTORS = ['pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const PROGRESS_OPTIONS = [25, 50, 75, 100];
const PENDING_STATUSES = ['pending', 'pending_advance', 'pending_review'];

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
      const updates = await listUpdates();
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
        return jsonResponse(200, { ok: true, updates: updated, storage: isSupabaseConfigured() ? 'supabase' : 'json' });
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
        return jsonResponse(200, { ok: true, update: saved || updatedRecord, storage: 'supabase' });
      }
      updates[index] = updatedRecord;
      await saveUpdates(updates);
      return jsonResponse(200, { ok: true, update: updates[index], storage: 'json' });
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao processar apontamentos setoriais.' });
  }
};
