
const { jsonResponse, requireSession, normalizeSectorValue } = require('./_auth');
const { readJson, writeJson } = require('./_githubStore');
const { isSupabaseConfigured, listStageUpdates, createStageUpdate, updateStageUpdate } = require('./_supabase');
const { findProjectAndSpool } = require('./_projectLookup');

const DATA_PATH = 'data/stage-updates.json';
const SUPPORTED_SECTORS = ['pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const PROGRESS_OPTIONS = [25, 50, 75, 100];

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
      const projectRowId = Number(body.projectRowId || 0);
      const spoolIso = String(body.spoolIso || '').trim();
      const progress = Number(body.progress || 0);
      const completionDate = String(body.completionDate || '').trim();
      const note = String(body.note || '').trim();
      const sector = session.role === 'admin'
        ? normalizeSectorValue(body.sector || session.sector)
        : getActorSector(session);

      if (!projectRowId || !spoolIso || !SUPPORTED_SECTORS.includes(sector)) {
        return jsonResponse(400, { ok: false, error: 'Informe BSP, spool e uma etapa válida.' });
      }
      if (!PROGRESS_OPTIONS.includes(progress)) {
        return jsonResponse(400, { ok: false, error: 'Selecione um avanço válido: 25%, 50%, 75% ou 100%.' });
      }
      const { project, spool } = await findProjectAndSpool(projectRowId, spoolIso);
      if (!project || !spool) {
        return jsonResponse(404, { ok: false, error: 'BSP ou spool não localizado para este apontamento.' });
      }
      const updates = await listUpdates();
      const pendingExists = updates.find((item) =>
        String(item.status || 'pending') === 'pending'
        && Number(item.projectRowId || 0) === projectRowId
        && String(item.spoolIso || '').trim().toLowerCase() === spoolIso.toLowerCase()
        && normalizeSectorValue(item.sector) === sector
      );
      if (pendingExists) {
        return jsonResponse(409, { ok: false, error: 'Já existe um apontamento pendente desta etapa para este spool.' });
      }
      const now = new Date().toISOString();
      const record = {
        id: `stg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
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
        status: 'pending',
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
        return jsonResponse(200, { ok: true, update: saved || record, storage: 'supabase' });
      }
      updates.unshift(record);
      await saveUpdates(updates);
      return jsonResponse(200, { ok: true, update: record, storage: 'json' });
    }

    if (event.httpMethod === 'PATCH') {
      if (!canValidate(session)) {
        return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode concluir apontamentos.' });
      }
      const body = JSON.parse(event.body || '{}');
      const id = String(body.id || '').trim();
      const resolutionNote = String(body.resolutionNote || '').trim();
      if (!id) return jsonResponse(400, { ok: false, error: 'Informe o apontamento para concluir.' });
      const updates = await listUpdates();
      const index = updates.findIndex((item) => String(item.id) === id);
      if (index < 0) return jsonResponse(404, { ok: false, error: 'Apontamento não encontrado.' });
      const updatedRecord = {
        ...updates[index],
        status: 'resolved',
        resolvedBy: session.username || '',
        resolvedByName: session.name || session.username || 'Usuário',
        resolvedAt: new Date().toISOString(),
        resolutionNote,
      };
      if (isSupabaseConfigured()) {
        const saved = await updateStageUpdate(id, {
          status: 'resolved',
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
