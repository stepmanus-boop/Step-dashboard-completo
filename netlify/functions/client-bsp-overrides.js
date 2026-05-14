const { jsonResponse, requireSession, normalizeText, normalizeSectorValue } = require('./_auth');
const {
  isSupabaseConfigured,
  getUserById,
  listClientBspOverrides,
  upsertClientBspOverride,
  deleteClientBspOverride,
} = require('./_supabase');
const { loadProjectPayload } = require('./_projectLookup');

async function hydrateSessionUser(session = {}) {
  if (!session?.sub) return session;
  if (!isSupabaseConfigured()) return session;
  try {
    const user = await getUserById(session.sub);
    return user ? { ...session, ...user, sub: session.sub } : session;
  } catch (_) {
    return session;
  }
}

function userHasProjectsScope(user = {}) {
  if (!user || user.role === 'admin') return false;
  const sectors = Array.isArray(user.alertSectors) ? user.alertSectors : [];
  return normalizeSectorValue(user.sector) === 'projetos' || sectors.map(normalizeSectorValue).includes('projetos');
}

function tokenizeNormalizedNames(values = []) {
  const set = new Set();
  const source = Array.isArray(values) ? values : [values];
  for (const value of source) {
    const normalized = normalizeText(value).trim();
    if (!normalized) continue;
    set.add(normalized);
    for (const part of normalized.split(/[^a-z0-9]+/)) {
      if (part) set.add(part);
    }
  }
  return set;
}

function projectBelongsToUser(project = {}, user = {}) {
  if (!project || !userHasProjectsScope(user)) return false;
  const pmValue = String(project.pm || '').trim();
  if (!pmValue) return false;
  const candidates = tokenizeNormalizedNames([
    user.name,
    user.username,
    String(user.username || '').split('@')[0],
    ...(Array.isArray(user.projectPmAliases) ? user.projectPmAliases : []),
  ]);
  if (!candidates.size) return false;
  const normalizedPm = normalizeText(pmValue).trim();
  const pmTokens = tokenizeNormalizedNames(pmValue.split(/[;,|/]+/));
  for (const candidate of candidates) {
    if (normalizedPm === candidate || normalizedPm.includes(candidate)) return true;
    if (pmTokens.has(candidate)) return true;
  }
  return false;
}

function canEditProject(project, user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return projectBelongsToUser(project, user);
}

function normalizeDateInput(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  return text.slice(0, 10);
}

function normalizeOverrideInput(body = {}, project = {}, user = {}) {
  const customFields = body.customFields && typeof body.customFields === 'object' && !Array.isArray(body.customFields)
    ? body.customFields
    : {};
  return {
    projectRowId: String(body.projectRowId || project.rowId || '').trim(),
    projectNumber: String(body.projectNumber || project.projectNumber || '').trim(),
    projectDisplay: String(body.projectDisplay || project.projectDisplay || project.projectNumber || '').trim(),
    clientKey: String(body.clientKey || '').trim(),
    clientName: String(body.clientName || project.client || '').trim(),
    vessel: String(body.vessel || project.vessel || project.unit || '').trim(),
    pm: String(body.pm || project.pm || '').trim(),
    fabricationStartOverride: normalizeDateInput(body.fabricationStartOverride),
    boilermakerFinishOverride: normalizeDateInput(body.boilermakerFinishOverride),
    weldingFinishOverride: normalizeDateInput(body.weldingFinishOverride),
    inspectionFinishOverride: normalizeDateInput(body.inspectionFinishOverride),
    thFinishOverride: normalizeDateInput(body.thFinishOverride),
    coatingFinishOverride: normalizeDateInput(body.coatingFinishOverride),
    projectFinishOverride: normalizeDateInput(body.projectFinishOverride),
    executiveStatus: String(body.executiveStatus || '').trim().slice(0, 160),
    executiveNote: String(body.executiveNote || '').trim().slice(0, 1000),
    delayReason: String(body.delayReason || '').trim().slice(0, 220),
    customFields,
    visibleToClient: body.visibleToClient !== false,
    createdBy: user.username || '',
    createdByName: user.name || user.username || '',
    updatedBy: user.username || '',
    updatedByName: user.name || user.username || '',
  };
}

async function findProject(projectRowId) {
  const payload = await loadProjectPayload({ allowFallback: true });
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  const target = String(projectRowId || '').trim();
  return projects.find((project) => String(project.rowId || '') === target || String(project.rowNumber || '') === target) || null;
}

function filterOverridesForUser(overrides = [], user = {}) {
  if (user.role === 'admin' || userHasProjectsScope(user)) return overrides;
  if (user.role === 'client') {
    const clientKey = normalizeText(user.clientKey || '');
    const clientName = normalizeText(user.clientName || user.clientKey || '');
    const allowed = new Set((Array.isArray(user.allowedClients) ? user.allowedClients : []).map((item) => normalizeText(item)).filter(Boolean));
    return overrides.filter((item) => {
      const itemClientKey = normalizeText(item.clientKey || '');
      const itemClientName = normalizeText(item.clientName || '');
      if (clientKey && itemClientKey && clientKey === itemClientKey) return true;
      if (clientName && itemClientName && clientName === itemClientName) return true;
      return allowed.has(itemClientKey) || allowed.has(itemClientName);
    });
  }
  return [];
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return jsonResponse(200, { ok: true });

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado. Configure SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY.' });
  }

  const user = await hydrateSessionUser(auth.session);

  try {
    if (event.httpMethod === 'GET') {
      const overrides = await listClientBspOverrides();
      return jsonResponse(200, { ok: true, overrides: filterOverridesForUser(overrides, user) });
    }

    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      const projectRowId = String(body.projectRowId || '').trim();
      if (!projectRowId) return jsonResponse(400, { ok: false, error: 'Informe a BSP que será ajustada.' });
      const project = await findProject(projectRowId);
      if (!project) return jsonResponse(404, { ok: false, error: 'BSP não encontrada na base operacional.' });
      if (!canEditProject(project, user)) return jsonResponse(403, { ok: false, error: 'Apenas PM vinculado à BSP ou administrador pode editar esta informação.' });
      const override = await upsertClientBspOverride(normalizeOverrideInput(body, project, user));
      return jsonResponse(200, { ok: true, override });
    }

    if (event.httpMethod === 'DELETE') {
      const id = event.queryStringParameters?.id || '';
      const projectRowId = event.queryStringParameters?.projectRowId || '';
      if (!id) return jsonResponse(400, { ok: false, error: 'Informe o ajuste que será removido.' });
      const project = projectRowId ? await findProject(projectRowId) : null;
      if (project && !canEditProject(project, user)) return jsonResponse(403, { ok: false, error: 'Apenas PM vinculado à BSP ou administrador pode remover este ajuste.' });
      if (!project && user.role !== 'admin') return jsonResponse(403, { ok: false, error: 'Informe a BSP para validar a permissão.' });
      await deleteClientBspOverride(id);
      return jsonResponse(200, { ok: true });
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    const message = String(error?.message || 'Falha ao processar ajustes executivos.');
    const missingTable = message.includes('client_bsp_overrides');
    return jsonResponse(missingTable ? 500 : 400, {
      ok: false,
      error: missingTable
        ? 'Tabela client_bsp_overrides não encontrada. Execute o SQL de criação no Supabase.'
        : message,
    });
  }
};
