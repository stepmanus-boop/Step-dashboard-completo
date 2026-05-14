const crypto = require('crypto');
const { jsonResponse } = require('./_auth');
const { buildPayload } = require('./projects');
const {
  isSupabaseConfigured,
  findClientApiKeyByHash,
  markClientApiKeyUsed,
} = require('./_supabase');

function sha256(value) {
  return crypto.createHash('sha256').update(String(value || '')).digest('hex');
}

const READ_ONLY_SCOPE = 'read:projects';
const READ_ONLY_ALLOWED_METHODS = 'GET, OPTIONS';

function readOnlyHeaders(extra = {}) {
  return {
    'access-control-allow-origin': '*',
    'access-control-allow-methods': READ_ONLY_ALLOWED_METHODS,
    'access-control-allow-headers': 'authorization, x-step-api-key, content-type',
    'allow': READ_ONLY_ALLOWED_METHODS,
    ...extra,
  };
}

function hasReadOnlyScope(key = {}) {
  const scopes = Array.isArray(key.scopes) ? key.scopes.map((scope) => String(scope || '').trim()) : [];
  return scopes.includes(READ_ONLY_SCOPE) && !scopes.some((scope) => /^write:|^edit:|^admin:|^delete:/i.test(scope));
}

function extractToken(event) {
  const headers = event.headers || {};
  const auth = headers.authorization || headers.Authorization || '';
  if (/^Bearer\s+/i.test(auth)) return auth.replace(/^Bearer\s+/i, '').trim();
  return String(headers['x-step-api-key'] || headers['X-Step-Api-Key'] || event.queryStringParameters?.api_key || '').trim();
}

function normalizeClientScopeValue(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

const CLIENT_SCOPE_GENERIC_WORDS = new Set([
  'de', 'da', 'do', 'das', 'dos', 'e', 'em', 'the', 'and', 'of', 'a', 'an',
  'sa', 's', 'ltda', 'ltd', 'llc', 'inc', 'corp', 'company', 'companhia',
  'brasil', 'brazil', 'global', 'international', 'internacional', 'energy',
  'energia', 'offshore', 'oil', 'gas', 'petroleo', 'petroleum', 'services',
  'service', 'servicos', 'solucoes', 'solutions', 'industrial', 'industria'
]);

function getClientPrimaryToken(value) {
  const normalized = normalizeClientScopeValue(value);
  if (!normalized) return '';
  const words = normalized.split(/\s+/).filter(Boolean);
  if (!words.length) return '';
  return words.find((word) => word.length >= 2 && !CLIENT_SCOPE_GENERIC_WORDS.has(word)) || words[0];
}

function projectBelongsToClientScope(project, key = {}) {
  const scopes = [key.clientKey, key.clientName, ...(Array.isArray(key.allowedClients) ? key.allowedClients : [])]
    .map(normalizeClientScopeValue)
    .filter(Boolean);
  if (!project || !scopes.length) return false;
  const client = normalizeClientScopeValue(project.client);
  if (!client) return false;
  const clientPrimary = getClientPrimaryToken(client);
  return scopes.some((scope) => {
    const scopePrimary = getClientPrimaryToken(scope);
    return client === scope || (clientPrimary && scopePrimary && clientPrimary === scopePrimary);
  });
}

function asNumber(value) {
  const num = Number(value || 0);
  return Number.isFinite(num) ? num : 0;
}

function getProjectItemCount(project) {
  const fromStats = Number(project?.spoolStats?.total || 0);
  if (fromStats) return fromStats;
  const fromQuantity = Number(project?.quantitySpools || 0);
  if (fromQuantity) return fromQuantity;
  return Array.isArray(project?.spools) ? project.spools.length : 0;
}

function buildSummary(projects = []) {
  const summary = projects.reduce((acc, project) => {
    acc.bsps += 1;
    acc.tags += getProjectItemCount(project);
    acc.weightKg += asNumber(project.kilos);
    acc.weldedKg += asNumber(project.weldedWeightKg);
    acc.m2 += asNumber(project.m2Painting);
    acc.progressTotal += asNumber(project.overallProgress);
    return acc;
  }, { bsps: 0, tags: 0, weightKg: 0, weldedKg: 0, m2: 0, progressTotal: 0 });
  return {
    bsps: summary.bsps,
    tags: summary.tags,
    weightKg: Number(summary.weightKg.toFixed(3)),
    weldedKg: Number(summary.weldedKg.toFixed(3)),
    m2: Number(summary.m2.toFixed(3)),
    averageProgress: summary.bsps ? Number((summary.progressTotal / summary.bsps).toFixed(2)) : 0,
  };
}

function sanitizeMilestones(milestones = []) {
  return (Array.isArray(milestones) ? milestones : []).map((item) => ({
    key: item.key || '',
    label: item.label || item.key || '',
    value: item.value || '',
    type: item.type || '',
  }));
}

function sanitizeSpool(spool = {}) {
  return {
    rowNumber: spool.rowNumber || null,
    iso: spool.iso || '',
    description: spool.description || '',
    kilos: asNumber(spool.kilos),
    m2Painting: asNumber(spool.m2Painting),
    stage: spool.stage || '',
    stagePercent: asNumber(spool.stagePercent),
    stageStatus: spool.stageStatus || '',
    individualProgress: asNumber(spool.individualProgress),
    overallProgress: asNumber(spool.overallProgress),
    milestones: sanitizeMilestones(spool.milestones),
  };
}

function sanitizeProject(project = {}, includeSpools = false) {
  const base = {
    rowNumber: project.rowNumber || null,
    projectPrefix: project.projectPrefix || '',
    projectNumber: project.projectNumber || '',
    projectDisplay: project.projectDisplay || '',
    projectType: project.projectType || '',
    client: project.client || '',
    vessel: project.vessel || '',
    className: project.className || '',
    quantitySpools: Number(project.quantitySpools || 0),
    itemCount: getProjectItemCount(project),
    kilos: asNumber(project.kilos),
    weldedWeightKg: asNumber(project.weldedWeightKg),
    m2Painting: asNumber(project.m2Painting),
    currentStage: project.currentStage || '',
    currentStagePercent: asNumber(project.currentStagePercent),
    currentStageStatus: project.currentStageStatus || '',
    individualProgress: asNumber(project.individualProgress),
    overallProgress: asNumber(project.overallProgress),
    projectStatus: project.projectStatus || '',
    jobProcessStatus: project.jobProcessStatus || '',
    summaryDrawing: project.summaryDrawing || '',
    finished: Boolean(project.finished),
    uiState: project.uiState || '',
    plannedStartDate: project.plannedStartDate || '',
    plannedFinishDate: project.plannedFinishDate || '',
    clientDisplayCode: project.clientDisplayCode || project.projectDisplay || '',
    milestones: sanitizeMilestones(project.milestones),
    spoolStats: project.spoolStats || {},
  };
  if (includeSpools) base.spools = (Array.isArray(project.spools) ? project.spools : []).map(sanitizeSpool);
  return base;
}

function buildUnits(projects = []) {
  const map = new Map();
  for (const project of projects) {
    const vessel = project.vessel || 'Unidade não informada';
    if (!map.has(vessel)) map.set(vessel, { vessel, bsps: 0, tags: 0, weightKg: 0, weldedKg: 0, m2: 0, progressTotal: 0 });
    const item = map.get(vessel);
    item.bsps += 1;
    item.tags += getProjectItemCount(project);
    item.weightKg += asNumber(project.kilos);
    item.weldedKg += asNumber(project.weldedWeightKg);
    item.m2 += asNumber(project.m2Painting);
    item.progressTotal += asNumber(project.overallProgress);
  }
  return Array.from(map.values()).map((item) => ({
    vessel: item.vessel,
    bsps: item.bsps,
    tags: item.tags,
    weightKg: Number(item.weightKg.toFixed(3)),
    weldedKg: Number(item.weldedKg.toFixed(3)),
    m2: Number(item.m2.toFixed(3)),
    averageProgress: item.bsps ? Number((item.progressTotal / item.bsps).toFixed(2)) : 0,
  })).sort((a, b) => b.bsps - a.bsps || a.vessel.localeCompare(b.vessel, 'pt-BR'));
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return jsonResponse(200, { ok: true, readOnly: true }, { headers: readOnlyHeaders() });
  }
  if (event.httpMethod !== 'GET') {
    return jsonResponse(405, {
      ok: false,
      readOnly: true,
      error: 'API somente leitura. Use apenas GET para consultar os dados; POST, PUT, PATCH e DELETE não são permitidos neste endpoint.',
    }, { headers: readOnlyHeaders() });
  }
  if (!isSupabaseConfigured()) return jsonResponse(500, { ok: false, error: 'Supabase não configurado.' });

  const token = extractToken(event);
  if (!token) return jsonResponse(401, { ok: false, error: 'Informe a API key no header Authorization: Bearer <token>.' });

  const key = await findClientApiKeyByHash(sha256(token));
  if (!key || key.active === false) return jsonResponse(401, { ok: false, error: 'API key inválida ou revogada.' }, { headers: readOnlyHeaders() });
  if (key.expiresAt && new Date(key.expiresAt).getTime() < Date.now()) return jsonResponse(401, { ok: false, error: 'API key expirada.' }, { headers: readOnlyHeaders() });
  if (!hasReadOnlyScope(key)) {
    return jsonResponse(403, {
      ok: false,
      readOnly: true,
      error: 'API key sem permissão de leitura válida. Esta API aceita somente o escopo read:projects.',
    }, { headers: readOnlyHeaders() });
  }

  const params = event.queryStringParameters || {};
  const includeSpools = String(params.includeSpools || params.include_spools || '') === '1';
  const format = String(params.format || 'full').toLowerCase();
  const bspFilter = normalizeClientScopeValue(params.bsp || params.project || '');
  const unitFilter = normalizeClientScopeValue(params.unit || params.vessel || '');

  const payload = await buildPayload({ force: false, preferCache: true });
  let projects = (Array.isArray(payload.projects) ? payload.projects : []).filter((project) => projectBelongsToClientScope(project, key));
  if (bspFilter) {
    projects = projects.filter((project) => normalizeClientScopeValue(`${project.projectDisplay || ''} ${project.projectNumber || ''} ${project.clientDisplayCode || ''}`).includes(bspFilter));
  }
  if (unitFilter) {
    projects = projects.filter((project) => normalizeClientScopeValue(project.vessel || '').includes(unitFilter));
  }

  markClientApiKeyUsed(key.id).catch(() => {});

  const summary = buildSummary(projects);
  const response = {
    ok: true,
    client: {
      key: key.clientKey || '',
      name: key.clientName || key.clientKey || '',
    },
    meta: {
      source: 'STEP Portal do Cliente API',
      accessMode: 'read_only',
      allowedMethods: ['GET'],
      permissions: [READ_ONLY_SCOPE],
      generatedAt: new Date().toISOString(),
      lastSync: payload.meta?.lastSync || null,
      version: payload.meta?.version || null,
      format,
      includeSpools,
    },
    summary,
    units: buildUnits(projects),
  };

  if (format !== 'summary') {
    response.projects = projects.map((project) => sanitizeProject(project, includeSpools));
  }

  return jsonResponse(200, response, {
    headers: {
      ...readOnlyHeaders({ 'cache-control': 'private, max-age=60, stale-while-revalidate=120' }),
    },
  });
};
