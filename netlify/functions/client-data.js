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
    // v37.13: preserva separadores como espaços. Ex.: BW_ENERGY_BR => "bw energy br".
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function stripClientRegionSuffix(value) {
  const normalized = normalizeClientScopeValue(value);
  if (!normalized) return '';
  return normalized
    .replace(/\s+(br|pt|brazil|brasil|portugal)$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function expandClientScopeAliases(rawValues = []) {
  const values = [];
  const add = (value) => {
    const normalized = normalizeClientScopeValue(value);
    if (!normalized) return;
    if (!values.some((item) => normalizeClientScopeValue(item) === normalized)) values.push(value);
  };

  rawValues.forEach(add);
  rawValues.forEach((value) => {
    const stripped = stripClientRegionSuffix(value);
    if (stripped) add(stripped);
  });

  const normalized = values.map((value) => normalizeClientScopeValue(value)).filter(Boolean);
  const compacted = normalized.map((value) => value.replace(/[^a-z0-9]+/g, ''));
  const hasToken = (...tokens) => tokens.some((token) => normalized.includes(token) || compacted.includes(String(token).replace(/[^a-z0-9]+/g, '')));

  // Alias explícito legado Portugal/SBM. Não habilita comparação por primeira palavra.
  if (hasToken('sbm', 'sbm pt', 'sbmpt')) {
    add('STEP PORTUGAL');
    add('STEP PORTUGAL PT');
    add('SBM Offshore');
    add('SBM OFFSHORE');
  }

  return values;
}

function getClientScopeValues(key = {}) {
  const configuredScopes = [key.clientKey, key.clientName, ...(Array.isArray(key.allowedClients) ? key.allowedClients : [])]
    .filter((value) => String(value || '').trim());
  return expandClientScopeAliases(configuredScopes)
    .map(normalizeClientScopeValue)
    .filter(Boolean)
    .filter((value, index, arr) => arr.indexOf(value) === index);
}

function projectBelongsToClientScope(project, key = {}) {
  const scopes = getClientScopeValues(key);
  if (!project || !scopes.length) return false;
  const client = normalizeClientScopeValue(project.client);
  if (!client) return false;

  // v37.13: comparação exata para não misturar empresas do mesmo grupo.
  // BW ENERGY não deve visualizar BW LNG, a menos que BW LNG esteja explicitamente em allowed_clients.
  if (scopes.includes(client)) return true;

  const clientWithoutRegion = stripClientRegionSuffix(client);
  return Boolean(clientWithoutRegion && scopes.includes(clientWithoutRegion));
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
    primary: spool.primary || '',
    projectRef: spool.projectRef || spool.projectDisplay || '',
    client: spool.client || '',
    vessel: spool.vessel || '',
    priority: spool.priority || '',
    pm: spool.pm || '',
    drawing: spool.drawing || '',
    lineNumber: spool.lineNumber || '',
    size: spool.size || '',
    quantitySpools: Number(spool.quantitySpools || 0),
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
    stageValues: spool.stageValues && typeof spool.stageValues === 'object' ? spool.stageValues : {},
  };
}

function sanitizeProject(project = {}, includeSpools = false) {
  const base = {
    rowNumber: project.rowNumber || null,
    primary: project.primary || '',
    projectPrefix: project.projectPrefix || '',
    projectNumber: project.projectNumber || '',
    projectDisplay: project.projectDisplay || '',
    projectType: project.projectType || '',
    priority: project.priority || '',
    client: project.client || '',
    vessel: project.vessel || '',
    pm: project.pm || '',
    lineNumber: project.lineNumber || '',
    size: project.size || '',
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
    replannedFinishDate: project.replannedFinishDate || '',
    replannedFinishSource: project.replannedFinishSource || '',
    clientDisplayCode: project.clientDisplayCode || project.projectDisplay || '',
    customerPo: project.customerPo || '',
    customerPoList: Array.isArray(project.customerPoList) ? project.customerPoList : [],
    customerPoDisplay: project.customerPoDisplay || '',
    clientFocalPoint: project.clientFocalPoint || '',
    clientFocalPointList: Array.isArray(project.clientFocalPointList) ? project.clientFocalPointList : [],
    clientFocalPointDisplay: project.clientFocalPointDisplay || '',
    milestones: sanitizeMilestones(project.milestones),
    stageValues: project.stageValues && typeof project.stageValues === 'object' ? project.stageValues : {},
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
    return jsonResponse(200, { ok: true });
  }
  if (event.httpMethod !== 'GET') return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  if (!isSupabaseConfigured()) return jsonResponse(500, { ok: false, error: 'Supabase não configurado.' });

  const token = extractToken(event);
  if (!token) return jsonResponse(401, { ok: false, error: 'Informe a API key no header Authorization: Bearer <token>.' });

  const key = await findClientApiKeyByHash(sha256(token));
  if (!key || key.active === false) return jsonResponse(401, { ok: false, error: 'API key inválida ou revogada.' });
  if (key.expiresAt && new Date(key.expiresAt).getTime() < Date.now()) return jsonResponse(401, { ok: false, error: 'API key expirada.' });

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
      'cache-control': 'private, max-age=60, stale-while-revalidate=120',
      'access-control-allow-origin': '*',
      'access-control-allow-headers': 'authorization, x-step-api-key, content-type',
    },
  });
};
