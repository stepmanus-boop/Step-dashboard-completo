// Netlify function: Yinson - Projetos em Desenvolvimento
// Busca a sheet correta pelo Smartsheet, identifica colunas de forma flexível
// e retorna a tabela: VESSEL + PRIORITY LEVEL + WORK ORDER + ID REPAIR + TAG NUMBER + SOB + DESCRIPTION + STATUS + IMAGEM.

const API_BASE = process.env.SMARTSHEET_API_BASE || 'https://api.smartsheet.com/2.0';
const REQUEST_TIMEOUT_MS = Math.max(2500, Number(process.env.YINSON_UNDER_DEV_REQUEST_TIMEOUT_MS || process.env.SMARTSHEET_REQUEST_TIMEOUT_MS || process.env.SMARTSHEET_FETCH_TIMEOUT_MS || 6500));
const UNDER_DEV_PAGE_SIZE = Math.max(50, Math.min(2500, Number(process.env.YINSON_UNDER_DEV_PAGE_SIZE || 2500)));
const YINSON_UNDER_DEV_LIST_MAX_PAGES = Math.max(1, Math.min(10, Number(process.env.YINSON_UNDER_DEV_LIST_MAX_PAGES || 6)));
const YINSON_UNDER_DEV_DISABLE_HEAVY_FALLBACKS = String(process.env.YINSON_UNDER_DEV_DISABLE_HEAVY_FALLBACKS || '1') !== '0';
const UNDER_DEV_MAX_PAGES = Math.max(1, Math.min(20, Number(process.env.YINSON_UNDER_DEV_MAX_PAGES || 4)));
const UNDER_DEV_CACHE_TTL_MS = Number(process.env.YINSON_UNDER_DEV_CACHE_TTL_MS || 60000);
const ROW_ATTACHMENT_FALLBACK_LIMIT = Number(process.env.YINSON_UNDER_DEV_ROW_ATTACHMENT_FALLBACK_LIMIT || 300);
let cachedPayload = null;
let cachedPayloadAt = 0;
let cachedPayloadKey = '';

const TOKEN = process.env.SMARTSHEET_API_KEY
  || process.env.SMARTSHEET_TOKEN
  || process.env.SMARTSHEET_ACCESS_TOKEN
  || process.env.SMARTSHEET_API_TOKEN
  || process.env.SMARTSHEET_BEARER_TOKEN
  || process.env.SMARTSHEET_PAT
  || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN
  || '';


const {
  readTrackingCache,
  writeTrackingCache,
  getTrackingCacheMeta,
  acquireTrackingRefreshLock,
  releaseTrackingRefreshLock,
  clearTrackingRefreshLock,
  touchTrackingCache,
  TRACKING_CACHE_REFRESH_LOCK_TTL_MS,
} = require('./_trackingCache');

const YINSON_UNDER_DEV_CACHE_KEY = 'yinson:under-dev:current';
const YINSON_UNDER_DEV_SCHEMA_VERSION = '37.63-priority-level';
const YINSON_UNDER_DEV_CACHE_MAX_AGE_MS = Math.max(
  15 * 60 * 1000,
  Number(process.env.YINSON_UNDER_DEV_CACHE_MAX_AGE_MS || 30 * 24 * 60 * 60 * 1000)
);
const YINSON_UNDER_DEV_REFRESH_LOCK_TTL_MS = Math.max(
  60 * 1000,
  Number(process.env.YINSON_UNDER_DEV_REFRESH_LOCK_TTL_MS || TRACKING_CACHE_REFRESH_LOCK_TTL_MS || 5 * 60 * 1000)
);
const YINSON_UNDER_DEV_CACHE_TIMEOUT_MS = Math.max(
  1500,
  Number(process.env.YINSON_UNDER_DEV_CACHE_TIMEOUT_MS || 4500)
);
const YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE = String(process.env.YINSON_UNDER_DEV_CACHE_ENABLED || '1') !== '0';

const YINSON_UNDER_DEV_PRIMARY_SHEET_ID = '7523600180400004';
const YINSON_UNDER_DEV_PRIMARY_SHEET_LINK = 'https://app.smartsheet.com/sheets/Vc4FwJfQ7XMpGfWv9HvWhJC6xf743r3FM7cMjQ41';
const YINSON_UNDER_DEV_PRIMARY_SHEET_HASH = 'Vc4FwJfQ7XMpGfWv9HvWhJC6xf743r3FM7cMjQ41';

const DEFAULT_CANDIDATE_SHEET_IDS = [
  // Sheet correta informada para Projetos em Desenvolvimento / Yinson.
  YINSON_UNDER_DEV_PRIMARY_SHEET_ID,
  // Fallback histórico usado em versões anteriores.
  '3612820139992964',
];

const FIELD_DEFINITIONS = {
  vessel: {
    label: 'VESSEL',
    aliases: ['vessel', 'vessels', 'platform', 'plataforma', 'unit', 'unidade', 'embarcacao', 'embarcação', 'navio'],
  },
  priorityLevel: {
    label: 'PRIORITY LEVEL',
    aliases: ['priority level', 'priority', 'priority lvl', 'prioritylevel', 'prioridade', 'nivel prioridade', 'nível prioridade', 'nivel de prioridade', 'nível de prioridade'],
  },
  workOrder: {
    label: 'WORK ORDER',
    aliases: ['work order', 'workorder', 'wo', 'w o', 'work order number', 'work order no', 'work order n', 'client wo number', 'client wo', 'wo number', 'wo no', 'wo n', 'ordem de trabalho', 'ordem trabalho'],
  },
  idRepair: {
    label: 'ID REPAIR',
    aliases: ['id repair', 'idrepair', 'repair id', 'repairid', 'repair', 'repair no', 'repair number', 'project', 'project code', 'project number', 'project no', 'project id', 'projeto', 'bsp', 'bsp project', 'id reparo', 'id do reparo', 'reparo id', 'id da reparacao', 'id da reparação'],
  },
  tagNumber: {
    label: 'TAG NUMBER',
    aliases: ['tag number', 'tagnumber', 'tag no', 'tag n', 'tag numero', 'tag número', 'numero tag', 'n tag', 'tag', 'line n', 'line no', 'line nº', 'line number', 'line num', 'line', 'linha'],
  },
  sob: {
    label: 'SOB',
    aliases: ['sob', 's o b', 's.o.b', 's o b ', 'sob number', 'sob no', 'sch', 'schedule'],
  },
  description: {
    label: 'DESCRIPTION',
    aliases: ['description', 'descricao', 'descrição', 'desc', 'descritivo', 'scope', 'escopo', 'observations', 'observation', 'observacoes', 'observações', 'obs', 'remarks', 'comments'],
  },
  status: {
    label: 'STATUS',
    aliases: ['status', 'situacao', 'situação', 'estado', 'stage', 'fase', 'class', 'classe', 'categoria'],
  },
  image: {
    label: 'IMAGEM',
    aliases: ['imagem', 'image', 'images', 'foto', 'fotos', 'photo', 'picture', 'anexo', 'attachment', 'attachments'],
  },
};

function json(statusCode, payload, extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store, no-cache, must-revalidate',
      ...extraHeaders,
    },
    body: JSON.stringify(payload),
  };
}

function normalizeTitle(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[º°ª]/g, ' ')
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compact(value) {
  return normalizeTitle(value).replace(/\s+/g, '');
}

function unique(values) {
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
}

function extractSmartsheetLinkHash(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const direct = text.match(/app\.smartsheet\.com\/sheets\/([A-Za-z0-9]+)/i);
  if (direct?.[1]) return direct[1];
  // Também aceita somente o hash colado na ENV/query.
  if (/^[A-Za-z0-9]{20,}$/.test(text) && !/^\d+$/.test(text)) return text;
  return '';
}

function getCandidateSheetLinkHashes(event) {
  const qs = event.queryStringParameters || {};
  return unique([
    qs.sheetLink,
    qs.sheetUrl,
    qs.permalink,
    qs.sheetHash,
    YINSON_UNDER_DEV_PRIMARY_SHEET_LINK,
    YINSON_UNDER_DEV_PRIMARY_SHEET_HASH,
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_URL,
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_LINK,
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_HASH,
    process.env.SMARTSHEET_UNDER_DEV_SHEET_URL,
    process.env.SMARTSHEET_UNDER_DEV_SHEET_LINK,
  ].map(extractSmartsheetLinkHash));
}

function buildCacheKey(event) {
  const qs = event.queryStringParameters || {};
  return [
    qs.sheetId || '',
    qs.sheetIds || '',
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_ID || '',
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_URL || '',
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_LINK || '',
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_HASH || '',
    process.env.SMARTSHEET_UNDER_DEV_SHEET_ID || '',
    process.env.SMARTSHEET_CLIENT_UNDER_DEV_SHEET_ID || '',
    process.env.SMARTSHEET_YINSON_SHEET_ID || '',
    process.env.SMARTSHEET_SHEET_ID_YINSON || '',
    process.env.SMARTSHEET_SHEET_IDS_LIST || '',
    process.env.SMARTSHEET_SHEET_ID_BR || '',
    process.env.SMARTSHEET_SHEET_ID || '',
    qs.rowAttachmentFallback || '',
  ].join('|');
}

function getCachedPayload(event, debug) {
  if (debug || !cachedPayload || !UNDER_DEV_CACHE_TTL_MS) return null;
  const key = buildCacheKey(event);
  const age = Date.now() - cachedPayloadAt;
  if (cachedPayloadKey === key && age >= 0 && age < UNDER_DEV_CACHE_TTL_MS) {
    return {
      ...cachedPayload,
      meta: {
        ...(cachedPayload.meta || {}),
        servedFromFunctionCache: true,
        functionCacheAgeMs: age,
      },
    };
  }
  return null;
}

function setCachedPayload(event, payload, debug) {
  if (debug || !payload?.ok || !UNDER_DEV_CACHE_TTL_MS) return;
  cachedPayloadKey = buildCacheKey(event);
  cachedPayloadAt = Date.now();
  cachedPayload = payload;
}

function getCandidateSheetIds(event) {
  const qs = event.queryStringParameters || {};
  const rawValues = [
    qs.sheetId,
    qs.sheetIds,
    YINSON_UNDER_DEV_PRIMARY_SHEET_ID,
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_CLIENT_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_YINSON_SHEET_ID,
    process.env.SMARTSHEET_SHEET_ID_YINSON,
    process.env.SMARTSHEET_SHEET_IDS_LIST,
    process.env.SMARTSHEET_SHEET_ID_BR,
    process.env.SMARTSHEET_SHEET_ID,
    ...DEFAULT_CANDIDATE_SHEET_IDS,
  ];

  return unique(rawValues.flatMap((item) => String(item || '').split(',')));
}

function getCandidateSheetNames() {
  return unique([
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_NAME,
    process.env.SMARTSHEET_UNDER_DEV_SHEET_NAME,
    process.env.SMARTSHEET_CLIENT_UNDER_DEV_SHEET_NAME,
    'Projetos em Desenvolvimento',
    'Projects under development',
    'Yinson',
  ]).map(normalizeTitle);
}

async function fetchJson(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    const text = await response.text();
    let payload = null;
    try { payload = text ? JSON.parse(text) : null; } catch (_) { payload = text; }
    if (!response.ok) {
      const error = new Error(`Smartsheet HTTP ${response.status}`);
      error.status = response.status;
      error.payload = payload;
      throw error;
    }
    return payload;
  } finally {
    clearTimeout(timeout);
  }
}

function scoreColumnForField(columnTitle, fieldKey) {
  const title = normalizeTitle(columnTitle);
  const compactTitle = compact(columnTitle);
  const definition = FIELD_DEFINITIONS[fieldKey];
  if (!title || !definition) return 0;

  let best = 0;
  for (const alias of definition.aliases) {
    const normalizedAlias = normalizeTitle(alias);
    const compactAlias = compact(alias);
    if (!normalizedAlias) continue;

    if (title === normalizedAlias) best = Math.max(best, 100);
    if (compactTitle === compactAlias) best = Math.max(best, 98);
    if (title.startsWith(`${normalizedAlias} `)) best = Math.max(best, 88);
    if (title.endsWith(` ${normalizedAlias}`)) best = Math.max(best, 82);
    if (title.includes(` ${normalizedAlias} `)) best = Math.max(best, 78);
    if (compactAlias.length >= 4 && compactTitle.includes(compactAlias)) best = Math.max(best, 72);
  }

  // Tratamentos específicos para evitar falso positivo e capturar variações comuns.
  if (fieldKey === 'sob' && compactTitle === 'sob') best = Math.max(best, 105);
  if (fieldKey === 'workOrder' && /(^|\s)wo($|\s|\d)/.test(title)) best = Math.max(best, 92);
  if (fieldKey === 'tagNumber' && title === 'tag') best = Math.max(best, 92);
  if (fieldKey === 'tagNumber' && /^line(\s|$)/.test(title)) best = Math.max(best, 88);
  if (fieldKey === 'idRepair' && title === 'id') best = Math.max(best, 55);
  if (fieldKey === 'description' && /(^|\s)obs(ervations?)?($|\s)/.test(title)) best = Math.max(best, 92);
  if (fieldKey === 'idRepair' && /(^|\s)(project|projeto|bsp)($|\s|\d)/.test(title)) best = Math.max(best, 92);
  if (fieldKey === 'status' && title === 'class') best = Math.max(best, 90);

  return best;
}

function mapColumns(columns = []) {
  const usedColumnIds = new Set();
  const mapped = {};
  const diagnostics = [];

  for (const fieldKey of Object.keys(FIELD_DEFINITIONS)) {
    const candidates = columns
      .map((column) => ({ column, score: scoreColumnForField(column.title, fieldKey) }))
      .filter((item) => item.score > 0 && !usedColumnIds.has(item.column.id))
      .sort((a, b) => b.score - a.score);

    const best = candidates[0];
    if (best) {
      mapped[fieldKey] = {
        id: best.column.id,
        title: best.column.title,
        score: best.score,
      };
      usedColumnIds.add(best.column.id);
      diagnostics.push({ field: fieldKey, label: FIELD_DEFINITIONS[fieldKey].label, matched: best.column.title, score: best.score });
    } else {
      diagnostics.push({ field: fieldKey, label: FIELD_DEFINITIONS[fieldKey].label, matched: null, score: 0 });
    }
  }

  const missing = Object.keys(FIELD_DEFINITIONS).filter((fieldKey) => !mapped[fieldKey]);
  return { mapped, missing, diagnostics };
}


function findColumnByAliases(columns = [], aliases = [], used = new Set()) {
  const normalizedAliases = aliases.map(normalizeTitle).filter(Boolean);
  const compactAliases = aliases.map(compact).filter(Boolean);
  let best = null;
  let bestScore = 0;
  for (const column of columns) {
    if (!column || used.has(column.id)) continue;
    const title = normalizeTitle(column.title);
    const compactTitle = compact(column.title);
    let score = 0;
    normalizedAliases.forEach((alias, index) => {
      const cAlias = compactAliases[index] || compact(alias);
      if (title === alias) score = Math.max(score, 100);
      if (compactTitle === cAlias) score = Math.max(score, 98);
      if (alias && title.includes(alias)) score = Math.max(score, 84);
      if (cAlias && cAlias.length >= 3 && compactTitle.includes(cAlias)) score = Math.max(score, 76);
    });
    if (score > bestScore) {
      bestScore = score;
      best = column;
    }
  }
  return best ? { column: best, score: bestScore } : null;
}

function applyYinsonKnownColumnFallback(sheet, mapping) {
  const columns = sheet?.columns || [];
  const mapped = { ...(mapping?.mapped || {}) };
  const diagnostics = [...(mapping?.diagnostics || [])];
  const used = new Set(Object.values(mapped).map((info) => info.id).filter(Boolean));
  const knownAliases = {
    vessel: ['Vessel', 'Vessel / Unit', 'Unidade', 'Unit'],
    priorityLevel: ['Priority Level', 'Priority', 'Priority LVL', 'Prioridade', 'Nível de Prioridade', 'Nivel de Prioridade'],
    workOrder: ['Client WO number', 'Client WO', 'Work Order', 'WO', 'WO Number', 'WO Nº'],
    idRepair: ['ID Repair', 'Repair', 'Project', 'Project Number', 'Project No', 'Projeto', 'BSP'],
    tagNumber: ['Tag Number', 'Tag', 'Line Nº', 'Line N', 'Line No', 'Line Number', 'Line Num', 'Line'],
    sob: ['SOB', 'SCH', 'Schedule'],
    description: ['Description', 'Observations', 'Observation', 'OBSERVATIONS', 'Comments', 'Remarks', 'Scope'],
    status: ['Status', 'Class', 'Classe', 'Stage', 'Situação'],
    image: ['Imagem', 'Image', 'Photo', 'Attachment', 'Attachments'],
  };

  for (const [fieldKey, aliases] of Object.entries(knownAliases)) {
    if (mapped[fieldKey]) continue;
    const match = findColumnByAliases(columns, aliases, used);
    if (match?.column?.id) {
      mapped[fieldKey] = { id: match.column.id, title: match.column.title, score: match.score, fallback: true };
      used.add(match.column.id);
      diagnostics.push({ field: fieldKey, label: FIELD_DEFINITIONS[fieldKey]?.label || fieldKey, matched: match.column.title, score: match.score, fallback: 'known-yinson-columns' });
    }
  }

  // Fallback final para a estrutura observada na planilha Yinson/WIP:
  // Client WO number | Class | Project | Client | Vessel | Line Nº | OBSERVATIONS | ...
  const byTitle = new Map(columns.map((column) => [compact(column.title), column]));
  const positional = {
    priorityLevel: ['prioritylevel', 'priority', 'prioridade', 'nivelprioridade', 'niveldeprioridade'],
    workOrder: ['clientwonumber', 'clientwo', 'workorder'],
    status: ['class', 'status'],
    idRepair: ['project', 'projeto', 'bsp'],
    vessel: ['vessel', 'unit', 'unidade'],
    tagNumber: ['lineno', 'linen', 'linenumber', 'tagnumber', 'tag'],
    description: ['observations', 'observation', 'descricao', 'description'],
    sob: ['sob', 'sch'],
  };
  for (const [fieldKey, keys] of Object.entries(positional)) {
    if (mapped[fieldKey]) continue;
    const column = keys.map((key) => byTitle.get(key)).find(Boolean);
    if (column && !used.has(column.id)) {
      mapped[fieldKey] = { id: column.id, title: column.title, score: 70, fallback: true };
      used.add(column.id);
      diagnostics.push({ field: fieldKey, label: FIELD_DEFINITIONS[fieldKey]?.label || fieldKey, matched: column.title, score: 70, fallback: 'compact-title' });
    }
  }

  const missing = Object.keys(FIELD_DEFINITIONS).filter((fieldKey) => !mapped[fieldKey]);
  return { mapped, missing, diagnostics };
}

function scoreSheet(sheet, mapping) {
  const matchedCount = Object.keys(mapping.mapped || {}).length;
  const title = normalizeTitle(sheet?.name || sheet?.title || '');
  let score = matchedCount * 100;
  if (title.includes('yinson')) score += 90;
  if (title.includes('desenvolvimento') || title.includes('development')) score += 80;
  if (title.includes('repair') || title.includes('reparo')) score += 30;
  return score;
}

function compactSmartsheetError(err) {
  const payload = err?.payload;
  const message = payload?.message || payload?.error || payload?.errorCode || payload?.refId || err?.message || 'erro';
  return { error: String(message).slice(0, 240), status: err?.status || null };
}

async function fetchSheetPagedBasic(sheetId, headers, pageSize = UNDER_DEV_PAGE_SIZE) {
  const encoded = encodeURIComponent(sheetId);
  let firstPage = null;
  const allRows = [];
  let page = 1;
  let totalPages = 1;

  while (page <= totalPages && page <= UNDER_DEV_MAX_PAGES) {
    const separator = '?';
    const url = `${API_BASE}/sheets/${encoded}${separator}pageSize=${encodeURIComponent(pageSize)}&page=${encodeURIComponent(page)}`;
    const payload = await fetchJson(url, { headers });
    if (!firstPage) firstPage = payload;
    if (Array.isArray(payload?.rows)) allRows.push(...payload.rows);

    const reportedTotalPages = Number(payload?.totalPages || payload?.page?.totalPages || 0);
    if (reportedTotalPages > 0) {
      totalPages = Math.min(reportedTotalPages, UNDER_DEV_MAX_PAGES);
    } else {
      const returned = Array.isArray(payload?.rows) ? payload.rows.length : 0;
      totalPages = returned >= pageSize ? page + 1 : page;
    }
    page += 1;
  }

  if (!firstPage) throw new Error(`Smartsheet sem resposta para sheet ${sheetId}`);
  return {
    ...firstPage,
    rows: allRows.length ? allRows : (firstPage.rows || []),
    __fetchMode: `paged-basic-${pageSize}`,
    __pagesLoaded: page - 1,
    __maxPages: UNDER_DEV_MAX_PAGES,
  };
}

async function fetchSheet(sheetId, headers) {
  // v37.40: o Yinson estava estourando o limite de 30s do Netlify porque tentava
  // abrir IDs candidatos em modo pesado e depois fazia várias tentativas em cascata.
  // Agora a leitura padrão é uma tentativa leve e paginada. Sem anexos/objectValue.
  // Se falhar, devolve erro rápido para permitir testar outro ID/hash ou manter cache.
  const errors = [];
  try {
    const sheet = await fetchSheetPagedBasic(sheetId, headers, UNDER_DEV_PAGE_SIZE);
    return { ...sheet, __fetchFallbackErrors: errors };
  } catch (err) {
    errors.push({ mode: `paged-basic-${UNDER_DEV_PAGE_SIZE}`, ...compactSmartsheetError(err) });
    const finalError = new Error(`Smartsheet HTTP ${err?.status || 'erro'} ao buscar sheet ${sheetId}`);
    finalError.status = err?.status || 500;
    finalError.payload = { attempts: errors };
    throw finalError;
  }
}

async function listSheets(headers) {
  const all = [];
  let page = 1;
  let totalPages = 1;
  const pageSize = 100;
  while (page <= totalPages && page <= 20) {
    const payload = await fetchJson(`${API_BASE}/sheets?pageSize=${pageSize}&page=${page}`, { headers });
    if (Array.isArray(payload?.data)) all.push(...payload.data);
    const reported = Number(payload?.totalPages || payload?.page?.totalPages || 0);
    if (reported > 0) {
      totalPages = Math.min(reported, 20);
    } else {
      totalPages = Array.isArray(payload?.data) && payload.data.length >= pageSize ? page + 1 : page;
    }
    page += 1;
  }
  return all;
}

function sheetInfoMatchesHash(sheetInfo, hashes = []) {
  const permalink = String(sheetInfo?.permalink || sheetInfo?.accessUrl || sheetInfo?.url || '');
  const text = `${permalink} ${sheetInfo?.name || ''}`;
  return hashes.find((hash) => hash && text.includes(hash)) || '';
}

async function findSheetInfosByHash(headers, hashes = [], tried = []) {
  const cleanHashes = unique(hashes).filter(Boolean);
  if (!cleanHashes.length) return [];

  const matches = [];
  let page = 1;
  let totalPages = 1;
  const pageSize = 100;

  while (page <= totalPages && page <= YINSON_UNDER_DEV_LIST_MAX_PAGES) {
    try {
      const payload = await fetchJson(`${API_BASE}/sheets?pageSize=${pageSize}&page=${page}`, { headers });
      const data = Array.isArray(payload?.data) ? payload.data : [];
      for (const sheetInfo of data) {
        const matchedHash = sheetInfoMatchesHash(sheetInfo, cleanHashes);
        if (matchedHash) matches.push({ sheetInfo, matchedHash, page });
      }
      tried.push({ source: 'sheet-hash-list', page, found: matches.length, rows: data.length });
      if (matches.length) return matches;

      const reported = Number(payload?.totalPages || payload?.page?.totalPages || 0);
      if (reported > 0) totalPages = Math.min(reported, YINSON_UNDER_DEV_LIST_MAX_PAGES);
      else totalPages = data.length >= pageSize ? page + 1 : page;
      page += 1;
    } catch (err) {
      tried.push({ source: 'sheet-hash-list', page, error: err.message, status: err.status || null });
      return matches;
    }
  }

  return matches;
}

async function resolveBestSheet(event, headers) {
  const debug = event.queryStringParameters?.debug === '1' || event.queryStringParameters?.debug === 'true';
  const tried = [];
  const candidates = [];
  const candidateIds = getCandidateSheetIds(event);
  const sheetHashes = getCandidateSheetLinkHashes(event);

  // v37.40: quando temos o link/hash da planilha, ele é mais confiável que o ID
  // numérico informado. Primeiro tentamos localizar pela lista de sheets acessíveis
  // e só depois abrimos a sheet encontrada. Isso evita gastar 20/30s tentando IDs
  // que retornam HTTP 500 antes de chegar na sheet correta.
  if (sheetHashes.length) {
    const hashMatches = await findSheetInfosByHash(headers, sheetHashes, tried);
    for (const { sheetInfo, matchedHash, page } of hashMatches.slice(0, 3)) {
      try {
        const sheet = await fetchSheet(sheetInfo.id, headers);
        const mapping = applyYinsonKnownColumnFallback(sheet, mapColumns(sheet.columns || []));
        const score = scoreSheet(sheet, mapping) + 1200;
        const item = { sheet, mapping, score, source: 'sheet-hash' };
        candidates.push(item);
        tried.push({
          sheetId: String(sheetInfo.id),
          name: sheet.name,
          fetchMode: sheet.__fetchMode || 'unknown',
          score,
          mapped: Object.keys(mapping.mapped).length,
          missing: mapping.missing,
          source: 'sheet-hash',
          matchedHash,
          page,
          permalink: sheetInfo.permalink || null,
          fallbackErrors: sheet.__fetchFallbackErrors || [],
        });
      } catch (err) {
        tried.push({
          sheetId: String(sheetInfo.id),
          name: sheetInfo.name,
          error: err.message,
          status: err.status || null,
          source: 'sheet-hash',
          matchedHash,
          page,
          permalink: sheetInfo.permalink || null,
          attempts: err.payload?.attempts || null,
        });
      }
    }
  }

  const bestHashCandidate = candidates.sort((a, b) => b.score - a.score)[0];
  if (bestHashCandidate && Object.keys(bestHashCandidate.mapping.mapped || {}).length >= 2) {
    if (debug) bestHashCandidate.debugTried = tried;
    return bestHashCandidate;
  }

  // Se o hash não localizar, tenta IDs explícitos, mas de forma limitada para não
  // estourar a function. O ID informado pode estar instável/sem acesso e devolver 500.
  for (const sheetId of candidateIds.slice(0, 4)) {
    try {
      const sheet = await fetchSheet(sheetId, headers);
      const mapping = applyYinsonKnownColumnFallback(sheet, mapColumns(sheet.columns || []));
      const score = scoreSheet(sheet, mapping);
      const item = { sheet, mapping, score, source: 'candidate-id' };
      candidates.push(item);
      tried.push({ sheetId, name: sheet.name, fetchMode: sheet.__fetchMode || 'unknown', score, mapped: Object.keys(mapping.mapped).length, missing: mapping.missing, fallbackErrors: sheet.__fetchFallbackErrors || [] });
    } catch (err) {
      tried.push({ sheetId, error: err.message, status: err.status || null, attempts: err.payload?.attempts || null });
    }
  }

  const bestCandidate = candidates.sort((a, b) => b.score - a.score)[0];
  if (bestCandidate && Object.keys(bestCandidate.mapping.mapped || {}).length >= 2) {
    if (debug) bestCandidate.debugTried = tried;
    return bestCandidate;
  }

  // Varredura por nome só como última alternativa e limitada.
  try {
    const sheetNames = getCandidateSheetNames();
    const listed = await listSheets(headers);
    const rankedList = listed
      .map((sheetInfo) => {
        const normalizedName = normalizeTitle(sheetInfo.name);
        let nameScore = 0;
        for (const name of sheetNames) {
          if (name && normalizedName.includes(name)) nameScore += 100;
        }
        if (normalizedName.includes('yinson')) nameScore += 100;
        if (normalizedName.includes('desenvolvimento') || normalizedName.includes('development')) nameScore += 80;
        if (normalizedName.includes('repair') || normalizedName.includes('reparo')) nameScore += 40;
        return { sheetInfo, nameScore };
      })
      .filter((item) => item.nameScore > 0)
      .sort((a, b) => b.nameScore - a.nameScore)
      .slice(0, 5);

    for (const { sheetInfo, nameScore } of rankedList) {
      if (candidateIds.includes(String(sheetInfo.id))) continue;
      try {
        const sheet = await fetchSheet(sheetInfo.id, headers);
        const mapping = applyYinsonKnownColumnFallback(sheet, mapColumns(sheet.columns || []));
        const score = scoreSheet(sheet, mapping) + nameScore;
        const item = { sheet, mapping, score, source: 'sheet-list' };
        candidates.push(item);
        tried.push({ sheetId: String(sheetInfo.id), name: sheet.name, fetchMode: sheet.__fetchMode || 'unknown', score, mapped: Object.keys(mapping.mapped).length, missing: mapping.missing, source: 'sheet-list', permalink: sheetInfo.permalink || null, fallbackErrors: sheet.__fetchFallbackErrors || [] });
      } catch (err) {
        tried.push({ sheetId: String(sheetInfo.id), name: sheetInfo.name, error: err.message, status: err.status || null, source: 'sheet-list', permalink: sheetInfo.permalink || null, attempts: err.payload?.attempts || null });
      }
    }
  } catch (err) {
    tried.push({ source: 'sheet-list', error: err.message, status: err.status || null });
  }

  const best = candidates.sort((a, b) => b.score - a.score)[0];
  const mappedCount = Object.keys(best?.mapping?.mapped || {}).length;
  const isPrimaryYinsonSheet = String(best?.sheet?.id || '') === YINSON_UNDER_DEV_PRIMARY_SHEET_ID;
  if (!best || mappedCount < (isPrimaryYinsonSheet ? 1 : 2)) {
    const error = new Error('Não foi possível localizar uma sheet com as colunas de Projetos em Desenvolvimento.');
    error.statusCode = 422;
    error.debug = { tried, sheetHashes, candidateIds: candidateIds.slice(0, 8) };
    throw error;
  }

  if (debug) best.debugTried = tried;
  return best;
}

function buildCellMap(row) {
  const map = new Map();
  (row.cells || []).forEach((cell) => map.set(cell.columnId, cell));
  return map;
}

function cellValue(cell) {
  if (!cell) return '';
  const value = cell.displayValue ?? cell.value ?? cell.objectValue ?? '';
  if (value && typeof value === 'object') return JSON.stringify(value);
  return String(value ?? '').trim();
}

function extractUrlsFromValue(value, output = []) {
  if (value == null) return output;
  if (typeof value === 'string') {
    const matches = value.match(/https?:\/\/[^\s"'<>]+/gi) || [];
    matches.forEach((url) => output.push({ url, name: 'Imagem' }));
    return output;
  }
  if (Array.isArray(value)) {
    value.forEach((item) => extractUrlsFromValue(item, output));
    return output;
  }
  if (typeof value === 'object') {
    const directUrl = value.url || value.href || value.link || value.downloadUrl || value.imageUrl;
    if (typeof directUrl === 'string' && /^https?:\/\//i.test(directUrl)) {
      output.push({ url: directUrl, name: value.name || value.title || value.fileName || 'Imagem' });
    }
    Object.values(value).forEach((item) => extractUrlsFromValue(item, output));
  }
  return output;
}

function imageRecordsFromCell(cell) {
  if (!cell) return [];
  const records = [];
  if (cell.hyperlink?.url) records.push({ url: cell.hyperlink.url, name: cell.displayValue || 'Imagem' });
  extractUrlsFromValue(cell.displayValue, records);
  extractUrlsFromValue(cell.value, records);
  extractUrlsFromValue(cell.objectValue, records);
  const seen = new Set();
  return records
    .filter((item) => item?.url && !seen.has(item.url) && seen.add(item.url))
    .map((item, index) => ({
      id: `cell-${index}`,
      name: item.name || 'Imagem',
      url: item.url,
      source: 'image-column',
    }));
}

function isImageAttachment(att) {
  const mime = String(att?.mimeType || '').toLowerCase();
  const name = String(att?.name || '').toLowerCase();
  return mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp|heic|heif)$/i.test(name);
}

async function fetchSheetAttachments(sheetId, headers) {
  try {
    const payload = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/attachments`, { headers });
    const byRow = {};
    for (const att of (payload?.data || [])) {
      if (!att.rowId) continue;
      if (!byRow[att.rowId]) byRow[att.rowId] = [];
      byRow[att.rowId].push(att);
    }
    return byRow;
  } catch (_) {
    return {};
  }
}

async function fetchRowAttachments(sheetId, rowId, headers) {
  try {
    const payload = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/rows/${encodeURIComponent(rowId)}/attachments`, { headers });
    return Array.isArray(payload?.data) ? payload.data : [];
  } catch (_) {
    return [];
  }
}

async function fetchAttachmentDetail(sheetId, att, headers) {
  try {
    // Endpoint por sheet/attachment costuma ser mais confiável para anexos de linha.
    const payload = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/attachments/${encodeURIComponent(att.id)}`, { headers });
    return { ...att, url: payload.url, urlExpiresAt: payload.urlExpiresAt };
  } catch (_) {
    try {
      const payload = await fetchJson(`${API_BASE}/attachments/${encodeURIComponent(att.id)}`, { headers });
      return { ...att, url: payload.url, urlExpiresAt: payload.urlExpiresAt };
    } catch (__) {
      return att;
    }
  }
}

async function mapWithConcurrency(items, limit, mapper) {
  const results = new Array(items.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, items.length || 1) }, async () => {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      results[index] = await mapper(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}


function isTruthyParam(value) {
  return ['1', 'true', 'yes', 'sim', 'force', 'refresh', 'manual'].includes(String(value || '').trim().toLowerCase());
}

function wantsForceRefresh(event) {
  const qs = event.queryStringParameters || {};
  return isTruthyParam(qs.force) || isTruthyParam(qs.refresh) || isTruthyParam(qs.update) || isTruthyParam(qs.sync);
}

function wantsNoPersistentCache(event) {
  const qs = event.queryStringParameters || {};
  return qs.cache === '0' || qs.cache === 'false' || qs.persistentCache === '0' || qs.persistentCache === 'false';
}

function clonePayload(payload) {
  return payload ? JSON.parse(JSON.stringify(payload)) : payload;
}

function decorateYinsonCachePayload(cacheEntry, extras = {}) {
  const payload = clonePayload(cacheEntry?.payload || cacheEntry || { ok: true, data: [] });
  payload.ok = payload.ok !== false;
  if (!Array.isArray(payload.data) && Array.isArray(payload.projects)) payload.data = payload.projects;
  if (!Array.isArray(payload.data)) payload.data = [];
  payload.meta = {
    ...(payload.meta || {}),
    servedFromPersistentCache: true,
    persistentCacheKey: YINSON_UNDER_DEV_CACHE_KEY,
    schemaVersion: payload?.meta?.schemaVersion || null,
    cacheUpdatedAt: cacheEntry?.updatedAt || payload?.meta?.cacheUpdatedAt || payload?.meta?.lastSync || null,
    cacheVersion: cacheEntry?.version || payload?.meta?.version || null,
    cacheAgeMs: cacheEntry?.updatedAt ? Math.max(0, Date.now() - new Date(cacheEntry.updatedAt).getTime()) : null,
    cacheProjectsCount: cacheEntry?.projectsCount || payload.data.length,
    ...extras,
  };
  return payload;
}

async function readYinsonPersistentCache(options = {}) {
  if (!YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE) return null;
  return readTrackingCache(YINSON_UNDER_DEV_CACHE_KEY, {
    maxAgeMs: options.maxAgeMs || YINSON_UNDER_DEV_CACHE_MAX_AGE_MS,
    timeoutMs: options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS,
  });
}

function isUsableYinsonCacheEntry(cacheEntry) {
  const payload = cacheEntry?.payload;
  if (!payload || payload.ok === false) return false;
  if (payload?.meta?.initializing || payload?.meta?.reason === 'refresh-lock-bootstrap') return false;
  const data = Array.isArray(payload.data) ? payload.data : (Array.isArray(payload.projects) ? payload.projects : []);
  return Array.isArray(data);
}

function isCurrentYinsonCacheEntry(cacheEntry) {
  const payload = cacheEntry?.payload || cacheEntry || {};
  if (payload?.meta?.schemaVersion === YINSON_UNDER_DEV_SCHEMA_VERSION) return true;
  const data = Array.isArray(payload.data) ? payload.data : (Array.isArray(payload.projects) ? payload.projects : []);
  if (!Array.isArray(data) || !data.length) return false;
  return data.every((row) => Object.prototype.hasOwnProperty.call(row || {}, 'priorityLevel'));
}

async function clearYinsonLock(reason = 'manual-yinson-lock-clear', timeoutMs = YINSON_UNDER_DEV_CACHE_TIMEOUT_MS) {
  if (!YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE) return false;
  return clearTrackingRefreshLock(YINSON_UNDER_DEV_CACHE_KEY, {
    reason,
    timeoutMs,
  });
}

async function writeYinsonPersistentCache(payload, options = {}) {
  if (!YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE || !payload?.ok) return false;
  const data = Array.isArray(payload.data) ? payload.data : [];
  const payloadToStore = {
    ...payload,
    data,
    // _trackingCache contabiliza projects_count usando payload.projects.
    // Mantemos também `data` porque é o contrato usado pela página Yinson.
    projects: data,
    meta: {
      ...(payload.meta || {}),
      persistentCacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      schemaVersion: YINSON_UNDER_DEV_SCHEMA_VERSION,
      lastSync: new Date().toISOString(),
      cacheSource: options.source || 'client-under-dev',
    },
  };
  return writeTrackingCache(YINSON_UNDER_DEV_CACHE_KEY, payloadToStore, {
    scope: 'yinson-under-development',
    source: options.source || 'client-under-dev',
    reason: options.reason || 'yinson-under-dev-refresh',
    version: payloadToStore.meta?.sheetVersion || payloadToStore.meta?.version || payloadToStore.meta?.sheetId || '',
    forceWrite: true,
    minWriteIntervalMs: 0,
    timeoutMs: options.timeoutMs || Math.max(7000, YINSON_UNDER_DEV_CACHE_TIMEOUT_MS),
  });
}

async function buildYinsonLivePayload(event, headers, debug) {
  const { sheet, mapping, score, source, debugTried } = await resolveBestSheet(event, headers);
  const mapped = mapping.mapped;
  const sheetId = String(sheet.id);
  const attachmentsByRow = await fetchSheetAttachments(sheetId, headers);

  const rowsWithData = (sheet.rows || []).filter((row) => {
    const cellMap = buildCellMap(row);
    return Object.values(mapped).some((field) => cellValue(cellMap.get(field.id)));
  });

  const hasSheetAttachmentIndex = Object.keys(attachmentsByRow || {}).length > 0;
  const qs = event.queryStringParameters || {};
  const rowFallbackDisabled = qs.rowAttachmentFallback !== '1' && qs.rowAttachmentFallback !== 'true';
  const useRowAttachmentFallback = !rowFallbackDisabled && rowsWithData.length <= ROW_ATTACHMENT_FALLBACK_LIMIT;

  const projects = await mapWithConcurrency(rowsWithData, useRowAttachmentFallback ? 6 : 1, async (row) => {
    const cellMap = buildCellMap(row);
    const getVal = (fieldKey) => cellValue(cellMap.get(mapped[fieldKey]?.id));

    let rowAttachments = Array.isArray(row.attachments) ? row.attachments : [];
    if (!rowAttachments.length) rowAttachments = attachmentsByRow[row.id] || [];
    if (!rowAttachments.length && useRowAttachmentFallback) {
      rowAttachments = await fetchRowAttachments(sheetId, row.id, headers);
    }

    const imageAttachments = rowAttachments.filter(isImageAttachment);
    const imageCellRecords = imageRecordsFromCell(cellMap.get(mapped.image?.id));

    return {
      rowId: row.id,
      vessel: getVal('vessel'),
      priorityLevel: getVal('priorityLevel'),
      workOrder: getVal('workOrder'),
      idRepair: getVal('idRepair'),
      tagNumber: getVal('tagNumber'),
      sob: getVal('sob'),
      description: getVal('description'),
      status: getVal('status'),
      images: [
        ...imageAttachments.map((att) => ({
          id: att.id,
          name: att.name,
          mimeType: att.mimeType,
          rowId: row.id,
          sheetId,
          source: 'row-attachment',
        })),
        ...imageCellRecords,
      ],
    };
  });

  const meta = {
    sheetId,
    sheetName: sheet.name,
    source,
    fetchMode: sheet.__fetchMode || 'unknown',
    score,
    totalRows: Array.isArray(sheet.rows) ? sheet.rows.length : 0,
    returnedRows: projects.length,
    mappedColumns: Object.fromEntries(Object.entries(mapped).map(([fieldKey, info]) => [fieldKey, { id: info.id, title: info.title, score: info.score }])),
    missingColumns: mapping.missing,
    availableColumns: (sheet.columns || []).map((col) => col.title),
    imageMode: 'metadata-on-click-row-fallback',
    usedSheetAttachmentIndex: hasSheetAttachmentIndex,
    usedRowAttachmentFallback: useRowAttachmentFallback,
    rowAttachmentFallbackLimit: ROW_ATTACHMENT_FALLBACK_LIMIT,
    cacheTtlMs: UNDER_DEV_CACHE_TTL_MS,
    schemaVersion: YINSON_UNDER_DEV_SCHEMA_VERSION,
    sheetVersion: sheet.version || sheet.modifiedAt || sheet.permalink || sheetId,
    lastSync: new Date().toISOString(),
  };
  if (debug) meta.tried = debugTried || [];

  return { ok: true, data: projects, projects, meta };
}

async function refreshYinsonUnderDevCache(event = {}, options = {}) {
  const debug = event.queryStringParameters?.debug === '1' || event.queryStringParameters?.debug === 'true';
  const headers = { Authorization: `Bearer ${TOKEN}` };
  const owner = String(options.owner || `yinson-${options.source || 'manual'}-${process.env.SITE_NAME || 'step'}-${Date.now()}`).slice(0, 180);
  let lock = null;
  let previousCache = null;

  if (!TOKEN) {
    previousCache = await readYinsonPersistentCache().catch(() => null);
    if (isUsableYinsonCacheEntry(previousCache)) {
      return {
        payload: decorateYinsonCachePayload(previousCache, {
          refreshSkipped: true,
          refreshError: 'Token Smartsheet não configurado.',
          refreshSource: options.source || 'manual',
        }),
        refreshed: false,
        cacheFallback: true,
      };
    }
    const error = new Error('Token Smartsheet não configurado. Configure SMARTSHEET_API_KEY ou SMARTSHEET_TOKEN no Netlify.');
    error.statusCode = 500;
    throw error;
  }

  previousCache = await readYinsonPersistentCache().catch(() => null);

  if (YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE) {
    lock = await acquireTrackingRefreshLock(YINSON_UNDER_DEV_CACHE_KEY, {
      ttlMs: options.lockTtlMs || YINSON_UNDER_DEV_REFRESH_LOCK_TTL_MS,
      owner,
      timeoutMs: options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS,
    });

    if (!lock?.acquired) {
      if (isUsableYinsonCacheEntry(previousCache)) {
        return {
          payload: decorateYinsonCachePayload(previousCache, {
            refreshSkipped: true,
            refreshLockReason: lock?.reason || 'lock-held',
            refreshLockUntil: lock?.lockedUntil || null,
            refreshSource: options.source || 'manual',
          }),
          refreshed: false,
          cacheFallback: true,
          lock,
        };
      }

      // v37.38: quando ainda não existe cache útil do Yinson, um lock preso não pode
      // impedir a primeira sincronização. Isso acontecia depois de uma tentativa que
      // criava a linha bootstrap, falhava no Smartsheet e deixava o usuário preso em
      // "lock-not-acquired". Para force/manual/scheduled, limpamos o lock da chave
      // Yinson e tentamos adquirir uma única vez novamente.
      const mayClearBootstrapLock = Boolean(
        wantsForceRefresh(event) ||
        String(options.source || '').includes('manual') ||
        String(options.source || '').includes('scheduled') ||
        String(options.reason || '').includes('cache-miss')
      );
      if (mayClearBootstrapLock) {
        const cleared = await clearYinsonLock(`clear-yinson-lock-before-${options.source || 'refresh'}`, options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS);
        if (cleared) {
          lock = await acquireTrackingRefreshLock(YINSON_UNDER_DEV_CACHE_KEY, {
            ttlMs: options.lockTtlMs || YINSON_UNDER_DEV_REFRESH_LOCK_TTL_MS,
            owner: `${owner}-retry`,
            timeoutMs: options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS,
          });
        }
      }

      if (!lock?.acquired) {
        // v37.39: bootstrap do Yinson não pode ficar preso em lock-not-acquired.
        // Quando ainda não existe cache útil, preferimos tentar uma sincronização
        // controlada sem lock em vez de deixar a tela permanentemente vazia.
        // Isso só acontece na primeira carga/recuperação; depois que o cache é
        // salvo, o fluxo volta a usar Supabase/cache normalmente.
        const canBypassInitialLock = !isUsableYinsonCacheEntry(previousCache);
        if (canBypassInitialLock) {
          lock = {
            acquired: false,
            bypassed: true,
            reason: lock?.reason || 'lock-not-acquired',
            original: lock || null,
          };
        } else {
          const error = new Error(`Atualização Yinson já está em andamento ou lock indisponível: ${lock?.reason || 'lock-held'}`);
          error.statusCode = 423;
          error.debug = { lock, attemptedLockClear: mayClearBootstrapLock };
          throw error;
        }
      }
    }
  }

  try {
    // Marca a tentativa para facilitar diagnóstico, mas sem destruir payload existente.
    if (isUsableYinsonCacheEntry(previousCache) && YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE) {
      await touchTrackingCache(YINSON_UNDER_DEV_CACHE_KEY, {
        reason: `${options.source || 'manual'}-yinson-refresh-started`,
        source: options.source || 'client-under-dev',
        scope: 'yinson-under-development',
        timeoutMs: options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS,
      }).catch(() => null);
    }

    const payload = await buildYinsonLivePayload(event, headers, debug);
    const persisted = await writeYinsonPersistentCache(payload, {
      source: options.source || 'client-under-dev',
      reason: options.reason || 'yinson-under-dev-refresh-success',
      timeoutMs: options.timeoutMs,
    });
    payload.meta = {
      ...(payload.meta || {}),
      persistentCacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      persistentCacheWritten: Boolean(persisted),
      servedFromPersistentCache: false,
      refreshSource: options.source || 'manual',
    };
    setCachedPayload(event, payload, debug);
    return { payload, refreshed: true, cacheFallback: false, lock };
  } catch (error) {
    if (isUsableYinsonCacheEntry(previousCache)) {
      return {
        payload: decorateYinsonCachePayload(previousCache, {
          staleCacheKept: true,
          refreshFailed: true,
          refreshError: String(error?.message || error).slice(0, 500),
          refreshErrorDetails: error?.payload || null,
          refreshDebug: error?.debug || null,
          refreshSource: options.source || 'manual',
        }),
        refreshed: false,
        cacheFallback: true,
        error,
        lock,
      };
    }
    throw error;
  } finally {
    if (lock?.acquired && lock.owner) {
      await releaseTrackingRefreshLock(YINSON_UNDER_DEV_CACHE_KEY, lock.owner, {
        reason: 'yinson-refresh-lock-released',
        timeoutMs: options.timeoutMs || YINSON_UNDER_DEV_CACHE_TIMEOUT_MS,
      }).catch(() => null);
    }
  }
}

exports.handler = async (event) => {
  const debug = event.queryStringParameters?.debug === '1' || event.queryStringParameters?.debug === 'true';
  const forceRefresh = wantsForceRefresh(event);
  const noPersistentCache = wantsNoPersistentCache(event);
  const qs = event.queryStringParameters || {};
  if (isTruthyParam(qs.clearLock) || isTruthyParam(qs.resetLock)) {
    const cleared = await clearYinsonLock('manual-query-clear-yinson-lock');
    return json(200, {
      ok: true,
      cacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      lockCleared: Boolean(cleared),
      meta: await getTrackingCacheMeta(YINSON_UNDER_DEV_CACHE_KEY).catch(() => null),
    });
  }

  const functionCached = (!debug && !forceRefresh) ? getCachedPayload(event, false) : null;
  if (functionCached) return json(200, functionCached);

  const persistentCache = (!noPersistentCache && !forceRefresh)
    ? await readYinsonPersistentCache().catch(() => null)
    : null;

  if (isUsableYinsonCacheEntry(persistentCache) && isCurrentYinsonCacheEntry(persistentCache) && !forceRefresh) {
    return json(200, decorateYinsonCachePayload(persistentCache));
  }

  try {
    const result = await refreshYinsonUnderDevCache(event, {
      source: forceRefresh ? 'client-under-dev-manual-refresh' : 'client-under-dev-cache-miss',
      reason: forceRefresh ? 'manual-refresh-yinson-under-dev' : 'cache-miss-yinson-under-dev',
    });
    return json(200, result.payload);
  } catch (err) {
    const latestCache = (!noPersistentCache)
      ? await readYinsonPersistentCache().catch(() => null)
      : null;
    if (isUsableYinsonCacheEntry(latestCache)) {
      return json(200, decorateYinsonCachePayload(latestCache, {
        refreshFailed: true,
        refreshError: String(err?.message || err).slice(0, 500),
        refreshErrorDetails: err?.payload || null,
        refreshDebug: err?.debug || null,
      }));
    }

    return json(err.statusCode || err.status || 500, {
      ok: false,
      error: err.message || 'Erro ao buscar dados da Smartsheet',
      details: err.payload || null,
      debug: {
        ...(err.debug || {}),
        cache: {
          key: YINSON_UNDER_DEV_CACHE_KEY,
          configured: YINSON_UNDER_DEV_ENABLE_PERSISTENT_CACHE,
          meta: await getTrackingCacheMeta(YINSON_UNDER_DEV_CACHE_KEY).catch(() => null),
        },
      },
    });
  }
};

exports.refreshYinsonUnderDevCache = refreshYinsonUnderDevCache;
exports.YINSON_UNDER_DEV_CACHE_KEY = YINSON_UNDER_DEV_CACHE_KEY;
