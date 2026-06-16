const { jsonResponse, requireSession } = require('./_auth');

const API_BASE = process.env.SMARTSHEET_API_BASE || 'https://api.smartsheet.com/2.0';
const TOKEN = process.env.SMARTSHEET_TOKEN || process.env.SMARTSHEET_ACCESS_TOKEN || process.env.SMARTSHEET_API_TOKEN || process.env.SMARTSHEET_BEARER_TOKEN || process.env.SMARTSHEET_PAT || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN || '';
const DRAWING_SHEET_ID = process.env.SMARTSHEET_DRAWING_CONTROL_SHEET_ID || process.env.SMARTSHEET_DRAWING_DOCUMENTATION_SHEET_ID || '';
const DRAWING_SHEET_URL = process.env.SMARTSHEET_DRAWING_CONTROL_SHEET_URL || process.env.SMARTSHEET_DRAWING_DOCUMENTATION_SHEET_URL || '';
const DRAWING_SHEET_NAME = process.env.SMARTSHEET_DRAWING_CONTROL_SHEET_NAME || process.env.SMARTSHEET_DRAWING_DOCUMENTATION_SHEET_NAME || 'DRAWING DOCUMENTATION CONTROL';
const DRAWING_CACHE_MS = Number(process.env.DRAWING_CONTROL_CACHE_MS || 10 * 60 * 1000);
const FETCH_TIMEOUT_MS = Number(process.env.SMARTSHEET_FETCH_TIMEOUT_MS || 15000);
const DRAWING_PAGE_SIZE = Math.max(100, Math.min(10000, Number(process.env.DRAWING_CONTROL_PAGE_SIZE || 5000)));

const drawingCache = global.__STEP_DRAWING_DETAIL_CACHE_V377__ || {
  sheetId: null,
  sheetName: null,
  metadataBySheet: {},
  detailBySheetAndTag: {},
};
global.__STEP_DRAWING_DETAIL_CACHE_V377__ = drawingCache;

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeName(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, ' ').trim();
}

function compact(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, '');
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function buildQuery(params = {}) {
  const entries = Object.entries(params).filter(([, value]) => value !== undefined && value !== null && String(value) !== '');
  if (!entries.length) return '';
  return `?${entries.map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`).join('&')}`;
}

async function apiFetch(path, options = {}) {
  if (!TOKEN) throw new Error('SMARTSHEET_TOKEN não configurado no Netlify.');
  const attempts = Math.max(1, Number(options.attempts || 2));
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), Math.max(3000, FETCH_TIMEOUT_MS));
    try {
      const response = await fetch(`${API_BASE}${path}`, {
        signal: controller.signal,
        headers: {
          Authorization: `Bearer ${TOKEN}`,
          'Content-Type': 'application/json',
        },
      });
      clearTimeout(timer);
      const bodyText = await response.text();
      if (!response.ok) {
        throw new Error(`Smartsheet ${response.status}: ${bodyText || 'resposta vazia'}`);
      }
      if (!String(bodyText || '').trim()) {
        throw new Error(`Smartsheet retornou resposta vazia em ${path}.`);
      }
      try {
        return JSON.parse(bodyText);
      } catch (_) {
        throw new Error(`Smartsheet retornou resposta não JSON em ${path}: ${String(bodyText).slice(0, 240)}`);
      }
    } catch (error) {
      clearTimeout(timer);
      lastError = error;
      const message = String(error?.message || error || '').toLowerCase();
      const transient = error?.name === 'AbortError' || message.includes('fetch failed') || message.includes('timeout') || message.includes('terminated') || message.includes('socket') || message.includes('econnreset');
      if (!transient || attempt >= attempts) throw error;
      await wait(400 * attempt);
    }
  }

  throw lastError || new Error('Falha ao consultar Smartsheet.');
}

function extractSmartsheetUrlToken(value = '') {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/\/sheets\/([^/?#]+)/i);
  return match ? match[1].trim() : '';
}

function isNumericSheetId(value = '') {
  return /^\d+$/.test(String(value || '').trim());
}

function getConfiguredDrawingSheetToken() {
  return extractSmartsheetUrlToken(DRAWING_SHEET_URL) || (!isNumericSheetId(DRAWING_SHEET_ID) ? String(DRAWING_SHEET_ID || '').trim() : '');
}

function sheetMatchesConfiguredToken(item, token) {
  if (!item || !token) return false;
  const needle = String(token || '').trim();
  return String(item.id || '') === needle
    || String(item.permalink || '').includes(needle)
    || String(item.workspace?.permalink || '').includes(needle);
}

async function resolveDrawingSheetId() {
  if (DRAWING_SHEET_ID && isNumericSheetId(DRAWING_SHEET_ID)) {
    drawingCache.sheetId = String(DRAWING_SHEET_ID).trim();
    return drawingCache.sheetId;
  }
  if (drawingCache.sheetId) return drawingCache.sheetId;

  const configuredToken = getConfiguredDrawingSheetToken();
  const target = normalizeName(DRAWING_SHEET_NAME);
  let page = 1;
  let selected = null;

  while (true) {
    const response = await apiFetch(`/sheets${buildQuery({ page, pageSize: 100 })}`);
    const items = response.data || [];

    selected = selected || (configuredToken ? items.find((item) => sheetMatchesConfiguredToken(item, configuredToken)) : null);
    selected = selected || items.find((item) => normalizeName(item.name) === target);
    selected = selected || items.find((item) => {
      const name = normalizeName(item.name);
      return name.includes(target) || target.includes(name) || name.includes('drawing documentation control') || name.includes('drawing control');
    });

    if (selected || !items.length || page >= (response.totalPages || 1)) break;
    page += 1;
  }

  if (!selected?.id) {
    throw new Error('Planilha Smartsheet de Drawing não encontrada. Use SMARTSHEET_DRAWING_CONTROL_SHEET_ID=2580648465590148 ou confira acesso do token.');
  }

  drawingCache.sheetId = String(selected.id);
  drawingCache.sheetName = selected.name || DRAWING_SHEET_NAME;
  return drawingCache.sheetId;
}

const BASE_COLUMNS = {
  client: ['Client'],
  projectNumber: ['Project Number'],
  taskName: ['Task Name (BSP, GASP, ICSP,SP)', 'Task Name', 'BSP'],
  vessel: ['UNIDADE', 'Unidade', 'Vessel'],
  pm: ['PM Responsible', 'PM'],
  po: ['Nº P.O.', 'N° P.O.', 'N P.O.', 'PO', 'P.O.'],
  title: ['Doc. Ref.Client / Title', 'Doc Ref Client Title', 'Drawing Title', 'Title', 'Description', 'Descrição'],
  currentStatus: ['Current Drawing Status', 'Drawing Status', 'Status'],
  drawingNumber: ['Drawing Number (Rev. A)', 'Drawing Number', 'Drawing Nº', 'Drawing No', 'TAG/ISO'],
  unit: ['UNIT', 'Unit', 'Unidade'],
  quantity: ['QUANTITY', 'Quantity', 'Qty', 'QTD', 'Peso', 'Weight', 'Kilos', 'KG'],
  drafterA: ['DWG Draftman (Rev. A)', 'DWG Draftsman (Rev. A)', 'Draftman (Rev. A)'],
  drafterHhA: ['DWG Draftman HH (Rev. A)', 'DWG Draftsman HH (Rev. A)', 'Draftman HH (Rev. A)'],
  reviewerA: ['Reviewer (Rev. A)'],
  reviewerHhA: ['Reviewer HH (Rev. A)'],
  approverA: ['Approver (Rev. A)'],
  approverHhA: ['Approver HH (Rev. A)'],
  startA: ['Drawings start date (Rev. A)', 'Drawing start date (Rev. A)', 'Drawings Start Date (Rev. A)'],
  reviewerApprovalA: ['Reviewer approval (Rev. A)', 'Reviewer Approval (Rev. A)'],
  approverApprovalA: ['Approver approval (Rev. A)', 'Approver Approval (Rev. A)'],
  internallySentA: ['Internally approved & sent to PM (Rev. A)', 'Internally Approved & Sent to PM (Rev. A)', 'Internally approved and sent to PM (Rev. A)'],
  pmApprovalA: ['PM approval (Rev. A)', 'Pm approval (Rev. A)', 'PM Approval (Rev. A)'],
  approvalDate: ['Approval Date', 'Approved Date'],
};

function getRevisionColumnCandidates(rev, field) {
  const r = String(rev || '').trim().toUpperCase();
  const map = {
    drafter: [`Draftman update (Rev. ${r})`, `Draftman Update (Rev. ${r})`, `Draftsman update (Rev. ${r})`, `DWG Draftman Update (Rev. ${r})`, `DWG Draftman update (Rev. ${r})`],
    drafterHh: [`Draftman Update HH (Rev. ${r})`, `Draftman update HH (Rev. ${r})`, `Draftsman Update HH (Rev. ${r})`, `DWG Draftman Update HH (Rev. ${r})`],
    reviewer: [`Reviewer Update (Rev. ${r})`, `Reviewer update (Rev. ${r})`],
    reviewerHh: [`Reviewer Update HH (Rev. ${r})`, `Reviewer update HH (Rev. ${r})`],
    approver: [`Approver Update (Rev. ${r})`, `Approver update (Rev. ${r})`],
    approverHh: [`Approver Update HH (Rev. ${r})`, `Approver update HH (Rev. ${r})`],
    clientComments: [
      `Last client comments received via PM (Rev. ${r})`,
      `Last client comments recieved via PM (Rev. ${r})`,
      `Last client comment received via PM (Rev. ${r})`,
      `Last comments received via PM (Rev. ${r})`,
    ],
    originReview: [`DWG Origin review (Rev. ${r})`, `DWG Origin Review (Rev. ${r})`, `Origin review (Rev. ${r})`],
    updateStart: [`Last revision Update start (Rev. ${r})`, `Last revision update start (Rev. ${r})`, `Revision Update start (Rev. ${r})`],
    reviewerApproval: [`Reviewer Update approval (Rev. ${r})`, `Reviewer update approval (Rev. ${r})`, `Reviewer approval (Rev. ${r})`],
    approverApproval: [`Approver Update approval (Rev. ${r})`, `Approver update approval (Rev. ${r})`, `Approver approval (Rev. ${r})`],
    internallySent: [`Internally approved & sent to PM (Rev. ${r})`, `Internally Approved & Sent to PM (Rev. ${r})`, `Internally approved and sent to PM (Rev. ${r})`],
    pmApproval: [`PM approval (Rev. ${r})`, `Pm approval (Rev. ${r})`, `PM Approval (Rev. ${r})`],
  };
  return map[field] || [];
}

function allCandidateColumnNames() {
  const names = [];
  for (const list of Object.values(BASE_COLUMNS)) names.push(...list);
  for (const rev of ['B', 'C', 'D', 'E', 'F']) {
    for (const field of ['drafter', 'drafterHh', 'reviewer', 'reviewerHh', 'approver', 'approverHh', 'clientComments', 'originReview', 'updateStart', 'reviewerApproval', 'approverApproval', 'internallySent', 'pmApproval']) {
      names.push(...getRevisionColumnCandidates(rev, field));
    }
  }
  return names;
}

function makeColumnLookup(columns = []) {
  const byId = new Map();
  const byNormalized = new Map();
  const byCompact = new Map();
  for (const column of columns || []) {
    if (!column?.id || !column?.title) continue;
    const id = String(column.id);
    const title = String(column.title);
    byId.set(id, title);
    const norm = normalizeName(title);
    const comp = compact(title);
    if (norm && !byNormalized.has(norm)) byNormalized.set(norm, column);
    if (comp && !byCompact.has(comp)) byCompact.set(comp, column);
  }
  return { byId, byNormalized, byCompact, columns };
}

function findColumn(columnsLookup, candidates = []) {
  for (const candidate of candidates) {
    const norm = normalizeName(candidate);
    const comp = compact(candidate);
    if (norm && columnsLookup.byNormalized.has(norm)) return columnsLookup.byNormalized.get(norm);
    if (comp && columnsLookup.byCompact.has(comp)) return columnsLookup.byCompact.get(comp);
  }
  return null;
}

function getColumnTitleByCandidates(columnsLookup, candidates = []) {
  const column = findColumn(columnsLookup, candidates);
  return column?.title || candidates[0] || '';
}

function buildSelectedColumnIds(columnsLookup) {
  const ids = [];
  const seen = new Set();
  for (const candidate of allCandidateColumnNames()) {
    const column = findColumn(columnsLookup, [candidate]);
    if (column?.id && !seen.has(String(column.id))) {
      seen.add(String(column.id));
      ids.push(String(column.id));
    }
  }
  return ids;
}

async function fetchSheetMetadata(sheetId, force = false) {
  const key = String(sheetId || '').trim();
  const cached = drawingCache.metadataBySheet[key];
  if (!force && cached?.columns?.length && cached.loadedAt && (Date.now() - cached.loadedAt) < DRAWING_CACHE_MS) return cached;

  // pageSize=1 traz o cabeçalho/colunas sem trazer a planilha inteira.
  const sheet = await apiFetch(`/sheets/${encodeURIComponent(key)}${buildQuery({ pageSize: 1, page: 1 })}`);
  const metadata = {
    sheetId: String(sheet.id || key),
    sheetName: sheet.name || DRAWING_SHEET_NAME,
    columns: sheet.columns || [],
    loadedAt: Date.now(),
  };
  drawingCache.metadataBySheet[key] = metadata;
  drawingCache.metadataBySheet[metadata.sheetId] = metadata;
  drawingCache.sheetName = metadata.sheetName;
  return metadata;
}

function mapRowValues(row, columnMap) {
  const values = {};
  for (const cell of row.cells || []) {
    const title = columnMap.get(String(cell.columnId));
    if (!title) continue;
    values[title] = { raw: cell.value ?? null, display: cell.displayValue ?? null };
  }
  return { id: row.id, rowNumber: row.rowNumber, values };
}

function getCellValue(row, key) {
  if (!row?.values) return '';
  const exact = row.values[key];
  if (exact) return exact.display ?? exact.raw ?? '';
  const target = normalizeName(key);
  const targetCompact = compact(key);
  for (const [title, cell] of Object.entries(row.values)) {
    if (normalizeName(title) === target || compact(title) === targetCompact) return cell?.display ?? cell?.raw ?? '';
  }
  return '';
}

function getFirstValue(row, keys = []) {
  for (const key of keys) {
    const value = getCellValue(row, key);
    if (value != null && String(value).trim()) return value;
  }
  return '';
}

function getValueByResolvedColumn(row, columnsLookup, candidates = []) {
  const title = getColumnTitleByCandidates(columnsLookup, candidates);
  return getFirstValue(row, [title, ...candidates]);
}

function excelSerialToDate(serial) {
  if (!Number.isFinite(serial) || serial < 1 || serial > 90000) return null;
  const excelEpoch = Date.UTC(1899, 11, 30);
  const date = new Date(excelEpoch + Math.round(serial) * 86400000);
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function formatDateValue(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'number') {
    const date = excelSerialToDate(value);
    return date ? date.toLocaleDateString('pt-BR', { timeZone: 'UTC' }) : '';
  }
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const date = excelSerialToDate(Number(raw));
    if (date) return date.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
  }
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
  const br = raw.match(/^(\d{2})\/(\d{2})\/(\d{2,4})$/);
  if (br) return raw;
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
  return raw;
}

function parseNumberValue(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  let str = String(value).trim().replace(/\s/g, '');
  if (!str) return null;
  if (str.includes(',') && str.includes('.')) {
    str = str.lastIndexOf(',') > str.lastIndexOf('.') ? str.replace(/\./g, '').replace(',', '.') : str.replace(/,/g, '');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  str = str.replace(/[^\d.-]/g, '');
  const num = Number(str);
  return Number.isFinite(num) ? num : null;
}

function rowHasAny(row, columnsLookup, candidates) {
  return candidates.some((key) => {
    const value = getValueByResolvedColumn(row, columnsLookup, [key]);
    return value != null && String(value).trim();
  });
}

function getRevisionValue(row, columnsLookup, rev, field) {
  return getValueByResolvedColumn(row, columnsLookup, getRevisionColumnCandidates(rev, field));
}

function getRevision(row, columnsLookup, rev) {
  const fields = ['drafter', 'drafterHh', 'reviewer', 'reviewerHh', 'approver', 'approverHh', 'clientComments', 'originReview', 'updateStart', 'reviewerApproval', 'approverApproval', 'internallySent', 'pmApproval'];
  const keys = fields.flatMap((field) => getRevisionColumnCandidates(rev, field));
  if (!rowHasAny(row, columnsLookup, keys)) return null;
  return {
    revision: rev,
    drafter: String(getRevisionValue(row, columnsLookup, rev, 'drafter') || '').trim(),
    drafterHh: String(getRevisionValue(row, columnsLookup, rev, 'drafterHh') || '').trim(),
    reviewer: String(getRevisionValue(row, columnsLookup, rev, 'reviewer') || '').trim(),
    reviewerHh: String(getRevisionValue(row, columnsLookup, rev, 'reviewerHh') || '').trim(),
    approver: String(getRevisionValue(row, columnsLookup, rev, 'approver') || '').trim(),
    approverHh: String(getRevisionValue(row, columnsLookup, rev, 'approverHh') || '').trim(),
    clientCommentsDate: formatDateValue(getRevisionValue(row, columnsLookup, rev, 'clientComments')),
    originReview: String(getRevisionValue(row, columnsLookup, rev, 'originReview') || '').trim(),
    updateStartDate: formatDateValue(getRevisionValue(row, columnsLookup, rev, 'updateStart')),
    reviewerApproval: String(getRevisionValue(row, columnsLookup, rev, 'reviewerApproval') || '').trim(),
    approverApproval: String(getRevisionValue(row, columnsLookup, rev, 'approverApproval') || '').trim(),
    internallySentToPmDate: formatDateValue(getRevisionValue(row, columnsLookup, rev, 'internallySent')),
    pmApproval: String(getRevisionValue(row, columnsLookup, rev, 'pmApproval') || '').trim(),
  };
}

function getProjectNumberFromDrawing(drawingNumber) {
  const match = String(drawingNumber || '').match(/BSP\s*-?\s*(\d{2}\s*-\s*\d+)/i) || String(drawingNumber || '').match(/(\d{2}\s*-\s*\d+)/);
  return match ? match[1].replace(/\s+/g, '') : '';
}

function getBspKeyFromDrawing(drawingNumber) {
  const number = getProjectNumberFromDrawing(drawingNumber);
  return number ? `BSP ${number}` : '';
}

function normalizeDrawingDetail(row, columnsLookup) {
  const drawingNumber = String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.drawingNumber) || '').trim();
  if (!drawingNumber) return null;

  const approvalDate = formatDateValue(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.approvalDate));
  const reviewerApprovalA = String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.reviewerApprovalA) || '').trim();
  const approverApprovalA = String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.approverApprovalA) || '').trim();
  const startA = formatDateValue(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.startA));

  return {
    rowId: String(row.id || ''),
    rowNumber: row.rowNumber || null,
    client: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.client) || '').trim(),
    projectNumberSource: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.projectNumber) || '').trim(),
    taskName: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.taskName) || '').trim(),
    vessel: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.vessel) || '').trim(),
    pmResponsible: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.pm) || '').trim(),
    poNumber: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.po) || '').trim(),
    drawingNumber,
    drawingKey: compact(drawingNumber),
    projectNumber: getProjectNumberFromDrawing(drawingNumber),
    bspKey: getBspKeyFromDrawing(drawingNumber),
    title: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.title) || '').trim(),
    currentDrawingStatus: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.currentStatus) || '').trim(),
    unit: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.unit) || '').trim(),
    quantity: parseNumberValue(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.quantity)),
    drawingStartDate: startA,
    drawingFinishDate: approvalDate,
    approvalDate,
    revA: {
      drafter: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.drafterA) || '').trim(),
      drafterHh: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.drafterHhA) || '').trim(),
      reviewer: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.reviewerA) || '').trim(),
      reviewerHh: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.reviewerHhA) || '').trim(),
      approver: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.approverA) || '').trim(),
      approverHh: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.approverHhA) || '').trim(),
      startDate: startA,
      reviewerApproval: reviewerApprovalA,
      approverApproval: approverApprovalA,
      internallySentToPmDate: formatDateValue(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.internallySentA)),
      pmApproval: String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.pmApprovalA) || '').trim(),
      approvalDate,
      approved: normalizeText(reviewerApprovalA) === 'approved' && normalizeText(approverApprovalA) === 'approved' && Boolean(approvalDate),
    },
    revisions: ['B', 'C', 'D', 'E', 'F'].map((rev) => getRevision(row, columnsLookup, rev)).filter(Boolean),
  };
}

function isDrawingMatch(found, wanted) {
  const f = compact(found);
  const w = compact(wanted);
  if (!f || !w) return false;
  if (f === w) return true;
  const fLoose = f.replace(/^bsp/, '');
  const wLoose = w.replace(/^bsp/, '');
  return fLoose === wLoose || f.includes(w) || w.includes(f);
}

async function fetchDrawingDetailFromSheet(sheetId, drawingNumber, force = false) {
  const key = String(sheetId || '').trim();
  if (!key) throw new Error('ID da planilha de Drawing vazio.');
  const wantedKey = compact(drawingNumber);
  const cacheKey = `${key}:${wantedKey}`;
  const cached = drawingCache.detailBySheetAndTag[cacheKey];
  if (!force && cached?.detail && cached.loadedAt && (Date.now() - cached.loadedAt) < DRAWING_CACHE_MS) return cached;

  const metadata = await fetchSheetMetadata(key, force);
  const columnsLookup = makeColumnLookup(metadata.columns);
  const drawingColumn = findColumn(columnsLookup, BASE_COLUMNS.drawingNumber);
  if (!drawingColumn?.id) {
    const available = metadata.columns.map((column) => column.title).slice(0, 120);
    throw new Error(`Coluna "Drawing Number (Rev. A)" não encontrada na planilha ${metadata.sheetName}. Colunas encontradas: ${available.join(' | ')}`);
  }

  const selectedIds = buildSelectedColumnIds(columnsLookup);
  if (!selectedIds.includes(String(drawingColumn.id))) selectedIds.unshift(String(drawingColumn.id));

  const columnMap = columnsLookup.byId;
  let page = 1;
  let scannedRows = 0;
  let closest = [];
  let detail = null;

  while (true) {
    const query = buildQuery({
      page,
      pageSize: DRAWING_PAGE_SIZE,
      columnIds: selectedIds.join(','),
      include: 'objectValue',
    });
    const sheet = await apiFetch(`/sheets/${encodeURIComponent(metadata.sheetId)}${query}`);
    const rows = sheet.rows || [];
    scannedRows += rows.length;

    for (const apiRow of rows) {
      const row = mapRowValues(apiRow, columnMap);
      const rowDrawingNumber = String(getValueByResolvedColumn(row, columnsLookup, BASE_COLUMNS.drawingNumber) || '').trim();
      if (!rowDrawingNumber) continue;

      const candidateKey = compact(rowDrawingNumber);
      if (candidateKey.includes(wantedKey.slice(0, Math.min(10, wantedKey.length))) || wantedKey.includes(candidateKey.slice(0, Math.min(10, candidateKey.length)))) {
        if (closest.length < 12) closest.push({ drawingNumber: rowDrawingNumber, rowNumber: row.rowNumber });
      }

      if (isDrawingMatch(rowDrawingNumber, drawingNumber)) {
        detail = normalizeDrawingDetail(row, columnsLookup);
        break;
      }
    }

    if (detail) break;
    const totalPages = Number(sheet.totalPages || 1);
    if (!rows.length || page >= totalPages) break;
    page += 1;
  }

  const result = {
    detail,
    sheetId: metadata.sheetId,
    sheetName: metadata.sheetName,
    scannedRows,
    selectedColumns: selectedIds.length,
    closest,
    loadedAt: Date.now(),
  };
  if (detail) drawingCache.detailBySheetAndTag[cacheKey] = result;
  return result;
}

async function listCandidateDrawingSheets() {
  const sheets = [];
  let page = 1;
  const target = normalizeName(DRAWING_SHEET_NAME);
  const token = getConfiguredDrawingSheetToken();
  while (true) {
    const response = await apiFetch(`/sheets${buildQuery({ page, pageSize: 100 })}`);
    const items = response.data || [];
    for (const item of items) {
      const name = normalizeName(item.name || '');
      let score = 0;
      if (token && sheetMatchesConfiguredToken(item, token)) score += 1000;
      if (name === target) score += 500;
      if (target && (name.includes(target) || target.includes(name))) score += 300;
      if (name.includes('drawing documentation control')) score += 260;
      if (name.includes('drawing') && name.includes('documentation')) score += 220;
      if (name.includes('drawing') && name.includes('control')) score += 180;
      if (score > 0 && item.id) sheets.push({ ...item, score });
    }
    if (!items.length || page >= (response.totalPages || 1)) break;
    page += 1;
  }
  const seen = new Set();
  return sheets
    .sort((a, b) => b.score - a.score || String(a.name || '').localeCompare(String(b.name || ''), 'pt-BR'))
    .filter((item) => {
      const id = String(item.id || '');
      if (!id || seen.has(id)) return false;
      seen.add(id);
      return true;
    })
    .slice(0, 8);
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return jsonResponse(204, { ok: true });
  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  try {
    const params = event.queryStringParameters || {};
    const drawingNumber = params.drawingNumber || params.tag || params.iso || '';
    if (!drawingNumber) return jsonResponse(400, { ok: false, error: 'Informe drawingNumber, tag ou iso.' });
    const force = params.force === '1' || params.refresh === '1';

    const scannedSheets = [];
    let primaryError = '';
    let result = null;

    try {
      const primarySheetId = await resolveDrawingSheetId();
      result = await fetchDrawingDetailFromSheet(primarySheetId, drawingNumber, force);
      scannedSheets.push({
        sheetId: result.sheetId,
        sheetName: result.sheetName,
        scannedRows: result.scannedRows,
        selectedColumns: result.selectedColumns,
        found: Boolean(result.detail),
      });
    } catch (error) {
      primaryError = error?.message || String(error);
    }

    if (!result?.detail && !DRAWING_SHEET_ID) {
      const candidates = await listCandidateDrawingSheets();
      for (const candidate of candidates) {
        if (result?.sheetId && String(candidate.id) === String(result.sheetId)) continue;
        try {
          const candidateResult = await fetchDrawingDetailFromSheet(String(candidate.id), drawingNumber, force);
          scannedSheets.push({
            sheetId: candidateResult.sheetId,
            sheetName: candidateResult.sheetName,
            scannedRows: candidateResult.scannedRows,
            selectedColumns: candidateResult.selectedColumns,
            found: Boolean(candidateResult.detail),
          });
          if (candidateResult.detail) {
            result = candidateResult;
            drawingCache.sheetId = candidateResult.sheetId;
            drawingCache.sheetName = candidateResult.sheetName;
            break;
          }
        } catch (error) {
          scannedSheets.push({ sheetId: String(candidate.id), sheetName: candidate.name || '', error: error?.message || String(error) });
        }
      }
    }

    if (!result?.detail) {
      return jsonResponse(404, {
        ok: false,
        error: 'Detalhamento de Drawing não encontrado para esta TAG/ISO.',
        drawingNumber,
        sheetId: result?.sheetId || drawingCache.sheetId || DRAWING_SHEET_ID || '',
        sheetName: result?.sheetName || drawingCache.sheetName || DRAWING_SHEET_NAME,
        scannedRows: result?.scannedRows || 0,
        selectedColumns: result?.selectedColumns || 0,
        closest: result?.closest || [],
        initialLoadError: primaryError,
        scannedSheets,
        hint: 'Confirme SMARTSHEET_DRAWING_CONTROL_SHEET_ID=2580648465590148 e se a TAG está na coluna Drawing Number (Rev. A).',
      });
    }

    drawingCache.sheetId = result.sheetId;
    drawingCache.sheetName = result.sheetName;

    return jsonResponse(200, {
      ok: true,
      detail: result.detail,
      meta: {
        sheetId: result.sheetId,
        sheetName: result.sheetName,
        scannedRows: result.scannedRows,
        selectedColumns: result.selectedColumns,
        loadedAt: new Date(result.loadedAt || Date.now()).toISOString(),
        cacheMs: DRAWING_CACHE_MS,
        scannedSheets,
        mode: 'targeted-column-scan',
      },
    });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error?.message || 'Falha ao consultar Drawing Documentation Control.' });
  }
};
