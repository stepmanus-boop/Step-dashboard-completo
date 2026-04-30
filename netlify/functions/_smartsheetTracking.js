const API_BASE = process.env.SMARTSHEET_API_BASE || 'https://api.smartsheet.com/2.0';
const SHEET_NAME = process.env.SMARTSHEET_SHEET_NAME || 'Progress Tracking Sheet - Piping Fabrication';
const SHEET_ID_ENV = process.env.SMARTSHEET_SHEET_ID || '';
const TOKEN = process.env.SMARTSHEET_TOKEN || process.env.SMARTSHEET_ACCESS_TOKEN || process.env.SMARTSHEET_API_TOKEN || process.env.SMARTSHEET_BEARER_TOKEN || process.env.SMARTSHEET_PAT || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN || '';

const TRACKING_PROGRESS_BY_SECTOR = {
  pintura: 'Surface preparation and/or coating',
  solda: 'Full welding execution',
  producao: 'Spool Assemble and tack weld',
  calderaria: 'Spool Assemble and tack weld',
  pre_montagem: 'Spool Assemble and tack weld',
  pendente_envio: 'Package and Delivered',
  logistica: 'Package and Delivered',
};

const TRACKING_DATE_BY_PROGRESS_COLUMN = {
  'Surface preparation and/or coating': 'Coating Finish Date',
  'Full welding execution': 'Welding Finish Date',
  'Hydro Test Pressure (QC)': 'TH Finish Date',
  'Final Dimensional Inpection/3D (QC)': 'Inspection Finish Date (QC)',
  'Non Destructive Examination (QC)': 'Inspection Finish Date (QC)',
  'Spool Assemble and tack weld': 'Boilermaker Finish Date',
  'HDG / FBE.  (PAINT)': 'HDG / FBE DATE RETORNO (PAINT)',
  'Final Inspection': 'Project Finish Date',
  'Package and Delivered': 'Project Finish Date',
};

const COLUMN_ALIASES = {
  'Package and Delivered': ['Package and Delivered', 'Package Delivered', 'Package & Delivered', 'Package'],
  'Final Dimensional Inpection/3D (QC)': ['Final Dimensional Inpection/3D (QC)', 'Final Dimensional Inspection/3D (QC)', 'Final Dimensional Inspection 3D QC'],
  'Non Destructive Examination (QC)': ['Non Destructive Examination (QC)', 'NDE', 'END'],
  'Hydro Test Pressure (QC)': ['Hydro Test Pressure (QC)', 'Hydro Test', 'TH'],
  'Surface preparation and/or coating': ['Surface preparation and/or coating', 'Surface Preparation Coating', 'Coating', 'Pintura'],
  'Spool Assemble and tack weld': ['Spool Assemble and tack weld', 'Spool Assemble', 'Tack weld', 'Pré Montagem', 'Pre Montagem'],
  'Full welding execution': ['Full welding execution', 'Welding Execution', 'Solda'],
  'Coating Finish Date': ['Coating Finish Date', 'Coating Finish', 'Data Pintura'],
  'Welding Finish Date': ['Welding Finish Date', 'Welding Finish'],
  'TH Finish Date': ['TH Finish Date', 'Hydro Test Finish Date'],
  'Inspection Finish Date (QC)': ['Inspection Finish Date (QC)', 'Inspection Finish Date QC'],
  'Boilermaker Finish Date': ['Boilermaker Finish Date', 'Boilermaker Finish'],
  'HDG / FBE DATE RETORNO (PAINT)': ['HDG / FBE DATE RETORNO (PAINT)', 'HDG FBE DATE RETORNO PAINT', 'HDG FBE RETORNO'],
  'Final Inspection': ['Final Inspection'],
  'Project Finish Date': ['Project Finish Date', 'Finish Date Project'],
  'Drawing': ['Drawing', 'Spool', 'ISO'],
  'Project': ['Project', 'BSP'],
};

function normalizeColumnTitle(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '')
    .toLowerCase();
}

function normalizeSectorValue(value) {
  const normalized = String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (['qualidade', 'inspecao', 'inspecao_qualidade', 'qc'].includes(normalized)) return 'inspecao';
  if (['logistica', 'pendente_envio', 'envio', 'preparando_envio', 'em_tratativa'].includes(normalized)) return 'pendente_envio';
  if (['producao', 'caldeiraria', 'calderaria'].includes(normalized)) return normalized === 'caldeiraria' ? 'calderaria' : normalized;
  return normalized;
}

function normalizeSpoolIdentity(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[–—−]/g, '-')
    .toUpperCase()
    .replace(/\([^)]*\)/g, '')
    .replace(/[^A-Z0-9]+/g, '');
}

function normalizeSpoolIdentityLoose(value) {
  const prepared = String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[–—−]/g, '-')
    .toUpperCase()
    .replace(/\([^)]*\)/g, ' ');
  const tokens = prepared.match(/[A-Z]+|\d+/g) || [];
  return tokens.map((token) => /^\d+$/.test(token) ? String(Number(token)) : token).join('');
}

function compactContainsMatch(a, b) {
  const variantsA = [normalizeSpoolIdentity(a), normalizeSpoolIdentityLoose(a)].filter(Boolean);
  const variantsB = [normalizeSpoolIdentity(b), normalizeSpoolIdentityLoose(b)].filter(Boolean);
  if (!variantsA.length || !variantsB.length) return false;

  for (const left of variantsA) {
    for (const right of variantsB) {
      if (!left || !right) continue;
      if (left === right) return true;
      const min = Math.min(left.length, right.length);
      if (min >= 8 && (left.includes(right) || right.includes(left))) return true;
    }
  }
  return false;
}

function parseNumberValue(input) {
  if (input == null || input === '') return null;
  if (typeof input === 'number') return Number.isFinite(input) ? input : null;
  let str = String(input).trim();
  if (!str) return null;
  str = str.replace(/\s/g, '');
  const hasComma = str.includes(',');
  const hasDot = str.includes('.');
  if (hasComma && hasDot) {
    str = str.lastIndexOf(',') > str.lastIndexOf('.') ? str.replace(/\./g, '').replace(',', '.') : str.replace(/,/g, '');
  } else if (hasComma) {
    str = str.replace(',', '.');
  }
  str = str.replace(/[^\d.-]/g, '');
  const num = Number(str);
  return Number.isFinite(num) ? num : null;
}

function parsePercentValue(rawValue, displayValue) {
  if (displayValue != null && String(displayValue).includes('%')) {
    const value = parseNumberValue(String(displayValue).replace('%', ''));
    return value == null ? null : value;
  }
  const parsed = parseNumberValue(rawValue ?? displayValue);
  if (parsed == null) return null;
  return parsed >= 0 && parsed <= 1 ? parsed * 100 : parsed;
}

function isEmptyCellValue(value) {
  if (value == null) return true;
  return String(value).trim() === '';
}

function getCell(row, column) {
  if (!row || !column) return null;
  return (row.cells || []).find((cell) => String(cell.columnId) === String(column.id)) || null;
}

function getCellDisplay(row, column) {
  const cell = getCell(row, column);
  if (!cell) return '';
  return cell.displayValue ?? cell.value ?? '';
}

function getCellPercent(row, column) {
  const cell = getCell(row, column);
  if (!cell) return null;
  return parsePercentValue(cell.value, cell.displayValue);
}

function getCellHasValue(row, column) {
  const cell = getCell(row, column);
  if (!cell) return false;
  return !isEmptyCellValue(cell.value ?? cell.displayValue);
}

function buildColumnIndex(sheet) {
  const columns = Array.isArray(sheet?.columns) ? sheet.columns : [];
  const byNorm = new Map();
  for (const column of columns) {
    const norm = normalizeColumnTitle(column.title);
    if (!byNorm.has(norm)) byNorm.set(norm, column);
  }
  return { columns, byNorm };
}

function findColumn(sheet, canonicalName, extraAliases = []) {
  const { columns, byNorm } = buildColumnIndex(sheet);
  const aliases = [canonicalName, ...(COLUMN_ALIASES[canonicalName] || []), ...(Array.isArray(extraAliases) ? extraAliases : [])]
    .map((item) => String(item || '').trim())
    .filter(Boolean);

  for (const alias of aliases) {
    const exact = byNorm.get(normalizeColumnTitle(alias));
    if (exact) return exact;
  }

  const canonicalNorm = normalizeColumnTitle(canonicalName);
  if (canonicalName === 'Package and Delivered') {
    const safe = columns.find((column) => {
      const norm = normalizeColumnTitle(column.title);
      return norm.includes('package') && (norm.includes('delivered') || norm.includes('delivery') || norm.includes('envio'));
    });
    if (safe) return safe;
    return byNorm.get('package') || null;
  }

  return columns.find((column) => {
    const norm = normalizeColumnTitle(column.title);
    return norm && canonicalNorm && norm.includes(canonicalNorm);
  }) || null;
}

async function apiFetch(path, options = {}) {
  if (!TOKEN) throw new Error('SMARTSHEET_TOKEN não configurado no Netlify.');
  const response = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    const message = await response.text().catch(() => '');
    throw new Error(`Smartsheet ${response.status}: ${message}`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

function normalizeName(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .trim()
    .toLowerCase();
}

const cache = global.__STEP_TRACKING_WRITE_CACHE__ || { sheetId: null, sheetName: null };
global.__STEP_TRACKING_WRITE_CACHE__ = cache;

async function resolveSheetId() {
  if (cache.sheetId) return cache.sheetId;
  if (SHEET_ID_ENV) {
    cache.sheetId = String(SHEET_ID_ENV);
    return cache.sheetId;
  }
  const target = normalizeName(SHEET_NAME);
  let page = 1;
  let fuzzyFound = null;
  while (true) {
    const response = await apiFetch(`/sheets?page=${page}&pageSize=100`);
    const items = response?.data || [];
    const exact = items.find((item) => normalizeName(item.name) === target);
    if (exact) {
      cache.sheetId = String(exact.id);
      cache.sheetName = exact.name;
      return cache.sheetId;
    }
    if (!fuzzyFound) {
      fuzzyFound = items.find((item) => normalizeName(item.name).includes(target) || target.includes(normalizeName(item.name)));
    }
    if (!items.length || page >= (response.totalPages || 1)) break;
    page += 1;
  }
  if (fuzzyFound) {
    cache.sheetId = String(fuzzyFound.id);
    cache.sheetName = fuzzyFound.name;
    return cache.sheetId;
  }
  throw new Error(`Sheet "${SHEET_NAME}" não encontrada.`);
}

async function fetchTrackingSheet() {
  const sheetId = await resolveSheetId();
  const sheet = await apiFetch(`/sheets/${sheetId}?includeAll=true`);
  return { sheetId, sheet };
}

function invalidateProjectCache() {
  if (global.__STEP_PROGRESS_CACHE__) {
    global.__STEP_PROGRESS_CACHE__.version = null;
    global.__STEP_PROGRESS_CACHE__.payload = null;
    global.__STEP_PROGRESS_CACHE__.lastSync = null;
  }
}

async function updateRows(sheetId, rowChangesMap) {
  const rows = Array.from(rowChangesMap.values())
    .map((entry) => ({ id: entry.id, cells: Array.from(entry.cells.values()) }))
    .filter((row) => row.id && row.cells.length);
  if (!rows.length) return { rows: [], responses: [] };

  const responses = [];
  const chunkSize = 400;
  for (let index = 0; index < rows.length; index += chunkSize) {
    const chunk = rows.slice(index, index + chunkSize);
    const response = await apiFetch('/sheets/' + sheetId + '/rows', {
      method: 'PUT',
      body: JSON.stringify(chunk),
    });
    responses.push(response);
  }
  invalidateProjectCache();
  return { rows, responses };
}

function addCellChange(rowChangesMap, rowId, column, value) {
  if (!rowId || !column?.id) return;
  const key = String(rowId);
  if (!rowChangesMap.has(key)) rowChangesMap.set(key, { id: rowId, cells: new Map() });
  const entry = rowChangesMap.get(key);
  const cellKey = String(column.id);
  const existing = entry.cells.get(cellKey);
  if (typeof value === 'number' && existing && typeof existing.value === 'number') {
    entry.cells.set(cellKey, { columnId: column.id, value: Math.max(existing.value, value), strict: false });
    return;
  }
  if (!existing || typeof value === 'number') {
    entry.cells.set(cellKey, { columnId: column.id, value, strict: false });
  }
}

function getDescendantRows(sheet, projectRowId) {
  const id = String(projectRowId || '').trim();
  if (!id) return [];
  const rows = Array.isArray(sheet?.rows) ? sheet.rows : [];
  const rowsById = new Map(rows.map((row) => [String(row.id), row]));
  return rows.filter((row) => {
    let parentId = row.parentId == null ? '' : String(row.parentId);
    const seen = new Set();
    while (parentId) {
      if (parentId === id) return true;
      if (seen.has(parentId)) return false;
      seen.add(parentId);
      const parent = rowsById.get(parentId);
      parentId = parent?.parentId == null ? '' : String(parent.parentId);
    }
    return false;
  });
}

function getMatchingTrackingRows(sheet, update) {
  const drawingColumn = findColumn(sheet, 'Drawing');
  const projectRows = getDescendantRows(sheet, update?.projectRowId);
  const sourceRows = projectRows.length ? projectRows : (Array.isArray(sheet?.rows) ? sheet.rows : []);
  const targets = [update?.spoolIso, update?.spoolDescription, update?.drawing]
    .map((item) => String(item || '').trim())
    .filter(Boolean);
  const matches = sourceRows.filter((row) => {
    const drawing = getCellDisplay(row, drawingColumn);
    if (!drawing) return false;
    return targets.some((target) => compactContainsMatch(drawing, target));
  });

  if (matches.length || projectRows.length) return matches;

  // Fallback controlado: só usa o Tracking inteiro quando o projeto não foi encontrado pelo parentId.
  return (Array.isArray(sheet?.rows) ? sheet.rows : []).filter((row) => {
    const drawing = getCellDisplay(row, drawingColumn);
    if (!drawing) return false;
    return targets.some((target) => compactContainsMatch(drawing, target));
  });
}

function resolveInspectionProgressColumn(sheet, row) {
  const columns = {
    finalInspection: findColumn(sheet, 'Final Inspection'),
    hydro: findColumn(sheet, 'Hydro Test Pressure (QC)'),
    nde: findColumn(sheet, 'Non Destructive Examination (QC)'),
    finalDimensional: findColumn(sheet, 'Final Dimensional Inpection/3D (QC)'),
    welding: findColumn(sheet, 'Full welding execution'),
    coating: findColumn(sheet, 'Surface preparation and/or coating'),
  };
  const welding = getCellPercent(row, columns.welding) ?? 0;
  const finalDimensional = getCellPercent(row, columns.finalDimensional) ?? 0;
  const nde = getCellPercent(row, columns.nde);
  const hydro = getCellPercent(row, columns.hydro) ?? 0;
  const coating = getCellPercent(row, columns.coating) ?? 0;
  const finalInspection = getCellPercent(row, columns.finalInspection) ?? 0;

  const active = [columns.finalDimensional, columns.nde, columns.hydro, columns.finalInspection]
    .filter(Boolean)
    .find((column) => {
      const value = getCellPercent(row, column);
      return value != null && value > 0 && value < 100;
    });
  if (active) return active;

  if (coating >= 100 && finalInspection < 100 && columns.finalInspection) return columns.finalInspection;
  if ((nde != null && nde >= 100 || finalDimensional >= 100) && hydro < 100 && columns.hydro) return columns.hydro;
  if (finalDimensional >= 100 && nde != null && nde < 100 && columns.nde) return columns.nde;
  if (welding >= 100 && finalDimensional < 100 && columns.finalDimensional) return columns.finalDimensional;
  return columns.finalDimensional || columns.nde || columns.hydro || columns.finalInspection || null;
}

function getCanonicalProgressColumnName(columnTitle) {
  const normalized = normalizeColumnTitle(columnTitle);
  if (!normalized) return '';
  const candidates = new Set([
    ...Object.keys(TRACKING_DATE_BY_PROGRESS_COLUMN),
    ...Object.values(TRACKING_PROGRESS_BY_SECTOR),
  ]);
  for (const candidate of candidates) {
    const aliases = [candidate, ...(COLUMN_ALIASES[candidate] || [])];
    if (aliases.some((alias) => normalizeColumnTitle(alias) === normalized)) return candidate;
  }
  return columnTitle || '';
}

function resolveProgressColumn(sheet, update, referenceRow) {
  const explicit = String(update?.trackingColumn || update?.progressColumn || '').trim();
  if (explicit) return findColumn(sheet, explicit);
  const sector = normalizeSectorValue(update?.sector);
  if (sector === 'inspecao') return resolveInspectionProgressColumn(sheet, referenceRow);
  const canonical = TRACKING_PROGRESS_BY_SECTOR[sector];
  return canonical ? findColumn(sheet, canonical) : null;
}

function getDateColumnNameForProgressColumn(progressColumnTitle) {
  const canonical = getCanonicalProgressColumnName(progressColumnTitle);
  const normalized = normalizeColumnTitle(canonical || progressColumnTitle);
  const key = Object.keys(TRACKING_DATE_BY_PROGRESS_COLUMN).find((candidate) => normalizeColumnTitle(candidate) === normalized);
  return key ? TRACKING_DATE_BY_PROGRESS_COLUMN[key] : '';
}

function getUpdateDate(update) {
  const raw = String(update?.completionDate || update?.completion_date || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(raw)) {
    const [d, m, y] = raw.split('/');
    return `${y}-${m}-${d}`;
  }
  if (raw) {
    const parsed = new Date(raw);
    if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);
  }
  const now = new Date();
  const formatter = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' });
  return formatter.format(now);
}

function isAdvanceUpdate(update) {
  const status = String(update?.status || '').trim().toLowerCase();
  return !status.includes('review');
}

function summarizeRowsProgress(rows, column) {
  const values = rows.map((row) => getCellPercent(row, column)).filter((value) => value != null);
  if (!values.length) return null;
  return Math.max(...values);
}

function buildStageUpdatePlan(sheet, update, options = {}) {
  const progress = Number(update?.progress || 0);
  const forceRewrite = Boolean(options.forceRewrite || update?.forceRewrite);
  const dateOnly = Boolean(options.dateOnly || update?.dateOnly);
  const rowChangesMap = options.rowChangesMap || new Map();
  const result = {
    id: update?.id || '',
    ok: false,
    success: false,
    trackingOk: false,
    needsTrackingUpdate: false,
    missingDate: false,
    specialNeedsUpdate: false,
    rowCount: 0,
    changedRows: 0,
    progressColumn: '',
    dateColumn: '',
    currentProgress: null,
    message: '',
  };

  if (!isAdvanceUpdate(update)) {
    result.message = 'Revisões devem ser tratadas manualmente pelo PCP.';
    return result;
  }
  if (![25, 50, 75, 100].includes(progress)) {
    result.message = 'Percentual inválido para atualização do Tracking.';
    return result;
  }

  const matchingRows = getMatchingTrackingRows(sheet, update);
  result.rowCount = matchingRows.length;
  if (!matchingRows.length) {
    result.message = 'Spool/drawing não localizado no Tracking.';
    return result;
  }

  const progressColumn = resolveProgressColumn(sheet, update, matchingRows[0]);
  if (!progressColumn) {
    result.message = 'Coluna de avanço não localizada no Tracking para este setor.';
    return result;
  }
  result.progressColumn = progressColumn.title;
  const desiredDecimal = progress / 100;
  const desiredPercent = progress;
  const currentProgress = summarizeRowsProgress(matchingRows, progressColumn);
  result.currentProgress = currentProgress == null ? null : Number(currentProgress.toFixed(2));
  const dateColumnName = getDateColumnNameForProgressColumn(progressColumn.title);
  const dateColumn = progress === 100 && dateColumnName ? findColumn(sheet, dateColumnName) : null;
  result.dateColumn = dateColumn?.title || dateColumnName || '';

  if (progress === 100 && dateColumnName && !dateColumn) {
    result.message = `Coluna de data "${dateColumnName}" não localizada no Tracking.`;
    return result;
  }

  const updateDate = getUpdateDate(update);
  let rowChangeCount = 0;
  let rowsWithDateMissing = 0;
  let rowsNeedingProgress = 0;
  let specialNeeds = 0;

  const isPainting100 = normalizeColumnTitle(getCanonicalProgressColumnName(progressColumn.title)) === normalizeColumnTitle('Surface preparation and/or coating') && progress === 100;
  const finalInspectionColumn = isPainting100 ? findColumn(sheet, 'Final Inspection') : null;
  const packageColumn = isPainting100 ? findColumn(sheet, 'Package and Delivered') : null;
  if (isPainting100 && (!finalInspectionColumn || !packageColumn)) {
    result.message = !finalInspectionColumn
      ? 'Coluna Final Inspection não localizada para alimentar a próxima etapa.'
      : 'Coluna Package and Delivered não localizada para alimentar 25% após Pintura 100%.';
    return result;
  }

  for (const row of matchingRows) {
    let changedThisRow = false;
    const current = getCellPercent(row, progressColumn);
    const currentSafe = current == null ? 0 : current;
    const canWriteProgress = current == null || currentSafe < desiredPercent || (forceRewrite && currentSafe <= desiredPercent);
    if (!dateOnly && canWriteProgress) {
      addCellChange(rowChangesMap, row.id, progressColumn, desiredDecimal);
      rowsNeedingProgress += 1;
      changedThisRow = true;
    }

    if (progress === 100 && dateColumn && !getCellHasValue(row, dateColumn)) {
      rowsWithDateMissing += 1;
      addCellChange(rowChangesMap, row.id, dateColumn, updateDate);
      changedThisRow = true;
    }

    if (isPainting100) {
      if (finalInspectionColumn) {
        const finalCurrent = getCellPercent(row, finalInspectionColumn);
        if (finalCurrent == null || finalCurrent < 25) {
          addCellChange(rowChangesMap, row.id, finalInspectionColumn, 0.25);
          specialNeeds += 1;
          changedThisRow = true;
        }
      }
      if (packageColumn) {
        const packageCurrent = getCellPercent(row, packageColumn);
        if (packageCurrent == null || packageCurrent < 25) {
          addCellChange(rowChangesMap, row.id, packageColumn, 0.25);
          specialNeeds += 1;
          changedThisRow = true;
        }
      }
    }

    if (changedThisRow) rowChangeCount += 1;
  }

  result.needsTrackingUpdate = rowsNeedingProgress > 0;
  result.missingDate = rowsWithDateMissing > 0;
  result.specialNeedsUpdate = specialNeeds > 0;
  result.changedRows = rowChangeCount;

  const alreadyAbove = currentProgress != null && currentProgress > desiredPercent;
  const alreadyEqual = currentProgress != null && Math.abs(currentProgress - desiredPercent) < 0.01;
  const progressOk = currentProgress != null && currentProgress >= desiredPercent && (!forceRewrite || alreadyAbove || alreadyEqual);
  result.trackingOk = progressOk && !result.missingDate && !result.specialNeedsUpdate;
  result.ok = true;
  result.success = true;
  if (rowChangeCount > 0) result.message = `Tracking atualizado em ${rowChangeCount} linha(s).`;
  else if (alreadyAbove) result.message = `Tracking já estava superior (${currentProgress.toFixed(0)}%).`;
  else if (alreadyEqual) result.message = 'Tracking já estava no mesmo avanço.';
  else result.message = 'Tracking conferido.';
  return result;
}

async function applyStageUpdatesToTracking(updates, options = {}) {
  const list = Array.isArray(updates) ? updates : [];
  if (!list.length) return { ok: false, results: [], errors: [], message: 'Nenhum apontamento informado.' };
  const { sheetId, sheet } = await fetchTrackingSheet();
  const rowChangesMap = new Map();
  const results = list.map((update) => buildStageUpdatePlan(sheet, update, { ...options, rowChangesMap }));
  const errors = results.filter((item) => !item.success).map((item) => ({ id: item.id, error: item.message }));

  if (!options.dryRun && rowChangesMap.size) {
    await updateRows(sheetId, rowChangesMap);
  }

  return {
    ok: errors.length === 0,
    sheetId,
    results,
    errors,
    changedRows: Array.from(rowChangesMap.values()).length,
    dryRun: Boolean(options.dryRun),
  };
}

async function listHistoryDatePendencies(updates) {
  const source = Array.isArray(updates) ? updates : [];
  const history100 = source.filter((item) => {
    const status = String(item?.status || '').trim().toLowerCase();
    return isAdvanceUpdate(item)
      && ['resolved', 'resolved_advance'].includes(status)
      && Number(item?.progress || 0) === 100;
  });
  if (!history100.length) return [];

  const { sheet } = await fetchTrackingSheet();
  const pendencies = [];

  for (const update of history100) {
    const rows = getMatchingTrackingRows(sheet, update);
    if (!rows.length) continue;
    const progressColumn = resolveProgressColumn(sheet, update, rows[0]);
    if (!progressColumn) continue;
    const dateColumnName = getDateColumnNameForProgressColumn(progressColumn.title);
    const dateColumn = dateColumnName ? findColumn(sheet, dateColumnName) : null;
    if (!dateColumn) continue;

    const isPainting100 = normalizeColumnTitle(getCanonicalProgressColumnName(progressColumn.title)) === normalizeColumnTitle('Surface preparation and/or coating');
    const finalInspectionColumn = isPainting100 ? findColumn(sheet, 'Final Inspection') : null;
    const packageColumn = isPainting100 ? findColumn(sheet, 'Package and Delivered') : null;

    const missingRows = rows.filter((row) => {
      const dateMissing = !getCellHasValue(row, dateColumn);
      const finalNeeds = isPainting100 && finalInspectionColumn && ((getCellPercent(row, finalInspectionColumn) ?? 0) < 25);
      const packageNeeds = isPainting100 && packageColumn && ((getCellPercent(row, packageColumn) ?? 0) < 25);
      return dateMissing || finalNeeds || packageNeeds;
    });
    if (!missingRows.length) continue;

    pendencies.push({
      id: String(update.id || ''),
      projectRowId: update.projectRowId,
      projectDisplay: update.projectDisplay || update.projectNumber || '',
      projectNumber: update.projectNumber || '',
      spoolIso: update.spoolIso || '',
      sector: update.sector || '',
      progress: Number(update.progress || 0),
      process: progressColumn.title,
      missingDateColumn: dateColumn.title,
      applyDate: getUpdateDate(update),
      rowCount: rows.length,
      affectedRows: missingRows.length,
      needsPaintingNextSteps: isPainting100,
      createdAt: update.createdAt || '',
      resolvedAt: update.resolvedAt || '',
    });
  }
  return pendencies;
}

module.exports = {
  TRACKING_PROGRESS_BY_SECTOR,
  TRACKING_DATE_BY_PROGRESS_COLUMN,
  normalizeColumnTitle,
  normalizeSpoolIdentity,
  normalizeSpoolIdentityLoose,
  parsePercentValue,
  fetchTrackingSheet,
  applyStageUpdatesToTracking,
  listHistoryDatePendencies,
};
