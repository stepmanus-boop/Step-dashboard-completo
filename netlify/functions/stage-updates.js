
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
  solda: ['Full welding execution'],
  producao: ['Spool Assemble and tack weld', 'Welding Preparation'],
  calderaria: ['Spool Assemble and tack weld', 'Welding Preparation', 'Material Separation', 'Material Release to Fabrication'],
  inspecao: ['Final Inspection', 'Hydro Test Pressure (QC)', 'Non Destructive Examination (QC)', 'Final Dimensional Inpection/3D (QC)', 'Initial Dimensional Inspection/3D'],
  pendente_envio: ['Package and Delivered', 'Final Inspection'],
};

function getTrackingUpdateColumnForSector(sector) {
  return TRACKING_UPDATE_COLUMN_BY_SECTOR[normalizeSectorValue(sector)] || '';
}

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

function findColumn(columns, columnTitle) {
  const target = normalizeColumnTitle(columnTitle);
  return (Array.isArray(columns) ? columns : []).find((column) => normalizeColumnTitle(column?.title) === target) || null;
}

function findColumnByCandidates(columns, candidates = []) {
  for (const title of candidates) {
    const column = findColumn(columns, title);
    if (column) return column;
  }

  const normalizedCandidates = candidates.map(normalizeColumnTitle).filter(Boolean);
  return (Array.isArray(columns) ? columns : []).find((column) => {
    const normalized = normalizeColumnTitle(column?.title);
    return normalizedCandidates.some((candidate) => normalized.includes(candidate) || candidate.includes(normalized));
  }) || null;
}

async function updateSmartsheetRowsWithProgressCells(sheetId, rows) {
  const cleanRows = (Array.isArray(rows) ? rows : [])
    .map((row) => ({
      id: Number(row?.id || 0),
      cells: (Array.isArray(row?.cells) ? row.cells : [])
        .filter((cell) => Number(cell?.columnId || 0))
        .map((cell) => ({
          columnId: Number(cell.columnId),
          value: getSmartsheetPercentCellValue(null, cell.progress),
          strict: false,
        })),
    }))
    .filter((row) => row.id && row.cells.length);

  if (!cleanRows.length) return { ok: true, count: 0 };

  await smartsheetFetch(`/sheets/${sheetId}/rows`, {
    method: 'PUT',
    body: JSON.stringify(cleanRows),
  });

  return { ok: true, count: cleanRows.length };
}

function getSmartsheetPercentCellValue(column, progress) {
  const value = Number(progress || 0);
  return value / 100;
}

async function updateSmartsheetRowsWithPercent(sheetId, column, rows) {
  const columnId = Number(column?.id || 0);
  const cleanRows = (Array.isArray(rows) ? rows : []).filter((row) => Number(row?.id || 0) && columnId);
  if (!cleanRows.length) return { ok: true, count: 0 };

  await smartsheetFetch(`/sheets/${sheetId}/rows`, {
    method: 'PUT',
    body: JSON.stringify(cleanRows.map((row) => {
      const progress = Number(row.progress || 0);
      return {
        id: Number(row.id),
        cells: [
          {
            columnId,
            value: getSmartsheetPercentCellValue(column, progress),
            strict: false,
          },
        ],
      };
    })),
  });

  return { ok: true, count: cleanRows.length };
}

async function updateSmartsheetRowsWithDate(sheetId, dateColumn, rows) {
  const dateColumnId = Number(dateColumn?.id || 0);
  const cleanRows = (Array.isArray(rows) ? rows : []).filter((row) => Number(row?.id || 0) && dateColumnId);
  if (!cleanRows.length) return { ok: true, count: 0 };

  await smartsheetFetch(`/sheets/${sheetId}/rows`, {
    method: 'PUT',
    body: JSON.stringify(cleanRows.map((row) => ({
      id: Number(row.id),
      cells: [
        {
          columnId: dateColumnId,
          value: normalizeDateForSmartsheet(row.completionDate),
          strict: false,
        },
      ],
    }))),
  });

  return { ok: true, count: cleanRows.length };
}

function makeUpdatedTrackingRecord(current, session, columnTitle, resolvedProgress = null, options = {}) {
  const now = new Date().toISOString();
  const finalProgress = Number(resolvedProgress == null ? current.progress || 0 : resolvedProgress);
  const actor = session.username || '';
  const actorName = session.name || session.username || 'PCP';
  return {
    ...current,
    status: 'resolved_advance',
    resolvedBy: actor,
    resolvedByName: actorName,
    resolvedAt: now,
    resolutionNote: options.higherCurrentProgress
      ? `Concluído automaticamente: o Tracking já estava em ${Math.round(Number(finalProgress || 0))}%, superior ao apontamento de ${Number(current.progress || 0)}%.`
      : 'Concluído automaticamente após conferência/atualização do Tracking pelo PCP.',
    trackingCheckedAt: now,
    trackingProgress: Number.isFinite(finalProgress) ? finalProgress : Number(current.progress || 0),
    trackingMatched: true,
    trackingStatus: 'matched',
    trackingUpdatedBy: options.skipUpdatedBy ? (current.trackingUpdatedBy || '') : actor,
    trackingUpdatedByName: options.skipUpdatedBy ? (current.trackingUpdatedByName || '') : actorName,
    trackingUpdatedAt: options.skipUpdatedBy ? (current.trackingUpdatedAt || '') : now,
    trackingUpdatedColumn: columnTitle,
  };
}

async function persistResolvedTrackingRecord(id, updatedRecord) {
  if (!isSupabaseConfigured()) return updatedRecord;

  const saved = await updateStageUpdate(id, {
    status: updatedRecord.status || 'resolved_advance',
    resolvedBy: updatedRecord.resolvedBy || '',
    resolvedByName: updatedRecord.resolvedByName || '',
    resolvedAt: updatedRecord.resolvedAt || new Date().toISOString(),
    resolutionNote: updatedRecord.resolutionNote || '',
  });

  return saved ? { ...updatedRecord, ...saved } : updatedRecord;
}

function getTrackingOverrideFromBody(body, id) {
  const targetId = String(id || '').trim();
  const items = Array.isArray(body?.items) ? body.items : [];
  const found = items.find((item) => String(item?.id || '').trim() === targetId);

  if (found) {
    return {
      ...found,
      rowId: found.rowId || found.trackingRowId || '',
      rowIds: Array.isArray(found.rowIds) ? found.rowIds : [],
      rows: Array.isArray(found.rows) ? found.rows : [],
      sheetId: found.sheetId || found.trackingSheetId || '',
      columnTitle: found.columnTitle || found.trackingUpdateColumn || '',
    };
  }

  return {
    id: targetId,
    rowId: body?.rowId || body?.trackingRowId || '',
    rowIds: Array.isArray(body?.rowIds) ? body.rowIds : [],
    rows: Array.isArray(body?.rows) ? body.rows : [],
    sheetId: body?.sheetId || body?.trackingSheetId || '',
    columnTitle: body?.columnTitle || body?.trackingUpdateColumn || '',
  };
}

async function updatePaintingCompletionNextSteps(sheetId, columns, rows) {
  const sourceRows = Array.isArray(rows) ? rows : [];
  const paintingRows = sourceRows.filter((row) => Number(row?.id || 0) && Number(row?.progress || 0) >= 100);
  if (!paintingRows.length) return { ok: true, finalInspection: 0, packageDelivered: 0 };

  const finalColumn = findColumnByCandidates(columns, [
    'Final Inspection',
    'Final Inspection ',
    'Final Inspec',
  ]);

  const packageColumn = findColumnByCandidates(columns, [
    'Package and Delivered',
    'Package Delivered',
    'Package & Delivered',
    'Package',
  ]);

  const rowsToUpdate = [];
  let finalInspection = 0;
  let packageDelivered = 0;

  for (const row of paintingRows) {
    const cells = [];

    const currentFinal = Number(row.finalInspectionProgress);
    if (finalColumn && (!Number.isFinite(currentFinal) || currentFinal < 25)) {
      cells.push({ columnId: finalColumn.id, progress: 25 });
      finalInspection += 1;
    }

    const currentPackage = Number(row.packageDeliveredProgress);
    if (packageColumn && (!Number.isFinite(currentPackage) || currentPackage < 25)) {
      cells.push({ columnId: packageColumn.id, progress: 25 });
      packageDelivered += 1;
    }

    if (cells.length) {
      rowsToUpdate.push({ id: row.id, cells });
    }
  }

  if (rowsToUpdate.length) {
    await updateSmartsheetRowsWithProgressCells(sheetId, rowsToUpdate);
  }

  return { ok: true, finalInspection, packageDelivered };
}

const TRACKING_DATE_COLUMN_BY_PROGRESS_COLUMN = {
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

function getTrackingDateColumnForProgressColumn(columnTitle) {
  return TRACKING_DATE_COLUMN_BY_PROGRESS_COLUMN[String(columnTitle || '').trim()] || '';
}

function normalizeDateForSmartsheet(value) {
  const raw = String(value || '').trim();
  if (!raw) return new Date().toISOString().slice(0, 10);

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

  const br = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (br) {
    const day = br[1].padStart(2, '0');
    const month = br[2].padStart(2, '0');
    const year = br[3].length === 2 ? `20${br[3]}` : br[3];
    return `${year}-${month}-${day}`;
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed.toISOString().slice(0, 10);

  return new Date().toISOString().slice(0, 10);
}

function hasTrackingDateValue(spool, dateColumnTitle) {
  if (!dateColumnTitle) return true;
  const value = spool?.stageValues?.[dateColumnTitle];
  if (value == null) return false;
  const text = String(value).trim();
  return Boolean(text && text !== 'N/A' && text !== 'Não' && text !== '-');
}

function normalizeSpoolIdentity(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '')
    .replace(/[–—]/g, '-')
    .toLowerCase();
}

function normalizeTrackingReference(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '')
    .toLowerCase();
}

function trackingReferenceKeys(value) {
  const key = normalizeTrackingReference(value);
  if (!key) return [];
  const withoutBsp = key.replace(/^bsp/, '');
  return Array.from(new Set([key, withoutBsp].filter(Boolean)));
}

function collectTrackingScopeFromUpdates(updates) {
  const projectKeys = new Set();
  const spoolKeys = new Set();

  for (const item of Array.isArray(updates) ? updates : []) {
    for (const key of [
      ...trackingReferenceKeys(item?.projectDisplay),
      ...trackingReferenceKeys(item?.projectNumber),
    ]) {
      projectKeys.add(key);
    }

    for (const key of trackingReferenceKeys(item?.spoolIso || item?.spoolDescription)) {
      spoolKeys.add(key);
    }
  }

  return { projectKeys, spoolKeys };
}

function projectMatchesTrackingScope(project, scope) {
  if (!scope || (!scope.projectKeys?.size && !scope.spoolKeys?.size)) return false;

  const projectValues = [
    project?.projectDisplay,
    project?.projectNumber,
    project?.project,
    project?.clientWoNumber,
    project?.projectName,
  ];

  return projectValues.some((value) =>
    trackingReferenceKeys(value).some((key) => scope.projectKeys.has(key))
  );
}

function spoolMatchesTrackingScope(spool, scope) {
  if (!scope || !scope.spoolKeys?.size) return false;
  return trackingReferenceKeys(spool?.iso || spool?.drawing || spool?.description || spool?.spoolIso)
    .some((key) => scope.spoolKeys.has(key));
}

function getSpoolIdentity(spool) {
  return normalizeSpoolIdentity(spool?.iso || spool?.drawing || spool?.spoolIso || spool?.description || '');
}

function getDuplicateTrackingRows(project, spool, columnTitle = '') {
  const baseKey = getSpoolIdentity(spool);
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const duplicates = spools.filter((item) => {
    const key = getSpoolIdentity(item);
    return key && baseKey && key === baseKey && Number(item?.rowId || 0);
  });

  const source = duplicates.length ? duplicates : (Number(spool?.rowId || 0) ? [spool] : []);

  return source.map((item) => {
    const stageValues = item?.stageValues || {};
    const currentProgress = columnTitle
      ? parseTrackingPercent(stageValues[columnTitle])
      : null;
    const finalInspectionProgress = parseTrackingPercent(stageValues['Final Inspection']);
    const packageDeliveredProgress = parseTrackingPercent(stageValues['Package and Delivered']);
    return {
      rowId: Number(item?.rowId || 0),
      spoolIso: item?.iso || item?.drawing || '',
      currentProgress: currentProgress == null ? null : Number(currentProgress),
      finalInspectionProgress: finalInspectionProgress == null ? null : Number(finalInspectionProgress),
      packageDeliveredProgress: packageDeliveredProgress == null ? null : Number(packageDeliveredProgress),
    };
  }).filter((item, index, arr) =>
    item.rowId && arr.findIndex((other) => Number(other.rowId) === Number(item.rowId)) === index
  );
}

function getMaxTrackingProgress(rows = []) {
  const values = (Array.isArray(rows) ? rows : [])
    .map((item) => Number(item?.currentProgress))
    .filter((value) => Number.isFinite(value));
  return values.length ? Math.max(...values) : null;
}

const TRACKING_DATE_PENDENCY_RULES = [
  { progressColumn: 'Surface preparation and/or coating', dateColumn: 'Coating Finish Date', label: 'Pintura 100% sem data' },
  { progressColumn: 'Full welding execution', dateColumn: 'Welding Finish Date', label: 'Solda 100% sem data' },
  { progressColumn: 'Spool Assemble and tack weld', dateColumn: 'Boilermaker Finish Date', label: 'Pré-montagem 100% sem data' },
  { progressColumn: 'Hydro Test Pressure (QC)', dateColumn: 'TH Finish Date', label: 'TH 100% sem data' },
  { progressColumn: 'Final Dimensional Inpection/3D (QC)', dateColumn: 'Inspection Finish Date (QC)', label: 'Inspeção final 100% sem data' },
  { progressColumn: 'Non Destructive Examination (QC)', dateColumn: 'Inspection Finish Date (QC)', label: 'END 100% sem data' },
  { progressColumn: 'HDG / FBE.  (PAINT)', dateColumn: 'HDG / FBE DATE RETORNO (PAINT)', label: 'HDG/FBE 100% sem data de retorno' },
  { progressColumn: 'Final Inspection', dateColumn: 'Project Finish Date', label: 'Final Inspection 100% sem Project Finish Date' },
  { progressColumn: 'Package and Delivered', dateColumn: 'Project Finish Date', label: 'Package 100% sem Project Finish Date' },
];

function buildDatePendencyId(rowId, progressColumn, dateColumn) {
  return `${rowId}::${normalizeColumnTitle(progressColumn)}::${normalizeColumnTitle(dateColumn)}`;
}

function getDatePendencySector(progressColumn) {
  const key = normalizeColumnTitle(progressColumn);
  if (key.includes('surface') || key.includes('coating') || key.includes('hdgfbe')) return 'pintura';
  if (key.includes('fullwelding')) return 'solda';
  if (key.includes('spoolassemble')) return 'calderaria';
  if (key.includes('hydro') || key.includes('inspection') || key.includes('nondestructive')) return 'inspecao';
  if (key.includes('package') || key.includes('finalinspection')) return 'pendente_envio';
  return '';
}

function findBestCompletionDateForPendency(updates, spoolIso, progressColumn, fallbackDate = '') {
  const sector = getDatePendencySector(progressColumn);
  const candidates = (Array.isArray(updates) ? updates : [])
    .filter((item) => String(item?.spoolIso || '').trim().toLowerCase() === String(spoolIso || '').trim().toLowerCase())
    .filter((item) => !sector || normalizeSectorValue(item?.sector) === sector)
    .filter((item) => Number(item?.progress || 0) >= 100)
    .sort((a, b) => new Date(b.resolvedAt || b.createdAt || 0) - new Date(a.resolvedAt || a.createdAt || 0));

  const candidate = candidates[0];
  return normalizeDateForSmartsheet(candidate?.completionDate || candidate?.resolvedAt || candidate?.createdAt || fallbackDate);
}

async function buildTrackingDatePendencies() {
  const payload = await loadProjectPayload();
  const updates = await listUpdates();
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  const sheetId = payload?.meta?.sheetId || '';
  const pendencyMap = new Map();

  const historyUpdates = (Array.isArray(updates) ? updates : [])
    .filter((item) => String(item?.status || '').trim().toLowerCase().startsWith('resolved'))
    .filter((item) => Number(item?.progress || 0) >= 100)
    .filter((item) => !String(item?.status || '').trim().toLowerCase().includes('review'));

  for (const update of historyUpdates) {
    const progressColumn = getTrackingUpdateColumnForSector(update?.sector);
    if (!progressColumn) continue;

    const dateColumn = getTrackingDateColumnForProgressColumn(progressColumn);
    if (!dateColumn) continue;

    let project = findProjectInPayload(projects, update?.projectRowId);
    let spool = project ? findSpoolInProject(project, update?.spoolIso) : null;

    if (!project || !spool) {
      for (const candidateProject of projects) {
        const candidateSpool = findSpoolInProject(candidateProject, update?.spoolIso);
        if (candidateSpool) {
          project = candidateProject;
          spool = candidateSpool;
          break;
        }
      }
    }

    if (!project || !spool) continue;

    const duplicateRows = getDuplicateTrackingRows(project, spool, progressColumn);
    if (!duplicateRows.length) continue;

    for (const duplicateRow of duplicateRows) {
      const rowId = Number(duplicateRow?.rowId || 0);
      if (!rowId) continue;

      const duplicateSpool = (Array.isArray(project?.spools) ? project.spools : [])
        .find((item) => Number(item?.rowId || 0) === rowId) || spool;
      const stageValues = duplicateSpool?.stageValues || {};
      const progress = parseTrackingPercent(stageValues[progressColumn]);

      if (progress == null || progress < 100) continue;
      if (hasTrackingDateValue(duplicateSpool, dateColumn)) continue;

      const key = buildDatePendencyId(rowId, progressColumn, dateColumn);
      if (pendencyMap.has(key)) continue;

      pendencyMap.set(key, {
        id: key,
        sheetId,
        rowId,
        rowIds: duplicateRows.map((row) => row.rowId),
        duplicateRows,
        projectDisplay: project?.projectDisplay || project?.projectNumber || update?.projectDisplay || update?.projectNumber || '',
        client: project?.client || update?.client || '',
        spoolIso: duplicateSpool?.iso || duplicateSpool?.drawing || update?.spoolIso || '',
        progressColumn,
        dateColumn,
        progress: Number(progress.toFixed(2)),
        label: `${progressColumn} 100% sem data`,
        suggestedDate: normalizeDateForSmartsheet(update?.completionDate || update?.resolvedAt || update?.createdAt || new Date().toISOString().slice(0, 10)),
      });
    }
  }

  return Array.from(pendencyMap.values()).sort((a, b) =>
    String(a.projectDisplay).localeCompare(String(b.projectDisplay), 'pt-BR', { numeric: true, sensitivity: 'base' })
    || String(a.spoolIso).localeCompare(String(b.spoolIso), 'pt-BR', { numeric: true, sensitivity: 'base' })
    || String(a.progressColumn).localeCompare(String(b.progressColumn), 'pt-BR', { numeric: true, sensitivity: 'base' })
  );
}

async function fixTrackingDatePendencies(ids = []) {
  const all = await buildTrackingDatePendencies();
  const selectedSet = new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean));
  const selected = selectedSet.size ? all.filter((item) => selectedSet.has(String(item.id))) : all;
  const errors = [];
  const fixed = [];

  if (!selected.length) {
    return { fixed, errors, pending: all };
  }

  const columnsCache = new Map();
  const getColumnsCached = async (sheetId) => {
    const key = String(sheetId || '');
    if (!columnsCache.has(key)) {
      columnsCache.set(key, await getSheetColumns(sheetId));
    }
    return columnsCache.get(key);
  };

  const bySheetAndDateColumn = new Map();
  const bySheetAndProgressColumn = new Map();
  const bySheetPaintingNextSteps = new Map();

  for (const item of selected) {
    const duplicateRows = Array.isArray(item.duplicateRows) && item.duplicateRows.length
      ? item.duplicateRows
      : (Array.isArray(item.rowIds) && item.rowIds.length
        ? item.rowIds.map((rowId) => ({ rowId }))
        : [{ rowId: item.rowId }]);

    const dateGroupKey = `${item.sheetId}::${normalizeColumnTitle(item.dateColumn)}`;
    if (!bySheetAndDateColumn.has(dateGroupKey)) {
      bySheetAndDateColumn.set(dateGroupKey, { sheetId: item.sheetId, dateColumn: item.dateColumn, rows: [] });
    }

    const progressGroupKey = `${item.sheetId}::${normalizeColumnTitle(item.progressColumn)}`;
    if (!bySheetAndProgressColumn.has(progressGroupKey)) {
      bySheetAndProgressColumn.set(progressGroupKey, { sheetId: item.sheetId, progressColumn: item.progressColumn, rows: [] });
    }

    if (item.progressColumn === 'Surface preparation and/or coating') {
      const paintingGroupKey = `${item.sheetId}::painting-next`;
      if (!bySheetPaintingNextSteps.has(paintingGroupKey)) {
        bySheetPaintingNextSteps.set(paintingGroupKey, { sheetId: item.sheetId, rows: [] });
      }
    }

    for (const duplicate of duplicateRows) {
      const rowId = Number(duplicate?.rowId || duplicate || 0);
      if (!rowId) continue;

      const commonRow = {
        id: rowId,
        progress: 100,
        completionDate: item.suggestedDate,
        itemId: item.id,
        finalInspectionProgress: duplicate?.finalInspectionProgress,
        packageDeliveredProgress: duplicate?.packageDeliveredProgress,
      };

      bySheetAndDateColumn.get(dateGroupKey).rows.push(commonRow);
      bySheetAndProgressColumn.get(progressGroupKey).rows.push(commonRow);

      if (item.progressColumn === 'Surface preparation and/or coating') {
        const paintingGroupKey = `${item.sheetId}::painting-next`;
        bySheetPaintingNextSteps.get(paintingGroupKey).rows.push(commonRow);
      }
    }
  }

  for (const group of bySheetAndProgressColumn.values()) {
    try {
      const columns = await getColumnsCached(group.sheetId);
      const progressColumn = findColumn(columns, group.progressColumn);
      if (!progressColumn) throw new Error(`Coluna de avanço "${group.progressColumn}" não encontrada.`);

      const uniqueRows = Array.from(new Map(group.rows.map((row) => [String(row.id), row])).values());
      await updateSmartsheetRowsWithPercent(group.sheetId, progressColumn, uniqueRows);
    } catch (error) {
      for (const row of group.rows) {
        errors.push({ id: row.itemId, error: error.message || 'Falha ao corrigir avanço.' });
      }
    }
  }

  for (const group of bySheetAndDateColumn.values()) {
    try {
      const columns = await getColumnsCached(group.sheetId);
      const dateColumn = findColumn(columns, group.dateColumn);
      if (!dateColumn) throw new Error(`Coluna de data "${group.dateColumn}" não encontrada.`);

      const uniqueRows = Array.from(new Map(group.rows.map((row) => [String(row.id), row])).values());
      await updateSmartsheetRowsWithDate(group.sheetId, dateColumn, uniqueRows);

      const fixedIds = new Set(group.rows.map((row) => String(row.itemId)));
      for (const item of selected) {
        if (fixedIds.has(String(item.id)) && !fixed.some((fixedItem) => String(fixedItem.id) === String(item.id))) {
          fixed.push(item);
        }
      }
    } catch (error) {
      for (const row of group.rows) {
        errors.push({ id: row.itemId, error: error.message || 'Falha ao corrigir data.' });
      }
    }
  }

  for (const group of bySheetPaintingNextSteps.values()) {
    try {
      const columns = await getColumnsCached(group.sheetId);
      const uniqueRows = Array.from(new Map(group.rows.map((row) => [String(row.id), row])).values());
      await updatePaintingCompletionNextSteps(group.sheetId, columns, uniqueRows);
    } catch (error) {
      for (const row of group.rows) {
        errors.push({ id: row.itemId, error: error.message || 'Falha ao alimentar Final Inspection/Package and Delivered.' });
      }
    }
  }

  const fixedIds = new Set(fixed.map((item) => String(item.id)));
  const pending = all.filter((item) => !fixedIds.has(String(item.id)));
  return { fixed, errors, pending };
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
  const duplicateRows = getDuplicateTrackingRows(project, spool, columnTitle);
  const maxCurrentProgress = getMaxTrackingProgress(duplicateRows);

  if (maxCurrentProgress != null && Number(maxCurrentProgress) > progress) {
    return {
      update: makeUpdatedTrackingRecord(
        applyTrackingVerification({
          ...update,
          trackingUpdatedBy: update.trackingUpdatedBy || '',
          trackingUpdatedByName: update.trackingUpdatedByName || '',
          trackingUpdatedAt: update.trackingUpdatedAt || '',
          trackingUpdatedColumn: columnTitle,
        }, project, spool, payload),
        session,
        columnTitle,
        maxCurrentProgress,
        { skipUpdatedBy: true, higherCurrentProgress: true }
      ),
      applied: false,
      alreadyUpdated: true,
      higherCurrentProgress: true,
      message: `Não atualizado: o Tracking já está em ${Math.round(maxCurrentProgress)}%, superior ao apontamento de ${progress}%.`,
      columnTitle,
      currentProgress: maxCurrentProgress,
    };
  }

  const sheetId = payload?.meta?.sheetId;
  if (!sheetId) {
    const err = new Error('Sheet ID do Tracking não localizado.');
    err.statusCode = 500;
    throw err;
  }

  const targetRows = duplicateRows.length ? duplicateRows : [{ rowId: Number(spool.rowId || 0), currentProgress: null }];
  const writableRows = targetRows
    .map((row) => ({
      id: Number(row.rowId || 0),
      progress,
      completionDate: update.completionDate,
      finalInspectionProgress: row.finalInspectionProgress,
      packageDeliveredProgress: row.packageDeliveredProgress,
    }))
    .filter((row) => row.id);

  if (!writableRows.length) {
    const err = new Error('Linha da spool no Tracking não localizada.');
    err.statusCode = 404;
    throw err;
  }

  const columns = await getSheetColumns(sheetId);
  const column = findColumn(columns, columnTitle);
  if (!column) {
    const err = new Error(`Coluna "${columnTitle}" não encontrada no Tracking.`);
    err.statusCode = 404;
    throw err;
  }

  await updateSmartsheetRowsWithPercent(sheetId, column, writableRows);

  const dateColumnTitle = getTrackingDateColumnForProgressColumn(columnTitle);
  if (progress >= 100 && dateColumnTitle) {
    const dateColumn = findColumn(columns, dateColumnTitle);
    if (dateColumn) {
      await updateSmartsheetRowsWithDate(sheetId, dateColumn, writableRows);
    }
  }

  if (columnTitle === 'Surface preparation and/or coating' && progress >= 100) {
    await updatePaintingCompletionNextSteps(sheetId, columns, writableRows);
  }

  const now = new Date().toISOString();
  return {
    update: makeUpdatedTrackingRecord({
      ...update,
      trackingCheckedAt: now,
      trackingProgress: progress,
      trackingMatched: true,
      trackingStatus: 'matched',
      trackingUpdatedBy: session.username || '',
      trackingUpdatedByName: session.name || session.username || 'PCP',
      trackingUpdatedAt: now,
      trackingUpdatedColumn: columnTitle,
      trackingRowIds: targetRows.map((row) => row.rowId).filter(Boolean),
      trackingDuplicateRows: targetRows,
    }, session, columnTitle, progress),
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

function applyTrackingVerification(update, project, spool, payload = null) {
  const progress = Number(update?.progress || 0);
  const trackingProgress = getTrackingProgressForSector(spool, update?.sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;
  const trackingUpdateColumn = getTrackingUpdateColumnForSector(update?.sector);
  const trackingDateColumn = getTrackingDateColumnForProgressColumn(trackingUpdateColumn);
  const duplicateRows = getDuplicateTrackingRows(project, spool, trackingUpdateColumn);
  const trackingMissingDate = Boolean(
    trackingMatched
    && Number(trackingProgress || 0) >= 100
    && trackingDateColumn
    && duplicateRows.some((row) => {
      const duplicate = (Array.isArray(project?.spools) ? project.spools : []).find((item) => Number(item?.rowId || 0) === Number(row.rowId));
      return duplicate && !hasTrackingDateValue(duplicate, trackingDateColumn);
    })
  );

  return {
    ...update,
    trackingCheckedAt: new Date().toISOString(),
    trackingProgress: trackingProgress == null ? null : Number(trackingProgress.toFixed(2)),
    trackingMatched,
    trackingStatus: trackingProgress == null
      ? 'not_found'
      : (trackingMatched ? 'matched' : 'waiting'),
    trackingMissingDate,
    trackingDateColumn,
    trackingSheetId: payload?.meta?.sheetId || update?.trackingSheetId || '',
    trackingRowId: spool?.rowId || update?.trackingRowId || '',
    trackingRowIds: duplicateRows.map((row) => row.rowId),
    trackingDuplicateRows: duplicateRows,
    trackingUpdateColumn,
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
    return applyTrackingVerification(item, project, spool, payload);
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
      const action = String(body.action || '').trim().toLowerCase();

      if (action === 'list_date_pendencies') {
        const pendencies = await buildTrackingDatePendencies();
        return jsonResponse(200, {
          ok: true,
          pendencies,
          count: pendencies.length,
        });
      }

      if (action === 'fix_date_pendencies') {
        const idsToFix = Array.isArray(body.ids) ? body.ids.map((id) => String(id || '').trim()).filter(Boolean) : [];
        const result = await fixTrackingDatePendencies(idsToFix);
        const remaining = Array.isArray(result.pending) ? result.pending : [];
        return jsonResponse(result.fixed.length ? 200 : 400, {
          ok: result.fixed.length > 0,
          error: result.fixed.length ? '' : (result.errors[0]?.error || 'Nenhuma pendência de data corrigida.'),
          fixed: result.fixed,
          fixedCount: result.fixed.length,
          errors: result.errors,
          pendencies: remaining,
          count: remaining.length,
        });
      }

      if (action === 'update_tracking') {
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
        const fastGroups = new Map();
        const pendingSaveIndexes = new Map();

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

          const override = getTrackingOverrideFromBody(body, id);
          const sheetId = String(override.sheetId || override.trackingSheetId || current.trackingSheetId || '').trim();
          const columnTitle = String(override.columnTitle || override.trackingUpdateColumn || current.trackingUpdateColumn || getTrackingUpdateColumnForSector(current.sector) || '').trim();
          const progress = Number(current.progress || 0);
          const overrideRows = Array.isArray(override.rows)
            ? override.rows.map((row) => ({
                rowId: Number(row?.rowId || row?.id || 0),
                currentProgress: Number(row?.currentProgress),
              })).filter((row) => row.rowId)
            : [];
          const overrideRowIds = Array.isArray(override.rowIds)
            ? override.rowIds.map((rowId) => Number(rowId || 0)).filter(Boolean)
            : [];
          const rowId = Number(override.rowId || override.trackingRowId || current.trackingRowId || 0);
          const targetRows = overrideRows.length
            ? overrideRows
            : (overrideRowIds.length ? overrideRowIds.map((id) => ({ rowId: id, currentProgress: Number(current.trackingProgress) })) : (rowId ? [{ rowId, currentProgress: Number(current.trackingProgress) }] : []));

          const maxCurrentProgress = getMaxTrackingProgress(targetRows);
          if (maxCurrentProgress != null && maxCurrentProgress > progress) {
            let updatedRecord = makeUpdatedTrackingRecord(current, session, columnTitle, maxCurrentProgress, { skipUpdatedBy: true, higherCurrentProgress: true });
            updatedRecord = await persistResolvedTrackingRecord(id, updatedRecord);
            if (!isSupabaseConfigured()) {
              updates[index] = updatedRecord;
            }
            updated.push(updatedRecord);
            results.push({
              id,
              applied: false,
              alreadyUpdated: true,
              higherCurrentProgress: true,
              progress,
              currentProgress: maxCurrentProgress,
              columnTitle,
              message: `Não atualizado: o Tracking já está em ${Math.round(maxCurrentProgress)}%, superior ao apontamento de ${progress}%.`,
            });
            continue;
          }

          if (sheetId && targetRows.length && columnTitle && [25, 50, 75, 100].includes(progress)) {
            const groupKey = `${sheetId}::${normalizeColumnTitle(columnTitle)}`;
            if (!fastGroups.has(groupKey)) {
              fastGroups.set(groupKey, { sheetId, columnTitle, rows: [] });
            }
            for (const targetRow of targetRows) {
              fastGroups.get(groupKey).rows.push({
                id: targetRow.rowId,
                progress,
                updateId: id,
                index,
                completionDate: current.completionDate,
                finalInspectionProgress: targetRow.finalInspectionProgress,
                packageDeliveredProgress: targetRow.packageDeliveredProgress,
              });
            }
            pendingSaveIndexes.set(id, { index, current, columnTitle });
            continue;
          }

          try {
            const result = await updateTrackingCellForStageUpdate(current, session);
            let updatedRecord = {
              ...current,
              ...result.update,
            };
            updatedRecord = await persistResolvedTrackingRecord(id, updatedRecord);

            if (!isSupabaseConfigured()) {
              updates[index] = updatedRecord;
            }

            updated.push(updatedRecord);
            results.push({
              id,
              applied: result.applied,
              alreadyUpdated: result.alreadyUpdated,
              higherCurrentProgress: Boolean(result.higherCurrentProgress),
              progress: Number(current.progress || 0),
              currentProgress: result.currentProgress,
              columnTitle: result.columnTitle,
              message: result.message || '',
            });
          } catch (error) {
            errors.push({ id, error: error.message || 'Falha ao atualizar Tracking.' });
          }
        }

        for (const group of fastGroups.values()) {
          try {
            const columns = await getSheetColumns(group.sheetId);
            const column = findColumn(columns, group.columnTitle);
            if (!column) {
              throw new Error(`Coluna "${group.columnTitle}" não encontrada no Tracking.`);
            }

            const maxByRowId = new Map();
            for (const row of group.rows) {
              const key = String(row.id);
              const existing = maxByRowId.get(key);
              if (!existing || Number(row.progress || 0) > Number(existing.progress || 0)) {
                maxByRowId.set(key, { id: row.id, progress: row.progress });
              }
            }

            const rowsToWrite = Array.from(maxByRowId.values()).map((row) => {
              const source = group.rows.find((item) => Number(item.id) === Number(row.id)) || row;
              return {
                ...row,
                completionDate: source.completionDate,
                finalInspectionProgress: source.finalInspectionProgress,
                packageDeliveredProgress: source.packageDeliveredProgress,
              };
            });
            await updateSmartsheetRowsWithPercent(group.sheetId, column, rowsToWrite);

            const dateColumnTitle = getTrackingDateColumnForProgressColumn(group.columnTitle);
            if (dateColumnTitle) {
              const dateColumn = findColumn(columns, dateColumnTitle);
              if (dateColumn) {
                await updateSmartsheetRowsWithDate(group.sheetId, dateColumn, rowsToWrite);
              }
            }

            if (group.columnTitle === 'Surface preparation and/or coating') {
              await updatePaintingCompletionNextSteps(group.sheetId, columns, rowsToWrite);
            }

            for (const row of group.rows) {
              const pending = pendingSaveIndexes.get(row.updateId);
              if (!pending) continue;

              const writtenProgress = Number(maxByRowId.get(String(row.id))?.progress || row.progress || 0);
              const higherCurrentProgress = writtenProgress > Number(row.progress || 0);
              let updatedRecord = makeUpdatedTrackingRecord(
                pending.current,
                session,
                group.columnTitle,
                writtenProgress,
                { higherCurrentProgress }
              );
              updatedRecord = await persistResolvedTrackingRecord(row.updateId, updatedRecord);
              if (!isSupabaseConfigured()) {
                updates[pending.index] = updatedRecord;
              }
              updated.push(updatedRecord);
              results.push({
                id: row.updateId,
                applied: !higherCurrentProgress,
                alreadyUpdated: higherCurrentProgress,
                higherCurrentProgress,
                progress: Number(row.progress || 0),
                currentProgress: writtenProgress,
                columnTitle: group.columnTitle,
                message: higherCurrentProgress
                  ? `Não atualizado individualmente: havia apontamento superior de ${writtenProgress}% para a mesma spool; foi mantido o maior avanço.`
                  : '',
              });
            }
          } catch (error) {
            for (const row of group.rows) {
              errors.push({ id: row.updateId, error: error.message || 'Falha ao atualizar Tracking.' });
            }
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
