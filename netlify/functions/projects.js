const fs = require('fs');
const fsPromises = require('fs/promises');
const path = require('path');
const { isSupabaseConfigured, getUserById } = require('./_supabase');
const { ensureIsoQrCodesForPayload } = require('./_isoQrCodes');
const { jsonResponse, requireSession } = require('./_auth');
const {
  readTrackingCache,
  writeTrackingCache,
  TRACKING_CACHE_MIN_WRITE_INTERVAL_MS,
} = require('./_trackingCache');
const API_BASE = process.env.SMARTSHEET_API_BASE || "https://api.smartsheet.com/2.0";
// Build Brasil: usa a mesma rotina estável do painel Portugal, porém apontando para as fontes BR.
const OPERATION_REGION = 'BR';
const SHEET_NAME = process.env.SMARTSHEET_SHEET_NAME_BR || process.env.SMARTSHEET_SHEET_NAME || "Progress Tracking Sheet - Piping Fabrication";
const SHEET_ID_ENV = process.env.SMARTSHEET_SHEET_ID_BR || process.env.SMARTSHEET_SHEET_ID || "";
const WIP_STEP_SHEET_NAME = process.env.SMARTSHEET_WIP_STEP_SHEET_NAME_BR || process.env.SMARTSHEET_WIP_STEP_SHEET_NAME || "Work in Progress - STEP";
const WIP_STEP_SHEET_ID_ENV = process.env.SMARTSHEET_WIP_STEP_SHEET_ID_BR || process.env.SMARTSHEET_WORK_IN_PROGRESS_BR_SHEET_ID || process.env.SMARTSHEET_WIP_STEP_SHEET_ID || process.env.SMARTSHEET_WORK_IN_PROGRESS_STEP_SHEET_ID || "";
const TOKEN = process.env.SMARTSHEET_TOKEN || process.env.SMARTSHEET_ACCESS_TOKEN || process.env.SMARTSHEET_API_TOKEN || process.env.SMARTSHEET_BEARER_TOKEN || process.env.SMARTSHEET_PAT || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN || "";
const PROJECTS_FAST_CACHE_MS = Number(process.env.PROJECTS_FAST_CACHE_MS || 30 * 60 * 1000); // v11 performance: 30 minutos default
const SESSION_HYDRATION_CACHE_MS = Number(process.env.SESSION_HYDRATION_CACHE_MS || 5 * 60 * 1000);

const cache = global.__STEP_PROGRESS_CACHE__ || {
  sheetId: null,
  sheetName: null,
  version: null,
  wipStepSheetId: null,
  wipStepSheetName: null,
  wipStepVersion: null,
  payload: null,
  lastSync: null,
  lastVersionCheck: null,
};
global.__STEP_PROGRESS_CACHE__ = cache;
if (cache.lastVersionCheck == null) cache.lastVersionCheck = null;

const sessionHydrationCache = global.__SESSION_HYDRATION_CACHE__ || {};
global.__SESSION_HYDRATION_CACHE__ = sessionHydrationCache;

// v36.84: cache específico para Portal do Cliente.
// Evita carregar a planilha operacional inteira quando o usuário logado é cliente.
const CLIENT_PROJECTS_FAST_CACHE_MS = Number(process.env.CLIENT_PROJECTS_FAST_CACHE_MS || 20 * 60 * 1000);
const CLIENT_SMARTSHEET_ROW_CHUNK_SIZE = Math.max(25, Number(process.env.CLIENT_SMARTSHEET_ROW_CHUNK_SIZE || 300));
const clientPayloadCache = global.__STEP_CLIENT_PROJECTS_CACHE__ || {};
global.__STEP_CLIENT_PROJECTS_CACHE__ = clientPayloadCache;

// v37.08: cache persistente no Supabase com modo seguro para plano Free.
// Mantém uma única linha de cache por ambiente, evita histórico/snapshots e
// impede que clientes criem várias linhas específicas no banco.
const PERSISTENT_CACHE_MAX_AGE_MS = Math.max(10 * 60 * 1000, Number(process.env.PERSISTENT_TRACKING_CACHE_MAX_AGE_MS || 12 * 60 * 60 * 1000));

// v37.15: API sempre responde pelo cache persistente primeiro quando possível.
// Quando a linha do Supabase passa do intervalo configurado, o frontend dispara
// /api/sync-tracking-cache em background. O endpoint usa lock para impedir que
// vários usuários atualizem o Smartsheet ao mesmo tempo.
const TRACKING_CACHE_AUTO_REFRESH_AFTER_MS = Math.max(
  60 * 1000,
  Number(process.env.TRACKING_CACHE_AUTO_REFRESH_AFTER_MS || process.env.TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || 15 * 60 * 1000)
);

// v37.11: não usar timeout curto para montar a primeira base real.
// O timeout de 8,5s da v37.10 era bom para evitar travamento, mas abortava a leitura
// completa do Smartsheet antes de popular o Supabase, deixando o painel zerado.
const SMARTSHEET_FETCH_TIMEOUT_MS = Math.max(2500, Number(process.env.SMARTSHEET_FETCH_TIMEOUT_MS || 12000));
const SMARTSHEET_FULL_SHEET_TIMEOUT_MS = Math.max(
  SMARTSHEET_FETCH_TIMEOUT_MS,
  Number(process.env.SMARTSHEET_FULL_SHEET_TIMEOUT_MS || process.env.SMARTSHEET_FIRST_SYNC_TIMEOUT_MS || 28000)
);
const SMARTSHEET_ROWS_FETCH_TIMEOUT_MS = Math.max(
  SMARTSHEET_FETCH_TIMEOUT_MS,
  Number(process.env.SMARTSHEET_ROWS_FETCH_TIMEOUT_MS || 20000)
);
const WIP_PO_FETCH_TIMEOUT_MS = Math.max(1500, Number(process.env.WIP_PO_FETCH_TIMEOUT_MS || 6500));
const SMARTSHEET_FETCH_RETRIES = Math.max(1, Number(process.env.SMARTSHEET_FETCH_RETRIES || 1));

// Fallback empacotado antigo NÃO pode alimentar a operação real por padrão.
// Ele deixava o login rápido, mas podia exibir base de 08/04/2026 com rowIds antigos,
// quebrando etapas, detalhe da BSP e imagens. Só habilitar manualmente para teste/dev.
const ALLOW_BUNDLED_FALLBACK = String(process.env.ALLOW_BUNDLED_FALLBACK || process.env.STEP_ALLOW_BUNDLED_FALLBACK || '0') === '1';
const REQUIRE_FRESH_WHEN_CACHE_EMPTY = String(process.env.REQUIRE_FRESH_WHEN_CACHE_EMPTY || '1') !== '0';

// v37.17: login/leitura normal não consulta mais Smartsheet.
// A base operacional deve vir do Supabase/cache. Smartsheet só é chamado por:
// - botão Atualizar, via /api/sync-tracking-cache?force=1&manual=1;
// - rotina agendada Netlify scheduled-sync-tracking-cache a cada 15 minutos;
// - opção explícita de emergência PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS=1.
const PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS = String(process.env.PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS || '0') === '1';
const PERSISTENT_CACHE_LOGIN_MAX_AGE_MS = Math.max(
  PERSISTENT_CACHE_MAX_AGE_MS,
  Number(process.env.PERSISTENT_TRACKING_CACHE_LOGIN_MAX_AGE_MS || 7 * 24 * 60 * 60 * 1000)
);

const smartsheetColumnsCache = global.__STEP_SMARTSHEET_COLUMNS_CACHE__ || {};
global.__STEP_SMARTSHEET_COLUMNS_CACHE__ = smartsheetColumnsCache;
const smartsheetIndexCache = global.__STEP_SMARTSHEET_INDEX_CACHE__ || {};
global.__STEP_SMARTSHEET_INDEX_CACHE__ = smartsheetIndexCache;
const wipPoMemoryCache = global.__STEP_WIP_PO_MEMORY_CACHE__ || {};
global.__STEP_WIP_PO_MEMORY_CACHE__ = wipPoMemoryCache;

const STAGE_ORDER = [
  { key: "Drawing Execution Advance%", label: "AG. Emissão de detalhamento", type: "percent" },
  { key: "Procuremnt Status %", label: "Verificando estoque", type: "percent" },
  { key: "Material Separation", label: "Separação de material", type: "percent" },
  { key: "Material Release to Fabrication", label: "Verificando estoque", type: "percent" },
  { key: "Fabrication Start Date", label: "Corte e Limpeza", type: "date" },
  { key: "Withdrew Material", label: "Withdrew Material", type: "percent", ignoredCurrentStage: true },
  { key: "Welding Preparation", label: "Pré - Montagem", type: "percent" },
  { key: "Spool Assemble and tack weld", label: "Pré - Montagem", type: "percent" },
  { key: "Boilermaker Finish Date", label: "Boilermaker Finish Date", type: "date", ignoredCurrentStage: true },
  { key: "Initial Dimensional Inspection/3D", label: "Inspeção Dimensional de Ajuste - 3D", type: "percent" },
  { key: "Full welding execution", label: "Solda", type: "percent" },
  { key: "Welding Finish Date", label: "Welding Finish Date", type: "date", optional: true },
  { key: "Final Dimensional Inpection/3D (QC)", label: "Inspeção Dimensional Final - 3D", type: "percent" },
  { key: "Non Destructive Examination (QC)", label: "Aguardando END", type: "percent", optional: true },
  { key: "Inspection Finish Date (QC)", label: "Inspection Finish Date (QC)", type: "date", optional: true },
  { key: "Hydro Test Pressure (QC)", label: "TH", type: "percent" },
  { key: "TH Finish Date", label: "TH Finish Date", type: "date", optional: true },
  { key: "HDG / FBE.  (PAINT)", label: "HDG / FBE. (PAINT)", type: "percent", optional: true },
  { key: "HDG / FBE DATE SAIDA (PAINT)", label: "HDG / FBE DATE SAIDA (PAINT)", type: "date", optional: true },
  { key: "HDG / FBE DATE RETORNO (PAINT)", label: "HDG / FBE DATE RETORNO (PAINT)", type: "date", optional: true },
  { key: "Surface preparation and/or coating", label: "Pintura", type: "percent" },
  { key: "Coating Finish Date", label: "Coating Finish Date", type: "date", optional: true },
  { key: "Final Inspection", label: "Unitização e Inspeção", type: "percent" },
  { key: "Package and Delivered", label: "Preparado para envio", type: "percent" },
  { key: "Project Finish Date", label: "Finalizado", type: "date" },
  { key: "Project Finished?", label: "Finalizado", type: "boolean" },
];


// v37.12-safe: aliases do Ponto Focal do cliente.
// Mantido fora do fluxo pesado para não alterar a performance da v37.11.
const CLIENT_FOCAL_POINT_ALIASES = [
  'Client Focal Point*',
  'Client Focal Point',
  'Client Focal Point *',
  'Focal Point',
  'Focal Point Cliente',
  'Ponto Focal Cliente',
  'Ponto Focal',
  'Client Responsible',
  'Client Representative',
  'Client Contact',
  'Responsible Client',
  'Contato Cliente',
  'Contato do Cliente',
  'Responsável Cliente',
  'Responsavel Cliente',
  'Cliente Responsável',
  'Cliente Responsavel',
];

const TRACKING_REQUIRED_COLUMN_TITLES = [
  'Project',
  'Primary',
  'PRIMARY',
  'Item',
  'Quantity Spools',
  'Drawing',
  'Project Type',
  'Client',
  'Vessel',
  'Priority',
  'PM',
  'Class',
  'Line Nº',
  'Line No',
  'Line Number',
  'Size',
  'OBSERVATIONS',
  'Kilos',
  'Peso (KG)',
  'Peso Soldado (KG)',
  'Quantity Juntas',
  'Quant. Juntas',
  'M2 Painting',
  'Start Date',
  'Finish Date',
  'Fabrication Start Date',
  'Project Finish Date',
  'Project Finished?',
  'Project Status',
  'PROJECT STATUS',
  'Overall Project Status',
  'Status',
  '% Overall Progress',
  '% Individual Progress',
  'Job Process Status',
  ...STAGE_ORDER.map((stage) => stage.key),
  ...CLIENT_FOCAL_POINT_ALIASES,
];

const TRACKING_REQUIRED_COLUMN_INCLUDES = [
  'Line',
  'Linha',
  'Quantity Juntas',
  'Quant Juntas',
  'Peso Soldado',
  'M2 Painting',
  'Overall Progress',
  'Individual Progress',
  'Project Status',
  'Client Focal Point',
  'Ponto Focal',
];

function normalizeColumnTitle(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "")
    .toLowerCase();
}

function getCellValue(row, key) {
  if (!row?.values) return { raw: null, display: null };
  if (row.values[key]) return row.values[key];

  const target = normalizeColumnTitle(key);
  if (!target) return { raw: null, display: null };

  for (const [title, cell] of Object.entries(row.values)) {
    if (normalizeColumnTitle(title) === target) {
      return cell || { raw: null, display: null };
    }
  }

  return { raw: null, display: null };
}

function textValue(row, key) {
  const cell = getCellValue(row, key);
  const value = cell.display ?? cell.raw;
  return value == null ? "" : String(value).trim();
}

function textValueAny(row, keys = []) {
  for (const key of Array.isArray(keys) ? keys : []) {
    const value = textValue(row, key);
    if (value && value !== 'N/A' && value !== '—') return value;
  }
  return '';
}

function parseNumberValue(input) {
  if (input == null || input === "") return null;
  if (typeof input === "number") return Number.isFinite(input) ? input : null;

  let str = String(input).trim();
  if (!str) return null;
  str = str.replace(/\s/g, "");

  const hasComma = str.includes(",");
  const hasDot = str.includes(".");

  if (hasComma && hasDot) {
    if (str.lastIndexOf(",") > str.lastIndexOf(".")) {
      str = str.replace(/\./g, "").replace(",", ".");
    } else {
      str = str.replace(/,/g, "");
    }
  } else if (hasComma) {
    str = str.replace(",", ".");
  }

  str = str.replace(/[^\d.-]/g, "");
  const num = Number(str);
  return Number.isFinite(num) ? num : null;
}

function parseNumber(row, key) {
  const cell = getCellValue(row, key);
  return parseNumberValue(cell.raw ?? cell.display);
}

function parsePercent(row, key) {
  const cell = getCellValue(row, key);
  const display = cell.display ?? "";
  const raw = cell.raw;

  if (typeof display === "string" && display.includes("%")) {
    const value = parseNumberValue(display.replace("%", ""));
    return value == null ? null : value;
  }

  const parsed = parseNumberValue(raw ?? display);
  if (parsed == null) return null;
  if (parsed >= 0 && parsed <= 1) return parsed * 100;
  return parsed;
}

function isTruthyValue(value) {
  if (value == null) return false;
  if (typeof value === "boolean") return value;
  const normalized = String(value).trim().toLowerCase();
  return ["true", "yes", "sim", "y", "1", "concluído", "concluido", "finalizado"].includes(normalized);
}

function excelSerialToDate(serial) {
  if (!Number.isFinite(serial)) return null;
  if (serial < 1 || serial > 90000) return null;
  const excelEpoch = Date.UTC(1899, 11, 30);
  const millis = excelEpoch + Math.round(serial) * 86400000;
  const date = new Date(millis);
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
}

function formatDateValue(value) {
  const parsedDate = parseDateObject(value);
  if (parsedDate) {
    return parsedDate.toLocaleDateString("pt-BR", { timeZone: "UTC" });
  }

  if (!value) return "";
  const raw = String(value).trim();
  if (!raw) return "";
  return raw;
}

function parseDateObject(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  }

  if (typeof value === "number") {
    return excelSerialToDate(value);
  }

  const raw = String(value).trim();
  if (!raw) return null;

  if (/^\d+(?:\.\d+)?$/.test(raw)) {
    const numericDate = excelSerialToDate(Number(raw));
    if (numericDate) return numericDate;
  }

  let match = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }

  match = raw.match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]) - 1;
    const year = Number(match[3]) >= 70 ? 1900 + Number(match[3]) : 2000 + Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }

  match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]) - 1;
    const day = Number(match[3]);
    return new Date(Date.UTC(year, month, day));
  }

  const date = new Date(raw);
  if (!Number.isNaN(date.getTime())) {
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  }

  return null;
}

function getWeekAnchor(year) {
  const jan1 = new Date(Date.UTC(year, 0, 1));
  const anchor = new Date(jan1);
  anchor.setUTCDate(jan1.getUTCDate() - jan1.getUTCDay());
  return anchor;
}

function getCurrentBrazilDateObject() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Sao_Paulo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const year = Number(parts.find((item) => item.type === "year")?.value || new Date().getUTCFullYear());
  const month = Number(parts.find((item) => item.type === "month")?.value || 1);
  const day = Number(parts.find((item) => item.type === "day")?.value || 1);
  return new Date(Date.UTC(year, month - 1, day));
}

function getCurrentBrazilYear() {
  return getCurrentBrazilDateObject().getUTCFullYear();
}

function formatProductionWeekLabel(weekNumber, weekYear) {
  return `Semana ${weekNumber} - ${weekYear}`;
}

function getProductionWeekLabel(value) {
  const date = parseDateObject(value);
  if (!date) return "";

  let weekYear = date.getUTCFullYear();
  const nextAnchor = getWeekAnchor(weekYear + 1);
  if (date >= nextAnchor) {
    weekYear += 1;
  } else {
    const currentAnchor = getWeekAnchor(weekYear);
    if (date < currentAnchor) weekYear -= 1;
  }

  const anchor = getWeekAnchor(weekYear);
  const diffDays = Math.floor((date - anchor) / 86400000);
  const weekNumber = Math.floor(diffDays / 7) + 1;
  return formatProductionWeekLabel(weekNumber, weekYear);
}

function hasDateValue(row, key) {
  const value = textValue(row, key);
  return Boolean(value && String(value).trim());
}

function isAwaitingShipment(row) {
  const coatingPercent = parsePercent(row, "Surface preparation and/or coating") ?? 0;
  const coatingDone = coatingPercent >= 100;
  const packageDelivered = parsePercent(row, "Package and Delivered") ?? 0;
  const projectFinished = isTruthyValue(textValue(row, "Project Finished?") || getCellValue(row, "Project Finished?").raw);
  return coatingDone && packageDelivered < 100 && !projectFinished;
}


function numberFromStageValue(stageValues, key) {
  const value = stageValues?.[key];
  if (value == null || value === "" || value === "N/A") return null;
  const parsed = parseNumberValue(value);
  return parsed == null ? null : parsed;
}

function hasStageValue(stageValues, key) {
  const value = stageValues?.[key];
  if (value == null) return false;
  const text = String(value).trim();
  return Boolean(text && text !== "N/A" && text !== "Não");
}

function isStageBooleanDone(stageValues, key) {
  return String(stageValues?.[key] || "").trim().toLowerCase() === "sim";
}

function pct(stageValues, key) {
  return numberFromStageValue(stageValues, key) ?? 0;
}

function isSpoolMaterialType(projectType, fallbackText = '') {
  const normalize = (value) => String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  const typeText = normalize(projectType).trim();
  if (typeText) {
    if (typeText.includes('spool')) return true;
    if (typeText.includes('support') || typeText.includes('suporte') || typeText === 'sup' || typeText.includes('structure') || typeText.includes('estrutura') || typeText.includes('frame')) return false;
  }

  const evidence = normalize(fallbackText);
  if (/\bspl\b/.test(evidence) || evidence.includes('spool')) return true;
  if (/\bsup\b/.test(evidence) || evidence.includes('support') || evidence.includes('suporte') || evidence.includes('structure') || evidence.includes('estrutura') || evidence.includes('frame')) return false;
  return false;
}

function fabricationStageItemsForType(projectType, fallbackText = '') {
  const includeHydro = isSpoolMaterialType(projectType, fallbackText);
  return [
    { keys: ["Welding Preparation", "Spool Assemble and tack weld"], weight: 10 },
    { keys: ["Initial Dimensional Inspection/3D"], weight: 8 },
    { keys: ["Full welding execution"], weight: 25 },
    { keys: ["Non Destructive Examination (QC)"], weight: 12 },
    { keys: ["Final Dimensional Inpection/3D (QC)"], weight: 8 },
    { keys: ["Hydro Test Pressure (QC)"], weight: 7, spoolOnly: true },
    { keys: ["Surface preparation and/or coating", "HDG / FBE.  (PAINT)"], weight: 15 },
  ].filter((item) => !item.spoolOnly || includeHydro);
}

const PRODUCTION_STAGE_EVIDENCE_KEYS = [
  "Drawing Execution Advance%",
  "Drawing",
  "Procuremnt Status %",
  "Procurement Status %",
  "Procurement",
  "Material Separation",
  "Material Release to Fabrication",
  "Welding Preparation",
  "Spool Assemble and tack weld",
  "Initial Dimensional Inspection/3D",
  "Full welding execution",
  "Non Destructive Examination (QC)",
  "Final Dimensional Inpection/3D (QC)",
  "Hydro Test Pressure (QC)",
  "Surface preparation and/or coating",
  "HDG / FBE.  (PAINT)",
  "Final Inspection",
  "Package and Delivered",
];

function hasStageProgressEvidence(stageValues, keys = PRODUCTION_STAGE_EVIDENCE_KEYS) {
  return keys.some((key) => {
    const value = stageValues?.[key];
    return value != null && value !== "" && value !== "N/A" && value !== "Não";
  });
}

function stageEvidenceValue(stageValues, keys) {
  for (const key of keys) {
    const value = stageValues?.[key];
    if (value != null && value !== "" && value !== "N/A" && value !== "Não") {
      return { hasEvidence: true, percent: pct(stageValues, key) };
    }
  }
  return { hasEvidence: false, percent: 0 };
}

function fabricationProgressFromStageValues(stageValues, projectType = '', fallbackText = '') {
  const painting = Math.max(pct(stageValues, "Surface preparation and/or coating"), pct(stageValues, "HDG / FBE.  (PAINT)"));
  if (painting >= 99.9) return 100;
  const stages = fabricationStageItemsForType(projectType, fallbackText);
  const totalWeight = stages.reduce((sum, item) => sum + item.weight, 0) || 100;
  return Math.max(0, Math.min(100, stages.reduce((sum, item) => {
    const value = item.keys.reduce((max, key) => Math.max(max, pct(stageValues, key)), 0);
    return sum + value * item.weight;
  }, 0) / totalWeight));
}

function productionStageSnapshotsFromValues(stageValues, projectType = '', fallbackText = '') {
  const engineering = stageEvidenceValue(stageValues, ["Drawing Execution Advance%", "Drawing"]);
  const procurementCandidates = [
    stageEvidenceValue(stageValues, ["Procuremnt Status %", "Procurement Status %", "Procurement"]),
    stageEvidenceValue(stageValues, ["Material Separation"]),
    stageEvidenceValue(stageValues, ["Material Release to Fabrication"]),
  ];
  const procurement = procurementCandidates.reduce((best, item) => {
    if (!item.hasEvidence) return best;
    if (!best.hasEvidence || item.percent > best.percent) return item;
    return best;
  }, { hasEvidence: false, percent: 0 });
  const fabricationEvidenceKeys = fabricationStageItemsForType(projectType, fallbackText).flatMap((item) => item.keys);
  const fabrication = {
    hasEvidence: hasStageProgressEvidence(stageValues, fabricationEvidenceKeys),
    percent: fabricationProgressFromStageValues(stageValues, projectType, fallbackText),
  };
  const packageDelivery = stageEvidenceValue(stageValues, ["Package and Delivered", "Final Inspection"]);
  return [
    { key: "engineering", percent: engineering.percent, weight: 15, hasEvidence: engineering.hasEvidence },
    { key: "procurement", percent: procurement.percent, weight: 15, hasEvidence: procurement.hasEvidence },
    { key: "fabrication", percent: fabrication.percent, weight: 65, hasEvidence: fabrication.hasEvidence },
    { key: "package", percent: packageDelivery.percent, weight: 5, hasEvidence: packageDelivery.hasEvidence },
  ];
}

function weightedOverallFromStageValues(stageValues, projectType = '', fallbackText = '') {
  const stages = productionStageSnapshotsFromValues(stageValues, projectType, fallbackText);
  const totalWeight = stages.reduce((sum, stage) => sum + stage.weight, 0) || 100;
  return Math.max(0, Math.min(100, stages.reduce((sum, stage) => sum + stage.percent * stage.weight, 0) / totalWeight));
}

function mergeProjectStageValuesFromSpools(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (!project || !spools.length) return project?.stageValues || {};

  const merged = { ...(project.stageValues || {}) };
  for (const stage of STAGE_ORDER) {
    if (stage.type !== 'percent') continue;
    const current = numberFromStageValue(merged, stage.key);
    if (current != null && current > 0) continue;

    let totalWeight = 0;
    let weightedSum = 0;
    for (const spool of spools) {
      const value = numberFromStageValue(spool?.stageValues, stage.key);
      if (value == null) continue;
      const weight = Number(spool?.kilos || 0) > 0 ? Number(spool.kilos) : 1;
      totalWeight += weight;
      weightedSum += value * weight;
    }
    if (totalWeight > 0) merged[stage.key] = Math.max(0, Math.min(100, weightedSum / totalWeight));
  }
  project.stageValues = merged;
  return merged;
}

function hasIncompleteProductionEvidence(stageValues, projectType = '', fallbackText = '') {
  return productionStageSnapshotsFromValues(stageValues, projectType, fallbackText).some((stage) => stage.hasEvidence && Number(stage.percent || 0) < 99.9);
}

function isSpoolFinishedByState(spool) {
  if (!spool || hasIncompleteProductionEvidence(spool.stageValues, spool.projectType, [spool.iso, spool.drawing, spool.description].filter(Boolean).join(' '))) return false;
  return Boolean(
    spool.finished
    || spool.projectFinishedFlag
    || spool.uiState === "completed"
    || spool.operationalState === "completed"
    || spool.flow?.state === "completed"
    || spool.flow?.status === "Finalizado"
    || Number(spool.overallProgress || 0) >= 99.9
  );
}

function paintingStatusFromPercent(value) {
  const percent = Number(value || 0);
  if (percent >= 100) return "Concluído";
  if (percent >= 90) return "Acabamento";
  if (percent >= 75) return "Intermediária";
  if (percent >= 50) return "J/F";
  if (percent >= 25) return "Aguardando início de pintura";
  if (percent > 0) return "Aguardando início de pintura";
  return "Pintura";
}

function makeFlow(status, sector, percent = 0, stageStatus = null, state = null) {
  const normalizedPercent = Number.isFinite(Number(percent)) ? Number(percent) : 0;
  const statusType = stageStatus || (normalizedPercent >= 100 ? "completed" : normalizedPercent > 0 ? "in_progress" : "waiting");
  const normalizedSector = String(sector || "Geral");
  let flowState = state;
  if (!flowState) {
    if (status === "Finalizado") flowState = "completed";
    else if (normalizedSector === "Logística" && ["Preparado para envio", "Aguardando envio"].includes(status)) flowState = "awaiting_shipment";
    else if (["Qualidade"].includes(normalizedSector)) flowState = "in_inspection";
    else if (["Produção", "Pintura", "Engenharia", "Suprimento"].includes(normalizedSector)) flowState = "in_production";
    else flowState = "not_started";
  }
  return { state: flowState, sector: normalizedSector, status, percent: normalizedPercent, stageStatus: statusType };
}

function deriveOperationalStage(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus, projectType = '', fallbackText = '') {
  const drawing = pct(stageValues, "Drawing Execution Advance%");
  const procurement = Math.max(pct(stageValues, "Procuremnt Status %"), pct(stageValues, "Material Release to Fabrication"));
  const materialSeparation = pct(stageValues, "Material Separation");
  const fabricationStarted = Boolean(fabricationStartDate || hasStageValue(stageValues, "Fabrication Start Date"));
  const withdrewMaterial = pct(stageValues, "Withdrew Material");
  const weldingPreparation = pct(stageValues, "Welding Preparation");
  const spoolAssemble = pct(stageValues, "Spool Assemble and tack weld");
  const boilermakerDone = hasStageValue(stageValues, "Boilermaker Finish Date");
  const dma3d = pct(stageValues, "Initial Dimensional Inspection/3D");
  const fullWelding = pct(stageValues, "Full welding execution");
  const finalDimensional = pct(stageValues, "Final Dimensional Inpection/3D (QC)");
  const nde = numberFromStageValue(stageValues, "Non Destructive Examination (QC)");
  const th = pct(stageValues, "Hydro Test Pressure (QC)");
  const coating = Number.isFinite(Number(coatingPercent)) ? Number(coatingPercent) : pct(stageValues, "Surface preparation and/or coating");
  const finalInspection = pct(stageValues, "Final Inspection");
  const packageDelivered = pct(stageValues, "Package and Delivered");
  const projectFinishDate = hasStageValue(stageValues, "Project Finish Date");
  const projectFinished = isStageBooleanDone(stageValues, "Project Finished?");
  const normalizedProjectStatus = String(projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
  const isHold = ["ON HOLD", "HOLD", "PAUSED", "EM ESPERA"].includes(normalizedProjectStatus);
  const includeHydro = isSpoolMaterialType(projectType, fallbackText);

  if (finished || projectFinished || projectFinishDate) return makeFlow("Finalizado", "Enviado", 100, "completed", "completed");

  const coatingCompleted = coating >= 100;
  const finalInspectionStarted = finalInspection > 0;
  const finalInspectionCompleted = finalInspection >= 100;

  if (coatingCompleted && finalInspectionCompleted && packageDelivered >= 100) {
    return makeFlow("Preparado para envio", "Logística", packageDelivered, "completed", "awaiting_shipment");
  }
  if (coatingCompleted && finalInspectionCompleted && packageDelivered >= 25) {
    return makeFlow("Aguardando envio", "Logística", packageDelivered, null, "awaiting_shipment");
  }
  if (coatingCompleted && finalInspectionCompleted && packageDelivered > 0) {
    return makeFlow("Preparado para envio", "Logística", packageDelivered, null, "awaiting_shipment");
  }
  if (coatingCompleted && finalInspectionCompleted) {
    return makeFlow("Preparado para envio", "Logística", finalInspection, "waiting", "awaiting_shipment");
  }
  if (coatingCompleted && finalInspectionStarted) {
    return makeFlow("Unitização e Inspeção", "Logística", finalInspection, null, "preparing_shipment");
  }
  if (coatingCompleted) {
    return makeFlow("Unitização e Inspeção", "Logística", coating, "waiting", "awaiting_shipment");
  }
  if (coating > 0) {
    return makeFlow(paintingStatusFromPercent(coating), "Pintura", coating, null, "in_production");
  }
  if (includeHydro && th >= 100) return makeFlow("Pintura", "Pintura", 0, "waiting", "in_production");
  if (includeHydro && th > 0) return makeFlow("TH", "Qualidade", th, null, "in_inspection");
  if (nde != null && nde > 0 && nde < 100) return makeFlow("Aguardando END", "Qualidade", nde, null, "in_inspection");
  if (finalDimensional >= 100) return includeHydro ? makeFlow("TH", "Qualidade", 0, "waiting", "in_inspection") : makeFlow("Pintura", "Pintura", 0, "waiting", "in_production");
  if (finalDimensional > 0) return makeFlow("Inspeção Dimensional Final - 3D", "Qualidade", finalDimensional, null, "in_inspection");
  if (fullWelding >= 100) return makeFlow("Inspeção Dimensional Final - 3D", "Qualidade", 0, "waiting", "in_inspection");
  if (fullWelding > 0) return makeFlow("Solda", "Produção", fullWelding, null, "in_production");
  if (dma3d >= 100) return makeFlow("Solda", "Produção", 0, "waiting", "in_production");
  if (dma3d > 0) return makeFlow("Inspeção Dimensional de Ajuste - 3D", "Qualidade", dma3d, null, "in_inspection");
  if (boilermakerDone || spoolAssemble >= 100) return makeFlow("Inspeção Dimensional de Ajuste - 3D", "Qualidade", 0, "waiting", "in_inspection");
  if (spoolAssemble > 0) return makeFlow("Pré - Montagem", "Produção", spoolAssemble, null, "in_production");
  if (weldingPreparation >= 100) return makeFlow("Pré - Montagem", "Produção", weldingPreparation, "in_progress", "in_production");
  if (weldingPreparation > 0 || withdrewMaterial > 0) return makeFlow("Pré - Montagem", "Produção", Math.max(weldingPreparation, withdrewMaterial), null, "in_production");
  if (fabricationStarted) return makeFlow("Corte e Limpeza", "Produção", 0, "in_progress", "in_production");
  if (materialSeparation >= 100) return makeFlow("Corte e Limpeza", "Produção", 0, "waiting", "in_production");
  if (materialSeparation > 0) return makeFlow("Separação de material", "Suprimento", materialSeparation, null, "in_production");
  if (procurement >= 100) return makeFlow("Separação de material", "Suprimento", 0, "waiting", "in_production");
  if (procurement > 0) return makeFlow("Verificando estoque", "Suprimento", procurement, null, "in_production");
  if (drawing >= 100) return makeFlow("Verificando estoque", "Suprimento", 0, "waiting", "in_production");
  if (drawing > 0) return makeFlow("AG. Emissão de detalhamento", "Engenharia", drawing, null, "in_production");
  if (isHold) return makeFlow("AG. Emissão de detalhamento", "Engenharia", 0, "waiting", "not_started");
  return makeFlow("AG. Emissão de detalhamento", "Engenharia", 0, "waiting", "not_started");
}

function normalizeFlowSortText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function getFlowSortWeight(flow) {
  const status = normalizeFlowSortText(flow?.status || "");
  const sector = normalizeFlowSortText(flow?.sector || "");

  if (status === "finalizado" || status.includes("project finished")) return 999;
  if (status.includes("package and delivered") || status === "delivered" || status === "enviado") return 130;
  if (status.includes("ag emissao de detalhamento") || status === "emissao de detalhamento") return 10;
  if (status.includes("verificando estoque") || status.includes("aguardando material")) return 20;
  if (status.includes("separacao de material") || status.includes("material separation")) return 30;
  if (status.includes("corte e limpeza") || status.includes("fabrication start")) return 40;
  if (status.includes("pre montagem") || status.includes("welding preparation")) return 50;
  if (status.includes("inspecao dimensional de ajuste") || status.includes("dma 3d")) return 60;
  if (status === "solda" || status.includes("full welding")) return 70;
  if (status.includes("inspecao dimensional final") || status.includes("inspection qc")) return 80;
  if (status.includes("aguardando end")) return 90;
  if (status === "th") return 100;
  if (status === "pintura") return 110;
  if (status.includes("aguardando inicio de pintura")) return 111;
  if (status === "j f" || status === "jf") return 112;
  if (status.includes("intermediaria")) return 113;
  if (status.includes("acabamento")) return 114;
  if (status === "concluido") return 115;
  if (status.includes("unitizacao e inspecao") || status.includes("unitizacao")) return 120;
  if (status.includes("preparado para envio") || status.includes("preparando para envio") || status.includes("aguardando envio")) return 130;

  if (sector === "engenharia") return 10;
  if (sector === "suprimento") return 20;
  if (sector === "producao") return 40;
  if (sector === "qualidade") return 80;
  if (sector === "pintura") return 110;
  if (sector === "logistica") return 120;
  return 500;
}

function summarizeFlowItems(items, fallbackFlow, fallbackQuantity = 1) {
  const source = Array.isArray(items) && items.length
    ? items.map((item) => ({
        ...item,
        flow: item.flow || {
          status: item.stage || item.currentStage || item.status || fallbackFlow?.status || "—",
          sector: item.operationalSector || fallbackFlow?.sector || "Geral",
          state: item.operationalState || item.uiState || fallbackFlow?.state || "not_started",
          percent: item.stagePercent || fallbackFlow?.percent || 0,
          stageStatus: item.stageStatus || fallbackFlow?.stageStatus || "waiting",
        },
        quantity: 1,
      }))
    : [{ flow: fallbackFlow || makeFlow("AG. Emissão de detalhamento", "Engenharia"), quantity: Number(fallbackQuantity || 1) }];

  const openItems = source.filter((item) => String(item.flow?.status || "") !== "Finalizado" && item.flow?.state !== "completed");
  const active = openItems.length ? openItems : source;
  const sortedActive = [...active].sort((a, b) => getFlowSortWeight(a.flow) - getFlowSortWeight(b.flow));
  const primary = sortedActive[0]?.flow || fallbackFlow || makeFlow("AG. Emissão de detalhamento", "Engenharia");
  const byStatus = new Map();
  const bySector = new Map();
  for (const item of source) {
    const flow = item.flow || primary;
    const quantity = Number(item.quantity || 1);
    const statusKey = flow.status || "—";
    const sectorKey = flow.sector || "Geral";
    byStatus.set(statusKey, (byStatus.get(statusKey) || 0) + quantity);
    bySector.set(sectorKey, (bySector.get(sectorKey) || 0) + quantity);
  }
  const statusBreakdown = Array.from(byStatus, ([label, count]) => ({ label, count })).sort((a, b) => getFlowSortWeight({ status: a.label }) - getFlowSortWeight({ status: b.label }));
  const sectorBreakdown = Array.from(bySector, ([label, count]) => ({ label, count })).sort((a, b) => String(a.label).localeCompare(String(b.label), "pt-BR"));
  const activeStatusBreakdown = statusBreakdown.filter((item) => item.label !== "Finalizado");
  const activeSectorBreakdown = sectorBreakdown.filter((item) => item.label !== "Logística" || active.some((sourceItem) => sourceItem.flow?.sector === "Logística"));
  const allFinished = source.length > 0 && source.every((item) => String(item.flow?.status || "") === "Finalizado" || item.flow?.state === "completed");
  const formatBreakdown = (rows, fallbackLabel) => {
    const clean = rows.filter((row) => row && row.label && Number(row.count || 0) > 0);
    if (!clean.length) return fallbackLabel || "—";
    if (clean.length === 1) return clean[0].label;
    return clean.map((row) => `${row.label}: ${row.count}`).join(" • ");
  };
  const flow = allFinished
    ? makeFlow("Finalizado", "Enviado", 100, "completed", "completed")
    : primary;
  return {
    flow,
    allFinished,
    statusSummary: allFinished ? "Finalizado" : formatBreakdown(activeStatusBreakdown, primary.status),
    sectorSummary: allFinished ? "Enviado" : formatBreakdown(activeSectorBreakdown.filter((row) => row.label !== "Logística" || primary.sector === "Logística"), primary.sector),
    statusBreakdown,
    sectorBreakdown,
  };
}

function parseProjectParts(projectText) {
  const cleaned = String(projectText || "").trim().replace(/\s+/g, " ");
  if (!cleaned) return { prefix: "", number: "", display: "" };

  const match = cleaned.match(/^(?:([A-Z]{2,5})[\s-]+)?(\d{2}-\d+(?:-\d+)*(?:-[A-Z0-9]+)?)$/i);
  if (match) {
    return { prefix: (match[1] || "").toUpperCase(), number: match[2], display: cleaned };
  }

  const loose = cleaned.match(/([A-Z]{2,5})?[\s-]*(\d{2}-\d+(?:-\d+)*(?:-[A-Z0-9]+)?)/i);
  if (loose) {
    const prefix = loose[1] ? loose[1].toUpperCase() : "";
    const number = loose[2];
    return { prefix, number, display: prefix ? `${prefix} ${number}` : number };
  }

  return { prefix: "", number: cleaned, display: cleaned };
}

function normalizeBspLookupKey(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';

  // Captura somente o código da BSP/BPP/B3D, sem deixar a PO entrar na chave.
  // Exemplos válidos:
  // BSP 25-732-03, 25-732-03, BPP 25-732-03, B3D 25-732-03.
  const match = raw.match(/(?:\b(?:BSP|BPP|B3D)\b\s*)?(\d{2}\s*-?\s*\d+(?:\s*-?\s*\d+)*)/i);
  const source = match ? match[1] : raw;
  return String(source || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\b(BSP|BPP|B3D)\b/gi, '')
    .replace(/[^0-9]+/g, '')
    .trim();
}

function getProjectBspLookupKeys(project) {
  const candidates = [
    project?.projectDisplay,
    project?.projectNumber,
    project?.projectCode,
    project?.project,
    project?.rawProject,
    project?.stageValues?.['Project BSP/BPP/B3D*'],
    project?.stageValues?.['Project BSP/BPP/B3D'],
  ].filter(Boolean);

  const keys = [];
  for (const candidate of candidates) {
    const key = normalizeBspLookupKey(candidate);
    if (key && !keys.includes(key)) keys.push(key);
  }
  return keys;
}

function getProjectPoDisplay(project) {
  const list = Array.isArray(project?.customerPoList) ? project.customerPoList.filter(Boolean) : [];
  if (!list.length) return 'Aguardando PO';
  if (list.length === 1) return `PO ${list[0]}`;
  return `POs ${list.join(' / ')}`;
}

function buildClientDisplayCode(project) {
  const bsp = String(project?.projectDisplay || project?.projectNumber || 'BSP').trim() || 'BSP';
  return `${bsp} - ${getProjectPoDisplay(project)}`;
}

function extractIsoDescription(drawingText) {
  const text = String(drawingText || "").trim();
  if (!text) return { iso: "", description: "" };
  const match = text.match(/^(.*?)\s*\((.*?)\)\s*$/);
  if (match) return { iso: match[1].trim(), description: match[2].trim() };
  return { iso: text, description: "" };
}

function stageStatusFromPercent(percent) {
  if (percent == null) return "ignored";
  if (percent >= 100) return "completed";
  if (percent > 0) return "in_progress";
  return "waiting";
}

const HDG_FBE_PAINT_PROGRESS_KEY = "HDG / FBE.  (PAINT)";
const HDG_FBE_PAINT_EXIT_KEY = "HDG / FBE DATE SAIDA (PAINT)";
const HDG_FBE_PAINT_RETURN_KEY = "HDG / FBE DATE RETORNO (PAINT)";

function isNotApplicableValue(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "n/a" || normalized === "na";
}

function getHdgFbePaintState(row) {
  const rawText = textValue(row, HDG_FBE_PAINT_PROGRESS_KEY);
  const percent = parsePercent(row, HDG_FBE_PAINT_PROGRESS_KEY);
  const ignored = isNotApplicableValue(rawText);
  const hasProgress = percent != null;
  const active = !ignored && (hasProgress || Boolean(rawText));
  return {
    rawText,
    percent,
    ignored,
    active,
    completed: active && hasProgress && percent >= 100,
    inProgress: active && hasProgress && percent > 0 && percent < 100,
  };
}

function shouldIgnoreOptionalPaintStage(row, stageKey) {
  if (![HDG_FBE_PAINT_PROGRESS_KEY, HDG_FBE_PAINT_EXIT_KEY, HDG_FBE_PAINT_RETURN_KEY].includes(stageKey)) return false;
  const paintState = getHdgFbePaintState(row);
  return paintState.ignored;
}

function buildStageValues(row) {
  const stageValues = {};
  const paintState = getHdgFbePaintState(row);
  for (const stage of STAGE_ORDER) {
    if (shouldIgnoreOptionalPaintStage(row, stage.key)) {
      stageValues[stage.key] = stage.type === "date" ? "" : "N/A";
      continue;
    }

    if (stage.type === "percent") {
      if (stage.key === HDG_FBE_PAINT_PROGRESS_KEY && paintState.ignored) {
        stageValues[stage.key] = "N/A";
        continue;
      }
      const value = parsePercent(row, stage.key);
      stageValues[stage.key] = value == null ? null : value;
      continue;
    }
    if (stage.type === "date") {
      const value = textValue(row, stage.key);
      stageValues[stage.key] = value ? formatDateValue(value) : "";
      continue;
    }
    if (stage.type === "boolean") {
      stageValues[stage.key] = isTruthyValue(textValue(row, stage.key) || getCellValue(row, stage.key).raw) ? "Sim" : "Não";
    }
  }
  return stageValues;
}

function deriveProgress(row) {
  const milestones = [];
  const completedStages = [];
  let currentStage = null;
  const paintState = getHdgFbePaintState(row);

  for (const stage of STAGE_ORDER) {
    if (shouldIgnoreOptionalPaintStage(row, stage.key)) {
      continue;
    }

    if (stage.type === "date") {
      const value = textValue(row, stage.key);
      if (value) {
        milestones.push({ key: stage.key, label: stage.label, value: formatDateValue(value), type: "date" });
      }
      continue;
    }

    if (stage.type === "boolean") {
      const truthy = isTruthyValue(textValue(row, stage.key) || getCellValue(row, stage.key).raw);
      milestones.push({ key: stage.key, label: stage.label, value: truthy ? "Sim" : "Não", type: "boolean" });
      if (truthy && !currentStage) {
        currentStage = { key: stage.key, label: stage.label, percent: 100, status: "completed", isAlert: false };
      }
      continue;
    }

    const percent = parsePercent(row, stage.key);
    const rawText = textValue(row, stage.key);
    if (stage.key === HDG_FBE_PAINT_PROGRESS_KEY && paintState.ignored) {
      continue;
    }

    const hasContent = percent != null || rawText;
    if (!hasContent && stage.optional) continue;
    if (!hasContent) continue;

    const status = stageStatusFromPercent(percent);
    if (status === "completed") {
      completedStages.push({ key: stage.key, label: stage.label, percent: 100, status });
      continue;
    }

    if (!currentStage) {
      currentStage = {
        key: stage.key,
        label: stage.label,
        percent: percent ?? 0,
        status,
        isAlert: status === "in_progress" || status === "waiting",
      };
    }
  }

  if (!currentStage) {
    currentStage = {
      key: "Package and Delivered",
      label: "Unitização e envio",
      percent: 100,
      status: "completed",
      isAlert: false,
    };
  }

  return { currentStage, completedStages, milestones };
}

function projectUiState(projectStatus, overallProgress, finished, fabricationStartDate, awaitingShipment = false) {
  if (finished) return "completed";
  if (awaitingShipment) return "awaiting_shipment";
  if (!fabricationStartDate && overallProgress <= 0) return "not_started";
  if (overallProgress <= 0 && /^on hold$/i.test(projectStatus || "")) return "not_started";
  if (overallProgress <= 0 && !fabricationStartDate) return "not_started";
  return "in_progress";
}

function getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus, projectType = '', fallbackText = '') {
  return deriveOperationalStage(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus, projectType, fallbackText);
}

function classifyStageSector(stageValue) {
  const stage = String(stageValue || '').toLowerCase();

  if (
    stage.includes('paint') ||
    stage.includes('coating') ||
    stage.includes('surface preparation') ||
    stage.includes('surface preparation and/or coating') ||
    stage.includes('hdg') ||
    stage.includes('fbe')
  ) {
    return 'Pintura';
  }

  if (
    stage.includes('inspection') ||
    stage.includes('nondestructive') ||
    stage.includes('non destructive') ||
    stage.includes('dimensional') ||
    stage.includes('hydro test') ||
    stage.includes('qc') ||
    stage.includes('th finish') ||
    stage.includes('final inspection')
  ) {
    return 'Inspeção';
  }

  if (
    stage.includes('welding') ||
    stage.includes('solda') ||
    stage.includes('spool assemble') ||
    stage.includes('tack weld')
  ) {
    return 'Solda';
  }

  if (
    stage.includes('boilermaker') ||
    stage.includes('caldeiraria') ||
    stage.includes('material release') ||
    stage.includes('material separation') ||
    stage.includes('withdrew material') ||
    stage.includes('drawing execution') ||
    stage.includes('procurement') ||
    stage.includes('fabrication')
  ) {
    return 'Calderaria';
  }

  return 'Geral';
}

function classifyAlertSector(project) {
  const stage = String(project?.currentStage || "").toLowerCase();
  const uiState = String(project?.uiState || project?.operationalState || "").toLowerCase();

  if (
    stage.includes("final inspection") ||
    stage.includes("unitização") ||
    stage.includes("unitizacao") ||
    stage.includes("package and delivered") ||
    stage.includes("envio") ||
    uiState === "awaiting_shipment"
  ) {
    return "Logística";
  }

  if (project?.operationalSector) return project.operationalSector;
  const stageValues = project?.stageValues || {};
  const flow = getOperationalFlow(
    stageValues,
    project?.fabricationStartDate,
    project?.coatingPercent,
    project?.finished,
    project?.projectStatus,
  );
  return flow.sector || 'Geral';
}

function buildAlertObservation(project, sector, diffDays) {
  const stageLabel = project?.currentStage || project?.jobProcessStatus || 'Etapa não identificada';
  const coatingPercent = Number(project?.coatingPercent || 0);
  const baseDaysText = diffDays < 0
    ? `O término planejado já venceu há ${Math.abs(diffDays)} dia(s).`
    : `Faltam ${diffDays} dia(s) para o término planejado.`;

  if (coatingPercent >= 100) {
    const coatingFinishDate = project?.stageValues?.["Coating Finish Date"] || project?.coatingFinishDate || "";
    const coatingFinishedText = coatingFinishDate
      ? ` A pintura já está em 100%, finalizada em ${coatingFinishDate}. Conferir envio.`
      : ' A pintura já está em 100%. Conferir envio.';
    return {
      title: diffDays < 0 ? 'Conferência em atraso' : 'Conferência pendente',
      message: `${baseDaysText}${coatingFinishedText}`,
    };
  }

  if (sector === 'Calderaria') {
    return {
      title: diffDays < 0 ? 'Calderaria em atraso' : 'Calderaria em atenção',
      message: `${baseDaysText} O projeto ainda está na Calderaria.`,
    };
  }

  if (sector === 'Solda') {
    return {
      title: diffDays < 0 ? 'Solda em atraso' : 'Solda em atenção',
      message: `${baseDaysText} O projeto ainda está em Solda.`,
    };
  }

  if (sector === 'Inspeção') {
    return {
      title: diffDays < 0 ? 'Inspeção em atraso' : 'Inspeção em atenção',
      message: `${baseDaysText} O projeto ainda está na Inspeção, preso em ${stageLabel}.`,
    };
  }

  if (sector === 'Pintura') {
    return {
      title: diffDays < 0 ? 'Pintura em atraso' : 'Pintura em atenção',
      message: `${baseDaysText} O projeto ainda está na Pintura.`,
    };
  }

  return {
    title: diffDays < 0 ? 'Prazo vencido' : 'Prazo próximo',
    message: `${baseDaysText} O projeto segue em andamento.`,
  };
}

function isSummaryRow(row) {
  const projectText = textValue(row, "Project");
  if (!projectText) return false;
  if (row.parentId) return false;

  const quantitySpools = parseNumber(row, "Quantity Spools");
  const drawing = textValue(row, "Drawing");
  const parts = parseProjectParts(projectText);

  if (!parts.number) return false;

  const hasClassicSummaryEvidence = quantitySpools != null || drawing === "ISO" || textValue(row, "Project Type");
  const hasSummaryIdentityEvidence = Boolean(
    textValue(row, "Client")
    || textValue(row, "Vessel")
    || textValue(row, "PM")
    || textValue(row, "Priority")
    || textValue(row, "Class")
    || textValue(row, "Start Date")
    || textValue(row, "Finish Date")
    || textValue(row, "Project Status")
    || textValue(row, "PROJECT STATUS")
    || textValue(row, "Overall Project Status")
    || textValue(row, "Status")
    || textValue(row, "% Overall Progress")
    || textValue(row, "% Individual Progress")
    || hasStageProgressEvidence(buildStageValues(row))
  );
  const drawingLooksLikeSpool = Boolean(drawing && drawing !== "ISO");

  // v36.71 - Portugal/cliente: algumas BSPs novas entram no Smartsheet sem o prefixo "BSP"
  // no campo Project (ex.: "26-7001"). Antes o painel só aceitava linha-mãe
  // quando havia prefixo, então a BSP ficava invisível. Agora uma linha sem prefixo
  // também pode abrir uma nova BSP quando houver evidência operacional na linha.
  const hasClientOrVessel = Boolean(textValue(row, "Client") || textValue(row, "Vessel"));
  const hasOperationalEvidence = Boolean(
    hasClassicSummaryEvidence
    || hasSummaryIdentityEvidence
    || textValue(row, "Item")
    || parseNumber(row, "Kilos") != null
    || parseNumber(row, "Peso (KG)") != null
    || parseNumber(row, "Peso Soldado (KG)") != null
    || parseNumber(row, "Quantity Juntas") != null
    || parseNumber(row, "Quant. Juntas") != null
    || textValue(row, "Line Nº")
    || textValue(row, "Line No")
    || textValue(row, "Line Number")
    || textValue(row, "Size")
    || textValue(row, "OBSERVATIONS")
  );

  if (parts.prefix && hasOperationalEvidence && !drawingLooksLikeSpool) return true;

  return Boolean(hasClientOrVessel && hasOperationalEvidence && !drawingLooksLikeSpool);
}

function isChildRow(row) {
  if (row.parentId) return true;
  const drawing = textValue(row, "Drawing");
  const projectText = textValue(row, "Project");
  const parts = parseProjectParts(projectText);
  return Boolean(!parts.prefix && parts.number && drawing && drawing !== "ISO");
}

function buildSpoolRow(row, parentSummary) {
  const drawingText = textValue(row, "Drawing");
  const parsedDrawing = extractIsoDescription(drawingText);
  const spoolProjectText = textValue(row, "Project") || textValue(parentSummary, "Project");
  const spoolQuantityRaw = parseNumber(row, "Quantity Spools");
  const progress = deriveProgress(row);
  const rowOverallProgress = parsePercent(row, "% Overall Progress");
  const rowIndividualProgress = parsePercent(row, "% Individual Progress");
  const overallProgress = rowOverallProgress ?? rowIndividualProgress ?? 0;
  const individualProgress = rowIndividualProgress ?? overallProgress;
  const projectFinishedFlag = isTruthyValue(getCellValue(row, "Project Finished?").raw);
  const fabricationStartDate = textValue(row, "Fabrication Start Date");
  const stageValues = buildStageValues(row);
  const projectType = textValue(row, "Project Type") || textValue(parentSummary, "Project Type");
  const typeFallbackText = [drawingText, parsedDrawing.iso, parsedDrawing.description].filter(Boolean).join(' ');
  const hasIncompleteStageEvidence = hasIncompleteProductionEvidence(stageValues, projectType, typeFallbackText);
  const finished = !hasIncompleteStageEvidence && (projectFinishedFlag || overallProgress >= 100 || hasStageValue(stageValues, "Project Finish Date"));
  const flow = getOperationalFlow(stageValues, fabricationStartDate, parsePercent(row, "Surface preparation and/or coating") ?? 0, finished, textValue(row, "PROJECT STATUS"), projectType, typeFallbackText);
  const awaitingShipment = flow.state === "awaiting_shipment";
  const uiState = uiStateFromFlow(flow, finished);
  const coatingPercent = parsePercent(row, "Surface preparation and/or coating") ?? 0;
  const weldingPercent = parsePercent(row, "Full welding execution") ?? 0;
  const weldingFinishDate = textValue(row, "Welding Finish Date");
  const weldedWeightKg = (() => {
    const kilos = parseNumber(row, "Kilos");
    if (kilos == null) return null;
    if (weldingPercent >= 100) return kilos;
    if (weldingPercent > 0) return (kilos * weldingPercent) / 100;
    return 0;
  })();
  const weldingWeek = weldingPercent >= 100 && weldingFinishDate ? getProductionWeekLabel(weldingFinishDate) : "";
  const observations = textValue(row, "OBSERVATIONS");
  const tratativaObservationMatches = getObservationTratativaMatches([{ source: parsedDrawing.iso || drawingText || `Linha ${row.rowNumber || row.id}`, text: observations }]);

  return {
    rowId: row.id,
    rowNumber: row.rowNumber,
    primary: textValue(row, "Primary") || textValue(row, "PRIMARY") || '',
    projectRef: spoolProjectText,
    projectDisplay: spoolProjectText,
    client: textValue(row, "Client") || textValue(parentSummary, "Client"),
    vessel: textValue(row, "Vessel") || textValue(parentSummary, "Vessel"),
    priority: textValue(row, "Priority") || textValue(parentSummary, "Priority"),
    lineNumber: textValue(row, "Line Nº") || textValue(row, "Line No") || textValue(row, "Line Number"),
    size: textValue(row, "Size"),
    quantitySpools: Number(spoolQuantityRaw || 0) > 0 ? spoolQuantityRaw : 1,
    iso: parsedDrawing.iso,
    description: parsedDrawing.description,
    drawing: drawingText,
    observations,
    tratativaObservationMatches,
    hasTratativaObservation: tratativaObservationMatches.length > 0,
    pm: textValue(row, "PM") || textValue(parentSummary, "PM"),
    projectType,
    operationalSector: flow.sector,
    operationalState: flow.state,
    currentStatus: flow.status,
    currentSector: flow.sector,
    flow,
    plannedStartDate: formatDateValue(textValue(row, "Start Date")),
    plannedFinishDate: formatDateValue(textValue(row, "Finish Date")),
    kilos: parseNumber(row, "Kilos"),
    weldedWeightKg,
    weldingWeek,
    coatingPercent,
    m2Painting: parseNumber(row, "M2 Painting"),
    stage: flow.status,
    stagePercent: flow.percent,
    stageStatus: flow.stageStatus,
    stageAlert: flow.stageStatus === "in_progress" || flow.stageStatus === "waiting",
    individualProgress,
    overallProgress,
    milestones: progress.milestones,
    stageValues,
    finished: finished,
    projectFinishedFlag,
    uiState,
  };
}

function normalizeSpoolIdentity(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getSpoolIdentityKey(spool) {
  const drawing = normalizeSpoolIdentity(spool?.drawing);
  if (drawing) return `drawing:${drawing}`;

  const iso = normalizeSpoolIdentity(spool?.iso);
  const description = normalizeSpoolIdentity(spool?.description);
  if (iso || description) return `iso:${iso}|desc:${description}`;

  return `row:${String(spool?.rowId || "")}`;
}

function getSpoolCompletenessScore(spool) {
  let score = 0;
  score += Number.isFinite(Number(spool?.stagePercent)) ? Number(spool.stagePercent) : 0;
  score += Number.isFinite(Number(spool?.overallProgress)) ? Number(spool.overallProgress) : 0;
  score += Number.isFinite(Number(spool?.individualProgress)) ? Number(spool.individualProgress) : 0;
  score += Number.isFinite(Number(spool?.kilos)) && Number(spool.kilos) > 0 ? 15 : 0;
  score += Number.isFinite(Number(spool?.weldedWeightKg)) && Number(spool.weldedWeightKg) > 0 ? 15 : 0;
  score += spool?.plannedStartDate ? 5 : 0;
  score += spool?.plannedFinishDate ? 5 : 0;
  score += spool?.weldingWeek ? 8 : 0;
  score += spool?.observations ? 3 : 0;
  score += spool?.stage && spool.stage !== '—' ? 5 : 0;
  score += spool?.uiState === 'completed' ? 10 : 0;
  score += spool?.uiState === 'in_progress' ? 6 : 0;
  return score;
}

function mergeUniqueObservationText(parts = []) {
  const values = [];
  const seen = new Set();
  for (const part of parts) {
    const value = String(part || "").trim();
    if (!value) continue;
    const key = normalizeStatusText(value);
    if (seen.has(key)) continue;
    seen.add(key);
    values.push(value);
  }
  return values.join(" | ");
}

function mergeSpoolObservationEvidence(primarySpool, secondarySpool) {
  if (!primarySpool || !secondarySpool) return primarySpool || secondarySpool;

  primarySpool.observations = mergeUniqueObservationText([primarySpool.observations, secondarySpool.observations]);

  const mergedMatches = [];
  const seenMatches = new Set();
  for (const match of [
    ...(Array.isArray(primarySpool.tratativaObservationMatches) ? primarySpool.tratativaObservationMatches : []),
    ...(Array.isArray(secondarySpool.tratativaObservationMatches) ? secondarySpool.tratativaObservationMatches : []),
  ]) {
    const label = String(match?.label || "").trim();
    const source = String(match?.source || "").trim();
    const text = String(match?.text || "").trim();
    if (!label || !text) continue;
    const key = `${label}|${source}|${normalizeStatusText(text)}`;
    if (seenMatches.has(key)) continue;
    seenMatches.add(key);
    mergedMatches.push({ label, source: source || "Tag", text });
  }

  if (!mergedMatches.length && primarySpool.observations) {
    const source = primarySpool.iso || primarySpool.drawing || primarySpool.description || `Linha ${primarySpool.rowNumber || primarySpool.rowId || ""}`.trim() || "Tag";
    mergedMatches.push(...getObservationTratativaMatches([{ source, text: primarySpool.observations }]));
  }

  primarySpool.tratativaObservationMatches = mergedMatches;
  primarySpool.hasTratativaObservation = mergedMatches.length > 0;

  return primarySpool;
}

function chooseBestSpoolRow(currentSpool, nextSpool) {
  if (!currentSpool) return nextSpool;
  const currentScore = getSpoolCompletenessScore(currentSpool);
  const nextScore = getSpoolCompletenessScore(nextSpool);

  let selected = currentSpool;
  let discarded = nextSpool;

  if (nextScore > currentScore) {
    selected = nextSpool;
    discarded = currentSpool;
  } else if (nextScore === currentScore) {
    const currentRowNumber = Number(currentSpool?.rowNumber || 0);
    const nextRowNumber = Number(nextSpool?.rowNumber || 0);
    if (nextRowNumber > currentRowNumber) {
      selected = nextSpool;
      discarded = currentSpool;
    }
  }

  return mergeSpoolObservationEvidence(selected, discarded);
}


function uiStateFromFlow(flow, allFinished = false) {
  if (allFinished || flow?.state === "completed") return "completed";
  if (flow?.state === "awaiting_shipment") return "awaiting_shipment";
  if (flow?.state === "not_started") return "not_started";
  return "in_progress";
}

function applyProjectSpoolRollup(project) {
  const spools = Array.isArray(project.spools) ? project.spools : [];
  const fallbackFlow = project.flow || makeFlow(project.currentStage || "AG. Emissão de detalhamento", project.operationalSector || "Engenharia", project.currentStagePercent || 0, project.currentStageStatus || "waiting", project.operationalState || project.uiState || "not_started");
  const summary = summarizeFlowItems(spools, fallbackFlow, project.quantitySpools || 1);
  const explicitFinished = Boolean(project.finished || project.projectFinishedFlag || hasProjectFinishDateMarker(project) || isProjectStatusFinished(project.projectStatus));
  const hasIncompleteStageEvidence = hasIncompleteProductionEvidence(project.stageValues, project.projectType, project.summaryDrawing) || spools.some((spool) => hasIncompleteProductionEvidence(spool.stageValues, spool.projectType || project.projectType, [spool.iso, spool.drawing, spool.description].filter(Boolean).join(' ')));
  const allSpoolsFinishedByEvidence = spools.length > 0 && spools.every(isSpoolFinishedByState);
  
  // v32.2: Cálculo de progresso baseado estritamente nas ISOs (spools)
  // Se houver spools, o progresso do projeto pai deve ser a média ponderada ou simples dos spools.
  if (spools.length > 0) {
    // Calcula o progresso ponderado das ISOs (spools).
    // Caso um spool não tenha o campo "kilos" definido ou seja zero, assume peso 1 para não ignorá-lo.
    const totalWeight = spools.reduce((sum, s) => sum + (s.kilos || 1), 0);
    const weightedOverall = spools.reduce((sum, s) => {
      const weight = s.kilos || 1;
      return sum + ((s.overallProgress || 0) * weight);
    }, 0) / (totalWeight || 1);
    const weightedIndividual = spools.reduce((sum, s) => {
      const weight = s.kilos || 1;
      return sum + ((s.individualProgress || 0) * weight);
    }, 0) / (totalWeight || 1);
    const weightedCoating = spools.reduce((sum, s) => {
      const weight = s.kilos || 1;
      return sum + ((s.coatingPercent || 0) * weight);
    }, 0) / (totalWeight || 1);
    const totalWeldedWeight = spools.reduce((sum, s) => sum + (s.weldedWeightKg || 0), 0);

    project.overallProgress = weightedOverall;
    project.individualProgress = weightedIndividual;
    project.coatingPercent = weightedCoating;
    project.weldedWeightKg = totalWeldedWeight;

    const spoolsWithStageEvidence = spools.filter((s) => hasStageProgressEvidence(s.stageValues));
    if (spoolsWithStageEvidence.length) {
      // Para cálculo por estágios, utilize peso 1 para spools sem kilos definidos.
      const stageTotalWeight = spoolsWithStageEvidence.reduce((sum, s) => sum + (s.kilos || 1), 0);
      project.overallProgress = spoolsWithStageEvidence.reduce((sum, s) => {
        const weight = s.kilos || 1;
        return sum + weightedOverallFromStageValues(
          s.stageValues,
          s.projectType || project.projectType,
          [s.iso, s.drawing, s.description].filter(Boolean).join(' ')
        ) * weight;
      }, 0) / (stageTotalWeight || 1);
    }
    
    // Atualiza estatísticas de spools baseadas no estado real de cada um
    project.spoolStats = spools.reduce((acc, s) => {
      acc.total += 1;
      if (s.uiState === "completed" || s.finished) acc.completed += 1;
      else if (s.uiState === "in_progress" || s.overallProgress > 0) acc.inProgress += 1;
      else acc.notStarted += 1;
      return acc;
    }, { total: 0, completed: 0, inProgress: 0, notStarted: 0 });
  }

  mergeProjectStageValuesFromSpools(project);

  if (hasStageProgressEvidence(project.stageValues)) {
    project.overallProgress = weightedOverallFromStageValues(project.stageValues, project.projectType, project.summaryDrawing);
  }

  const finalFinished = summary.allFinished || (explicitFinished && !hasIncompleteStageEvidence && (spools.length === 0 || allSpoolsFinishedByEvidence));
  const finalFlow = finalFinished ? makeFlow("Finalizado", "Enviado", 100, "completed", "completed") : summary.flow;
  
  project.demandSummary = summary;
  project.statusSummary = finalFinished ? "Finalizado" : summary.statusSummary;
  project.sectorSummary = finalFinished ? "Enviado" : summary.sectorSummary;
  project.statusBreakdown = summary.statusBreakdown;
  project.sectorBreakdown = summary.sectorBreakdown;
  project.flow = finalFlow;
  project.currentStage = finalFinished ? "Finalizado" : summary.statusSummary;
  project.currentStageGroup = finalFinished ? "Enviado" : summary.sectorSummary;
  project.currentStagePercent = finalFinished ? 100 : (spools.length > 0 ? project.overallProgress : finalFlow.percent);
  project.currentStageStatus = finalFinished ? "completed" : (finalFlow.stageStatus || "waiting");
  project.currentStageAlert = !finalFinished && ["in_progress", "waiting"].includes(project.currentStageStatus);
  project.operationalSector = finalFinished ? "Enviado" : summary.sectorSummary;
  project.operationalState = finalFlow.state;
  project.finished = finalFinished;
  project.uiState = uiStateFromFlow(finalFlow, finalFinished);

  // Correção v32.1/v32.2: Se o projeto está finalizado, forçamos todos os indicadores a 100%
  // Isso garante que mesmo que a planilha tenha dados parciais nas ISOs, o status "Finalizado" prevaleça.
  if (finalFinished) {
    project.overallProgress = 100;
    project.individualProgress = 100;
    project.currentStagePercent = 100;
    project.coatingPercent = 100;
    
    if (project.kilos != null) {
      project.weldedWeightKg = project.kilos;
    }

    if (project.spoolStats) {
      project.spoolStats.completed = project.spoolStats.total;
      project.spoolStats.inProgress = 0;
      project.spoolStats.notStarted = 0;
    }

    if (Array.isArray(project.spools)) {
      project.spools.forEach(spool => {
        spool.finished = true;
        spool.uiState = 'completed';
        spool.individualProgress = 100;
        spool.overallProgress = 100;
        spool.stagePercent = 100;
        spool.stageStatus = 'completed';
        if (spool.kilos != null) {
          spool.weldedWeightKg = spool.kilos;
        }
        if (spool.flow) {
          spool.flow.percent = 100;
          spool.flow.status = 'Finalizado';
          spool.flow.state = 'completed';
          spool.flow.stageStatus = 'completed';
        }
      });
    }

    if (project.demandSummary) {
      project.demandSummary.allFinished = true;
      project.demandSummary.statusSummary = "Finalizado";
      project.demandSummary.sectorSummary = "Enviado";
      if (project.demandSummary.flow) {
        project.demandSummary.flow.percent = 100;
        project.demandSummary.flow.status = 'Finalizado';
        project.demandSummary.flow.state = 'completed';
      }
    }
  }

  return project;
}

function buildProject(summaryRow, childRows) {
  const projectText = textValue(summaryRow, "Project");
  const parts = parseProjectParts(projectText);
  const progress = deriveProgress(summaryRow);
  const overallProgress = parsePercent(summaryRow, "% Overall Progress") ?? 0;
  const individualProgress = parsePercent(summaryRow, "% Individual Progress") ?? overallProgress;
  const projectFinishedFlag = isTruthyValue(getCellValue(summaryRow, "Project Finished?").raw);
  const projectStatus = textValue(summaryRow, "Project Status") || textValue(summaryRow, "PROJECT STATUS") || textValue(summaryRow, "Overall Project Status") || textValue(summaryRow, "Status");
  const observations = textValue(summaryRow, "OBSERVATIONS");
  const coatingPercent = parsePercent(summaryRow, "Surface preparation and/or coating") ?? 0;
  const fabricationStartDate = textValue(summaryRow, "Fabrication Start Date");
  const projectType = textValue(summaryRow, "Project Type");
  const clientFocalPointFromTracking = textValueAny(summaryRow, CLIENT_FOCAL_POINT_ALIASES);
  const summaryDrawing = textValue(summaryRow, "Drawing");
  const stageValues = buildStageValues(summaryRow);
  const summaryHasIncompleteStageEvidence = hasIncompleteProductionEvidence(stageValues, projectType, summaryDrawing);
  const summaryFinished = !summaryHasIncompleteStageEvidence && (projectFinishedFlag || overallProgress >= 100 || hasStageValue(stageValues, "Project Finish Date"));
  const flow = getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, summaryFinished, projectStatus, projectType, summaryDrawing);
  const awaitingShipment = flow.state === "awaiting_shipment";
  const uiState = uiStateFromFlow(flow, summaryFinished) || projectUiState(projectStatus, overallProgress, summaryFinished, fabricationStartDate, awaitingShipment);
  const weldingPercent = parsePercent(summaryRow, "Full welding execution") ?? 0;
  const weldingFinishDate = textValue(summaryRow, "Welding Finish Date");
  const spools = childRows.map((row) => buildSpoolRow(row, summaryRow));
  const summaryWeldedWeightKg = (() => {
    const kilos = parseNumber(summaryRow, "Kilos");
    if (kilos == null) return null;
    if (weldingPercent >= 100) return kilos;
    if (weldingPercent > 0) return (kilos * weldingPercent) / 100;
    return 0;
  })();
  const weldedWeightKg = spools.length
    ? spools.reduce((total, spool) => total + (spool.weldingWeek ? (spool.weldedWeightKg || 0) : 0), 0)
    : summaryWeldedWeightKg;
  const weldingWeek = weldingPercent >= 100 && weldingFinishDate ? getProductionWeekLabel(weldingFinishDate) : "";

  const spoolStats = spools.reduce((acc, spool) => {
    acc.total += 1;
    if (spool.uiState === "completed") acc.completed += 1;
    else if (spool.uiState === "in_progress") acc.inProgress += 1;
    else acc.notStarted += 1;
    return acc;
  }, { total: 0, completed: 0, inProgress: 0, notStarted: 0 });

  const quantitySpoolsRaw = parseNumber(summaryRow, "Quantity Spools");
  const quantitySpools = Number(quantitySpoolsRaw || 0) > 0
    ? quantitySpoolsRaw
    : spools.length;

  const operationalSector = flow.sector;

  const project = {
    rowId: summaryRow.id,
    rowNumber: summaryRow.rowNumber,
    primary: textValue(summaryRow, "Primary") || textValue(summaryRow, "PRIMARY") || '',
    projectPrefix: parts.prefix,
    projectNumber: parts.number,
    projectDisplay: parts.display || projectText,
    customerPo: '',
    customerPoList: [],
    customerPoDisplay: 'Aguardando PO',
    customerPoStatus: 'waiting',
    clientFocalPoint: clientFocalPointFromTracking,
    clientFocalPointList: clientFocalPointFromTracking ? collectWipNameValues(clientFocalPointFromTracking) : [],
    clientFocalPointDisplay: clientFocalPointFromTracking,
    clientDisplayCode: `${parts.display || projectText || 'BSP'} - Aguardando PO`,
    priority: textValue(summaryRow, "Priority"),
    lineNumber: textValue(summaryRow, "Line Nº") || textValue(summaryRow, "Line No") || textValue(summaryRow, "Line Number"),
    size: textValue(summaryRow, "Size"),
    quantitySpools,
    kilos: parseNumber(summaryRow, "Kilos"),
    weldedWeightKg,
    weldingWeek,
    coatingPercent,
    m2Painting: parseNumber(summaryRow, "M2 Painting"),
    currentStage: flow.status,
    currentStagePercent: flow.percent,
    currentStageStatus: flow.stageStatus,
    currentStageAlert: flow.stageStatus === "in_progress" || flow.stageStatus === "waiting",
    individualProgress: spools.length > 0 ? 0 : individualProgress, // Será calculado no rollup se houver spools
    overallProgress: spools.length > 0 ? 0 : overallProgress,       // Será calculado no rollup se houver spools
    projectStatus,
    observations,
    jobProcessStatus: textValue(summaryRow, "Job Process Status") || progress.currentStage.label,
    summaryDrawing,
    projectType,
    fabricationStartDate: formatDateValue(textValue(summaryRow, "Fabrication Start Date")),
    plannedStartDate: formatDateValue(textValue(summaryRow, "Start Date")),
    plannedFinishDate: formatDateValue(textValue(summaryRow, "Finish Date")),
    shipmentDate: stageValues["Project Finish Date"] || formatDateValue(textValue(summaryRow, "Project Finish Date")) || "",
    client: textValue(summaryRow, "Client"),
    pm: textValue(summaryRow, "PM"),
    vessel: textValue(summaryRow, "Vessel"),
    className: textValue(summaryRow, "Class"),
    milestones: progress.milestones,
    stageValues,
    finished: summaryFinished,
    projectFinishedFlag,
    uiState,
    operationalSector,
    operationalState: flow.state,
    currentStatus: flow.status,
    currentSector: flow.sector,
    statusSummary: flow.status,
    sectorSummary: flow.sector,
    flow,
    spools,
    spoolStats,
  };
  return decorateProjectTratativaObservation(applyProjectSpoolRollup(project));
}

function mapApiRows(sheet) {
  const columnMap = new Map((sheet.columns || []).map((column) => [column.id, column.title]));
  return (sheet.rows || []).map((row) => {
    const values = {};
    for (const cell of row.cells || []) {
      const title = columnMap.get(cell.columnId);
      if (!title) continue;
      values[title] = { raw: cell.value ?? null, display: cell.displayValue ?? null };
    }
    return {
      id: row.id,
      rowNumber: row.rowNumber,
      parentId: row.parentId ?? null,
      siblingId: row.siblingId ?? null,
      expanded: row.expanded ?? null,
      values,
    };
  });
}

function buildProjects(rows) {
  const projects = [];
  const rowsById = new Map(rows.map((row) => [row.id, row]));
  const childrenByParent = new Map();

  for (const row of rows) {
    if (row.parentId && rowsById.has(row.parentId)) {
      if (!childrenByParent.has(row.parentId)) childrenByParent.set(row.parentId, []);
      childrenByParent.get(row.parentId).push(row);
    }
  }

  function getLeafChildRows(parentId) {
    const directChildren = childrenByParent.get(parentId) || [];
    const leafRows = [];
    for (const child of directChildren) {
      const descendants = getLeafChildRows(child.id);
      if (descendants.length) leafRows.push(...descendants);
      else leafRows.push(child);
    }
    return leafRows;
  }

  let currentSummary = null;

  for (const row of rows) {
    const currentProjectNumber = currentSummary ? parseProjectParts(textValue(currentSummary, "Project")).number : '';
    const rowProjectNumber = parseProjectParts(textValue(row, "Project")).number;

    // v36.71 - quando a planilha está em formato plano, as linhas seguintes da mesma BSP
    // podem não vir como children do Smartsheet e também podem não ter coluna Drawing.
    // Se a linha atual tem o mesmo número da última BSP aberta, ela deve entrar como item/tag
    // da BSP atual, não abrir uma BSP duplicada nem ser ignorada.
    if (currentSummary && rowProjectNumber && currentProjectNumber && rowProjectNumber === currentProjectNumber && !row.parentId) {
      const lastProject = projects[projects.length - 1];
      if (lastProject) {
        const spool = buildSpoolRow(row, currentSummary);
        lastProject.spools.push(spool);
        lastProject.spoolStats.total += 1;
        if (spool.uiState === "completed") lastProject.spoolStats.completed += 1;
        else if (spool.uiState === "in_progress") lastProject.spoolStats.inProgress += 1;
        else lastProject.spoolStats.notStarted += 1;
        continue;
      }
    }

    if (isSummaryRow(row)) {
      const projectChildren = getLeafChildRows(row.id);
      currentSummary = row;
      projects.push(buildProject(row, projectChildren));
      continue;
    }

    if (!currentSummary) continue;
    if (!isChildRow(row)) continue;

    if (!rowProjectNumber || rowProjectNumber !== currentProjectNumber) continue;

    const lastProject = projects[projects.length - 1];
    if (!lastProject) continue;
    const spool = buildSpoolRow(row, currentSummary);
    lastProject.spools.push(spool);
    lastProject.spoolStats.total += 1;
    if (spool.uiState === "completed") lastProject.spoolStats.completed += 1;
    else if (spool.uiState === "in_progress") lastProject.spoolStats.inProgress += 1;
    else lastProject.spoolStats.notStarted += 1;
  }

  for (const project of projects) {
    const uniqueMap = new Map();
    for (const spool of project.spools) {
      const key = getSpoolIdentityKey(spool);
      const currentSpool = uniqueMap.get(key);
      uniqueMap.set(key, chooseBestSpoolRow(currentSpool, spool));
    }
    const unique = Array.from(uniqueMap.values()).sort((a, b) => (Number(a?.rowNumber || 0) - Number(b?.rowNumber || 0)));
    project.spools = unique;
    project.spoolStats = unique.reduce((acc, spool) => {
      acc.total += 1;
      if (spool.uiState === "completed") acc.completed += 1;
      else if (spool.uiState === "in_progress") acc.inProgress += 1;
      else acc.notStarted += 1;
      return acc;
    }, { total: 0, completed: 0, inProgress: 0, notStarted: 0 });
    applyProjectSpoolRollup(project);
  }

  return projects;
}

function getProjectAlert(project, today = getCurrentBrazilDateObject()) {
  if (!project.fabricationStartDate) return null;
  if (hasProjectFinishedMarker(project)) return null;
  if (project?.uiState === "completed" || project?.operationalState === "completed") return null;

  const plannedFinish = parseDateObject(project.plannedFinishDate);
  if (!plannedFinish) return null;

  const diffDays = Math.floor((plannedFinish - today) / 86400000);
  const coatingPercent = Number(project.coatingPercent || 0);
  const sector = classifyAlertSector(project);
  const observation = buildAlertObservation(project, sector, diffDays);

  if (coatingPercent < 100 && diffDays <= 5) {
    return {
      projectDisplay: project.projectDisplay,
      projectNumber: project.projectNumber,
      projectRowId: project.rowId,
      client: project.client,
      sector,
      plannedFinishDate: project.plannedFinishDate,
      daysRemaining: diffDays,
      type: diffDays < 0 ? "overdue" : "deadline",
      title: observation.title,
      message: observation.message,
      coatingPercent,
      currentStage: project.currentStage,
    };
  }

  if (coatingPercent >= 100 && diffDays <= 3) {
    return {
      projectDisplay: project.projectDisplay,
      projectNumber: project.projectNumber,
      projectRowId: project.rowId,
      client: project.client,
      sector,
      plannedFinishDate: project.plannedFinishDate,
      daysRemaining: diffDays,
      type: diffDays < 0 ? "conference_overdue" : "conference",
      title: observation.title,
      message: observation.message,
      coatingPercent,
      currentStage: project.currentStage,
    };
  }

  return null;
}

function buildAlerts(projects) {
  const alerts = projects
    .map((project) => getProjectAlert(project))
    .filter(Boolean)
    .sort((a, b) => {
      if (a.daysRemaining !== b.daysRemaining) return a.daysRemaining - b.daysRemaining;
      return String(a.projectDisplay || "").localeCompare(String(b.projectDisplay || ""), "pt-BR");
    });

  const signature = alerts
    .map((alert) => [alert.projectDisplay, alert.type, alert.plannedFinishDate, alert.daysRemaining].join("|"))
    .join("||");

  return { alerts, signature };
}

function normalizeStatusText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function compactStatusText(value) {
  return normalizeStatusText(value).replace(/[^A-Z0-9]+/g, "");
}


const TRATATIVA_OBSERVATION_RULES = [
  {
    label: "Revisão de P.O",
    aliases: ["Revisão de P.O", "Revisao de P.O", "Revisão de PO", "Revisao de PO", "Revisão PO", "Revisao PO"],
  },
  {
    label: "Aguardando liberação para envio",
    aliases: ["Aguardando liberação para envio", "Aguardando liberacao para envio", "Aguardando liberação p/ envio", "Aguardando liberacao p envio", "Aguardando liberação envio", "Aguardando liberacao envio"],
  },
  {
    label: "Entrega parcial",
    aliases: ["Entrega parcial"],
  },
];

function observationTextMatchesRule(text, aliases = []) {
  const normalized = normalizeStatusText(text || "");
  const compact = compactStatusText(text || "");
  if (!normalized && !compact) return false;
  return aliases.some((alias) => {
    const normalizedAlias = normalizeStatusText(alias || "");
    const compactAlias = compactStatusText(alias || "");
    return Boolean(
      (normalizedAlias && normalized.includes(normalizedAlias))
      || (compactAlias && compact.includes(compactAlias))
    );
  });
}

function getObservationTratativaMatches(contexts = []) {
  const matches = [];
  const seen = new Set();
  for (const context of contexts) {
    const text = String(context?.text || "").trim();
    if (!text) continue;
    for (const rule of TRATATIVA_OBSERVATION_RULES) {
      if (!observationTextMatchesRule(text, rule.aliases)) continue;
      const source = String(context?.source || "BSP").trim() || "BSP";
      const key = `${rule.label}|${source}|${text}`;
      if (seen.has(key)) continue;
      seen.add(key);
      matches.push({ label: rule.label, source, text });
    }
  }
  return matches;
}

function getProjectObservationTratativaMatches(project) {
  if (!project) return [];
  const contexts = [{ source: "BSP", text: project.observations }];
  if (Array.isArray(project.spools)) {
    project.spools.forEach((spool, index) => {
      const source = spool?.iso || spool?.drawing || spool?.description || `Tag ${index + 1}`;
      contexts.push({ source: `Tag ${source}`, text: spool?.observations });
    });
  }
  return getObservationTratativaMatches(contexts);
}

function decorateProjectTratativaObservation(project) {
  const matches = getProjectObservationTratativaMatches(project);
  project.tratativaObservationMatches = matches;
  project.tratativaObservationReason = matches.map((item) => `${item.label} • ${item.source}: ${item.text}`).join(" | ");
  project.hasTratativaObservation = matches.length > 0;
  if (matches.length) {
    project.statusPresentationOverride = { text: "Em tratativa", state: "in_progress", reason: project.tratativaObservationReason };
  }
  return project;
}

function isMeaningfulFinishValue(value) {
  if (value == null) return false;
  const raw = String(value).trim();
  if (!raw) return false;
  const compact = compactStatusText(raw);
  return !["NA", "N/A", "NONE", "NULL", "FALSE", "NO", "NAO", "0"].includes(compact);
}

function isProjectStatusFinished(value) {
  const normalized = normalizeStatusText(value);
  const compact = compactStatusText(value);
  return ["FINALIZADO", "CONCLUIDO", "FINISHED", "COMPLETED", "DELIVERED", "ENTREGUE", "ENVIADO"].includes(compact)
    || normalized.includes("PROJECT FINISHED")
    || normalized.includes("PROJETO FINALIZADO");
}

function hasProjectFinishDateMarker(project) {
  if (!project) return false;
  const values = [project.projectFinishDate, project.finishDate, project.finishedDate, project.shipmentDate, project.stageValues?.["Project Finish Date"], project.stageValues?.["PROJECT FINISH DATE"]];
  return values.some(isMeaningfulFinishValue);
}

function hasProjectFinishedBooleanMarker(project) {
  if (!project) return false;
  const values = [project.finished, project.projectFinishedFlag, project.stageValues?.["Project Finished?"], project.stageValues?.["PROJECT FINISHED?"]];
  return values.some((value) => {
    if (typeof value === "boolean") return value;
    const compact = compactStatusText(value);
    return ["TRUE", "YES", "SIM", "Y", "1", "FINALIZADO", "CONCLUIDO", "COMPLETED", "FINISHED"].includes(compact);
  });
}

function areAllProjectSpoolsFinished(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  return spools.length > 0 && spools.every((spool) => !hasIncompleteProductionEvidence(spool?.stageValues, spool?.projectType, [spool?.iso, spool?.drawing, spool?.description].filter(Boolean).join(' ')) && Boolean(
    spool?.finished
    || spool?.projectFinishedFlag
    || spool?.uiState === "completed"
    || spool?.operationalState === "completed"
    || spool?.flow?.state === "completed"
    || spool?.flow?.status === "Finalizado"
    || isProjectStatusFinished(spool?.projectStatus)
    || isProjectStatusFinished(spool?.status)
    || isProjectStatusFinished(spool?.currentStatus)
    || isMeaningfulFinishValue(spool?.stageValues?.["Project Finish Date"])
  ));
}

function hasProjectFinishedMarker(project) {
  if (!project) return false;
  if (hasIncompleteProductionEvidence(project.stageValues, project.projectType, project.summaryDrawing) || (Array.isArray(project.spools) && project.spools.some((spool) => hasIncompleteProductionEvidence(spool.stageValues, spool.projectType || project.projectType, [spool.iso, spool.drawing, spool.description].filter(Boolean).join(' '))))) return false;
  return Boolean(
    hasProjectFinishedBooleanMarker(project)
    || hasProjectFinishDateMarker(project)
    || project.uiState === "completed"
    || project.operationalState === "completed"
    || project.flow?.state === "completed"
    || isProjectStatusFinished(project.projectStatus)
    || isProjectStatusFinished(project.status)
    || isProjectStatusFinished(project.currentStage)
    || isProjectStatusFinished(project.currentStatus)
    || isProjectStatusFinished(project.statusSummary)
    || isProjectStatusFinished(project.flow?.status)
    || areAllProjectSpoolsFinished(project)
  );
}

function getOpenFlowItemsForStats(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const source = spools.length
    ? spools.map((spool) => ({ flow: spool.flow || { status: spool.stage, sector: spool.operationalSector, state: spool.operationalState }, spool }))
    : [{ flow: project?.flow || { status: project?.currentStage, sector: project?.operationalSector, state: project?.operationalState }, spool: null }];
  return source.filter((item) => item.flow?.state !== "completed" && item.flow?.status !== "Finalizado");
}

function isProjectStatusOnHold(projectStatus) {
  const normalized = String(projectStatus || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
  const compact = normalized.replace(/[^A-Z0-9]+/g, "");
  return compact === "ONHOLD"
    || compact === "HOLD"
    || compact === "PAUSADO"
    || compact === "PAUSED"
    || compact === "EMESPERA"
    || normalized.includes("ON HOLD")
    || normalized.includes("EM HOLD")
    || normalized.includes("PROJETO EM HOLD")
    || normalized.includes("HOLD CONFORME")
    || normalized.includes("HOLD")
    || normalized.includes("EM ESPERA")
    || normalized.includes("PAUSADO")
    || normalized.includes("PAUSED")
    || normalized.includes("PARALISADO")
    || normalized.includes("SUSPENSO");
}

function isProjectStatusPending(projectStatus) {
  const compact = compactStatusText(projectStatus);
  return compact === "PENDING";
}

function getProjectStatusTexts(project) {
  if (!project) return [];
  const texts = [project.projectStatus, project["PROJECT STATUS"], project.status, project.currentStatus, project.currentStage, project.statusSummary, project.sectorSummary, project.operationalState, project.uiState, project.flow?.status, project.flow?.state];
  if (project.stageValues && typeof project.stageValues === "object") {
    texts.push(project.stageValues["Project Finished?"], project.stageValues["Project Finish Date"]);
  }
  if (Array.isArray(project.spools)) {
    project.spools.forEach((spool) => {
      texts.push(spool?.projectStatus, spool?.["PROJECT STATUS"], spool?.status, spool?.currentStatus, spool?.stage, spool?.uiState, spool?.operationalState, spool?.flow?.status, spool?.flow?.state);
      if (spool?.stageValues && typeof spool.stageValues === "object") {
        texts.push(spool.stageValues["Project Finished?"], spool.stageValues["Project Finish Date"]);
      }
    });
  }
  return texts.filter((value) => value != null && String(value).trim() !== "");
}

function isProjectPending(project) {
  return getProjectStatusTexts(project).some((value) => isProjectStatusPending(value));
}

function isProjectExcludedFromTotal(project) {
  return isProjectOnHold(project) || isProjectPending(project) || hasProjectFinishedMarker(project);
}

function getProjectRootJobProcessStatus(project) {
  const candidates = [
    project?.jobProcessStatus,
    project?.stageValues?.["Job Process Status"],
    project?.stageValues?.["JOB PROCESS STATUS"],
    project?.currentStage,
  ];
  const value = candidates.find((item) => item != null && String(item).trim() !== "");
  return value == null ? "" : String(value).trim();
}

function isProjectRootJobStatusDelivered(project) {
  const normalized = normalizeStatusText(getProjectRootJobProcessStatus(project));
  return normalized.includes("PACKAGE AND DELIVERED") || normalized.includes("PROJECT FINISHED");
}

function isProjectRootJobStatusNotStarted(project) {
  const normalized = normalizeStatusText(getProjectRootJobProcessStatus(project));
  return normalized.includes("FABRICATION NOT STARTED");
}

function getProjectCardBucketFromFlow(flow) {
  const sector = normalizeFlowSortText(flow?.sector || "");
  const weight = getFlowSortWeight(flow || {});

  if (["logistica", "envio", "enviado"].includes(sector) || (weight >= 120 && weight < 500)) return "awaiting";
  if (["pintura", "painting", "coating"].includes(sector) || (weight >= 110 && weight < 500)) return "painting";
  if (["qualidade", "inspecao", "inspection", "quality"].includes(sector) || (weight >= 80 && weight < 500)) return "inspection";
  if (["producao", "solda", "calderaria", "production", "welding", "fabrication"].includes(sector) || (weight >= 40 && weight < 500)) return "production";
  return "not_started";
}

function getProjectDelayedStageStats(project) {
  const openItems = getOpenFlowItemsForStats(project);
  const fallbackFlow = project?.flow || { status: getProjectRootJobProcessStatus(project), sector: project?.operationalSector || project?.currentStageGroup || project?.currentSector };
  const source = openItems.length ? openItems : [{ flow: fallbackFlow, spool: null }];
  const sorted = [...source].sort((a, b) => getFlowSortWeight(a.flow || {}) - getFlowSortWeight(b.flow || {}));
  const delayedBucket = getProjectCardBucketFromFlow(sorted[0]?.flow || fallbackFlow);
  const matchingItems = source.filter((item) => getProjectCardBucketFromFlow(item.flow || fallbackFlow) === delayedBucket);
  const itemCount = source.some((item) => item.spool)
    ? matchingItems.length
    : Number(project?.quantitySpools || matchingItems.length || 1);
  return {
    bucket: delayedBucket,
    tagCount: Math.max(0, Number(itemCount || 0)),
    flow: sorted[0]?.flow || fallbackFlow,
  };
}

function getProjectExclusiveCardBucket(project) {
  // v37.70: classificação exclusiva pela etapa mais atrasada dos ISOs abertos.
  // A Etapa Atual continua mostrando todas as etapas, mas os cards contam a menor etapa.
  // Ex.: se a BSP tem ISO em Pintura e ISO em Solda, o card da BSP entra em Produção/Solda.
  if (!project) return "unknown";
  if (isProjectOnHold(project)) return "hold";
  if (isProjectPending(project)) return "pending";
  if (hasProjectFinishedMarker(project)) return "finished";
  const delayed = getProjectDelayedStageStats(project);
  if (delayed.bucket === "not_started") return "not_started";
  return "started";
}

function isProjectStartedForStats(project) {
  if (!project) return { started: false, tags: 0 };
  const bucket = getProjectExclusiveCardBucket(project);
  if (bucket !== "started") return { started: false, tags: 0 };
  return { started: true, tags: Number(project.quantitySpools || project.spools?.length || 1) };
}

function isAwaitingShipmentStatsFlow(flow) {
  const normalized = normalizeStatusText(flow?.status || "").replace(/[^A-Z0-9]+/g, "");
  return normalized === "AGUARDANDOENVIO";
}

function getAwaitingShipmentTagsForStats(project) {
  if (!project || isProjectExcludedFromTotal(project)) return 0;
  const openItems = getOpenFlowItemsForStats(project);
  const awaitingItems = openItems.filter((item) => isAwaitingShipmentStatsFlow(item.flow));
  return awaitingItems.length;
}

function getProjectHoldContextTexts(project) {
  if (!project) return [];
  const texts = [
    project.projectStatus,
    project.status,
    project.jobProcessStatus,
    project.currentStage,
    project.currentStageGroup,
    project.currentStatus,
    project.statusSummary,
    project.sectorSummary,
    project.operationalState,
    project.observations,
    project.note,
    project.notes,
    project.summaryDrawing,
  ];

  if (project.stageValues && typeof project.stageValues === "object") {
    texts.push(...Object.values(project.stageValues));
  }

  if (Array.isArray(project.spools)) {
    project.spools.forEach((spool) => {
      texts.push(
        spool?.observations,
        spool?.currentStatus,
        spool?.stage,
        spool?.stageStatus,
        spool?.operationalState,
        spool?.drawing,
        spool?.description
      );
      if (spool?.stageValues && typeof spool.stageValues === "object") {
        texts.push(...Object.values(spool.stageValues));
      }
    });
  }

  return texts.filter((value) => value != null && String(value).trim() !== "");
}

function isProjectOnHold(project) {
  return getProjectHoldContextTexts(project).some((value) => isProjectStatusOnHold(value));
}

function buildStats(projects) {
  const sourceProjects = Array.isArray(projects) ? projects : [];
  const activeProjects = sourceProjects.filter((project) => !isProjectExcludedFromTotal(project));
  const stats = {
    // v37.68: Total de Projetos deve contar a quantidade real de linhas raiz/projetos
    // retornados do Smartsheet. On Hold, Não iniciado e Enviado são recortes, não exclusões.
    totalProjects: sourceProjects.length,
    totalSpools: 0,
    totalWeightKg: 0,
    totalWeldedWeightKg: 0,
    totalPaintingM2: 0,
    completed: 0,
    completedTags: 0,
    startedProjects: 0,
    startedTags: 0,
    inProgress: 0,
    inProgressTags: 0,
    inspectionProjects: 0,
    inspectionTags: 0,
    paintingProjects: 0,
    paintingTags: 0,
    awaitingShipment: 0,
    awaitingShipmentTags: 0,
    notStarted: 0,
    notStartedTags: 0,
    notStartedHold: 0,
    notStartedHoldTags: 0,
    averageOverallProgress: 0,
  };

  let progressAccumulator = 0;

  for (const project of projects) {
    const tags = Number(project.quantitySpools || project.spools?.length || 0);
    const spools = Array.isArray(project.spools) ? project.spools : [];
    const isFinishedProject = hasProjectFinishedMarker(project);
    stats.totalSpools += tags;
    stats.totalWeightKg += project.kilos || 0;
    stats.totalWeldedWeightKg += project.weldedWeightKg || 0;
    const openPaintingM2 = spools.length
      ? spools.filter((spool) => spool.flow?.state !== "completed" && spool.flow?.status !== "Finalizado").reduce((total, spool) => total + Number(spool.m2Painting || 0), 0)
      : 0;
    stats.totalPaintingM2 += isFinishedProject ? 0 : (openPaintingM2 > 0 ? openPaintingM2 : Number(project.m2Painting || 0));
    const cardBucket = getProjectExclusiveCardBucket(project);

    if (cardBucket === "hold") {
      stats.notStartedHold += 1;
      stats.notStartedHoldTags += tags;
      continue;
    }

    if (cardBucket === "pending") {
      continue;
    }

    if (cardBucket === "finished") {
      stats.completed += 1;
      stats.completedTags += tags;
      continue;
    }

    if (cardBucket === "not_started") {
      stats.notStarted += 1;
      stats.notStartedTags += tags;
      continue;
    }

    stats.startedProjects += 1;
    stats.startedTags += tags;
    progressAccumulator += project.overallProgress || 0;

    const delayedStage = getProjectDelayedStageStats(project);
    const delayedTags = delayedStage.tagCount || tags || 1;
    if (delayedStage.bucket === "production") { stats.inProgress += 1; stats.inProgressTags += delayedTags; }
    if (delayedStage.bucket === "inspection") { stats.inspectionProjects += 1; stats.inspectionTags += delayedTags; }
    if (delayedStage.bucket === "painting") { stats.paintingProjects += 1; stats.paintingTags += delayedTags; }
    if (delayedStage.bucket === "awaiting") { stats.awaitingShipment += 1; stats.awaitingShipmentTags += delayedTags; }
  }

  stats.averageOverallProgress = activeProjects.length ? progressAccumulator / activeProjects.length : 0;
  return stats;
}

function isTransientSmartsheetError(error) {
  const message = String(error?.message || error || '').toLowerCase();
  return error?.name === 'AbortError'
    || message.includes('fetch failed')
    || message.includes('terminated')
    || message.includes('timeout')
    || message.includes('econnreset')
    || message.includes('socket')
    || message.includes('network');
}

async function apiFetch(path, options = {}) {
  const timeoutMs = Math.max(2500, Number(options.timeoutMs || SMARTSHEET_FETCH_TIMEOUT_MS));
  const attempts = Math.max(1, Number(options.retries || SMARTSHEET_FETCH_RETRIES));
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const response = await fetch(`${API_BASE}${path}`, {
        signal: controller.signal,
        headers: {
          Authorization: `Bearer ${TOKEN}`,
          "Content-Type": "application/json",
        },
      });
      clearTimeout(timer);

      if (!response.ok) {
        const message = await response.text().catch(() => '');
        throw new Error(`Smartsheet ${response.status}: ${message}`);
      }

      return response.json();
    } catch (error) {
      clearTimeout(timer);
      lastError = error;
      const transient = isTransientSmartsheetError(error);
      if (!transient || attempt >= attempts) break;
      await new Promise((resolve) => setTimeout(resolve, 250 * attempt));
    }
  }

  throw lastError || new Error('Falha ao consultar Smartsheet.');
}

function normalizeName(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase();
}

async function resolveSheetId() {
  if (cache.sheetId) return cache.sheetId;
  if (SHEET_ID_ENV) {
    cache.sheetId = SHEET_ID_ENV;
    return cache.sheetId;
  }

  const target = normalizeName(SHEET_NAME);
  let page = 1;
  let fuzzyFound = null;

  while (true) {
    const response = await apiFetch(`/sheets?page=${page}&pageSize=100`);
    const items = response.data || [];

    const exactFound = items.find((item) => normalizeName(item.name) === target);
    if (exactFound) {
      cache.sheetId = String(exactFound.id);
      cache.sheetName = exactFound.name;
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

  throw new Error(`Sheet "${SHEET_NAME}" não encontrada. Defina SMARTSHEET_SHEET_ID ou confira SMARTSHEET_SHEET_NAME.`);
}

async function resolveWipStepSheetId() {
  if (cache.wipStepSheetId) return cache.wipStepSheetId;
  if (WIP_STEP_SHEET_ID_ENV) {
    cache.wipStepSheetId = String(WIP_STEP_SHEET_ID_ENV);
    return cache.wipStepSheetId;
  }

  const target = normalizeName(WIP_STEP_SHEET_NAME);
  let page = 1;
  let fuzzyFound = null;

  while (true) {
    const response = await apiFetch(`/sheets?page=${page}&pageSize=100`);
    const items = response.data || [];
    const eligible = items.filter((item) => !normalizeName(item.name).includes('portugal'));

    const exactFound = eligible.find((item) => normalizeName(item.name) === target);
    if (exactFound) {
      cache.wipStepSheetId = String(exactFound.id);
      cache.wipStepSheetName = exactFound.name;
      return cache.wipStepSheetId;
    }

    if (!fuzzyFound) {
      fuzzyFound = eligible.find((item) => normalizeName(item.name).includes(target) || target.includes(normalizeName(item.name)));
    }

    if (!items.length || page >= (response.totalPages || 1)) break;
    page += 1;
  }

  if (fuzzyFound) {
    cache.wipStepSheetId = String(fuzzyFound.id);
    cache.wipStepSheetName = fuzzyFound.name;
    return cache.wipStepSheetId;
  }

  return null;
}

function collectWipPoValues(value) {
  const raw = String(value || '').trim();
  if (!raw || raw === 'N/A') return [];
  return raw
    .split(/[\n;,\/|]+/)
    .map((item) => String(item || '').trim())
    .filter(Boolean)
    .filter((item, index, arr) => arr.indexOf(item) === index);
}

function findTextValueByNormalizedColumn(row, exactNames = [], includesNames = []) {
  if (!row?.values) return '';
  const exact = new Set(exactNames.map(normalizeColumnTitle).filter(Boolean));
  const includes = includesNames.map(normalizeColumnTitle).filter(Boolean);

  for (const [title, cell] of Object.entries(row.values)) {
    const normalized = normalizeColumnTitle(title);
    if (!normalized) continue;
    const isExact = exact.has(normalized);
    const isIncluded = includes.some((needle) => normalized.includes(needle));
    if (!isExact && !isIncluded) continue;
    const value = cell?.display ?? cell?.raw;
    if (value != null && String(value).trim() && String(value).trim() !== 'N/A') {
      return String(value).trim();
    }
  }

  return '';
}

function getWipProjectRef(row) {
  return findTextValueByNormalizedColumn(
    row,
    [
      'Project BSP/BPP/B3D*',
      'Project BSP/BPP/B3D',
      'Project BSP/BPP/B3D *',
      'Project',
    ],
    [
      'Project BSP',
      'BSP/BPP/B3D',
      'BSP BPP B3D',
    ]
  );
}

function getWipCustomerPo(row) {
  return findTextValueByNormalizedColumn(
    row,
    [
      'Customer PO*',
      'Customer PO',
      'Customer PO *',
      'PO Cliente',
    ],
    [
      'Customer PO',
      'Cliente PO',
    ]
  );
}

function getWipClientFocalPoint(row) {
  return findTextValueByNormalizedColumn(
    row,
    [
      'Client Focal Point*',
      'Client Focal Point',
      'Client Focal Point *',
      'Focal Point',
      'Focal Point Cliente',
      'Ponto Focal Cliente',
      'Ponto Focal',
      'Client Responsible',
      'Client Representative',
      'Client Contact',
      'Responsible Client',
      'Contato Cliente',
      'Contato do Cliente',
      'Responsável Cliente',
      'Responsavel Cliente',
      'Cliente Responsável',
      'Cliente Responsavel',
    ],
    [
      'Client Focal Point',
      'Focal Point',
      'Ponto Focal',
      'Client Responsible',
      'Client Representative',
      'Client Contact',
      'Responsible Client',
      'Contato Cliente',
      'Responsavel Cliente',
      'Responsável Cliente',
    ]
  );
}

function getWipReplannedFinishDate(row) {
  return findTextValueByNormalizedColumn(
    row,
    [
      'Deadline Date as Agreeded with Client*',
      'Deadline Date as Agreeded with Client',
      'Deadline Date as Agreeded with Client *',
      'Deadline Date as Agreed with Client*',
      'Deadline Date as Agreed with Client',
      'Deadline Date as Agreed with Client *',
      'Deadline as Agreeded with Client',
      'Deadline Agreed with Client',
      'Data Replanejada',
      'Data replanejada',
      'Replanejado',
      'Replanned Date',
      'Replanned Finish Date',
    ],
    [
      'Deadline Date as Agreeded',
      'Deadline Date as Agreed',
      'Agreeded with Client',
      'Agreed with Client',
      'Data Replanejada',
      'Replanejado',
      'Replanned',
    ]
  );
}

function collectWipNameValues(value) {
  const raw = String(value || '').trim();
  if (!raw || raw === 'N/A') return [];
  return raw
    .split(/[\n;,|]+/)
    .map((item) => String(item || '').trim())
    .filter(Boolean)
    .filter((item, index, arr) => arr.findIndex((other) => other.toLowerCase() === item.toLowerCase()) === index);
}

function buildWipPoMap(rows) {
  const poMap = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    const projectRef = getWipProjectRef(row);
    const poValue = getWipCustomerPo(row);
    const key = normalizeBspLookupKey(projectRef);
    if (!key) continue;
    const values = collectWipPoValues(poValue);
    if (!values.length) continue;
    const current = poMap.get(key) || [];
    for (const value of values) {
      if (!current.includes(value)) current.push(value);
    }
    poMap.set(key, current);
  }
  return poMap;
}

function buildWipClientFocalMap(rows) {
  const focalMap = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    const projectRef = getWipProjectRef(row);
    const focalValue = getWipClientFocalPoint(row);
    const key = normalizeBspLookupKey(projectRef);
    if (!key) continue;
    const values = collectWipNameValues(focalValue);
    if (!values.length) continue;
    const current = focalMap.get(key) || [];
    for (const value of values) {
      if (!current.some((item) => String(item).toLowerCase() === String(value).toLowerCase())) current.push(value);
    }
    focalMap.set(key, current);
  }
  return focalMap;
}

function buildWipReplannedFinishMap(rows) {
  const replannedMap = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    const projectRef = getWipProjectRef(row);
    const rawDeadline = getWipReplannedFinishDate(row);
    const key = normalizeBspLookupKey(projectRef);
    if (!key || !rawDeadline) continue;

    const parsed = parseDateObject(rawDeadline);
    if (!parsed) continue;
    const formatted = formatDateValue(parsed);
    const current = replannedMap.get(key);
    if (!current || parsed > current.date) {
      replannedMap.set(key, {
        value: formatted,
        raw: String(rawDeadline).trim(),
        date: parsed,
        source: 'Work in Progress - STEP | Deadline Date as Agreeded with Client*',
      });
    }
  }
  return replannedMap;
}

function getPoListForProject(project, poMap) {
  const keys = getProjectBspLookupKeys(project);
  for (const key of keys) {
    const direct = poMap.get(key);
    if (direct?.length) return direct;
  }

  // Fallback seguro para casos em que a BSP venha com prefixo/sufixo diferente.
  for (const key of keys) {
    for (const [mapKey, values] of poMap.entries()) {
      if (!values?.length) continue;
      if (mapKey === key || (mapKey.length >= 6 && key.length >= 6 && (mapKey.endsWith(key) || key.endsWith(mapKey)))) {
        return values;
      }
    }
  }

  return [];
}

function getFocalPointListForProject(project, focalMap) {
  const keys = getProjectBspLookupKeys(project);
  for (const key of keys) {
    const direct = focalMap.get(key);
    if (direct?.length) return direct;
  }

  for (const key of keys) {
    for (const [mapKey, values] of focalMap.entries()) {
      if (!values?.length) continue;
      if (mapKey === key || (mapKey.length >= 6 && key.length >= 6 && (mapKey.endsWith(key) || key.endsWith(mapKey)))) {
        return values;
      }
    }
  }

  return [];
}

function getReplannedFinishForProject(project, replannedMap) {
  const keys = getProjectBspLookupKeys(project);
  for (const key of keys) {
    const direct = replannedMap.get(key);
    if (direct?.value) return direct;
  }

  for (const key of keys) {
    for (const [mapKey, item] of replannedMap.entries()) {
      if (!item?.value) continue;
      if (mapKey === key || (mapKey.length >= 6 && key.length >= 6 && (mapKey.endsWith(key) || key.endsWith(mapKey)))) {
        return item;
      }
    }
  }

  return null;
}

async function fetchWipStepPoMap(options = {}) {
  try {
    const sheetId = await resolveWipStepSheetId();
    if (!sheetId) return { poMap: new Map(), focalMap: new Map(), replannedFinishMap: new Map(), version: null, sheetName: WIP_STEP_SHEET_NAME, available: false };
    const wipFetchOptions = { timeoutMs: Number(options.timeoutMs || WIP_PO_FETCH_TIMEOUT_MS), retries: 1 };
    const version = await fetchSheetVersion(sheetId, wipFetchOptions);
    const memoryKey = `${sheetId}:${version || 'no-version'}`;
    const memoryCached = wipPoMemoryCache[memoryKey];
    if (memoryCached && Date.now() - Number(memoryCached.savedAt || 0) <= PROJECTS_FAST_CACHE_MS) {
      return memoryCached.data;
    }

    // v36.84: leitura leve da base de PO. Não carrega mais a WIP inteira com todas as colunas.
    const columns = await fetchSheetColumns(sheetId, wipFetchOptions);
    const columnIds = uniqueColumnIds([
      findColumnId(columns, ['Project BSP/BPP/B3D*', 'Project BSP/BPP/B3D', 'Project BSP/BPP/B3D *', 'Project'], ['Project BSP', 'BSP/BPP/B3D', 'BSP BPP B3D']),
      findColumnId(columns, ['Customer PO*', 'Customer PO', 'Customer PO *', 'PO Cliente'], ['Customer PO', 'Cliente PO']),
      findColumnId(columns, CLIENT_FOCAL_POINT_ALIASES, ['Client Focal Point', 'Focal Point', 'Ponto Focal', 'Client Responsible', 'Client Representative', 'Client Contact', 'Responsible Client', 'Contato Cliente', 'Responsavel Cliente', 'Responsável Cliente']),
      findColumnId(columns, [
        'Deadline Date as Agreeded with Client*',
        'Deadline Date as Agreeded with Client',
        'Deadline Date as Agreeded with Client *',
        'Deadline Date as Agreed with Client*',
        'Deadline Date as Agreed with Client',
        'Deadline Date as Agreed with Client *',
        'Deadline as Agreeded with Client',
        'Deadline Agreed with Client',
        'Data Replanejada',
        'Data replanejada',
        'Replanejado',
        'Replanned Date',
        'Replanned Finish Date',
      ], [
        'Deadline Date as Agreeded',
        'Deadline Date as Agreed',
        'Agreeded with Client',
        'Agreed with Client',
        'Data Replanejada',
        'Replanejado',
        'Replanned',
      ]),
    ]);

    const sheet = columnIds.length
      ? await fetchSheetWithColumns(sheetId, columnIds, wipFetchOptions)
      : await fetchFullSheet(sheetId, { timeoutMs: Math.max(WIP_PO_FETCH_TIMEOUT_MS, Number(wipFetchOptions.timeoutMs || 0)), retries: 1 });
    let rows = mapApiRows(sheet);

    const projectKeys = options?.projectKeys instanceof Set ? options.projectKeys : null;
    if (projectKeys?.size) {
      rows = rows.filter((row) => {
        const key = normalizeBspLookupKey(getWipProjectRef(row));
        if (!key) return false;
        if (projectKeys.has(key)) return true;
        for (const projectKey of projectKeys) {
          if (projectKey && key.length >= 6 && projectKey.length >= 6 && (key.endsWith(projectKey) || projectKey.endsWith(key))) return true;
        }
        return false;
      });
    }

    const data = { poMap: buildWipPoMap(rows), focalMap: buildWipClientFocalMap(rows), replannedFinishMap: buildWipReplannedFinishMap(rows), version, sheetName: sheet.name || cache.wipStepSheetName || WIP_STEP_SHEET_NAME, available: true };
    wipPoMemoryCache[memoryKey] = { savedAt: Date.now(), data };
    Object.keys(wipPoMemoryCache).forEach((key) => {
      if (Date.now() - Number(wipPoMemoryCache[key]?.savedAt || 0) > PROJECTS_FAST_CACHE_MS * 2) delete wipPoMemoryCache[key];
    });
    return data;
  } catch (error) {
    console.warn('Não foi possível carregar Work in Progress - STEP para vínculo de PO:', error.message);
    return { poMap: new Map(), focalMap: new Map(), replannedFinishMap: new Map(), version: null, sheetName: WIP_STEP_SHEET_NAME, available: false, error: error.message };
  }
}

function enrichProjectsWithCustomerPo(projects, poMap, focalMap = new Map(), replannedFinishMap = new Map()) {
  const map = poMap instanceof Map ? poMap : new Map();
  const focalPointMap = focalMap instanceof Map ? focalMap : new Map();
  const replannedMap = replannedFinishMap instanceof Map ? replannedFinishMap : new Map();
  return (Array.isArray(projects) ? projects : []).map((project) => {
    const list = getPoListForProject(project, map).filter(Boolean);
    const uniqueList = Array.from(new Set(list));
    const existingFocalList = [
      ...collectWipNameValues(project.clientFocalPoint),
      ...collectWipNameValues(project.clientFocalPointDisplay),
      ...(Array.isArray(project.clientFocalPointList) ? project.clientFocalPointList.flatMap((value) => collectWipNameValues(value)) : []),
    ];
    const focalList = [
      ...existingFocalList,
      ...getFocalPointListForProject(project, focalPointMap).filter(Boolean),
    ];
    const uniqueFocalList = [];
    for (const focal of focalList) {
      const value = String(focal || '').trim();
      if (value && !uniqueFocalList.some((item) => String(item).toLowerCase() === value.toLowerCase())) uniqueFocalList.push(value);
    }
    const replanned = getReplannedFinishForProject(project, replannedMap);
    project.customerPoList = uniqueList;
    project.customerPo = uniqueList[0] || '';
    project.customerPoDisplay = uniqueList.length ? getProjectPoDisplay(project) : 'Aguardando PO';
    project.customerPoStatus = uniqueList.length ? 'found' : 'waiting';
    project.clientFocalPointList = uniqueFocalList;
    project.clientFocalPoint = uniqueFocalList[0] || '';
    project.clientFocalPointDisplay = uniqueFocalList.join(', ');
    project.replannedFinishDate = replanned?.value || '';
    project.replannedFinishRaw = replanned?.raw || '';
    project.replannedFinishSource = replanned?.source || '';
    project.replannedFinishStatus = replanned?.value ? 'found' : 'none';
    project.clientDisplayCode = buildClientDisplayCode(project);
    return project;
  });
}

async function fetchSheetVersion(sheetId, options = {}) {
  const versionData = await apiFetch(`/sheets/${sheetId}/version`, options);
  return versionData.version;
}

async function fetchFullSheet(sheetId, options = {}) {
  return apiFetch(`/sheets/${sheetId}?includeAll=true`, {
    timeoutMs: options.timeoutMs || SMARTSHEET_FULL_SHEET_TIMEOUT_MS,
    retries: options.retries || 1,
  });
}

function getRequiredTrackingColumnIds(columns = []) {
  const exact = new Set(TRACKING_REQUIRED_COLUMN_TITLES.map(normalizeColumnTitle).filter(Boolean));
  const includes = TRACKING_REQUIRED_COLUMN_INCLUDES.map(normalizeColumnTitle).filter(Boolean);
  return uniqueColumnIds((columns || [])
    .filter((column) => {
      const normalized = normalizeColumnTitle(column?.title);
      if (!normalized) return false;
      return exact.has(normalized) || includes.some((needle) => normalized.includes(needle));
    })
    .map((column) => column.id));
}

async function fetchOperationalTrackingSheet(sheetId, options = {}) {
  try {
    const columns = await fetchSheetColumns(sheetId, {
      timeoutMs: Math.max(1500, Math.min(Number(options.timeoutMs || SMARTSHEET_FETCH_TIMEOUT_MS), 6000)),
      retries: options.retries || 1,
    });
    const columnIds = getRequiredTrackingColumnIds(columns);
    const projectColumnId = findColumnId(columns, ['Project']);
    if (!columnIds.length || (projectColumnId && !columnIds.includes(projectColumnId))) {
      return fetchFullSheet(sheetId, options);
    }
    return fetchSheetWithColumns(sheetId, columnIds, options);
  } catch (error) {
    console.warn('[projects] Leitura otimizada por colunas falhou; tentando sheet completa:', error?.message || error);
    return fetchFullSheet(sheetId, options);
  }
}

async function fetchSheetColumns(sheetId, options = {}) {
  const cacheKey = String(sheetId || '');
  const cached = smartsheetColumnsCache[cacheKey];
  if (cached?.columns && Date.now() - Number(cached.savedAt || 0) <= 6 * 60 * 60 * 1000) {
    return cached.columns;
  }
  const response = await apiFetch(`/sheets/${sheetId}/columns?includeAll=true`, options);
  const columns = Array.isArray(response?.data) ? response.data : (Array.isArray(response?.columns) ? response.columns : []);
  smartsheetColumnsCache[cacheKey] = { savedAt: Date.now(), columns };
  return columns;
}

function findColumnId(columns = [], exactNames = [], includesNames = []) {
  const exact = new Set((exactNames || []).map(normalizeColumnTitle).filter(Boolean));
  const includes = (includesNames || []).map(normalizeColumnTitle).filter(Boolean);

  const exactFound = columns.find((column) => exact.has(normalizeColumnTitle(column?.title)));
  if (exactFound?.id) return String(exactFound.id);

  const includesFound = columns.find((column) => {
    const normalized = normalizeColumnTitle(column?.title);
    return normalized && includes.some((needle) => normalized.includes(needle));
  });
  return includesFound?.id ? String(includesFound.id) : '';
}

function uniqueColumnIds(ids = []) {
  return Array.from(new Set((ids || []).map((id) => String(id || '').trim()).filter(Boolean)));
}

async function fetchSheetWithColumns(sheetId, columnIds = [], options = {}) {
  const ids = uniqueColumnIds(columnIds);
  if (!ids.length) return fetchFullSheet(sheetId, options);
  return apiFetch(`/sheets/${sheetId}?includeAll=true&columnIds=${encodeURIComponent(ids.join(','))}`, {
    timeoutMs: options.timeoutMs || SMARTSHEET_ROWS_FETCH_TIMEOUT_MS,
    retries: options.retries || 1,
  });
}

async function fetchSheetRowsByIds(sheetId, rowIds = [], columnIds = [], options = {}) {
  const ids = Array.from(new Set((rowIds || []).map((id) => String(id || '').trim()).filter(Boolean)));
  if (!ids.length) {
    return { name: cache.sheetName || SHEET_NAME, columns: [], rows: [] };
  }

  const chunks = [];
  for (let i = 0; i < ids.length; i += CLIENT_SMARTSHEET_ROW_CHUNK_SIZE) {
    chunks.push(ids.slice(i, i + CLIENT_SMARTSHEET_ROW_CHUNK_SIZE));
  }

  let mergedSheet = null;
  const allRows = [];
  for (const chunk of chunks) {
    const params = [`rowIds=${encodeURIComponent(chunk.join(','))}`];
    const colIds = uniqueColumnIds(columnIds);
    if (colIds.length) params.push(`columnIds=${encodeURIComponent(colIds.join(','))}`);
    const sheet = await apiFetch(`/sheets/${sheetId}?${params.join('&')}`, {
      timeoutMs: options.timeoutMs || SMARTSHEET_ROWS_FETCH_TIMEOUT_MS,
      retries: options.retries || 1,
    });
    if (!mergedSheet) mergedSheet = sheet;
    allRows.push(...(Array.isArray(sheet?.rows) ? sheet.rows : []));
  }

  return {
    ...(mergedSheet || {}),
    rows: allRows.sort((a, b) => Number(a?.rowNumber || 0) - Number(b?.rowNumber || 0)),
  };
}

function isWarmPayloadCache() {
  const lastCheck = Number(cache.lastVersionCheck || 0);
  return Boolean(cache.payload && lastCheck > 0 && Date.now() - lastCheck <= PROJECTS_FAST_CACHE_MS);
}

function cloneCachedPayloadWithMeta(extraMeta = {}) {
  if (!cache.payload) return null;
  return {
    ...cache.payload,
    meta: {
      ...(cache.payload.meta || {}),
      ...extraMeta,
    },
  };
}


function payloadProjectHasResolvedCustomerPo(project = {}) {
  const values = [
    project.customerPo,
    project.customerPoDisplay,
    ...(Array.isArray(project.customerPoList) ? project.customerPoList : []),
  ];
  return values.some((value) => {
    const text = String(value || '').trim();
    return text && !/aguardando\s+po/i.test(text);
  });
}

function isPayloadOperationallyComplete(payload = {}) {
  const projects = Array.isArray(payload.projects) ? payload.projects : [];
  if (!projects.length) return true;

  // Quando a base complementar de PO respondeu, o payload pode ser usado em cache.
  // Algumas BSPs isoladas podem continuar sem PO por falta de correspondência real.
  if (payload.meta?.wipStepPoAvailable === true) return true;

  // Snapshots antigos tinham todas as BSPs sem PO. Eles não podem ser devolvidos como
  // resposta rápida, senão a guia anônima abre o portal sem PO até o usuário dar F5.
  return projects.some(payloadProjectHasResolvedCustomerPo);
}


function getFullPersistentCacheKey() {
  // v37.08 FREE PLAN SAFE: uma única linha no Supabase por ambiente.
  // Isso evita crescimento de banco e elimina cache por cliente.
  return `projects:${OPERATION_REGION}:current`;
}

function getClientPersistentCacheScope(session = {}) {
  return getClientScopeValues(session).sort().join('|') || normalizeClientScopeValue(session.username || session.name || session.sub || 'client');
}

function getClientPersistentCacheKey(session = {}) {
  // Compatibilidade nominal: cliente não grava linha própria no Supabase.
  // A leitura é feita pelo cache completo e filtrada no backend.
  return getFullPersistentCacheKey();
}

function clonePersistentPayload(payload, reason = 'persistent-cache', cacheInfo = {}) {
  if (!payload) return null;
  const ageMs = cacheInfo.ageMs == null ? null : Number(cacheInfo.ageMs);
  const autoRefreshRecommended = ageMs != null && ageMs >= TRACKING_CACHE_AUTO_REFRESH_AFTER_MS;
  return {
    ...payload,
    meta: {
      ...(payload.meta || {}),
      servedFromPersistentCache: true,
      stale: autoRefreshRecommended,
      cacheReason: reason,
      persistentCacheServedAt: new Date().toISOString(),
      persistentCacheUpdatedAt: cacheInfo.updatedAt || null,
      persistentCacheAgeMs: ageMs,
      persistentCacheAutoRefreshAfterMs: TRACKING_CACHE_AUTO_REFRESH_AFTER_MS,
      persistentCacheLastWriteReason: cacheInfo.lastWriteReason || null,
      persistentCacheRefreshStartedAt: cacheInfo.refreshStartedAt || null,
      persistentCacheRefreshLockUntil: cacheInfo.refreshLockUntil || null,
      persistentCacheRefreshLockOwner: cacheInfo.refreshLockOwner || null,
      autoRefreshRecommended,
    },
  };
}


function cloneBundledFallbackPayload(reason = 'bundled-fallback') {
  if (!ALLOW_BUNDLED_FALLBACK) {
    console.warn('[projects] Fallback empacotado bloqueado em produção:', reason);
    return null;
  }
  const fallbackPayload = readBundledFallbackPayloadSync();
  if (!fallbackPayload || !Array.isArray(fallbackPayload.projects) || !fallbackPayload.projects.length) return null;
  return {
    ...fallbackPayload,
    ok: true,
    meta: {
      ...(fallbackPayload.meta || {}),
      servedFromBundledFallback: true,
      fallbackMayBeIncomplete: !isPayloadOperationallyComplete(fallbackPayload),
      stale: true,
      cacheReason: reason,
      fallbackServedAt: new Date().toISOString(),
      operationalUseBlockedByDefault: true,
    },
  };
}

async function readBestAvailableFullPayload(reason = 'best-available-cache') {
  const staleMemoryPayload = cloneCachedPayloadWithMeta({ stale: true, staleReason: reason, cacheReason: reason });
  if (staleMemoryPayload && Array.isArray(staleMemoryPayload.projects) && staleMemoryPayload.projects.length && isPayloadOperationallyComplete(staleMemoryPayload)) {
    return staleMemoryPayload;
  }

  const persistentPayload = await readPersistentFullPayload(`persistent-cache-${reason}`);
  if (persistentPayload) return persistentPayload;

  const bundledPayload = cloneBundledFallbackPayload(`bundled-fallback-${reason}`);
  if (bundledPayload) return bundledPayload;

  return null;
}

async function readBestAvailablePayloadForSession(session = {}, reason = 'best-available-cache-client') {
  if (!session || session.role !== 'client') return readBestAvailableFullPayload(reason);

  const scopedMemoryPayload = cache.payload
    ? scopePayloadForSession(cloneCachedPayloadWithMeta({ stale: true, staleReason: reason, cacheReason: reason }) || cache.payload, session)
    : null;
  if (scopedMemoryPayload && isUsableClientPayload(scopedMemoryPayload, session)) return scopedMemoryPayload;

  const persistentPayload = await readPersistentPayloadForSession(session, `persistent-cache-${reason}`);
  if (persistentPayload) return persistentPayload;

  const bundledPayload = cloneBundledFallbackPayload(`bundled-fallback-${reason}`);
  if (bundledPayload) {
    const scopedBundledPayload = scopePayloadForSession(bundledPayload, session);
    if (isUsableClientPayload(scopedBundledPayload, session)) {
      scopedBundledPayload.meta = {
        ...(scopedBundledPayload.meta || {}),
        servedFromBundledFallback: true,
        stale: true,
        cacheReason: `bundled-fallback-${reason}`,
        fallbackServedAt: new Date().toISOString(),
      };
      return scopedBundledPayload;
    }
  }

  return null;
}

async function readPersistentFullPayload(reason = 'persistent-cache', options = {}) {
  const maxAgeMs = options.loginRead ? PERSISTENT_CACHE_LOGIN_MAX_AGE_MS : PERSISTENT_CACHE_MAX_AGE_MS;
  const cached = await readTrackingCache(getFullPersistentCacheKey(), { maxAgeMs });
  const payload = cached?.payload;
  if (!payload || !Array.isArray(payload.projects) || !payload.projects.length) return null;
  if (!isPayloadOperationallyComplete(payload)) return null;
  const updatedAtMs = cached.updatedAt ? new Date(cached.updatedAt).getTime() : 0;
  return clonePersistentPayload(payload, reason, {
    updatedAt: cached.updatedAt || null,
    ageMs: updatedAtMs > 0 ? Date.now() - updatedAtMs : null,
    lastWriteReason: cached.lastWriteReason || null,
    refreshStartedAt: cached.refreshStartedAt || null,
    refreshLockUntil: cached.refreshLockUntil || null,
    refreshLockOwner: cached.refreshLockOwner || null,
  });
}

async function readPersistentPayloadForSession(session = {}, reason = 'persistent-cache-client', options = {}) {
  // v37.08 FREE PLAN SAFE: não existe mais cache persistente separado por cliente.
  // O backend lê uma única base operacional e filtra a carteira do cliente antes de responder.
  const fullPayload = await readPersistentFullPayload(reason, options);
  if (!fullPayload) return null;
  if (!session || session.role !== 'client') return fullPayload;
  const scoped = scopePayloadForSession(fullPayload, session);
  // v37.14: se o cliente possui escopo explícito e a carteira filtrada é vazia,
  // devolve a resposta vazia rapidamente pelo cache persistente. Isso evita cair no
  // Smartsheet ao vivo a cada login quando o cliente não possui BSPs válidas ou quando
  // o filtro exato remove empresas do mesmo grupo, como BW ENERGY x BW LNG.
  if (isUsableClientPayload(scoped, session)) {
    scoped.meta = {
      ...(scoped.meta || {}),
      servedFromPersistentCache: true,
      stale: true,
      cacheReason: reason,
      persistentCacheMode: 'single-full-cache-filtered-by-backend',
      persistentCacheServedAt: new Date().toISOString(),
      emptyClientScopeResolved: Array.isArray(scoped.projects) && scoped.projects.length === 0,
    };
    return scoped;
  }
  return null;
}

function persistFullPayloadInBackground(payload, options = {}) {
  if (!payload || !Array.isArray(payload.projects) || !payload.projects.length) return Promise.resolve(false);
  return writeTrackingCache(getFullPersistentCacheKey(), payload, {
    scope: 'single-current-cache',
    source: String(options.source || `projects-${OPERATION_REGION}`),
    version: payload?.meta?.version || '',
    reason: String(options.reason || 'full-refresh-single-row-free-plan'),
    forceWrite: Boolean(options.forceWrite),
  }).catch(() => null);
}

function persistClientPayloadInBackground(session = {}, payload) {
  // v37.08 FREE PLAN SAFE: não grava cache por cliente no Supabase.
  // Cliente usa a base única persistente, filtrada no backend, ou cache em memória/local.
  return false;
}

function buildCacheUnavailableError() {
  const error = new Error('Cache operacional do Supabase ainda não disponível. Use o botão Atualizar ou aguarde a rotina agendada de 15 minutos.');
  error.code = 'PERSISTENT_CACHE_UNAVAILABLE';
  return error;
}

async function readLoginCacheOnlyPayloadForSession(session = {}, reason = 'login-cache-only') {
  if (session?.role === 'client') {
    return readPersistentPayloadForSession(session, reason, { loginRead: true });
  }
  return readPersistentFullPayload(reason, { loginRead: true });
}


async function fetchCurrentTrackingSheetVersion(options = {}) {
  const sheetId = await resolveSheetId();
  const version = await fetchSheetVersion(sheetId, options);
  return { sheetId, version };
}

async function buildPayload(options = {}) {
  if (!TOKEN) {
    const fallbackPayload = cloneBundledFallbackPayload('missing-smartsheet-token');
    if (fallbackPayload) return fallbackPayload;
    throw new Error("SMARTSHEET_TOKEN não configurado.");
  }
  const force = Boolean(options.force);
  const preferCache = Boolean(options.preferCache);
  const bypassMemoryCache = Boolean(options.bypassMemoryCache || options.persistentFirst);

  // v37.20: quando a leitura é cache-only/preferCache, prioriza o Supabase persistente
  // antes do cache em memória da Function. Isso evita exibir uma hora antiga logo após
  // o botão Atualizar ou após a rotina agendada gravar uma base nova em outra instância.
  if (!force && preferCache) {
    const persistentPayload = await readPersistentFullPayload('persistent-cache-login-cache-only', { loginRead: true });
    if (persistentPayload) return persistentPayload;
  }

  // Mantém o Portal do Cliente e o painel interno rápidos em F5/login:
  // se a função Netlify ainda estiver aquecida e uma base completa foi validada há pouco,
  // responde pelo cache em memória sem reler Tracking + base de PO.
  if (!force && !bypassMemoryCache && isWarmPayloadCache() && isPayloadOperationallyComplete(cache.payload)) {
    return cache.payload;
  }

  // Caminho crítico de login: se já há um payload válido em memória, devolver primeiro
  // a última versão conhecida evita que a tela fique bloqueada por checagem de versão
  // do Smartsheet. O botão "Atualizar agora" continua usando force=1 e busca fresco.
  if (!force && preferCache && !bypassMemoryCache && cache.payload && isPayloadOperationallyComplete(cache.payload)) {
    return cloneCachedPayloadWithMeta({
      ...(cache.payload.meta || {}),
      servedFromFastCache: true,
      cacheReason: 'prefer-cache',
    }) || cache.payload;
  }

  // v37.17/v37.20: se o Supabase persistente não existir, só então tenta fallback permitido.
  // O login normal não chama Smartsheet quando PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS=0.
  if (!force && preferCache) {
    const bundledPayload = cloneBundledFallbackPayload('bundled-fallback-prefer-cache');
    if (bundledPayload) return bundledPayload;
    if (!PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS) throw buildCacheUnavailableError();
  }

  const sheetId = await resolveSheetId();
  let version;
  try {
    version = await fetchSheetVersion(sheetId);
    cache.lastVersionCheck = Date.now();
  } catch (error) {
    const fallbackPayload = !force ? await readBestAvailableFullPayload('version-check-failed') : null;
    if (fallbackPayload) return fallbackPayload;
    throw error;
  }

  // Preserva a velocidade: se a base principal não mudou, não força leitura da base complementar de PO.
  // O botão Atualizar agora usa force=1 para recalcular tudo quando necessário.
  if (!force && cache.payload && cache.version === version && isPayloadOperationallyComplete(cache.payload)) {
    return cache.payload;
  }

  const wipPoPromise = fetchWipStepPoMap({ timeoutMs: Number(options.wipPoTimeoutMs || WIP_PO_FETCH_TIMEOUT_MS) }).catch((error) => ({
    poMap: new Map(),
    focalMap: new Map(),
    replannedFinishMap: new Map(),
    version: cache.wipStepVersion || null,
    sheetName: cache.wipStepSheetName || WIP_STEP_SHEET_NAME,
    available: false,
    error: error.message,
  }));

  let wipPoData;
  let sheet;
  try {
    [wipPoData, sheet] = await Promise.all([wipPoPromise, fetchOperationalTrackingSheet(sheetId, { timeoutMs: Number(options.fullSheetTimeoutMs || SMARTSHEET_FULL_SHEET_TIMEOUT_MS) })]);
  } catch (error) {
    const fallbackPayload = !force ? await readBestAvailableFullPayload('sheet-fetch-failed') : null;
    if (fallbackPayload) return fallbackPayload;
    throw error;
  }
  const rows = mapApiRows(sheet);
  const projects = enrichProjectsWithCustomerPo(buildProjects(rows), wipPoData.poMap, wipPoData.focalMap, wipPoData.replannedFinishMap);
  const stats = buildStats(projects);
  const alertData = buildAlerts(projects);

  const payload = {
    ok: true,
    meta: {
      sheetId,
      sheetName: sheet.name || cache.sheetName || SHEET_NAME,
      version,
      wipStepSheetName: wipPoData.sheetName || WIP_STEP_SHEET_NAME,
      wipStepVersion: wipPoData.version || null,
      wipStepPoAvailable: Boolean(wipPoData.available),
      lastSync: new Date().toISOString(),
      stageOrder: STAGE_ORDER.filter((stage) => !stage.ignoredCurrentStage).map((stage) => ({
        key: stage.key,
        label: stage.label,
        type: stage.type,
        optional: Boolean(stage.optional),
      })),
      currentWeek: getProductionWeekLabel(getCurrentBrazilDateObject()),
      alertSignature: alertData.signature,
    },
    stats,
    alerts: alertData.alerts,
    projects,
  };

  cache.sheetId = sheetId;
  cache.sheetName = payload.meta.sheetName;
  cache.version = version;
  cache.wipStepSheetName = payload.meta.wipStepSheetName;
  cache.wipStepVersion = payload.meta.wipStepVersion;
  cache.lastSync = payload.meta.lastSync;
  cache.lastVersionCheck = Date.now();
  cache.payload = payload;

  // v37.07/v37.15: Supabase é o cache persistente confiável entre cold starts.
  const persistentWritePromise = persistFullPayloadInBackground(payload, {
    forceWrite: Boolean(options.forcePersistentCacheWrite),
    reason: options.persistentWriteReason || (options.forcePersistentCacheWrite ? 'manual-or-scheduled-refresh-force-write' : 'full-refresh-single-row-free-plan'),
    source: options.persistentWriteSource || `projects-${OPERATION_REGION}`,
  });
  if (options.waitForPersistentCacheWrite) {
    await persistentWritePromise;
  }

  // v37.64: QR Codes de ISO são criados automaticamente quando o cache do Tracking é atualizado.
  // Não bloqueia o painel se a tabela ainda não foi criada ou se o Supabase oscilar.
  const isoQrSyncPromise = ensureIsoQrCodesForPayload(payload, {
    source: options.persistentWriteSource || `projects-${OPERATION_REGION}`,
    baseUrl: options.baseUrl || '',
    timeoutMs: Number(process.env.ISO_QR_SUPABASE_TIMEOUT_MS || 9000),
  }).catch((error) => {
    console.warn('[iso-qr] Geração automática ignorada:', error?.message || error);
    return null;
  });
  if (options.waitForIsoQrSync || options.waitForPersistentCacheWrite) {
    await isoQrSyncPromise;
  }

  // Mantém snapshot local apenas como compatibilidade operacional; Netlify pode ter filesystem efêmero.
  savePayloadToDisk(payload).catch(err => console.error('[v32] Erro ao salvar snapshot:', err.message));

  return payload;
}

function resolveFallbackPath() {
  const candidates = [
    path.resolve(__dirname, 'fallback-projects.json'),
    path.resolve(__dirname, '..', 'data', 'fallback-projects.json'),
    path.resolve(__dirname, '..', '..', 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'data', 'fallback-projects.json'),
  ];
  return candidates.find((filePath) => fs.existsSync(filePath)) || candidates[0];
}

function readBundledFallbackPayloadSync() {
  const candidates = [
    path.resolve(__dirname, 'fallback-projects.json'),
    path.resolve(__dirname, '..', 'data', 'fallback-projects.json'),
    path.resolve(__dirname, '..', '..', 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'data', 'fallback-projects.json'),
  ];
  for (const filePath of candidates) {
    try {
      if (fs.existsSync(filePath)) {
        const raw = fs.readFileSync(filePath, 'utf8');
        return JSON.parse(raw);
      }
    } catch (err) {
      console.error('[v36.91] Erro ao ler fallback em disco:', filePath, err.message);
    }
  }
  return null;
}

async function savePayloadToDisk(payload) {
  try {
    const filePath = resolveFallbackPath();
    const dirPath = path.dirname(filePath);
    if (!fs.existsSync(dirPath)) {
      await fsPromises.mkdir(dirPath, { recursive: true });
    }
    // Adiciona marcação de que este é um snapshot persistido
    const diskPayload = {
      ...payload,
      meta: {
        ...(payload.meta || {}),
        source: 'disk-snapshot',
        persistedAt: new Date().toISOString()
      }
    };
    await fsPromises.writeFile(filePath, JSON.stringify(diskPayload, null, 2), 'utf8');
    console.log('[v32] Snapshot salvo com sucesso em:', filePath);
  } catch (err) {
    console.error('[v32] Erro ao salvar snapshot em disco:', err.message);
  }
}


function normalizeClientScopeValue(value) {
  const normalized = String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    // v37.14: preserva separadores como espaços. Ex.: BW_ENERGY_BR => "bw energy br".
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
  return normalized;
}

const CLIENT_SCOPE_GENERIC_WORDS = new Set([
  'de', 'da', 'do', 'das', 'dos', 'e', 'em', 'the', 'and', 'of', 'a', 'an',
  'sa', 's', 'ltda', 'ltd', 'llc', 'inc', 'corp', 'company', 'companhia',
  'brasil', 'brazil', 'global', 'international', 'internacional', 'energy',
  'energia', 'offshore', 'oil', 'gas', 'petroleo', 'petroleum', 'services',
  'service', 'servicos', 'solucoes', 'solutions', 'industrial', 'industria',
  'cliente', 'client', 'portal', 'usuario', 'user', 'login'
]);

function stripClientRegionSuffix(value) {
  const normalized = normalizeClientScopeValue(value);
  if (!normalized) return '';
  return normalized
    .replace(/\s+(br|pt|brazil|brasil|portugal)$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function addUniqueClientScopeValue(values, value) {
  const normalized = normalizeClientScopeValue(value);
  if (!normalized) return;
  if (!values.some((item) => normalizeClientScopeValue(item) === normalized)) values.push(value);
}

function expandClientScopeAliases(rawValues = []) {
  const values = [];
  const add = (value) => addUniqueClientScopeValue(values, value);

  rawValues.forEach(add);

  // Também aceita clientKey cadastrado com sufixo/região, sem transformar isso em filtro amplo.
  // Ex.: BW_ENERGY_BR libera apenas "BW ENERGY", nunca "BW LNG".
  rawValues.forEach((value) => {
    const stripped = stripClientRegionSuffix(value);
    if (stripped) add(stripped);
  });

  const normalized = values.map((value) => normalizeClientScopeValue(value)).filter(Boolean);
  const compacted = normalized.map((value) => value.replace(/[^a-z0-9]+/g, ''));
  const hasToken = (...tokens) => tokens.some((token) => normalized.includes(token) || compacted.includes(String(token).replace(/[^a-z0-9]+/g, '')));

  // v36.79: segurança para o Portal Portugal.
  // Algumas bases do Tracking vêm com Client = "STEP PORTUGAL", enquanto o login do cliente é SBM/SBM_PT.
  // Esta expansão é explícita e limitada ao cliente SBM, sem liberar grupos por palavra parecida.
  if (hasToken('sbm', 'sbm pt', 'sbmpt')) {
    add('STEP PORTUGAL');
    add('STEP PORTUGAL PT');
    add('SBM Offshore');
    add('SBM OFFSHORE');
  }

  return values;
}

function isValidClientScopeValue(value) {
  const normalized = normalizeClientScopeValue(value);
  if (!normalized || normalized.length < 2) return false;
  const words = normalized.split(/\s+/).filter(Boolean);
  if (words.length === 1 && CLIENT_SCOPE_GENERIC_WORDS.has(words[0])) return false;
  return true;
}

function getClientScopeValues(session = {}) {
  // v37.14: escopo de cliente agora é estrito por nome exato/alias explícito.
  // Não usamos mais a primeira palavra como fallback, pois BW ENERGY e BW LNG pertencem ao grupo BW,
  // mas são clientes diferentes e não podem compartilhar carteira por semelhança.
  const configuredScopes = [
    session.clientKey,
    session.clientName,
    ...(Array.isArray(session.allowedClients) ? session.allowedClients : []),
  ].filter((value) => String(value || '').trim());

  // Fallback legado apenas se a sessão realmente não trouxer clientKey/clientName/allowedClients.
  // Isso evita que nome de usuário seja usado como escopo quando já existe cliente cadastrado.
  const rawValues = configuredScopes.length
    ? configuredScopes
    : [session.name, session.username, session.rawUsername].filter((value) => String(value || '').trim());

  return expandClientScopeAliases(rawValues)
    .map((value) => normalizeClientScopeValue(value))
    .filter(Boolean)
    .filter(isValidClientScopeValue)
    .filter((value, index, arr) => arr.indexOf(value) === index);
}

function projectBelongsToClientScope(project, session = {}) {
  if (!project) return false;
  const scopeValues = getClientScopeValues(session);
  if (!scopeValues.length) return false;

  const client = normalizeClientScopeValue(project.client);
  if (!client) return false;

  // v37.14: comparação exata. Agrupamentos só podem ocorrer se o cliente estiver cadastrado
  // explicitamente em allowed_clients no Supabase. Ex.: BW ENERGY não enxerga BW LNG por "BW".
  if (scopeValues.includes(client)) return true;

  const clientWithoutRegion = stripClientRegionSuffix(client);
  if (clientWithoutRegion && scopeValues.includes(clientWithoutRegion)) return true;

  return false;
}


function mergeHydratedClientSession(session = {}, freshUser = {}) {
  if (!freshUser) return session;
  return {
    ...session,
    ...freshUser,
    sub: session.sub,
    username: freshUser.username || session.username,
    role: freshUser.role || session.role,
    clientKey: freshUser.clientKey || session.clientKey || '',
    clientName: freshUser.clientName || session.clientName || freshUser.clientKey || '',
    clientLogoUrl: freshUser.clientLogoUrl || session.clientLogoUrl || '',
    clientPlatformImageUrl: freshUser.clientPlatformImageUrl || session.clientPlatformImageUrl || '',
    clientPlatformImages: freshUser.clientPlatformImages || session.clientPlatformImages || {},
    allowedClients: Array.isArray(freshUser.allowedClients) && freshUser.allowedClients.length
      ? freshUser.allowedClients
      : (Array.isArray(session.allowedClients) ? session.allowedClients : []),
  };
}

function pruneSessionHydrationCache(now = Date.now()) {
  Object.keys(sessionHydrationCache).forEach((key) => {
    const savedAt = Number(sessionHydrationCache[key]?.savedAt || 0);
    if (!savedAt || now - savedAt > SESSION_HYDRATION_CACHE_MS * 2) {
      delete sessionHydrationCache[key];
    }
  });
}

async function hydrateClientSession(session = {}) {
  if (!session || session.role !== 'client') return session;

  // Sessões antigas podem estar sem clientKey/clientName no cookie.
  // O frontend mostra o usuário correto via /api/auth-me, mas /api/projects usa o cookie.
  // Para evitar carteira vazia no Portal do Cliente, reidratamos a sessão pelo Supabase.
  if (!isSupabaseConfigured() || !session.sub) return session;

  const cacheKey = String(session.sub);
  const now = Date.now();
  const cached = sessionHydrationCache[cacheKey];
  if (cached?.user && now - Number(cached.savedAt || 0) <= SESSION_HYDRATION_CACHE_MS) {
    return mergeHydratedClientSession(session, cached.user);
  }

  try {
    const freshUser = await getUserById(session.sub);
    if (!freshUser) return session;
    sessionHydrationCache[cacheKey] = { savedAt: now, user: freshUser };
    pruneSessionHydrationCache(now);
    return mergeHydratedClientSession(session, freshUser);
  } catch (error) {
    // Se o Supabase oscilar, usa a última hidratação válida quando existir; caso contrário, mantém o cookie.
    if (cached?.user) return mergeHydratedClientSession(session, cached.user);
    return session;
  }
}

function scopePayloadForSession(payload, session = {}) {
  if (!payload || session.role !== 'client') return payload;
  const projects = (Array.isArray(payload.projects) ? payload.projects : []).filter((project) => projectBelongsToClientScope(project, session));
  const stats = buildStats(projects);
  const alertData = buildAlerts(projects);
  return {
    ...payload,
    stats,
    alerts: alertData.alerts,
    projects,
    meta: {
      ...(payload.meta || {}),
      clientPortal: true,
      clientName: session.clientName || session.clientKey || session.name || 'Cliente',
      clientKey: session.clientKey || '',
      clientLogoUrl: session.clientLogoUrl || '',
      alertSignature: alertData.signature,
    },
  };
}


function clientTextBelongsToScope(clientText, session = {}) {
  return projectBelongsToClientScope({ client: clientText }, session);
}

function collectClientRelevantRowIds(indexRows = [], session = {}) {
  const rowsById = new Map(indexRows.map((row) => [String(row.id), row]));
  const childrenByParent = new Map();
  const parentById = new Map();

  for (const row of indexRows) {
    const id = String(row.id || '');
    const parentId = row.parentId ? String(row.parentId) : '';
    if (!id || !parentId) continue;
    parentById.set(id, parentId);
    if (!childrenByParent.has(parentId)) childrenByParent.set(parentId, []);
    childrenByParent.get(parentId).push(id);
  }

  const matched = new Set();
  for (const row of indexRows) {
    const client = textValue(row, 'Client');
    if (clientTextBelongsToScope(client, session)) matched.add(String(row.id));
  }

  const include = new Set();
  const includeAncestors = (id) => {
    let current = String(id || '');
    let guard = 0;
    while (current && rowsById.has(current) && guard < 100) {
      include.add(current);
      current = parentById.get(current) || '';
      guard += 1;
    }
  };
  const includeDescendants = (id) => {
    const stack = [String(id || '')];
    while (stack.length) {
      const current = stack.pop();
      if (!current || include.has(current)) continue;
      include.add(current);
      const children = childrenByParent.get(current) || [];
      for (const childId of children) stack.push(childId);
    }
  };

  for (const id of matched) {
    includeAncestors(id);
    includeDescendants(id);
  }

  return Array.from(include);
}

function collectProjectLookupKeySet(projects = []) {
  const keys = new Set();
  for (const project of Array.isArray(projects) ? projects : []) {
    getProjectBspLookupKeys(project).forEach((key) => { if (key) keys.add(key); });
  }
  return keys;
}

function getClientPayloadCacheKey(sheetId, version, session = {}) {
  const scope = getClientScopeValues(session).sort().join('|') || normalizeClientScopeValue(session.username || session.name || session.sub || 'client');
  return `${sheetId}:${version || 'no-version'}:${scope}`;
}

function getCachedClientPayload(cacheKey) {
  const cached = clientPayloadCache[cacheKey];
  if (!cached?.payload) return null;
  if (Date.now() - Number(cached.savedAt || 0) > CLIENT_PROJECTS_FAST_CACHE_MS) return null;
  return cached.payload;
}

function isNonEmptyClientPayload(payload = {}) {
  return Boolean(payload?.meta?.clientPortal && Array.isArray(payload.projects) && payload.projects.length > 0);
}

function hasExplicitClientScope(session = {}) {
  if (!session || session.role !== 'client') return false;
  const explicitSession = {
    clientKey: session.clientKey,
    clientName: session.clientName,
    allowedClients: Array.isArray(session.allowedClients) ? session.allowedClients : [],
  };
  return getClientScopeValues(explicitSession).length > 0;
}

function isUsableClientPayload(payload = {}, session = {}) {
  if (!payload?.meta?.clientPortal || !Array.isArray(payload.projects)) return false;
  if (payload.projects.length > 0) return true;
  return hasExplicitClientScope(session);
}

function setCachedClientPayload(cacheKey, payload) {
  // v37.04: nunca cachear carteira vazia do Portal do Cliente.
  // Se a primeira consulta ocorrer com sessão ainda não hidratada ou Smartsheet oscilando,
  // cachear 0 BSP(s) deixa o cliente preso com cards "--" mesmo depois da API voltar.
  if (payload?.meta?.clientPortal && (!Array.isArray(payload.projects) || payload.projects.length === 0) && !payload?.meta?.emptyClientScopeResolved) {
    console.warn('[client-cache] Payload vazio do Portal do Cliente não foi cacheado:', cacheKey);
    return;
  }

  clientPayloadCache[cacheKey] = { savedAt: Date.now(), payload };
  const keys = Object.keys(clientPayloadCache);
  if (keys.length > 60) {
    keys.sort((a, b) => Number(clientPayloadCache[a]?.savedAt || 0) - Number(clientPayloadCache[b]?.savedAt || 0))
      .slice(0, Math.max(0, keys.length - 60))
      .forEach((key) => delete clientPayloadCache[key]);
  }
}

async function buildClientPayload(session = {}, options = {}) {
  if (!TOKEN) {
    const fallbackPayload = await readBestAvailablePayloadForSession(session, 'missing-smartsheet-token-client');
    if (fallbackPayload) return fallbackPayload;
    throw new Error('SMARTSHEET_TOKEN não configurado.');
  }
  const force = Boolean(options.force);
  const preferCache = Boolean(options.preferCache);
  const bypassMemoryCache = Boolean(options.bypassMemoryCache || options.persistentFirst);

  // v37.20: Portal do Cliente deve ler primeiro a linha atual do Supabase quando
  // preferCache=true. Assim o horário exibido acompanha imediatamente o cache persistente,
  // sem reaproveitar uma carteira antiga em memória da Function.
  if (!force && preferCache) {
    const persistentClientPayload = await readPersistentPayloadForSession(session, 'persistent-cache-client-login-cache-only', { loginRead: true });
    if (persistentClientPayload) return persistentClientPayload;
  }

  // Se o payload completo já estiver aquecido, só filtra em memória.
  if (!force && !bypassMemoryCache && isWarmPayloadCache() && isPayloadOperationallyComplete(cache.payload)) {
    const scopedWarmPayload = scopePayloadForSession(cache.payload, session);
    if (isUsableClientPayload(scopedWarmPayload, session)) return scopedWarmPayload;
  }

  // v37.17/v37.20: Portal do Cliente não consulta Smartsheet no login.
  // Se o Supabase persistente não estiver pronto, tenta apenas fallback seguro já existente.
  if (!force && preferCache) {
    const fallbackPayload = await readBestAvailablePayloadForSession(session, 'client-prefer-cache');
    if (fallbackPayload) return fallbackPayload;
    if (!PROJECTS_ALLOW_SMARTSHEET_ON_CACHE_MISS) throw buildCacheUnavailableError();
  }

  const sheetId = await resolveSheetId();
  let version = null;
  try {
    version = await fetchSheetVersion(sheetId);
    cache.lastVersionCheck = Date.now();
  } catch (error) {
    const fallbackPayload = !force ? await readBestAvailablePayloadForSession(session, 'client-version-check-failed') : null;
    if (fallbackPayload) return fallbackPayload;
    throw error;
  }

  const cacheKey = getClientPayloadCacheKey(sheetId, version, session);
  if (!force) {
    const cached = getCachedClientPayload(cacheKey);
    // v37.04: cache de cliente só é aceito quando tem pelo menos 1 BSP.
    // Isso evita reaproveitar uma resposta vazia gerada por sessão incompleta ou leitura parcial.
    if (cached && isUsableClientPayload(cached, session) && (preferCache || isPayloadOperationallyComplete(cached))) return cached;
  }

  let columns;
  try {
    columns = await fetchSheetColumns(sheetId);
  } catch (error) {
    const fallbackPayload = !force ? await readBestAvailablePayloadForSession(session, 'client-columns-failed') : null;
    if (fallbackPayload) return fallbackPayload;
    throw error;
  }
  const projectColumnId = findColumnId(columns, ['Project']);
  const clientColumnId = findColumnId(columns, ['Client']);
  if (!clientColumnId) {
    throw new Error('Coluna Client não encontrada na planilha principal.');
  }

  const indexColumnIds = uniqueColumnIds([projectColumnId, clientColumnId]);
  const indexCacheKey = `${sheetId}:${version || 'no-version'}:${indexColumnIds.join(',')}`;
  let indexSheet = smartsheetIndexCache[indexCacheKey]?.sheet || null;
  if (!indexSheet) {
    try {
      indexSheet = await fetchSheetWithColumns(sheetId, indexColumnIds);
    } catch (error) {
      const fallbackPayload = !force ? await readBestAvailablePayloadForSession(session, 'client-index-failed') : null;
      if (fallbackPayload) return fallbackPayload;
      throw error;
    }
    smartsheetIndexCache[indexCacheKey] = { savedAt: Date.now(), sheet: indexSheet };
    Object.keys(smartsheetIndexCache).forEach((key) => {
      if (Date.now() - Number(smartsheetIndexCache[key]?.savedAt || 0) > PROJECTS_FAST_CACHE_MS * 2) delete smartsheetIndexCache[key];
    });
  }
  const indexRows = mapApiRows(indexSheet);
  const relevantRowIds = collectClientRelevantRowIds(indexRows, session);

  if (!relevantRowIds.length) {
    const emptyPayload = {
      ok: true,
      meta: {
        sheetId,
        sheetName: indexSheet.name || cache.sheetName || SHEET_NAME,
        version,
        lastSync: new Date().toISOString(),
        stageOrder: STAGE_ORDER.filter((stage) => !stage.ignoredCurrentStage).map((stage) => ({
          key: stage.key,
          label: stage.label,
          type: stage.type,
          optional: Boolean(stage.optional),
        })),
        currentWeek: getProductionWeekLabel(getCurrentBrazilDateObject()),
        clientPortal: true,
        clientOptimized: true,
        clientName: session.clientName || session.clientKey || session.name || 'Cliente',
        clientKey: session.clientKey || '',
        clientLogoUrl: session.clientLogoUrl || '',
        alertSignature: '',
        emptyClientScopeResolved: hasExplicitClientScope(session),
      },
      stats: buildStats([]),
      alerts: [],
      projects: [],
    };
    setCachedClientPayload(cacheKey, emptyPayload);
    return emptyPayload;
  }

  let clientSheet;
  try {
    clientSheet = await fetchSheetRowsByIds(sheetId, relevantRowIds);
  } catch (error) {
    const fallbackPayload = !force ? await readBestAvailablePayloadForSession(session, 'client-rows-failed') : null;
    if (fallbackPayload) return fallbackPayload;
    throw error;
  }
  const rows = mapApiRows(clientSheet).sort((a, b) => Number(a.rowNumber || 0) - Number(b.rowNumber || 0));
  let projects = buildProjects(rows).filter((project) => projectBelongsToClientScope(project, session));
  const projectKeys = collectProjectLookupKeySet(projects);
  const wipPoData = await fetchWipStepPoMap({ projectKeys, timeoutMs: WIP_PO_FETCH_TIMEOUT_MS });
  projects = enrichProjectsWithCustomerPo(projects, wipPoData.poMap, wipPoData.focalMap, wipPoData.replannedFinishMap);

  const stats = buildStats(projects);
  const alertData = buildAlerts(projects);
  const payload = {
    ok: true,
    meta: {
      sheetId,
      sheetName: clientSheet.name || indexSheet.name || cache.sheetName || SHEET_NAME,
      version,
      wipStepSheetName: wipPoData.sheetName || WIP_STEP_SHEET_NAME,
      wipStepVersion: wipPoData.version || null,
      wipStepPoAvailable: Boolean(wipPoData.available),
      lastSync: new Date().toISOString(),
      stageOrder: STAGE_ORDER.filter((stage) => !stage.ignoredCurrentStage).map((stage) => ({
        key: stage.key,
        label: stage.label,
        type: stage.type,
        optional: Boolean(stage.optional),
      })),
      currentWeek: getProductionWeekLabel(getCurrentBrazilDateObject()),
      alertSignature: alertData.signature,
      clientPortal: true,
      clientOptimized: true,
      clientRowsLoaded: rows.length,
      clientName: session.clientName || session.clientKey || session.name || 'Cliente',
      clientKey: session.clientKey || '',
      clientLogoUrl: session.clientLogoUrl || '',
    },
    stats,
    alerts: alertData.alerts,
    projects,
  };

  setCachedClientPayload(cacheKey, payload);
  persistClientPayloadInBackground(session, payload);
  return payload;
}

exports.handler = async (event) => {
  const warmup = String(event.queryStringParameters?.warmup || "") === "1";

  // Pré-aquecimento seguro: permite carregar/atualizar o cache do Smartsheet antes do login,
  // mas não retorna projetos, estatísticas, alertas ou dados operacionais para a tela pública.
  if (warmup) {
    if (event.httpMethod && event.httpMethod !== 'GET') {
      return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
    }

    try {
      const cachedPayload = await readPersistentFullPayload('warmup-cache-only', { loginRead: true });
      return jsonResponse(200, {
        ok: true,
        warmed: Boolean(cachedPayload),
        cacheOnly: true,
        cacheUpdatedAt: cachedPayload?.meta?.persistentCacheUpdatedAt || cachedPayload?.meta?.lastSync || null,
      }, {
        headers: {
          'cache-control': 'no-store',
        },
      });
    } catch (error) {
      console.warn('[projects-warmup] Falha ao pré-aquecer cache:', error?.message || error);
      return jsonResponse(200, {
        ok: false,
        warmed: false,
        error: 'Aquecimento não concluído.',
      }, {
        headers: {
          'cache-control': 'no-store',
        },
      });
    }
  }

  const auth = requireSession(event);
  if (!auth.ok) {
    return jsonResponse(401, { ok: false, error: 'Faça login para visualizar o painel.' });
  }

  try {
    const force = String(event.queryStringParameters?.force || "") === "1";
    // v37.17: leitura normal sempre prioriza Supabase/cache. O Smartsheet não é chamado
    // no login; somente botão manual/rotina agendada atualizam a linha persistente.
    const preferCache = !force || String(event.queryStringParameters?.preferCache || "") === "1";
    const session = await hydrateClientSession(auth.session);

    // v36.84: cliente não carrega mais a planilha completa.
    // Primeiro lê apenas coluna Client/Project, identifica as linhas do cliente,
    // depois busca só essas linhas + POs compatíveis. Admin/setores continuam com visão completa.
    const bypassMemoryCache = !force && preferCache;
    const payload = session.role === 'client'
      ? await buildClientPayload(session, { force, preferCache, bypassMemoryCache })
      : await buildPayload({ force, preferCache, bypassMemoryCache });

    return jsonResponse(200, payload, {
      headers: {
        'cache-control': 'private, max-age=60, stale-while-revalidate=120',
      },
    });
  } catch (error) {
    try {
      const fallbackPayload = await readBestAvailablePayloadForSession(auth.session, 'handler-catch-fallback');
      if (fallbackPayload) {
        return jsonResponse(200, fallbackPayload, {
          headers: {
            'cache-control': 'private, max-age=30, stale-while-revalidate=300',
            'x-step-fallback': '1',
          },
        });
      }
    } catch (_fallbackError) {}
    const status = error?.code === 'PERSISTENT_CACHE_UNAVAILABLE' ? 503 : 500;
    return jsonResponse(status, {
      ok: false,
      error: error.message,
      cacheOnly: error?.code === 'PERSISTENT_CACHE_UNAVAILABLE',
      needsManualRefresh: error?.code === 'PERSISTENT_CACHE_UNAVAILABLE',
    });
  }
};
exports.buildPayload = buildPayload;
exports.fetchCurrentTrackingSheetVersion = fetchCurrentTrackingSheetVersion;
