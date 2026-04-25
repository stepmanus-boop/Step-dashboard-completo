const { jsonResponse, requireSession } = require('./_auth');
const API_BASE = process.env.SMARTSHEET_API_BASE || "https://api.smartsheet.com/2.0";
const SHEET_NAME = process.env.SMARTSHEET_SHEET_NAME || "Progress Tracking Sheet - Piping Fabrication";
const SHEET_ID_ENV = process.env.SMARTSHEET_SHEET_ID || "";
const TOKEN = process.env.SMARTSHEET_TOKEN || "5pP36OjBaD1W2HWyxf6aoGxXasPvEl8gbqOmQ";

const cache = global.__STEP_PROGRESS_CACHE__ || {
  sheetId: null,
  sheetName: null,
  version: null,
  payload: null,
  lastSync: null,
};
global.__STEP_PROGRESS_CACHE__ = cache;

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

function getCellValue(row, key) {
  return row.values[key] || { raw: null, display: null };
}

function textValue(row, key) {
  const cell = getCellValue(row, key);
  const value = cell.display ?? cell.raw;
  return value == null ? "" : String(value).trim();
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
    else if (normalizedSector === "Logística" && status === "Preparado para envio") flowState = "awaiting_shipment";
    else if (["Qualidade"].includes(normalizedSector)) flowState = "in_inspection";
    else if (["Produção", "Pintura", "Engenharia", "Suprimento"].includes(normalizedSector)) flowState = "in_production";
    else flowState = "not_started";
  }
  return { state: flowState, sector: normalizedSector, status, percent: normalizedPercent, stageStatus: statusType };
}

function deriveOperationalStage(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus) {
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

  if (finished || projectFinished || projectFinishDate) return makeFlow("Finalizado", "Logística", 100, "completed", "completed");
  if (packageDelivered >= 100) return makeFlow("Preparado para envio", "Logística", packageDelivered, "completed", "awaiting_shipment");
  if (packageDelivered > 0) return makeFlow("Preparado para envio", "Logística", packageDelivered, null, "awaiting_shipment");
  if (finalInspection >= 100) return makeFlow("Preparado para envio", "Logística", finalInspection, "waiting", "awaiting_shipment");
  if (finalInspection > 0) return makeFlow("Unitização e Inspeção", "Logística", finalInspection, null, "awaiting_shipment");
  if (coating >= 100) return makeFlow("Unitização e Inspeção", "Logística", coating, "waiting", "awaiting_shipment");
  if (coating > 0) return makeFlow(paintingStatusFromPercent(coating), "Pintura", coating, null, "in_production");
  if (th >= 100) return makeFlow("Pintura", "Pintura", 0, "waiting", "in_production");
  if (th > 0) return makeFlow("TH", "Qualidade", th, null, "in_inspection");
  if (nde != null && nde > 0 && nde < 100) return makeFlow("Aguardando END", "Qualidade", nde, null, "in_inspection");
  if (finalDimensional >= 100) return makeFlow("TH", "Qualidade", 0, "waiting", "in_inspection");
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

function getFlowSortWeight(flow) {
  const status = String(flow?.status || "");
  const sector = String(flow?.sector || "");
  if (status === "Finalizado") return 999;
  if (sector === "Engenharia") return 10;
  if (sector === "Suprimento") return 20;
  if (status === "Corte e Limpeza") return 30;
  if (status === "Pré - Montagem") return 40;
  if (status === "Inspeção Dimensional de Ajuste - 3D") return 50;
  if (status === "Solda") return 60;
  if (status === "Inspeção Dimensional Final - 3D") return 70;
  if (status === "Aguardando END") return 75;
  if (status === "TH") return 80;
  if (sector === "Pintura") return 90;
  if (status === "Unitização e Inspeção") return 100;
  if (status === "Preparado para envio") return 110;
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
    ? makeFlow("Finalizado", "Logística", 100, "completed", "completed")
    : primary;
  return {
    flow,
    allFinished,
    statusSummary: allFinished ? "Finalizado" : formatBreakdown(activeStatusBreakdown, primary.status),
    sectorSummary: allFinished ? "Logística" : formatBreakdown(activeSectorBreakdown.filter((row) => row.label !== "Logística" || primary.sector === "Logística"), primary.sector),
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

function getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus) {
  return deriveOperationalStage(stageValues, fabricationStartDate, coatingPercent, finished, projectStatus);
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

  return Boolean(parts.prefix && parts.number && (quantitySpools != null || drawing === "ISO" || textValue(row, "Project Type")));
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
  const progress = deriveProgress(row);
  const rowOverallProgress = parsePercent(row, "% Overall Progress");
  const rowIndividualProgress = parsePercent(row, "% Individual Progress");
  const overallProgress = rowOverallProgress ?? rowIndividualProgress ?? 0;
  const individualProgress = rowIndividualProgress ?? overallProgress;
  const projectFinishedFlag = isTruthyValue(getCellValue(row, "Project Finished?").raw);
  const fabricationStartDate = textValue(row, "Fabrication Start Date");
  const stageValues = buildStageValues(row);
  const finished = projectFinishedFlag || overallProgress >= 100 || hasStageValue(stageValues, "Project Finish Date");
  const flow = getOperationalFlow(stageValues, fabricationStartDate, parsePercent(row, "Surface preparation and/or coating") ?? 0, finished, textValue(row, "PROJECT STATUS"));
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

  return {
    rowId: row.id,
    rowNumber: row.rowNumber,
    iso: parsedDrawing.iso,
    description: parsedDrawing.description,
    drawing: drawingText,
    observations: textValue(row, "OBSERVATIONS"),
    pm: textValue(row, "PM") || textValue(parentSummary, "PM"),
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

function chooseBestSpoolRow(currentSpool, nextSpool) {
  if (!currentSpool) return nextSpool;
  const currentScore = getSpoolCompletenessScore(currentSpool);
  const nextScore = getSpoolCompletenessScore(nextSpool);
  if (nextScore > currentScore) return nextSpool;
  if (nextScore < currentScore) return currentSpool;

  const currentRowNumber = Number(currentSpool?.rowNumber || 0);
  const nextRowNumber = Number(nextSpool?.rowNumber || 0);
  if (nextRowNumber > currentRowNumber) return nextSpool;
  return currentSpool;
}


function uiStateFromFlow(flow, allFinished = false) {
  if (allFinished || flow?.state === "completed") return "completed";
  if (flow?.state === "awaiting_shipment") return "awaiting_shipment";
  if (flow?.state === "not_started") return "not_started";
  return "in_progress";
}

function applyProjectSpoolRollup(project) {
  const fallbackFlow = project.flow || makeFlow(project.currentStage || "AG. Emissão de detalhamento", project.operationalSector || "Engenharia", project.currentStagePercent || 0, project.currentStageStatus || "waiting", project.operationalState || project.uiState || "not_started");
  const summary = summarizeFlowItems(project.spools || [], fallbackFlow, project.quantitySpools || 1);
  project.demandSummary = summary;
  project.statusSummary = summary.statusSummary;
  project.sectorSummary = summary.sectorSummary;
  project.statusBreakdown = summary.statusBreakdown;
  project.sectorBreakdown = summary.sectorBreakdown;
  project.flow = summary.flow;
  project.currentStage = summary.statusSummary;
  project.currentStageGroup = summary.sectorSummary;
  project.currentStagePercent = summary.flow.percent;
  project.currentStageStatus = summary.allFinished ? "completed" : (summary.flow.stageStatus || "waiting");
  project.currentStageAlert = !summary.allFinished && ["in_progress", "waiting"].includes(project.currentStageStatus);
  project.operationalSector = summary.sectorSummary;
  project.operationalState = summary.flow.state;
  project.finished = summary.allFinished;
  project.uiState = uiStateFromFlow(summary.flow, summary.allFinished);
  return project;
}

function buildProject(summaryRow, childRows) {
  const projectText = textValue(summaryRow, "Project");
  const parts = parseProjectParts(projectText);
  const progress = deriveProgress(summaryRow);
  const overallProgress = parsePercent(summaryRow, "% Overall Progress") ?? 0;
  const individualProgress = parsePercent(summaryRow, "% Individual Progress") ?? overallProgress;
  const projectFinishedFlag = isTruthyValue(getCellValue(summaryRow, "Project Finished?").raw);
  const projectStatus = textValue(summaryRow, "PROJECT STATUS") || textValue(summaryRow, "Overall Project Status") || textValue(summaryRow, "Status");
  const coatingPercent = parsePercent(summaryRow, "Surface preparation and/or coating") ?? 0;
  const fabricationStartDate = textValue(summaryRow, "Fabrication Start Date");
  const stageValues = buildStageValues(summaryRow);
  const summaryFinished = projectFinishedFlag || overallProgress >= 100 || hasStageValue(stageValues, "Project Finish Date");
  const flow = getOperationalFlow(stageValues, fabricationStartDate, coatingPercent, summaryFinished, projectStatus);
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

  const operationalSector = flow.sector;

  const project = {
    rowId: summaryRow.id,
    rowNumber: summaryRow.rowNumber,
    projectPrefix: parts.prefix,
    projectNumber: parts.number,
    projectDisplay: parts.display || projectText,
    quantitySpools: parseNumber(summaryRow, "Quantity Spools") ?? spools.length,
    kilos: parseNumber(summaryRow, "Kilos"),
    weldedWeightKg,
    weldingWeek,
    coatingPercent,
    m2Painting: parseNumber(summaryRow, "M2 Painting"),
    currentStage: flow.status,
    currentStagePercent: flow.percent,
    currentStageStatus: flow.stageStatus,
    currentStageAlert: flow.stageStatus === "in_progress" || flow.stageStatus === "waiting",
    individualProgress,
    overallProgress,
    projectStatus,
    jobProcessStatus: textValue(summaryRow, "Job Process Status") || progress.currentStage.label,
    summaryDrawing: textValue(summaryRow, "Drawing"),
    projectType: textValue(summaryRow, "Project Type"),
    fabricationStartDate: formatDateValue(textValue(summaryRow, "Fabrication Start Date")),
    plannedStartDate: formatDateValue(textValue(summaryRow, "Start Date")),
    plannedFinishDate: formatDateValue(textValue(summaryRow, "Finish Date")),
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
  return applyProjectSpoolRollup(project);
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

  let currentSummary = null;

  for (const row of rows) {
    if (isSummaryRow(row)) {
      const directChildren = childrenByParent.get(row.id) || [];
      currentSummary = row;
      projects.push(buildProject(row, directChildren));
      continue;
    }

    if (!currentSummary) continue;
    if (!isChildRow(row)) continue;

    const currentProjectNumber = parseProjectParts(textValue(currentSummary, "Project")).number;
    const childProjectNumber = parseProjectParts(textValue(row, "Project")).number;
    if (!childProjectNumber || childProjectNumber !== currentProjectNumber) continue;

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

function hasProjectFinishedMarker(project) {
  return Boolean(project?.finished || project?.uiState === "completed" || project?.operationalState === "completed");
}

function getOpenFlowItemsForStats(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const source = spools.length
    ? spools.map((spool) => ({ flow: spool.flow || { status: spool.stage, sector: spool.operationalSector, state: spool.operationalState }, spool }))
    : [{ flow: project?.flow || { status: project?.currentStage, sector: project?.operationalSector, state: project?.operationalState }, spool: null }];
  return source.filter((item) => item.flow?.state !== "completed" && item.flow?.status !== "Finalizado");
}

function buildStats(projects) {
  const stats = {
    totalProjects: projects.length,
    totalSpools: 0,
    totalWeightKg: 0,
    totalWeldedWeightKg: 0,
    totalPaintingM2: 0,
    completed: 0,
    completedTags: 0,
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
    stats.totalSpools += tags;
    stats.totalWeightKg += project.kilos || 0;
    stats.totalWeldedWeightKg += project.weldedWeightKg || 0;
    const openPaintingM2 = spools.length
      ? spools.filter((spool) => spool.flow?.state !== "completed" && spool.flow?.status !== "Finalizado").reduce((total, spool) => total + Number(spool.m2Painting || 0), 0)
      : 0;
    stats.totalPaintingM2 += project.finished ? 0 : (openPaintingM2 > 0 ? openPaintingM2 : Number(project.m2Painting || 0));
    progressAccumulator += project.overallProgress || 0;

    const normalizedProjectStatus = String(project?.projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
    const isOnHold = ["ON HOLD", "HOLD", "EM ESPERA", "PAUSED"].includes(normalizedProjectStatus);

    if (project.finished) {
      stats.completed += 1;
      stats.completedTags += tags;
      continue;
    }

    const openItems = getOpenFlowItemsForStats(project);
    const countSector = (sector) => openItems.filter((item) => item.flow?.sector === sector).length;
    const producaoTags = countSector("Produção");
    const qualidadeTags = countSector("Qualidade");
    const pinturaTags = countSector("Pintura");
    const logisticaTags = countSector("Logística");
    const preStartTags = openItems.filter((item) => ["Engenharia", "Suprimento"].includes(item.flow?.sector)).length;

    if (producaoTags) { stats.inProgress += 1; stats.inProgressTags += producaoTags; }
    if (qualidadeTags) { stats.inspectionProjects += 1; stats.inspectionTags += qualidadeTags; }
    if (pinturaTags) { stats.paintingProjects += 1; stats.paintingTags += pinturaTags; }
    if (logisticaTags) { stats.awaitingShipment += 1; stats.awaitingShipmentTags += logisticaTags; }
    if (preStartTags || (!openItems.length && !project.finished)) {
      stats.notStarted += 1;
      stats.notStartedTags += preStartTags || tags;
      if (isOnHold) {
        stats.notStartedHold += 1;
        stats.notStartedHoldTags += preStartTags || tags;
      }
    }
  }

  stats.averageOverallProgress = projects.length ? progressAccumulator / projects.length : 0;
  return stats;
}

async function apiFetch(path) {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const message = await response.text();
    throw new Error(`Smartsheet ${response.status}: ${message}`);
  }

  return response.json();
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

async function fetchSheetVersion(sheetId) {
  const versionData = await apiFetch(`/sheets/${sheetId}/version`);
  return versionData.version;
}

async function fetchFullSheet(sheetId) {
  return apiFetch(`/sheets/${sheetId}?includeAll=true`);
}

async function buildPayload() {
  if (!TOKEN) throw new Error("SMARTSHEET_TOKEN não configurado.");

  const sheetId = await resolveSheetId();
  const version = await fetchSheetVersion(sheetId);

  if (cache.payload && cache.version === version) {
    return cache.payload;
  }

  const sheet = await fetchFullSheet(sheetId);
  const rows = mapApiRows(sheet);
  const projects = buildProjects(rows);
  const stats = buildStats(projects);
  const alertData = buildAlerts(projects);

  const payload = {
    ok: true,
    meta: {
      sheetId,
      sheetName: sheet.name || cache.sheetName || SHEET_NAME,
      version,
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
  cache.lastSync = payload.meta.lastSync;
  cache.payload = payload;

  return payload;
}

exports.handler = async (event) => {
  const auth = requireSession(event);
  if (!auth.ok) {
    return jsonResponse(401, { ok: false, error: 'Faça login para visualizar o painel.' });
  }

  try {
    const payload = await buildPayload();
    return jsonResponse(200, payload);
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message });
  }
};
exports.buildPayload = buildPayload;
