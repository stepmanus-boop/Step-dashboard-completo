const { jsonResponse, requireSession } = require('./_auth');
const API_BASE = process.env.SMARTSHEET_API_BASE || "https://api.smartsheet.com/2.0";
const SHEET_NAME = process.env.SMARTSHEET_SHEET_NAME || "Progress Tracking Sheet - Piping Fabrication";
const SHEET_ID_ENV = process.env.SMARTSHEET_SHEET_ID || "";
const WIP_STEP_SHEET_NAME = process.env.SMARTSHEET_WIP_STEP_SHEET_NAME || "Work in Progress - STEP";
const WIP_STEP_SHEET_ID_ENV = process.env.SMARTSHEET_WIP_STEP_SHEET_ID || process.env.SMARTSHEET_WORK_IN_PROGRESS_STEP_SHEET_ID || "";
const TOKEN = process.env.SMARTSHEET_TOKEN || process.env.SMARTSHEET_ACCESS_TOKEN || process.env.SMARTSHEET_API_TOKEN || process.env.SMARTSHEET_BEARER_TOKEN || process.env.SMARTSHEET_PAT || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN || "5pP36OjBaD1W2HWyxf6aoGxXasPvEl8gbqOmQ";

const cache = global.__STEP_PROGRESS_CACHE__ || {
  sheetId: null,
  sheetName: null,
  version: null,
  wipStepSheetId: null,
  wipStepSheetName: null,
  wipStepVersion: null,
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
    else if (normalizedSector === "Logística" && ["Preparado para envio", "Aguardando envio"].includes(status)) flowState = "awaiting_shipment";
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

  if (status === "finalizado") return 999;
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
  const explicitFinished = Boolean(project.finished || project.projectFinishedFlag || hasProjectFinishDateMarker(project) || isProjectStatusFinished(project.projectStatus));
  const finalFinished = explicitFinished || summary.allFinished;
  const finalFlow = finalFinished ? makeFlow("Finalizado", "Enviado", 100, "completed", "completed") : summary.flow;
  project.demandSummary = summary;
  project.statusSummary = finalFinished ? "Finalizado" : summary.statusSummary;
  project.sectorSummary = finalFinished ? "Enviado" : summary.sectorSummary;
  project.statusBreakdown = summary.statusBreakdown;
  project.sectorBreakdown = summary.sectorBreakdown;
  project.flow = finalFlow;
  project.currentStage = finalFinished ? "Finalizado" : summary.statusSummary;
  project.currentStageGroup = finalFinished ? "Enviado" : summary.sectorSummary;
  project.currentStagePercent = finalFlow.percent;
  project.currentStageStatus = finalFinished ? "completed" : (finalFlow.stageStatus || "waiting");
  project.currentStageAlert = !finalFinished && ["in_progress", "waiting"].includes(project.currentStageStatus);
  project.operationalSector = finalFinished ? "Enviado" : summary.sectorSummary;
  project.operationalState = finalFlow.state;
  project.finished = finalFinished;
  project.uiState = uiStateFromFlow(finalFlow, finalFinished);
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

  const quantitySpoolsRaw = parseNumber(summaryRow, "Quantity Spools");
  const quantitySpools = Number(quantitySpoolsRaw || 0) > 0
    ? quantitySpoolsRaw
    : spools.length;

  const operationalSector = flow.sector;

  const project = {
    rowId: summaryRow.id,
    rowNumber: summaryRow.rowNumber,
    projectPrefix: parts.prefix,
    projectNumber: parts.number,
    projectDisplay: parts.display || projectText,
    customerPo: '',
    customerPoList: [],
    customerPoDisplay: 'Aguardando PO',
    customerPoStatus: 'waiting',
    clientDisplayCode: `${parts.display || projectText || 'BSP'} - Aguardando PO`,
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
    individualProgress,
    overallProgress,
    projectStatus,
    observations: textValue(summaryRow, "OBSERVATIONS"),
    jobProcessStatus: textValue(summaryRow, "Job Process Status") || progress.currentStage.label,
    summaryDrawing: textValue(summaryRow, "Drawing"),
    projectType: textValue(summaryRow, "Project Type"),
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
    if (isSummaryRow(row)) {
      const projectChildren = getLeafChildRows(row.id);
      currentSummary = row;
      projects.push(buildProject(row, projectChildren));
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
  return spools.length > 0 && spools.every((spool) => Boolean(
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

function isProjectStartedForStats(project) {
  if (!project || isProjectExcludedFromTotal(project)) return { started: false, tags: 0 };
  const openItems = getOpenFlowItemsForStats(project);
  const startedItems = openItems.filter((item) => {
    const sector = String(item.flow?.sector || "").trim();
    return ["Produção", "Qualidade", "Pintura", "Logística", "Enviado"].includes(sector);
  });
  if (startedItems.length) return { started: true, tags: startedItems.length };
  const statusText = normalizeStatusText([project.currentStage, project.currentStatus, project.statusSummary, project.operationalSector, project.currentSector, project.flow?.status, project.flow?.sector].filter(Boolean).join(" "));
  const textualStarted = ["CORTE", "FABRICATION", "PRE", "SOLD", "WELD", "INSPEC", "TH", "PINT", "PAINT", "COATING", "UNITIZ", "ENVIO", "LOGIST"].some((term) => statusText.includes(term));
  if (textualStarted) return { started: true, tags: Number(project.quantitySpools || 1) };
  const progress = Number(project.overallProgress || project.currentStagePercent || 0);
  return progress > 0 ? { started: true, tags: Number(project.quantitySpools || 1) } : { started: false, tags: 0 };
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
  const activeProjects = (Array.isArray(projects) ? projects : []).filter((project) => !isProjectExcludedFromTotal(project));
  const stats = {
    totalProjects: activeProjects.length,
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
    const isOnHold = isProjectOnHold(project);
    const isPending = isProjectPending(project);

    if (isOnHold) {
      stats.notStartedHold += 1;
      stats.notStartedHoldTags += tags;
      continue;
    }

    if (isPending) {
      continue;
    }

    if (isFinishedProject) {
      stats.completed += 1;
      stats.completedTags += tags;
      continue;
    }

    progressAccumulator += project.overallProgress || 0;

    const openItems = getOpenFlowItemsForStats(project);
    const startedSnapshot = isProjectStartedForStats(project);
    if (startedSnapshot.started) {
      stats.startedProjects += 1;
      stats.startedTags += Number(startedSnapshot.tags || tags || 0);
    }
    const countSector = (sector) => openItems.filter((item) => item.flow?.sector === sector).length;
    const producaoTags = countSector("Produção");
    const qualidadeTags = countSector("Qualidade");
    const pinturaTags = countSector("Pintura");
    const logisticaTags = getAwaitingShipmentTagsForStats(project);
    const preStartTags = openItems.filter((item) => ["Engenharia", "Suprimento"].includes(item.flow?.sector)).length;

    if (producaoTags) { stats.inProgress += 1; stats.inProgressTags += producaoTags; }
    if (qualidadeTags) { stats.inspectionProjects += 1; stats.inspectionTags += qualidadeTags; }
    if (pinturaTags) { stats.paintingProjects += 1; stats.paintingTags += pinturaTags; }
    if (logisticaTags) { stats.awaitingShipment += 1; stats.awaitingShipmentTags += logisticaTags; }
    if (preStartTags || (!openItems.length && !project.finished)) {
      stats.notStarted += 1;
      stats.notStartedTags += preStartTags || tags;
    }
  }

  stats.averageOverallProgress = activeProjects.length ? progressAccumulator / activeProjects.length : 0;
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

async function fetchWipStepPoMap() {
  try {
    const sheetId = await resolveWipStepSheetId();
    if (!sheetId) return { poMap: new Map(), version: null, sheetName: WIP_STEP_SHEET_NAME, available: false };
    const version = await fetchSheetVersion(sheetId);
    const sheet = await fetchFullSheet(sheetId);
    const rows = mapApiRows(sheet);
    return { poMap: buildWipPoMap(rows), version, sheetName: sheet.name || cache.wipStepSheetName || WIP_STEP_SHEET_NAME, available: true };
  } catch (error) {
    console.warn('Não foi possível carregar Work in Progress - STEP para vínculo de PO:', error.message);
    return { poMap: new Map(), version: null, sheetName: WIP_STEP_SHEET_NAME, available: false, error: error.message };
  }
}

function enrichProjectsWithCustomerPo(projects, poMap) {
  const map = poMap instanceof Map ? poMap : new Map();
  return (Array.isArray(projects) ? projects : []).map((project) => {
    const list = getPoListForProject(project, map).filter(Boolean);
    const uniqueList = Array.from(new Set(list));
    project.customerPoList = uniqueList;
    project.customerPo = uniqueList[0] || '';
    project.customerPoDisplay = uniqueList.length ? getProjectPoDisplay(project) : 'Aguardando PO';
    project.customerPoStatus = uniqueList.length ? 'found' : 'waiting';
    project.clientDisplayCode = buildClientDisplayCode(project);
    return project;
  });
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
  const wipPoData = await fetchWipStepPoMap();

  if (cache.payload && cache.version === version && cache.wipStepVersion === wipPoData.version) {
    return cache.payload;
  }

  const sheet = await fetchFullSheet(sheetId);
  const rows = mapApiRows(sheet);
  const projects = enrichProjectsWithCustomerPo(buildProjects(rows), wipPoData.poMap);
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
  cache.payload = payload;

  return payload;
}


function normalizeClientScopeValue(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, ' ')
    .trim()
    .toLowerCase();
}

function getClientScopeAliases(session = {}) {
  const values = [
    session.clientKey,
    session.clientName,
    session.name,
    session.username,
    session.email,
    ...(Array.isArray(session.allowedClients) ? session.allowedClients : []),
  ];
  const aliases = new Set();
  values.forEach((value) => {
    const normalized = normalizeClientScopeValue(value);
    if (!normalized) return;
    aliases.add(normalized);
    normalized.split(/\s+/).forEach((part) => {
      if (part && part.length >= 3) aliases.add(part);
    });
  });
  return aliases;
}

function projectBelongsToClientScope(project, session = {}) {
  if (!project) return false;
  const aliases = getClientScopeAliases(session);
  if (!aliases.size) return false;
  const client = normalizeClientScopeValue(project.client);
  if (!client) return false;
  for (const alias of aliases) {
    if (!alias) continue;
    if (client === alias || client.includes(alias) || alias.includes(client)) return true;
  }
  return false;
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
      clientName: session.clientName || session.clientKey || session.name || session.username || 'Cliente',
      clientKey: session.clientKey || session.clientName || session.name || session.username || '',
      clientLogoUrl: session.clientLogoUrl || '',
      alertSignature: alertData.signature,
    },
  };
}

exports.handler = async (event) => {
  const auth = requireSession(event);
  if (!auth.ok) {
    return jsonResponse(401, { ok: false, error: 'Faça login para visualizar o painel.' });
  }

  const query = event?.queryStringParameters || {};
  const fastMode = query.fast === '1' || query.fast === 'true';

  try {
    if (fastMode && cache.payload) {
      const payload = scopePayloadForSession(cache.payload, auth.session);
      return jsonResponse(200, {
        ...payload,
        meta: {
          ...(payload.meta || {}),
          fastCache: true,
          cacheServedAt: new Date().toISOString(),
        },
      }, {
        headers: {
          'cache-control': 'private, max-age=30, stale-while-revalidate=180',
        },
      });
    }

    const payload = scopePayloadForSession(await buildPayload(), auth.session);
    return jsonResponse(200, payload, {
      headers: {
        'cache-control': 'private, max-age=60, stale-while-revalidate=120',
      },
    });
  } catch (error) {
    if (cache.payload) {
      const payload = scopePayloadForSession(cache.payload, auth.session);
      return jsonResponse(200, {
        ...payload,
        meta: {
          ...(payload.meta || {}),
          staleFallback: true,
          fallbackReason: error.message,
          cacheServedAt: new Date().toISOString(),
        },
      }, {
        headers: {
          'cache-control': 'private, max-age=15, stale-while-revalidate=120',
        },
      });
    }
    return jsonResponse(500, { ok: false, error: error.message });
  }
};
exports.buildPayload = buildPayload;
