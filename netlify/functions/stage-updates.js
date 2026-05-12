const { jsonResponse, requireSession, normalizeSectorValue } = require('./_auth');
const { readJson, writeJson } = require('./_githubStore');
const { isSupabaseConfigured, listStageUpdates, createStageUpdate, updateStageUpdate, deleteStageUpdates } = require('./_supabase');
const { findProjectAndSpool, loadProjectPayload } = require('./_projectLookup');
const { applyStageUpdatesToTracking, listHistoryDatePendencies } = require('./_smartsheetTracking');

const DATA_PATH = 'data/stage-updates.json';
const SUPPORTED_SECTORS = ['engenharia', 'suprimento', 'pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const PROGRESS_OPTIONS = [25, 50, 75, 100];
const PENDING_STATUSES = ['pending', 'pending_advance', 'pending_review'];
const RESOLVED_STATUSES = ['resolved', 'resolved_advance', 'resolved_review'];

const TRACKING_FIELDS_BY_SECTOR = {
  engenharia: ['Drawing Execution Advance%'],
  suprimento: ['Material Separation', 'Procuremnt Status %', 'Material Release to Fabrication'],
  pintura: ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)'],
  inspecao: ['Final Inspection', 'Hydro Test Pressure (QC)', 'Non Destructive Examination (QC)', 'Final Dimensional Inpection/3D (QC)', 'Initial Dimensional Inspection/3D'],
  pendente_envio: ['Package and Delivered', 'Final Inspection'],
  producao: ['Spool Assemble and tack weld', 'Welding Preparation'],
  calderaria: ['Spool Assemble and tack weld', 'Welding Preparation', 'Material Separation', 'Material Release to Fabrication'],
  solda: ['Full welding execution'],
};

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

function normalizeStageWorkspaceText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[–—−]/g, '-')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function inferStageSectorFromOwnText(value) {
  const text = normalizeStageWorkspaceText(value);
  if (!text || text.includes('finalizado') || text.includes('concluido') || text.includes('enviado')) return '';
  if (text.includes('package and delivered') || text.includes('unitizacao') || text.includes('preparado para envio') || text.includes('logistica')) return 'pendente_envio';
  if (text.includes('pintura') || text.includes('inicio de pintura') || text.includes('intermediaria') || text.includes('coating') || text.includes('paint') || text.includes('surface preparation')) return 'pintura';
  if (text.includes('inspecao') || text.includes('inspection') || text.includes('dimensional') || text.includes('qualidade') || text.includes('nde') || text.includes('end') || /\bth\b/.test(text) || text.includes('hydro')) return 'inspecao';
  if (text.includes('solda') || text.includes('full welding')) return 'solda';
  if (text.includes('pre montagem') || text.includes('spool assemble') || text.includes('tack weld') || text.includes('welding preparation') || text.includes('boilermaker') || text.includes('calderaria')) return 'calderaria';
  if (text.includes('corte') || text.includes('limpeza') || text.includes('fabrication start') || text.includes('producao')) return 'producao';
  if (text.includes('separacao de material') || text.includes('material separation') || text.includes('estoque') || text.includes('procure') || text.includes('suprimento')) return 'suprimento';
  if (text.includes('detalhamento') || text.includes('drawing') || text.includes('engenharia')) return 'engenharia';
  return '';
}



function normalizeCompetenceSector(value) {
  const normalized = normalizeSectorValue(value);
  if (['qualidade', 'quality', 'qc'].includes(normalized)) return 'inspecao';
  if (['logistica', 'logistics', 'expedicao', 'shipping'].includes(normalized)) return 'pendente_envio';
  if (['engineering'].includes(normalized)) return 'engenharia';
  if (['supply', 'supply_chain', 'procurement', 'suprimentos'].includes(normalized)) return 'suprimento';
  return normalized;
}

function hasStageValue(stageValues, key) {
  const value = stageValues?.[key];
  if (value == null) return false;
  const text = String(value).trim();
  return Boolean(text && text !== 'N/A' && text.toLowerCase() !== 'não' && text.toLowerCase() !== 'nao');
}

function getStagePercent(stageValues, key) {
  return parseTrackingPercent(stageValues?.[key]) ?? 0;
}

function getSpoolStageLabel(project, spool) {
  return spool?.currentStatus
    || spool?.stage
    || spool?.flow?.status
    || project?.currentStage
    || project?.statusSummary
    || project?.flow?.status
    || 'Etapa não identificada';
}


function normalizeQualityCompetencies(input = []) {
  const rawValues = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const allowed = new Set(['dimensional_inicial', 'dimensional_final', 'nde', 'th', 'final_inspection_qc']);
  const aliases = {
    inicial: 'dimensional_inicial',
    dimensional_inicial: 'dimensional_inicial',
    initial_dimensional: 'dimensional_inicial',
    dimensional_final: 'dimensional_final',
    final_dimensional: 'dimensional_final',
    nde: 'nde',
    end: 'nde',
    th: 'th',
    hydro: 'th',
    hydro_test: 'th',
    final_inspection: 'final_inspection_qc',
    final_inspection_qc: 'final_inspection_qc',
  };
  const seen = new Set();
  const values = [];
  for (const value of rawValues) {
    const key = String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase().replace(/[\s-]+/g, '_').replace(/__+/g, '_');
    const normalized = aliases[key] || key;
    if (!allowed.has(normalized) || seen.has(normalized)) continue;
    seen.add(normalized);
    values.push(normalized);
  }
  return values;
}

function qualityCompetencyLabel(value) {
  const labels = {
    dimensional_inicial: 'Inspeção Dimensional Inicial',
    dimensional_final: 'Inspeção Dimensional Final',
    nde: 'END / NDE',
    th: 'TH',
    final_inspection_qc: 'Final Inspection QC',
  };
  return labels[value] || value || 'não identificada';
}

function getSpoolQualityCompetence(project, spool) {
  const stageValues = spool?.stageValues || project?.stageValues || {};
  const text = normalizeStageWorkspaceText([
    spool?.currentStatus,
    spool?.stage,
    spool?.flow?.status,
    spool?.currentSector,
    spool?.operationalSector,
    spool?.flow?.sector,
    project?.currentStage,
    project?.sectorSummary,
  ].filter(Boolean).join(' '));

  const th = getStagePercent(stageValues, 'Hydro Test Pressure (QC)');
  const nde = parseTrackingPercent(stageValues?.['Non Destructive Examination (QC)']);
  const finalDimensional = getStagePercent(stageValues, 'Final Dimensional Inpection/3D (QC)');
  const initialDimensional = getStagePercent(stageValues, 'Initial Dimensional Inspection/3D');
  const finalInspection = getStagePercent(stageValues, 'Final Inspection');

  if (text.includes('hydro') || text.includes('teste hidrostatico') || /\bth\b/.test(text) || text.includes('aguardando em th') || (th > 0 && th < 100)) return 'th';
  if (text.includes('nde') || /\bend\b/.test(text) || text.includes('non destructive') || (nde != null && nde > 0 && nde < 100)) return 'nde';
  if (text.includes('final dimensional') || text.includes('final dimensional inspection') || text.includes('final dimensional inpection') || finalDimensional > 0) return 'dimensional_final';
  if (text.includes('initial dimensional') || text.includes('inspecao dimensional inicial') || initialDimensional > 0) return 'dimensional_inicial';
  if (text.includes('final inspection') || text.includes('inspection finish') || finalInspection > 0) return 'final_inspection_qc';

  return 'dimensional_final';
}

function ensureQualityCompetenceAllowed(project, spool, sector, session) {
  if (normalizeCompetenceSector(sector) !== 'inspecao') return;
  if (!session || session.role === 'admin' || isPcpUser(session)) return;
  const competencies = normalizeQualityCompetencies(session.qualityCompetencies || []);
  if (!competencies.length) return;
  const competence = getSpoolQualityCompetence(project, spool);
  if (!competence || competencies.includes(competence)) return;
  const err = new Error(`Este spool pertence à competência ${qualityCompetencyLabel(competence)}, mas seu usuário não possui permissão para apontar esta demanda.`);
  err.statusCode = 403;
  throw err;
}

function getSpoolCompetenceSector(project, spool) {
  const stageValues = spool?.stageValues || project?.stageValues || {};
  const text = normalizeStageWorkspaceText([
    spool?.currentStatus,
    spool?.stage,
    spool?.flow?.status,
    spool?.currentSector,
    spool?.operationalSector,
    spool?.flow?.sector,
    project?.currentStage,
    project?.currentStatus,
    project?.currentSector,
    project?.operationalSector,
    project?.sectorSummary,
    project?.statusSummary,
  ].filter(Boolean).join(' '));

  const finished = Boolean(spool?.finished || spool?.projectFinishedFlag)
    || text.includes('finalizado');
  if (finished) return '';

  const ownTextSector = inferStageSectorFromOwnText(spoolOwnText);
  if (ownTextSector) return ownTextSector;

  const coating = Math.max(
    getStagePercent(stageValues, 'Surface preparation and/or coating'),
    getStagePercent(stageValues, 'HDG / FBE.  (PAINT)'),
    parseTrackingPercent(spool?.coatingPercent) ?? 0
  );
  const packageDelivered = getStagePercent(stageValues, 'Package and Delivered');
  const th = getStagePercent(stageValues, 'Hydro Test Pressure (QC)');
  const nde = parseTrackingPercent(stageValues?.['Non Destructive Examination (QC)']);
  const finalDimensional = getStagePercent(stageValues, 'Final Dimensional Inpection/3D (QC)');
  const fullWelding = getStagePercent(stageValues, 'Full welding execution');
  const initialDimensional = getStagePercent(stageValues, 'Initial Dimensional Inspection/3D');
  const spoolAssemble = getStagePercent(stageValues, 'Spool Assemble and tack weld');
  const weldingPreparation = getStagePercent(stageValues, 'Welding Preparation');
  const withdrewMaterial = getStagePercent(stageValues, 'Withdrew Material');
  const materialSeparation = getStagePercent(stageValues, 'Material Separation');
  const procurement = Math.max(
    getStagePercent(stageValues, 'Procuremnt Status %'),
    getStagePercent(stageValues, 'Material Release to Fabrication')
  );
  const drawing = getStagePercent(stageValues, 'Drawing Execution Advance%');
  const fabricationStarted = Boolean(spool?.fabricationStartDate || hasStageValue(stageValues, 'Fabrication Start Date'));
  const boilermakerDone = hasStageValue(stageValues, 'Boilermaker Finish Date');
  const projectFinishDate = hasStageValue(stageValues, 'Project Finish Date');

  if (projectFinishDate || packageDelivered >= 100) return '';

  // Prioridade para o setor explicitamente exibido na demanda.
  // Isso evita bloquear apontamento quando o painel já classificou a demanda como Pintura,
  // mas alguns campos técnicos intermediários ainda carregam percentuais de Qualidade/TH/END.
  if (text.includes('pintura') || text.includes('paint') || text.includes('coating') || text.includes('surface preparation') || text.includes('acabamento') || text.includes('intermediaria') || text === 'j f') return 'pintura';
  if (text.includes('package and delivered') || text.includes('unitizacao') || text.includes('preparado para envio') || text.includes('logistica') || text.includes('aguardando envio')) return 'pendente_envio';

  if (coating >= 100) return 'pendente_envio';
  if (coating > 0 || th >= 100) return 'pintura';
  if (fullWelding > 0 && fullWelding < 100) return 'solda';
  if (th > 0 || (nde != null && nde > 0) || finalDimensional >= 100 || finalDimensional > 0 || fullWelding >= 100 || initialDimensional > 0 || boilermakerDone || spoolAssemble >= 100) {
    if (initialDimensional >= 100 && fullWelding <= 0) return 'solda';
    return 'inspecao';
  }
  if (fullWelding > 0 || initialDimensional >= 100) return 'solda';
  if (spoolAssemble > 0 || weldingPreparation > 0 || weldingPreparation >= 100 || withdrewMaterial > 0) return 'calderaria';
  if (fabricationStarted || materialSeparation >= 100) return 'producao';
  if (materialSeparation > 0 || procurement > 0 || procurement >= 100 || drawing >= 100) return 'suprimento';

  if (text.includes('final inspection') || text.includes('hydro') || /\bth\b/.test(text) || text.includes('dimensional') || text.includes('inspection') || text.includes('inspecao') || text.includes('qualidade') || text.includes('nde') || text.includes('end')) return 'inspecao';
  if (text.includes('full welding') || text.includes('solda') || text === 'solda') return 'solda';
  if (text.includes('pre montagem') || text.includes('spool assemble') || text.includes('tack weld') || text.includes('welding preparation') || text.includes('boilermaker') || text.includes('calderaria')) return 'calderaria';
  if (text.includes('corte') || text.includes('limpeza') || text.includes('fabrication start') || text.includes('producao')) return 'producao';
  if (text.includes('separacao de material') || text.includes('material separation') || text.includes('estoque') || text.includes('procure') || text.includes('suprimento')) return 'suprimento';
  if (text.includes('detalhamento') || text.includes('drawing') || text.includes('engenharia')) return 'engenharia';

  return normalizeCompetenceSector(spool?.currentSector || spool?.operationalSector || spool?.flow?.sector || project?.currentSector || project?.operationalSector || project?.sectorSummary || 'engenharia');
}

function ensureSpoolReleasedForSector(project, spool, sector, session = null) {
  const actorSector = normalizeCompetenceSector(sector);

  const directText = normalizeStageWorkspaceText([
    spool?.currentStatus,
    spool?.stage,
    spool?.currentSector,
    spool?.operationalSector,
    spool?.flow?.status,
    spool?.flow?.sector,
  ].filter(Boolean).join(' '));
  const directSector = inferStageSectorFromOwnText(directText);

  if (directSector) {
    if (directSector !== actorSector) {
      const err = new Error(`Este spool pertence ao setor ${sectorLabel(directSector)}, não ao setor ${sectorLabel(actorSector)}.`);
      err.statusCode = 403;
      throw err;
    }
    ensureQualityCompetenceAllowed(project, spool, actorSector, session);
    return;
  }

  const competenceSector = getSpoolCompetenceSector(project, spool);
  if (!competenceSector || competenceSector !== actorSector) {
    const err = new Error(`Este spool ainda não está liberado para apontamento do setor ${sectorLabel(actorSector)}. Etapa atual: ${getSpoolStageLabel(project, spool)}. Setor responsável: ${sectorLabel(competenceSector) || 'não identificado'}.`);
    err.statusCode = 403;
    throw err;
  }
  ensureQualityCompetenceAllowed(project, spool, actorSector, session);
}

function sectorLabel(value) {
  const normalized = normalizeCompetenceSector(value);
  const labels = {
    engenharia: 'Engenharia',
    suprimento: 'Suprimento',
    pintura: 'Pintura',
    inspecao: 'Qualidade',
    pendente_envio: 'Logística',
    producao: 'Produção',
    calderaria: 'Calderaria',
    solda: 'Solda',
    pcp: 'PCP',
  };
  return labels[normalized] || value || '';
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

function isPendingStatus(status) {
  return PENDING_STATUSES.includes(String(status || 'pending').trim().toLowerCase());
}

function isResolvedStatus(status) {
  return RESOLVED_STATUSES.includes(String(status || '').trim().toLowerCase());
}

function isReviewStatus(status) {
  return String(status || '').trim().toLowerCase().includes('review');
}

function applyTrackingVerification(update, project, spool) {
  const progress = Number(update?.progress || 0);
  const trackingProgress = getTrackingProgressForSector(spool, update?.sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;
  return {
    ...update,
    trackingCheckedAt: new Date().toISOString(),
    trackingProgress: trackingProgress == null ? null : Number(trackingProgress.toFixed(2)),
    trackingMatched,
    trackingStatus: trackingProgress == null ? 'not_found' : (trackingMatched ? 'matched' : 'waiting'),
  };
}

async function enrichUpdatesWithTracking(updates) {
  const list = Array.isArray(updates) ? updates : [];
  if (!list.length) return [];
  let payload = null;
  try {
    payload = await loadProjectPayload({ allowFallback: false });
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
    return applyTrackingVerification(item, project, spool);
  });
}


async function autoResolveTrackingMatchedUpdates(enrichedUpdates, rawUpdates, session) {
  if (!canValidate(session)) return { changed: false, updates: Array.isArray(enrichedUpdates) ? enrichedUpdates : [] };

  const candidates = (Array.isArray(enrichedUpdates) ? enrichedUpdates : []).filter((item) =>
    isPendingStatus(item?.status)
    && !isReviewStatus(item?.status)
    && item?.trackingMatched === true
  );

  if (!candidates.length) return { changed: false, updates: Array.isArray(enrichedUpdates) ? enrichedUpdates : [] };

  let changed = false;
  const resolutionNote = 'Concluído automaticamente pelo PCP porque o Tracking já estava OK.';
  for (const item of candidates) {
    const saved = await resolveStageUpdateRecord(item.id, rawUpdates, session, resolutionNote);
    if (saved) changed = true;
  }

  if (changed && !isSupabaseConfigured()) await saveUpdates(rawUpdates);

  const refreshed = await enrichUpdatesWithTracking(rawUpdates);
  return { changed, updates: refreshed };
}

function canValidate(session) {
  const sector = normalizeSectorValue(session?.sector);
  return session?.role === 'admin' || sector === 'pcp';
}

function isPcpUser(session) {
  return Boolean(session && normalizeSectorValue(session.sector) === 'pcp');
}

function canCreate(session) {
  const sector = normalizeSectorValue(session?.sector);
  return SUPPORTED_SECTORS.includes(sector) || isPcpUser(session) || session?.role === 'admin';
}

function getActorSector(session) {
  return normalizeSectorValue(session?.sector);
}

function getRequestedStageSector(payload, session) {
  const actorSector = getActorSector(session);
  const requested = normalizeSectorValue(payload?.sector || payload?.targetSector || '');
  if (session?.role === 'admin') return normalizeCompetenceSector(requested || actorSector);
  if (isPcpUser(session)) return normalizeCompetenceSector(requested);
  return normalizeCompetenceSector(actorSector);
}

async function listUpdates() {
  if (isSupabaseConfigured()) {
    try {
      return await listStageUpdates();
    } catch (error) {
      // Proteção operacional: se o Supabase falhar temporariamente ou a tabela
      // stage_updates estiver com schema/cache instável, não derruba a tela de apontamentos.
      // O sistema carrega o fallback JSON para manter PCP/setores funcionando.
      console.warn('Falha ao listar stage_updates no Supabase; usando fallback JSON:', error.message);
    }
  }
  const rows = await readJson(DATA_PATH, []);
  return Array.isArray(rows) ? rows : [];
}

async function saveUpdates(rows) {
  return writeJson(DATA_PATH, rows, 'chore: atualiza apontamentos setoriais');
}

function normalizeUpdateForJson(record) {
  return { ...record, updatedAt: new Date().toISOString() };
}

async function resolveStageUpdateRecord(id, updates, session, resolutionNote = '') {
  const index = updates.findIndex((item) => String(item.id) === String(id));
  if (index < 0) return null;
  const current = updates[index];
  const currentStatus = String(current?.status || 'pending').trim().toLowerCase();
  const nextStatus = currentStatus === 'pending_review' ? 'resolved_review' : 'resolved_advance';
  const resolvedAt = new Date().toISOString();
  const updatedRecord = normalizeUpdateForJson({
    ...current,
    status: nextStatus,
    resolvedBy: session.username || '',
    resolvedByName: session.name || session.username || 'Usuário',
    resolvedAt,
    resolutionNote,
  });
  if (isSupabaseConfigured()) {
    const saved = await updateStageUpdate(id, {
      status: nextStatus,
      resolvedBy: updatedRecord.resolvedBy,
      resolvedByName: updatedRecord.resolvedByName,
      resolvedAt: updatedRecord.resolvedAt,
      resolutionNote: updatedRecord.resolutionNote,
    });
    return saved || updatedRecord;
  }
  updates[index] = updatedRecord;
  return updatedRecord;
}

async function createSingleUpdate(payload, session, existingUpdates = null) {
  const projectRowId = Number(payload.projectRowId || 0);
  const spoolIso = String(payload.spoolIso || '').trim();
  const progress = Number(payload.progress || 0);
  const completionDate = String(payload.completionDate || '').trim();
  const note = String(payload.note || '').trim();
  const actionType = String(payload.actionType || 'advance').trim().toLowerCase() === 'review' ? 'review' : 'advance';
  const actorSector = getActorSector(session);
  const sector = getRequestedStageSector(payload, session);
  const pcpDelegated = isPcpUser(session) && SUPPORTED_SECTORS.includes(sector);

  if (!projectRowId || !spoolIso || !SUPPORTED_SECTORS.includes(sector)) {
    throw new Error(pcpDelegated || isPcpUser(session)
      ? 'Selecione o setor que o PCP irá apontar.'
      : 'Informe BSP, spool e uma etapa válida.');
  }
  if (!PROGRESS_OPTIONS.includes(progress)) {
    throw new Error('Selecione um avanço válido: 25%, 50%, 75% ou 100%.');
  }
  const { project, spool } = await findProjectAndSpool(projectRowId, spoolIso, { allowFallback: false });
  if (!project || !spool) {
    const err = new Error('BSP ou spool não localizado para este apontamento.');
    err.statusCode = 404;
    throw err;
  }
  ensureSpoolReleasedForSector(project, spool, sector, session);
  const trackingProgress = getTrackingProgressForSector(spool, sector);
  const trackingMatched = trackingProgress != null && trackingProgress >= progress;
  const updates = existingUpdates || await listUpdates();
  const pendingExists = updates.find((item) =>
    isPendingStatus(item.status)
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
  const pcpDelegationNote = pcpDelegated
    ? `Apontamento realizado pelo PCP ${session.name || session.username || 'usuário'} em nome do setor ${sectorLabel(sector)}.`
    : '';
  const finalNote = pcpDelegationNote
    ? `${pcpDelegationNote}${note ? `\n${note}` : ''}`
    : note;
  const record = {
    id: `stg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    projectRowId,
    projectNumber: project.projectNumber || project.projectDisplay || `Projeto ${projectRowId}`,
    projectDisplay: project.projectDisplay || project.projectNumber || `Projeto ${projectRowId}`,
    client: project.client || '',
    spoolIso,
    spoolDescription: spool.description || spool.drawing || '',
    sector,
    progress,
    completionDate: completionDate || (progress === 100 ? now.slice(0, 10) : ''),
    note: finalNote,
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

function getUpdatesByIds(updates, ids) {
  const cleanIds = new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean));
  return updates.filter((item) => cleanIds.has(String(item.id || '')));
}

function getAlreadyResolvedByIds(updates, ids) {
  const cleanIds = new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean));
  return (Array.isArray(updates) ? updates : []).filter((item) =>
    cleanIds.has(String(item?.id || '')) && isResolvedStatus(item?.status)
  );
}

function alreadyProcessedResponse(updates, ids, message = 'Apontamento já estava concluído ou processado.') {
  const resolved = getAlreadyResolvedByIds(updates, ids);
  return jsonResponse(200, {
    ok: true,
    alreadyProcessed: true,
    updates: resolved,
    errors: [],
    message,
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}

async function updateTrackingAndResolve(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode atualizar o Tracking.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe os apontamentos para atualizar.' });

  const forceRewrite = Boolean(body.forceRewrite || body.rewrite);
  const dateOnly = Boolean(body.dateOnly);
  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => {
    if (dateOnly) return isResolvedStatus(item.status) && Number(item.progress || 0) === 100 && !isReviewStatus(item.status);
    return isPendingStatus(item.status) && !isReviewStatus(item.status);
  });
  if (!selected.length) {
    const alreadyResolved = getAlreadyResolvedByIds(updates, ids);
    if (alreadyResolved.length) {
      return alreadyProcessedResponse(updates, ids, 'Este apontamento já estava concluído. A tela será sincronizada novamente.');
    }
    return jsonResponse(404, { ok: false, error: 'Nenhum apontamento elegível encontrado para atualizar o Tracking.' });
  }

  const trackingResult = await applyStageUpdatesToTracking(selected, { forceRewrite, dateOnly });
  const successResultIds = new Set((trackingResult.results || []).filter((item) => item.success).map((item) => String(item.id || '')));
  const resolutionNoteBase = dateOnly
    ? 'Pendência de data do histórico corrigida no Smartsheet/Tracking.'
    : (forceRewrite ? 'Tracking regravado e apontamento validado automaticamente.' : 'Tracking atualizado e apontamento validado automaticamente.');

  const resolved = [];
  if (!dateOnly) {
    for (const item of selected) {
      if (!successResultIds.has(String(item.id))) continue;
      const saved = await resolveStageUpdateRecord(item.id, updates, session, resolutionNoteBase);
      if (saved) resolved.push(saved);
    }
    if (!isSupabaseConfigured() && resolved.length) await saveUpdates(updates);
  }

  const hasSuccess = (trackingResult.results || []).some((item) => item.success);
  const hasErrors = Array.isArray(trackingResult.errors) && trackingResult.errors.length > 0;
  return jsonResponse(hasSuccess ? 200 : 400, {
    ok: hasSuccess,
    partial: hasSuccess && hasErrors,
    tracking: trackingResult,
    updates: resolved,
    errors: trackingResult.errors || [],
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}


async function deleteStageUpdateRecords(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode remover apontamentos.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe os apontamentos para remover.' });

  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => isPendingStatus(item.status));
  if (!selected.length) return jsonResponse(404, { ok: false, error: 'Nenhum apontamento pendente encontrado para remover.' });
  const selectedIds = selected.map((item) => String(item.id));

  if (isSupabaseConfigured()) {
    await deleteStageUpdates(selectedIds);
  } else {
    const selectedSet = new Set(selectedIds);
    const remaining = updates.filter((item) => !selectedSet.has(String(item.id || '')));
    await saveUpdates(remaining);
  }

  return jsonResponse(200, {
    ok: true,
    removed: selected,
    removedCount: selected.length,
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}

async function concludeTrackingOkOnly(body, session) {
  if (!canValidate(session)) {
    return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode concluir apontamentos.' });
  }
  const ids = Array.isArray(body.ids)
    ? body.ids.map((id) => String(id || '').trim()).filter(Boolean)
    : [String(body.id || '').trim()].filter(Boolean);
  const resolutionNote = String(body.resolutionNote || '').trim();
  if (!ids.length) return jsonResponse(400, { ok: false, error: 'Informe o apontamento para concluir.' });

  const updates = await listUpdates();
  const selected = getUpdatesByIds(updates, ids).filter((item) => isPendingStatus(item.status));
  if (!selected.length) {
    const alreadyResolved = getAlreadyResolvedByIds(updates, ids);
    if (alreadyResolved.length) {
      return alreadyProcessedResponse(updates, ids, 'Este apontamento já estava concluído. A tela será sincronizada novamente.');
    }
    return jsonResponse(404, { ok: false, error: 'Apontamento não encontrado.' });
  }

  const reviewItems = selected.filter((item) => isReviewStatus(item.status));
  const advanceItems = selected.filter((item) => !isReviewStatus(item.status));
  let eligibleAdvance = [];
  let blockedErrors = [];

  if (advanceItems.length) {
    const dryRun = await applyStageUpdatesToTracking(advanceItems, { dryRun: true });
    eligibleAdvance = advanceItems.filter((item) => {
      const result = (dryRun.results || []).find((entry) => String(entry.id) === String(item.id));
      return result?.trackingOk === true;
    });
    blockedErrors = (dryRun.results || [])
      .filter((entry) => entry?.trackingOk !== true)
      .map((entry) => ({ id: entry.id, error: entry.message || 'Tracking ainda pendente de atualização.' }));
  }

  const toResolve = [...reviewItems, ...eligibleAdvance];
  if (!toResolve.length) {
    return jsonResponse(409, {
      ok: false,
      error: 'Nenhum item está com Tracking OK para concluir. Atualize ou regrave o Tracking primeiro.',
      errors: blockedErrors,
    });
  }

  const resolved = [];
  const note = resolutionNote || 'Concluído pelo PCP após conferência de Tracking OK.';
  for (const item of toResolve) {
    const saved = await resolveStageUpdateRecord(item.id, updates, session, note);
    if (saved) resolved.push(saved);
  }
  if (!isSupabaseConfigured() && resolved.length) await saveUpdates(updates);

  return jsonResponse(200, {
    ok: blockedErrors.length === 0,
    partial: blockedErrors.length > 0,
    updates: resolved,
    errors: blockedErrors,
    storage: isSupabaseConfigured() ? 'supabase' : 'json',
  });
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return jsonResponse(200, { ok: true });
  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  const session = auth.session;

  try {
    if (event.httpMethod === 'GET') {
      const mode = String(event.queryStringParameters?.mode || '').trim();
      let updates = [];
      let listWarning = '';
      try {
        updates = await listUpdates();
      } catch (error) {
        listWarning = error.message || 'Não foi possível carregar o histórico de apontamentos.';
        console.warn('Falha ao listar apontamentos; seguindo com lista vazia:', listWarning);
        updates = [];
      }
      if (mode === 'history-date-pending') {
        if (!canValidate(session)) return jsonResponse(403, { ok: false, error: 'Apenas PCP ou administrador pode consultar pendências.' });
        const pendencies = await listHistoryDatePendencies(updates);
        return jsonResponse(200, { ok: true, pendencies });
      }
      let enriched = updates;
      let autoResolved = { changed: false, updates };
      let warning = '';
      try {
        enriched = await enrichUpdatesWithTracking(updates);
        autoResolved = await autoResolveTrackingMatchedUpdates(enriched, updates, session);
      } catch (error) {
        warning = error.message || 'Não foi possível cruzar apontamentos com o Tracking agora.';
        console.warn('Falha ao enriquecer apontamentos com Tracking:', warning);
      }
      return jsonResponse(200, {
        ok: true,
        updates: autoResolved.updates,
        warning: [listWarning, warning].filter(Boolean).join(' | '),
        autoResolvedCount: autoResolved.changed
          ? autoResolved.updates.filter((item) => isResolvedStatus(item?.status) && String(item?.resolutionNote || '').includes('Tracking já estava OK')).length
          : 0,
        permissions: {
          canCreate: canCreate(session),
          canValidate: canValidate(session),
          canPcpPointAsSector: isPcpUser(session),
          sector: getActorSector(session),
        },
        progressOptions: PROGRESS_OPTIONS,
      });
    }

    if (event.httpMethod === 'POST') {
      if (!canCreate(session)) {
        return jsonResponse(403, { ok: false, error: 'Seu perfil não pode lançar apontamentos setoriais.' });
      }
      const body = JSON.parse(event.body || '{}');
      const items = Array.isArray(body.items) ? body.items : null;
      if (items && items.length) {
        let baseUpdates = [];
        try {
          baseUpdates = isSupabaseConfigured() ? [] : await listUpdates();
        } catch (_) {
          baseUpdates = [];
        }
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
        if (!isSupabaseConfigured() && created.length) await saveUpdates(baseUpdates);
        return jsonResponse(created.length ? 200 : 400, {
          ok: created.length > 0,
          updates: created,
          errors,
          storage: isSupabaseConfigured() ? 'supabase' : 'json',
        });
      }
      try {
        let updates = null;
        try {
          updates = isSupabaseConfigured() ? null : await listUpdates();
        } catch (_) {
          updates = [];
        }
        const saved = await createSingleUpdate(body, session, updates);
        if (!isSupabaseConfigured()) await saveUpdates(updates);
        return jsonResponse(200, { ok: true, update: saved, storage: isSupabaseConfigured() ? 'supabase' : 'json' });
      } catch (error) {
        return jsonResponse(error.statusCode || 400, { ok: false, error: error.message || 'Falha ao enviar apontamento.' });
      }
    }

    if (event.httpMethod === 'PUT') {
      const body = JSON.parse(event.body || '{}');
      const action = String(body.action || '').trim().toLowerCase();
      if (action === 'update-tracking') return updateTrackingAndResolve(body, session);
      if (action === 'fix-history-dates') return updateTrackingAndResolve({ ...body, dateOnly: true, forceRewrite: true }, session);
      return jsonResponse(400, { ok: false, error: 'Ação de atualização não reconhecida.' });
    }

    if (event.httpMethod === 'DELETE') {
      const body = JSON.parse(event.body || '{}');
      return deleteStageUpdateRecords(body, session);
    }

    if (event.httpMethod === 'PATCH') {
      const body = JSON.parse(event.body || '{}');
      return concludeTrackingOkOnly(body, session);
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao processar apontamentos setoriais.' });
  }
};
