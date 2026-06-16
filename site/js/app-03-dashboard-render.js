/* STEP Dashboard v37.11 - Dashboard operacional, cards, exportações, tabela e modal. Arquivo gerado a partir da divisão segura do app.js. */
function handleClientDashboardClick(event) {
  const spoolPanelRow = event.target.closest('[data-client-spool-panel]');
  if (spoolPanelRow) {
    event.preventDefault();
    event.stopPropagation();
    const projectId = spoolPanelRow.dataset.clientSpoolProjectId || state.clientPortal.selectedProjectId;
    const project = state.projects.find((item) => String(item.rowId) === String(projectId));
    const spool = findClientSpoolByPanelKey(project, spoolPanelRow.dataset.clientSpoolPanel || '');
    if (project && spool) openClientSpoolIndividualPanel(project, spool);
    return;
  }
  const reportButton = event.target.closest('[data-client-download-report]');
  if (reportButton) {
    event.preventDefault();
    event.stopPropagation();
    const project = state.projects.find((item) => String(item.rowId) === String(reportButton.dataset.clientDownloadReport));
    downloadClientTrackingReport(project);
    return;
  }
  const apiButton = event.target.closest('[data-client-open-api]');
  if (apiButton) {
    openClientApiModal();
    return;
  }

  // Botão Projetos em Desenvolvimento (client under development)
  const underDevButton = event.target.closest('[data-client-open-under-dev]');
  if (underDevButton) {
    openClientUnderDevPage();
    return;
  }
  const macroButton = event.target.closest('[data-client-open-macro-dashboard]');
  if (macroButton) {
    openClientMacroExecutive();
    return;
  }
  const vesselButton = event.target.closest('[data-client-vessel]');
  if (vesselButton) {
    // O clique simples é atrasado para não cancelar o evento nativo de duplo clique.
    // Antes, o primeiro clique renderizava novamente os cards da unidade e o segundo clique
    // caía em outro elemento, impedindo a abertura dos gráficos por unidade.
    const unitKey = vesselButton.dataset.clientVessel || '';
    if (Number(event.detail || 0) >= 2) {
      event.preventDefault();
      openClientUnitExecutiveByKey(unitKey);
      return;
    }
    window.clearTimeout(state.clientPortal.vesselClickTimer);
    state.clientPortal.vesselClickTimer = window.setTimeout(() => {
      state.clientPortal.selectedVesselKey = unitKey;
      state.clientPortal.selectedProjectId = null;
      renderClientDashboard();
      state.clientPortal.vesselClickTimer = null;
    }, 260);
    return;
  }
  const clearButton = event.target.closest('#client-clear-vessel');
  if (clearButton) {
    window.clearTimeout(state.clientPortal.vesselClickTimer);
    state.clientPortal.vesselClickTimer = null;
    state.clientPortal.selectedVesselKey = '';
    state.clientPortal.selectedProjectId = null;
    renderClientDashboard();
    return;
  }
  const row = event.target.closest('[data-client-project-id]');
  if (row) {
    const projectId = Number(row.dataset.clientProjectId || 0);
    window.clearTimeout(state.clientPortal.rowClickTimer);
    state.clientPortal.rowClickTimer = window.setTimeout(() => {
      state.clientPortal.selectedProjectId = projectId;
      renderClientBspPanel();
      state.clientPortal.rowClickTimer = null;
    }, 230);
    return;
  }
  const analyticsButton = event.target.closest('[data-client-open-analytics]');
  if (analyticsButton) {
    const project = state.projects.find((item) => String(item.rowId) === String(analyticsButton.dataset.clientOpenAnalytics));
    if (project) openClientBspExecutive(project);
  }
}

function handleClientDashboardDblClick(event) {
  if (event.target.closest('[data-client-download-report]')) {
    event.preventDefault();
    event.stopPropagation();
    return;
  }
  const vesselButton = event.target.closest('[data-client-vessel]');
  if (vesselButton && isClientUser()) {
    event.preventDefault();
    event.stopPropagation();
    openClientUnitExecutiveByKey(vesselButton.dataset.clientVessel || '');
    return;
  }

  const row = event.target.closest('[data-client-project-id]');
  if (!row || !isClientUser()) return;
  event.preventDefault();
  event.stopPropagation();
  window.clearTimeout(state.clientPortal.rowClickTimer);
  state.clientPortal.rowClickTimer = null;
  const projectId = Number(row.dataset.clientProjectId || 0);
  const project = state.projects.find((item) => String(item.rowId) === String(row.dataset.clientProjectId));
  if (!project) return;
  state.clientPortal.selectedProjectId = projectId;
  renderClientBspPanel();
  openClientBspExecutive(project);
}

function incrementTrailingNumberLabel(value, index) {
  const text = String(value || '').trim();
  const nextNumber = String(index).padStart(2, '0');
  if (!text) return `Item ${nextNumber}`;

  const patterns = [
    /(\bSP\s*[- ]*)(\d{1,3})(\s*)$/i,
    /(\bSPL\s*[- ]*)(\d{1,3})(\s*)$/i,
    /(\bSPOOL\s*[- ]*)(\d{1,3})(\s*)$/i,
    /([\s-])(\d{1,3})(\s*)$/,
  ];

  for (const pattern of patterns) {
    if (pattern.test(text)) {
      return text.replace(pattern, (match, prefix, number, suffix = '') => {
        const width = Math.max(String(number || '').length, 2);
        return `${prefix}${String(index).padStart(width, '0')}${suffix}`;
      });
    }
  }

  return `${text} - Item ${nextNumber}`;
}

function getDisplaySpoolsForProject(project, sourceSpools = null) {
  const spools = Array.isArray(sourceSpools) ? sourceSpools : (Array.isArray(project?.spools) ? project.spools : []);
  // Regra operacional: exibir/exportar somente itens realmente cadastrados no Tracking.
  // Mesmo que a BSP tenha uma quantidade informada, o sistema não deve criar ISO/SPOOL virtual.
  return spools.filter((spool) => spool && !spool.isVirtualQuantityItem);
}

function getPendingSpools(project) {
  return (project?.spools || []).filter((spool) => {
    const total = Number(spool.kilos || 0);
    const welded = Number(spool.weldedWeightKg || 0);
    return total > welded + 0.0001;
  });
}

function getBacklogItemCount(project) {
  return getPendingSpools(project).length;
}

function formatBacklogItemText(project) {
  const count = getBacklogItemCount(project);
  return `${formatNumber(count)} ${count === 1 ? "produto em produção" : "produtos em produção"}`;
}

function getProjectWeldedWeightKg(project) {
  if (!project) return 0;
  const directValue = Number(project.weldedWeightKg || 0);
  if (Number.isFinite(directValue) && directValue > 0) return directValue;
  const spools = Array.isArray(project.spools) ? project.spools : [];
  return spools.reduce((total, spool) => total + Number(spool?.weldedWeightKg || 0), 0);
}

function getProjectHoldTagCount(project) {
  return Number(project?.quantitySpools || (Array.isArray(project?.spools) ? project.spools.length : 0) || 0);
}

function getProjectHoldReason(project) {
  if (!project) return '—';
  const statusCandidates = [
    project.projectStatus,
    project.status,
    project.jobProcessStatus,
    project.currentStatus,
    project.statusSummary,
  ].filter(Boolean);

  const statusMatch = statusCandidates.find((value) => isProjectStatusOnHold(value));
  if (statusMatch) return String(statusMatch).trim();

  const noteCandidates = [
    project.observations,
    project.note,
    project.notes,
    project.summaryDrawing,
    ...(Array.isArray(project.spools) ? project.spools.flatMap((spool) => [
      spool?.observations,
      spool?.currentStatus,
      spool?.stage,
      spool?.stageStatus,
      spool?.drawing,
      spool?.description,
    ]) : []),
  ].filter(Boolean);

  const noteMatch = noteCandidates.find((value) => isProjectStatusOnHold(value));
  return noteMatch ? String(noteMatch).trim() : 'On Hold identificado';
}

function getProjectDrillMode() {
  return state.projectDrill?.mode || 'total';
}

function getProjectDelayedCardTagCount(project, expectedBucket) {
  const delayedStage = getProjectDelayedStageStats(project);
  return delayedStage.bucket === expectedBucket ? Number(delayedStage.tagCount || 0) : 0;
}

function getProjectProductionTags(project) {
  return getProjectDelayedCardTagCount(project, 'production');
}

function getProjectInspectionTags(project) {
  return getProjectDelayedCardTagCount(project, 'inspection');
}

function getProjectPaintingTags(project) {
  return getProjectDelayedCardTagCount(project, 'painting');
}

function getProjectAwaitingTagsForCard(project) {
  return getProjectDelayedCardTagCount(project, 'awaiting');
}

function getProjectStartedTagsForDrill(project) {
  return getProjectStartedTagsForSummary(project);
}

function getProjectNotStartedTagsForDrill(project) {
  return getProjectNotStartedTagsForSummary(project);
}

function getProjectOpenPaintingM2(project) {
  if (!project || isProjectFinishedForTotal(project) || isProjectOnHold(project) || isProjectPending(project)) return 0;
  const spools = Array.isArray(project.spools) ? project.spools : [];
  const openPaintingM2 = spools.length
    ? spools
        .filter((spool) => spool?.flow?.state !== 'completed' && spool?.flow?.status !== 'Finalizado')
        .reduce((total, spool) => total + Number(spool?.m2Painting || 0), 0)
    : 0;
  return openPaintingM2 > 0 ? openPaintingM2 : Number(project.m2Painting || 0);
}

function getProjectDrillMetric(project, mode = getProjectDrillMode()) {
  switch (mode) {
    case 'started': return getProjectStartedTagsForDrill(project);
    case 'not-started': return getProjectNotStartedTagsForDrill(project);
    case 'hold': return getProjectHoldTagCount(project);
    case 'production': return getProjectProductionTags(project);
    case 'inspection': return getProjectInspectionTags(project);
    case 'painting': return getProjectPaintingTags(project);
    case 'awaiting': return getProjectAwaitingTagsForCard(project);
    case 'total-weight': return Number(project?.kilos || 0);
    case 'welded': return getProjectWeldedWeightKg(project);
    case 'backlog': return getBacklogKg(project);
    case 'painting-m2': return getProjectOpenPaintingM2(project);
    case 'total':
    default: return getProjectItemCount(project);
  }
}

function getProjectDrillMetricUnit(mode = getProjectDrillMode()) {
  if (['total-weight', 'welded', 'backlog'].includes(mode)) return 'kg';
  if (mode === 'painting-m2') return 'm²';
  return '';
}

function getProjectDrillMetricLabel(mode = getProjectDrillMode()) {
  const labels = {
    total: 'projeto(s)',
    started: 'projeto(s) iniciado(s)',
    'not-started': 'projeto(s) não iniciado(s)',
    hold: 'projeto(s) em On Hold',
    production: 'projeto(s) em produção',
    inspection: 'projeto(s) em qualidade',
    painting: 'projeto(s) em pintura',
    awaiting: 'projeto(s) aguardando envio',
    'total-weight': 'kg programado',
    welded: 'kg soldado',
    backlog: 'kg pendente de solda',
    'painting-m2': 'm² programada',
  };
  return labels[mode] || 'projeto(s)';
}

function formatProjectDrillMetric(value, mode = getProjectDrillMode()) {
  if (['total-weight', 'welded', 'backlog'].includes(mode)) return `${formatNumber(value, 0)} kg`;
  if (mode === 'painting-m2') return `${formatNumber(value, 3)} m²`;
  return formatNumber(value, 0);
}

function getProjectDrillSource(mode = getProjectDrillMode()) {
  const source = getStatsProjectsSource();
  if (!Array.isArray(source)) return [];

  if (mode === 'hold') {
    return source.filter((project) => isProjectOnHold(project));
  }

  if (mode === 'started') {
    return source.filter((project) => getProjectStartedTagsForDrill(project) > 0);
  }

  if (mode === 'not-started') {
    return source.filter((project) => getProjectNotStartedTagsForDrill(project) > 0);
  }

  const active = source.filter((project) => !isProjectExcludedFromTotal(project));

  if (mode === 'production') {
    return active.filter((project) => getProjectProductionTags(project) > 0);
  }

  if (mode === 'inspection') {
    return active.filter((project) => getProjectInspectionTags(project) > 0);
  }

  if (mode === 'painting') {
    return active.filter((project) => getProjectPaintingTags(project) > 0);
  }

  if (mode === 'awaiting') {
    return active.filter((project) => getProjectAwaitingTagsForCard(project) > 0);
  }

  if (mode === 'total-weight') {
    return active.filter((project) => Number(project.kilos || 0) > 0);
  }

  if (mode === 'welded') {
    return active.filter((project) => getProjectWeldedWeightKg(project) > 0);
  }

  if (mode === 'backlog') {
    return active.filter((project) => getBacklogKg(project) > 0);
  }

  if (mode === 'painting-m2') {
    return active.filter((project) => getProjectOpenPaintingM2(project) > 0);
  }

  return source;
}

function getProjectDrillClientLabel(project) {
  const label = getProjectClientLabel(project);
  return label && label !== '—' ? label : 'Cliente não informado';
}

function getProjectDrillVesselLabel(project) {
  const label = getProjectVesselLabel(project);
  return label && label !== '—' ? label : 'Unidade não informada';
}

function createProjectDrillKey(label) {
  return normalizeText(label || 'nao informado') || 'nao informado';
}

function getProjectDrillClientGroups(projects = getProjectDrillSource()) {
  const groups = new Map();

  projects.forEach((project) => {
    const label = getProjectDrillClientLabel(project);
    const key = createProjectDrillKey(label);
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        label,
        count: 0,
        projects: [],
        totalWeightKg: 0,
        weldedWeightKg: 0,
        backlogKg: 0,
        paintingM2: 0,
        totalTags: 0,
        startedTags: 0,
        notStartedTags: 0,
        holdTags: 0,
        productionTags: 0,
        inspectionTags: 0,
        paintingTags: 0,
        awaitingTags: 0,
        metricValue: 0,
        vesselCount: 0,
      });
    }

    const group = groups.get(key);
    group.count += 1;
    group.projects.push(project);
    group.totalWeightKg += Number(project.kilos || 0);
    group.weldedWeightKg += getProjectWeldedWeightKg(project);
    group.backlogKg += getBacklogKg(project);
    group.paintingM2 += getProjectOpenPaintingM2(project);
    group.totalTags += getProjectItemCount(project);
    group.startedTags += getProjectStartedTagsForDrill(project);
    group.notStartedTags += getProjectNotStartedTagsForDrill(project);
    group.holdTags += getProjectHoldTagCount(project);
    group.productionTags += getProjectProductionTags(project);
    group.inspectionTags += getProjectInspectionTags(project);
    group.paintingTags += getProjectPaintingTags(project);
    group.awaitingTags += getProjectAwaitingTagsForCard(project);
    group.metricValue += Number(getProjectDrillMetric(project) || 0);
  });

  groups.forEach((group) => {
    const vesselKeys = new Set(group.projects.map((project) => createProjectDrillKey(getProjectDrillVesselLabel(project))));
    group.vesselCount = vesselKeys.size;
  });

  return Array.from(groups.values()).sort((a, b) => {
    const mode = getProjectDrillMode();
    if (['welded', 'total-weight', 'backlog', 'painting-m2'].includes(mode)) {
      return b.metricValue - a.metricValue || a.label.localeCompare(b.label, 'pt-BR');
    }
    return b.count - a.count || b.metricValue - a.metricValue || a.label.localeCompare(b.label, 'pt-BR');
  });
}

function getProjectDrillSelectedClientGroup() {
  const groups = getProjectDrillClientGroups();
  if (!state.projectDrill.selectedClientKey) return null;
  return groups.find((group) => group.key === state.projectDrill.selectedClientKey) || null;
}

function getProjectDrillVesselGroups(clientGroup) {
  const projects = Array.isArray(clientGroup?.projects) ? clientGroup.projects : [];
  const groups = new Map();

  projects.forEach((project) => {
    const label = getProjectDrillVesselLabel(project);
    const key = createProjectDrillKey(label);
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        label,
        count: 0,
        projects: [],
        totalWeightKg: 0,
        weldedWeightKg: 0,
        backlogKg: 0,
        paintingM2: 0,
        totalTags: 0,
        startedTags: 0,
        notStartedTags: 0,
        holdTags: 0,
        productionTags: 0,
        inspectionTags: 0,
        paintingTags: 0,
        awaitingTags: 0,
        metricValue: 0,
      });
    }

    const group = groups.get(key);
    group.count += 1;
    group.projects.push(project);
    group.totalWeightKg += Number(project.kilos || 0);
    group.weldedWeightKg += getProjectWeldedWeightKg(project);
    group.backlogKg += getBacklogKg(project);
    group.paintingM2 += getProjectOpenPaintingM2(project);
    group.totalTags += getProjectItemCount(project);
    group.startedTags += getProjectStartedTagsForDrill(project);
    group.notStartedTags += getProjectNotStartedTagsForDrill(project);
    group.holdTags += getProjectHoldTagCount(project);
    group.productionTags += getProjectProductionTags(project);
    group.inspectionTags += getProjectInspectionTags(project);
    group.paintingTags += getProjectPaintingTags(project);
    group.awaitingTags += getProjectAwaitingTagsForCard(project);
    group.metricValue += Number(getProjectDrillMetric(project) || 0);
  });

  return Array.from(groups.values()).sort((a, b) => b.count - a.count || b.metricValue - a.metricValue || a.label.localeCompare(b.label, 'pt-BR'));
}

function getProjectDrillSelectedVesselGroup(clientGroup) {
  const vesselGroups = getProjectDrillVesselGroups(clientGroup);
  if (!state.projectDrill.selectedVesselKey) return null;
  return vesselGroups.find((group) => group.key === state.projectDrill.selectedVesselKey) || null;
}

function setProjectDrillLevel({ clientKey = '', vesselKey = '' } = {}) {
  state.projectDrill.selectedClientKey = clientKey;
  state.projectDrill.selectedVesselKey = vesselKey;
  renderProjectDrillPanel();
  renderPcpDeadlineCards();
}

function closeProjectDrillPanel() {
  state.projectDrill.open = false;
  state.projectDrill.mode = 'total';
  state.projectDrill.selectedClientKey = '';
  state.projectDrill.selectedVesselKey = '';
  getProjectDrillCards().forEach((card) => {
    if (card) card.setAttribute('aria-expanded', 'false');
  });
  renderProjectDrillPanel();
}

function openProjectDrillPanel(mode = 'total') {
  state.projectDrill.open = true;
  state.projectDrill.mode = mode || 'total';
  state.projectDrill.selectedClientKey = '';
  state.projectDrill.selectedVesselKey = '';
  getProjectDrillCards().forEach((card) => {
    if (card) card.setAttribute('aria-expanded', card === getProjectDrillTriggerCard(state.projectDrill.mode) ? 'true' : 'false');
  });
  renderProjectDrillPanel();
  if (projectDrillPanelEl) {
    projectDrillPanelEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function getProjectDrillTriggerCard(mode = getProjectDrillMode()) {
  const cards = {
    total: totalProjectsCardEl,
    started: startedProjectsCardEl,
    'not-started': notStartedCardEl,
    hold: onHoldCardEl,
    production: productionCardEl,
    inspection: inspectionCardEl,
    painting: paintingCardEl,
    awaiting: awaitingShipmentCardEl,
    'total-weight': totalWeightCardEl,
    welded: weldedWeightCardEl,
    backlog: backlogWeldingCardEl,
    'painting-m2': paintingM2CardEl,
  };
  return cards[mode] || totalProjectsCardEl;
}

function getProjectDrillLabels(mode = getProjectDrillMode(), clientGroup = null, vesselGroup = null) {
  const metricLabel = getProjectDrillMetricLabel(mode);
  const metricValue = clientGroup ? formatProjectDrillMetric(clientGroup.metricValue || 0, mode) : '';
  const labels = {
    started: {
      listTitle: 'Projetos iniciados por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos que já iniciaram produção, qualidade, pintura ou logística.',
      clientTitle: 'Projetos iniciados',
      kicker: 'Projetos iniciados',
    },
    'not-started': {
      listTitle: 'Projetos não iniciados por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos que ainda estão em engenharia, suprimento ou sem início operacional.',
      clientTitle: 'Projetos não iniciados',
      kicker: 'Não iniciados',
    },
    hold: {
      listTitle: 'Projetos em On Hold por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos em On Hold vinculados a ele.',
      clientTitle: 'On Hold',
      kicker: 'On Hold por cliente',
    },
    production: {
      listTitle: 'Projetos em produção por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos com itens em corte, pré-montagem, solda ou caldeiraria.',
      clientTitle: 'Projetos em produção',
      kicker: 'Produção por cliente',
    },
    inspection: {
      listTitle: 'Projetos em qualidade por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos com itens em inspeção, END, TH ou qualidade.',
      clientTitle: 'Projetos em qualidade',
      kicker: 'Qualidade por cliente',
    },
    painting: {
      listTitle: 'Projetos em pintura por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos com itens em pintura ou coating.',
      clientTitle: 'Projetos em pintura',
      kicker: 'Pintura por cliente',
    },
    awaiting: {
      listTitle: 'Projetos preparados para envio por cliente',
      listSubtitle: 'Clique em um cliente para visualizar apenas os projetos com status Aguardando envio.',
      clientTitle: 'Preparados para envio',
      kicker: 'Preparados para envio',
    },
    'total-weight': {
      listTitle: 'Peso total programado por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos e o peso programado de cada BSP.',
      clientTitle: 'Peso programado',
      kicker: 'Peso programado por cliente',
    },
    welded: {
      listTitle: 'Peso total soldado por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos e o peso soldado de cada BSP.',
      clientTitle: 'Peso soldado',
      kicker: 'Peso soldado por cliente',
    },
    backlog: {
      listTitle: 'Peso pendente de solda por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos e o peso ainda pendente de solda.',
      clientTitle: 'Pendente de solda',
      kicker: 'Pendente de solda por cliente',
    },
    'painting-m2': {
      listTitle: '(M²) Programada por cliente',
      listSubtitle: 'Clique em um cliente para visualizar os projetos e a área operacional programada.',
      clientTitle: '(M²) Programada',
      kicker: 'Área operacional por cliente',
    },
  };

  if (mode !== 'total') {
    const config = labels[mode] || labels.started;
    if (clientGroup) {
      return {
        title: `${config.clientTitle} • ${clientGroup.label}`,
        subtitle: `${formatNumber(clientGroup.count)} projeto(s) neste cliente • ${metricValue} ${metricLabel}. Dê 2 cliques na BSP para abrir o detalhamento completo.`,
        kicker: config.kicker,
      };
    }
    return {
      title: config.listTitle,
      subtitle: config.listSubtitle,
      kicker: config.kicker,
    };
  }

  if (vesselGroup && clientGroup) {
    return {
      title: `${clientGroup.label} • ${vesselGroup.label}`,
      subtitle: `${formatNumber(vesselGroup.projects.length)} BSP(s) vinculada(s) a esta unidade/obra. Dê 2 cliques na BSP para abrir o detalhamento completo.`,
      kicker: 'Visão por cliente',
    };
  }

  if (clientGroup) {
    return {
      title: `Unidades / obras de ${clientGroup.label}`,
      subtitle: 'Clique em uma unidade para visualizar as BSPs vinculadas.',
      kicker: 'Visão por cliente',
    };
  }

  return {
    title: 'Projetos por cliente',
    subtitle: 'Selecione um cliente para abrir as unidades/obras vinculadas a ele.',
    kicker: 'Visão por cliente',
  };
}

function renderProjectDrillClientCards(clientGroups, mode) {
  if (!clientGroups.length) {
    return '<div class="project-drill-empty">Nenhum projeto disponível para detalhar.</div>';
  }

  const buildSmall = (group) => {
    if (mode === 'total') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.totalTags)} tag(s) • ${formatNumber(group.totalWeightKg, 0)} kg programado`;
    if (mode === 'started') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.startedTags)} tag(s) iniciada(s) • ${formatNumber(group.totalWeightKg, 0)} kg programado`;
    if (mode === 'not-started') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.notStartedTags)} tag(s) não iniciada(s) • ${formatNumber(group.totalWeightKg, 0)} kg programado`;
    if (mode === 'hold') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.holdTags)} tag(s) em On Hold • ${formatNumber(group.weldedWeightKg, 0)} kg soldado`;
    if (mode === 'production') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.productionTags)} tag(s) em produção • ${formatNumber(group.weldedWeightKg, 0)} kg soldado`;
    if (mode === 'inspection') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.inspectionTags)} tag(s) em qualidade • ${formatNumber(group.weldedWeightKg, 0)} kg soldado`;
    if (mode === 'painting') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.paintingTags)} tag(s) em pintura • ${formatNumber(group.paintingM2, 3)} m²`;
    if (mode === 'awaiting') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.awaitingTags || 0)} tag(s) aguardando envio • ${formatNumber(group.weldedWeightKg, 0)} kg soldado`;
    if (mode === 'total-weight') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.totalTags)} tag(s) • ${formatNumber(group.totalWeightKg, 0)} kg programado`;
    if (mode === 'welded') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.totalTags)} tag(s) • ${formatNumber(group.weldedWeightKg, 0)} kg soldado`;
    if (mode === 'backlog') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.totalTags)} tag(s) • ${formatNumber(group.backlogKg, 0)} kg pendente`;
    if (mode === 'painting-m2') return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.paintingTags)} tag(s) em pintura • ${formatNumber(group.paintingM2, 3)} m²`;
    return `${formatNumber(group.vesselCount)} unidade(s) • ${formatNumber(group.totalTags)} tag(s)`;
  };

  return `
    <div class="project-drill-grid project-drill-grid--clients">
      ${clientGroups.map((group) => `
        <button type="button" class="project-drill-card" data-drill-client="${escapeHtml(group.key)}">
          <span class="project-drill-label">${escapeHtml(group.label)}</span>
          <strong>${formatNumber(group.count)}</strong>
          <small>${buildSmall(group)}</small>
        </button>
      `).join('')}
    </div>
  `;
}

function renderProjectDrillProjectsTable(projects, mode) {
  const rows = [...projects].sort((a, b) => {
    if (['total-weight', 'welded', 'backlog', 'painting-m2'].includes(mode)) {
      return getProjectDrillMetric(b, mode) - getProjectDrillMetric(a, mode) || compareProjectsByPlannedFinishDate(a, b);
    }
    if (mode === 'awaiting') return getProjectAwaitingTagsForCard(b) - getProjectAwaitingTagsForCard(a) || compareProjectsByPlannedFinishDate(a, b);
    return compareProjectsByPlannedFinishDate(a, b);
  });

  const extraHeadMap = {
    total: '<th>Itens</th><th>Término planejado</th><th>Data de envio</th><th>% Geral</th>',
    started: '<th>Tags iniciadas</th><th>Itens</th><th>Término planejado</th><th>% Geral</th>',
    'not-started': '<th>Tags não iniciadas</th><th>Itens</th><th>Término planejado</th><th>% Geral</th>',
    hold: '<th>Tags</th><th>Motivo On Hold</th>',
    production: '<th>Tags produção</th><th>Peso programado</th><th>Peso soldado</th>',
    inspection: '<th>Tags qualidade</th><th>Peso programado</th><th>Peso soldado</th>',
    painting: '<th>Tags pintura</th><th>Área m²</th><th>Peso programado</th>',
    awaiting: '<th>Tags aguardando envio</th><th>Data de envio</th>',
    'total-weight': '<th>Peso programado</th><th>Peso soldado</th><th>Pendente solda</th>',
    welded: '<th>Peso programado</th><th>Peso soldado</th><th>Pendente solda</th>',
    backlog: '<th>Peso programado</th><th>Peso soldado</th><th>Pendente solda</th>',
    'painting-m2': '<th>Área m²</th><th>Itens</th><th>Término planejado</th>',
  };

  const getExtraCells = (project) => {
    if (mode === 'started') {
      return `<td>${formatNumber(getProjectStartedTagsForDrill(project))}</td><td>${formatNumber(getProjectItemCount(project))}</td><td>${escapeHtml(project.plannedFinishDate || '—')}</td><td>${formatPercent(project.overallProgress)}</td>`;
    }
    if (mode === 'not-started') {
      return `<td>${formatNumber(getProjectNotStartedTagsForDrill(project))}</td><td>${formatNumber(getProjectItemCount(project))}</td><td>${escapeHtml(project.plannedFinishDate || '—')}</td><td>${formatPercent(project.overallProgress)}</td>`;
    }
    if (mode === 'hold') {
      return `<td>${formatNumber(getProjectHoldTagCount(project))}</td><td>${escapeHtml(getProjectHoldReason(project))}</td>`;
    }
    if (mode === 'production') {
      return `<td>${formatNumber(getProjectProductionTags(project))}</td><td>${formatNumber(project.kilos || 0, 0)} kg</td><td>${formatNumber(getProjectWeldedWeightKg(project), 0)} kg</td>`;
    }
    if (mode === 'inspection') {
      return `<td>${formatNumber(getProjectInspectionTags(project))}</td><td>${formatNumber(project.kilos || 0, 0)} kg</td><td>${formatNumber(getProjectWeldedWeightKg(project), 0)} kg</td>`;
    }
    if (mode === 'painting') {
      return `<td>${formatNumber(getProjectPaintingTags(project))}</td><td>${formatNumber(getProjectOpenPaintingM2(project), 3)} m²</td><td>${formatNumber(project.kilos || 0, 0)} kg</td>`;
    }
    if (mode === 'awaiting') {
      return `<td>${formatNumber(getProjectAwaitingTagsForCard(project))}</td><td>${escapeHtml(getProjectShipmentDate(project) || '—')}</td>`;
    }
    if (['total-weight', 'welded', 'backlog'].includes(mode)) {
      return `<td>${formatNumber(project.kilos || 0, 0)} kg</td><td>${formatNumber(getProjectWeldedWeightKg(project), 0)} kg</td><td>${formatNumber(getBacklogKg(project), 0)} kg</td>`;
    }
    if (mode === 'painting-m2') {
      return `<td>${formatNumber(getProjectOpenPaintingM2(project), 3)} m²</td><td>${formatNumber(getProjectItemCount(project))}</td><td>${escapeHtml(project.plannedFinishDate || '—')}</td>`;
    }
    return `<td>${formatNumber(getProjectItemCount(project))}</td><td>${escapeHtml(project.plannedFinishDate || '—')}</td><td>${escapeHtml(getProjectShipmentDate(project))}</td><td>${formatPercent(project.overallProgress)}</td>`;
  };

  return `
    <div class="project-drill-table-shell">
      <table class="project-drill-table">
        <thead>
          <tr>
            <th>BSP / Projeto</th>
            <th>Tipo</th>
            <th>Etapa atual</th>
            <th>Status</th>
            ${extraHeadMap[mode] || extraHeadMap.total}
          </tr>
        </thead>
        <tbody>
          ${rows.map((project) => {
            const statusPresentation = getProjectStatusPresentation(project);
            return `
              <tr data-drill-project-id="${project.rowId}">
                <td>${escapeHtml(project.projectDisplay || '—')}</td>
                <td><span class="type-pill">${escapeHtml(getProjectTypeLabel(project))}</span></td>
                <td>${escapeHtml(getProjectCurrentStageDisplay(project))}</td>
                <td><span class="cell-status cell-status--${statusPresentation.state}">${escapeHtml(statusPresentation.text)}</span></td>
                ${getExtraCells(project)}
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderProjectDrillPanel() {
  if (!projectDrillPanelEl || !projectDrillContentEl) return;

  if (!state.projectDrill.open) {
    projectDrillPanelEl.classList.add('hidden');
    if (projectDrillContentEl) projectDrillContentEl.innerHTML = '';
    return;
  }

  const mode = getProjectDrillMode();
  const clientGroups = getProjectDrillClientGroups();
  const clientGroup = getProjectDrillSelectedClientGroup();

  if (state.projectDrill.selectedClientKey && !clientGroup) {
    state.projectDrill.selectedClientKey = '';
    state.projectDrill.selectedVesselKey = '';
  }

  const refreshedClientGroup = getProjectDrillSelectedClientGroup();
  const vesselGroup = mode === 'total' && refreshedClientGroup ? getProjectDrillSelectedVesselGroup(refreshedClientGroup) : null;

  if (mode === 'total' && state.projectDrill.selectedVesselKey && !vesselGroup) {
    state.projectDrill.selectedVesselKey = '';
  }
  if (mode !== 'total') {
    state.projectDrill.selectedVesselKey = '';
  }

  const activeClientGroup = getProjectDrillSelectedClientGroup();
  const activeVesselGroup = mode === 'total' && activeClientGroup ? getProjectDrillSelectedVesselGroup(activeClientGroup) : null;
  const showingClients = !activeClientGroup;
  const showingVessels = mode === 'total' && activeClientGroup && !activeVesselGroup;
  const showingProjects = Boolean(activeClientGroup && (mode !== 'total' || activeVesselGroup));
  const labels = getProjectDrillLabels(mode, activeClientGroup, activeVesselGroup);

  projectDrillPanelEl.classList.remove('hidden');
  getProjectDrillCards().forEach((card) => {
    if (card) card.setAttribute('aria-expanded', card === getProjectDrillTriggerCard(mode) ? 'true' : 'false');
  });

  const kickerEl = projectDrillPanelEl.querySelector('.project-drill-kicker');
  if (kickerEl) kickerEl.textContent = labels.kicker;
  if (projectDrillTitleEl) projectDrillTitleEl.textContent = labels.title;
  if (projectDrillSubtitleEl) projectDrillSubtitleEl.textContent = labels.subtitle;

  if (projectDrillBackEl) projectDrillBackEl.classList.toggle('hidden', showingClients);

  if (projectDrillBreadcrumbEl) {
    if (showingClients) {
      projectDrillBreadcrumbEl.classList.add('hidden');
      projectDrillBreadcrumbEl.innerHTML = '';
    } else {
      projectDrillBreadcrumbEl.classList.remove('hidden');
      const crumbs = [
        '<button type="button" class="project-drill-crumb" data-drill-client="">Clientes</button>',
        `<button type="button" class="project-drill-crumb" data-drill-client="${escapeHtml(activeClientGroup.key)}">${escapeHtml(activeClientGroup.label)}</button>`,
      ];
      if (activeVesselGroup) {
        crumbs.push(`<span>${escapeHtml(activeVesselGroup.label)}</span>`);
      }
      projectDrillBreadcrumbEl.innerHTML = crumbs.join('<span class="project-drill-separator">›</span>');
    }
  }

  if (showingClients) {
    projectDrillContentEl.innerHTML = renderProjectDrillClientCards(clientGroups, mode);
    return;
  }

  if (showingVessels) {
    const vesselGroups = getProjectDrillVesselGroups(activeClientGroup);
    projectDrillContentEl.innerHTML = `
      <div class="project-drill-grid project-drill-grid--vessels">
        ${vesselGroups.map((group) => `
          <button type="button" class="project-drill-card project-drill-card--vessel" data-drill-vessel="${escapeHtml(group.key)}">
            <span class="project-drill-label">${escapeHtml(group.label)}</span>
            <strong>${formatNumber(group.count)}</strong>
            <small>1 unidade • ${formatNumber(group.totalTags)} tag(s) • ${formatNumber(group.weldedWeightKg, 0)} kg soldado</small>
          </button>
        `).join('')}
      </div>
    `;
    return;
  }

  if (showingProjects) {
    const projects = mode === 'total' ? activeVesselGroup.projects : activeClientGroup.projects;
    projectDrillContentEl.innerHTML = renderProjectDrillProjectsTable(projects, mode);
  }
}


function pcpDeadlineDayMs() {
  return 24 * 60 * 60 * 1000;
}

function getPcpDeadlineDate(project) {
  const directCandidates = [
    project?.plannedFinishDate,
    project?.plannedEndDate,
    project?.projectFinishDate,
    project?.finishDate,
    project?.deadline,
  ];
  const stageValues = project?.stageValues || {};
  const stageCandidates = [
    stageValues['Término planejado'],
    stageValues['Termino planejado'],
    stageValues['Planned Finish Date'],
    stageValues['Project Finish Date'],
    stageValues['Finish Date'],
  ];
  for (const value of [...directCandidates, ...stageCandidates]) {
    const parsed = parseClientSafeDateObject(value, { minYear: 2020, maxYear: 2055 }) || parseDateObject(value);
    if (parsed && !Number.isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function formatPcpDeadlineDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '—';
  const day = String(date.getUTCDate()).padStart(2, '0');
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const year = date.getUTCFullYear();
  return `${day}/${month}/${year}`;
}

function getPcpDeadlineDiffDays(project) {
  const date = getPcpDeadlineDate(project);
  if (!date) return null;
  const today = getCurrentBrazilDate();
  return Math.floor((date.getTime() - today.getTime()) / pcpDeadlineDayMs());
}

function getPcpDeadlineStatusText(diffDays) {
  if (diffDays == null) return 'Sem data';
  if (diffDays < 0) return `Vencida há ${Math.abs(diffDays)} dia${Math.abs(diffDays) === 1 ? '' : 's'}`;
  if (diffDays === 0) return 'Vence hoje';
  return `Vence em ${diffDays} dia${diffDays === 1 ? '' : 's'}`;
}

function isPcpDeadlineOpenProject(project) {
  if (!project) return false;
  if (isProjectFinalizedForDisplay(project)) return false;
  return Boolean(getPcpDeadlineDate(project));
}

function getPcpDeadlineProjects(mode = 'overdue') {
  const source = getVisibleProjectsSource();
  return source
    .filter((project) => isPcpDeadlineOpenProject(project))
    .map((project) => ({ project, diffDays: getPcpDeadlineDiffDays(project), deadlineDate: getPcpDeadlineDate(project) }))
    .filter((item) => {
      if (item.diffDays == null) return false;
      if (mode === 'overdue') return item.diffDays < 0;
      if (mode === 'due-7') return item.diffDays >= 0 && item.diffDays <= 7;
      return item.diffDays >= 0 && item.diffDays <= 15;
    })
    .sort((a, b) => {
      if (a.diffDays !== b.diffDays) return a.diffDays - b.diffDays;
      return String(a.project?.client || '').localeCompare(String(b.project?.client || ''), 'pt-BR')
        || String(a.project?.projectDisplay || '').localeCompare(String(b.project?.projectDisplay || ''), 'pt-BR');
    });
}

function getPcpDeadlineClientGroups(mode = 'overdue') {
  const map = new Map();
  for (const item of getPcpDeadlineProjects(mode)) {
    const label = String(item.project?.client || 'Sem cliente informado').trim() || 'Sem cliente informado';
    const key = normalizeText(label) || 'sem-cliente';
    if (!map.has(key)) map.set(key, { key, label, items: [] });
    map.get(key).items.push(item);
  }
  return Array.from(map.values()).sort((a, b) => b.items.length - a.items.length || a.label.localeCompare(b.label, 'pt-BR'));
}


function getPcpReminderDateKey(date = getCurrentBrazilDate()) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return new Date().toISOString().slice(0, 10);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

function getPcpReminderStorageKey() {
  const userKey = normalizeText(state.user?.sub || state.user?.username || state.user?.name || state.user?.email || 'pcp') || 'pcp';
  return `step-pcp-deadline-reminder-v37.30:${userKey}`;
}

function buildPcpDeadlineReminderSignature(items = []) {
  return (Array.isArray(items) ? items : [])
    .map((item) => `${item.project?.rowId || item.project?.projectNumber || item.project?.projectDisplay || ''}:${formatPcpDeadlineDate(item.deadlineDate)}:${item.diffDays}`)
    .filter(Boolean)
    .sort()
    .join('|');
}

function detectPcpDeadlineReminder() {
  if (!isPcpStageUser()) return;
  if (!Array.isArray(state.projects) || !state.projects.length) return;
  const items = getPcpDeadlineProjects('due-7');
  if (!items.length) return;
  const signature = buildPcpDeadlineReminderSignature(items);
  if (!signature) return;
  const todayKey = getPcpReminderDateKey();
  const storageKey = getPcpReminderStorageKey();
  try {
    const previous = JSON.parse(window.localStorage.getItem(storageKey) || '{}');
    if (previous?.date === todayKey && previous?.signature === signature) return;
    window.localStorage.setItem(storageKey, JSON.stringify({ date: todayKey, signature, count: items.length, notifiedAt: new Date().toISOString() }));
  } catch {}
  const clientCount = getPcpDeadlineClientGroups('due-7').length;
  const first = items[0]?.project || null;
  window.setTimeout(() => queueAttentionPopup({
    kind: 'pcp-deadline-7',
    dedupeKey: `pcp-deadline-7:${todayKey}:${signature}`,
    title: 'Lembrete PCP: BSPs a vencer',
    meta: `${formatNumber(items.length)} BSP(s) em até 7 dias • ${formatNumber(clientCount)} cliente(s)`,
    message: `${first?.projectDisplay || first?.projectNumber || 'Há BSPs'} com término planejado próximo. Abra a lista para acompanhar por cliente e PM responsável.`,
    actionLabel: 'Abrir BSPs a vencer',
  }), 700);
}

function getPcpDeadlinePm(project) {
  const candidates = [
    project?.pm,
    project?.projectManager,
    project?.responsiblePm,
    project?.pmResponsible,
    project?.stageValues?.PM,
    project?.stageValues?.['PM Responsável'],
  ];
  for (const value of candidates) {
    const text = String(value || '').trim();
    if (text && text !== '—' && text.toLowerCase() !== 'n/a') return text;
  }
  return '—';
}

function renderPcpDeadlineCards() {
  const cards = [
    { id: 'pcp-overdue-card', valueId: 'stat-pcp-overdue', subId: 'stat-pcp-overdue-subtext', mode: 'overdue' },
    { id: 'pcp-due-soon-card', valueId: 'stat-pcp-due-soon', subId: 'stat-pcp-due-soon-subtext', mode: 'due-soon' },
  ];
  const visible = isPcpStageUser();
  for (const config of cards) {
    const card = document.getElementById(config.id);
    const valueEl = document.getElementById(config.valueId);
    const subEl = document.getElementById(config.subId);
    if (!card) continue;
    card.classList.toggle('hidden', !visible);
    card.setAttribute('aria-hidden', visible ? 'false' : 'true');
    if (!visible) continue;
    const items = getPcpDeadlineProjects(config.mode);
    const clientCount = getPcpDeadlineClientGroups(config.mode).length;
    if (valueEl) valueEl.textContent = formatNumber(items.length);
    if (subEl) {
      subEl.textContent = config.mode === 'overdue'
        ? `${formatNumber(clientCount)} cliente(s) com prazo vencido`
        : `${formatNumber(clientCount)} cliente(s) nos próximos 15 dias`;
    }
  }
}

function getPcpDeadlineGroupPmSummary(items = []) {
  const counts = new Map();
  for (const item of Array.isArray(items) ? items : []) {
    const pm = getPcpDeadlinePm(item?.project || {});
    if (!pm || pm === '—') continue;
    counts.set(pm, (counts.get(pm) || 0) + 1);
  }
  const ranked = Array.from(counts.entries()).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
  if (!ranked.length) return 'PM não informado';
  if (ranked.length === 1) return ranked[0][0];
  return `${ranked[0][0]} + ${ranked.length - 1} PM${ranked.length - 1 === 1 ? '' : 's'}`;
}

function getPcpDeadlineGroupSummary(group, mode = 'overdue') {
  const items = Array.isArray(group?.items) ? group.items : [];
  const total = items.length;
  const overdue = items.filter((item) => Number(item.diffDays) < 0).length;
  const dueToday = items.filter((item) => Number(item.diffDays) === 0).length;
  const nextSeven = items.filter((item) => Number(item.diffDays) >= 0 && Number(item.diffDays) <= 7).length;
  const earliest = items.slice().sort((a, b) => Number(a.diffDays ?? 9999) - Number(b.diffDays ?? 9999))[0] || null;
  const units = new Set(items.map((item) => String(item.project?.vessel || item.project?.unit || '').trim()).filter(Boolean));
  return {
    total,
    overdue,
    dueToday,
    nextSeven,
    units: units.size,
    pmSummary: getPcpDeadlineGroupPmSummary(items),
    earliestText: earliest ? getPcpDeadlineStatusText(earliest.diffDays) : 'Sem data',
    earliestDate: earliest ? formatPcpDeadlineDate(earliest.deadlineDate) : '—',
    urgencyClass: overdue ? 'is-overdue' : (dueToday ? 'is-today' : 'is-due'),
  };
}

function renderPcpDeadlineModalContent(mode = 'overdue') {
  const groups = getPcpDeadlineClientGroups(mode);
  if (!groups.length) {
    const emptyText = mode === 'overdue' ? 'vencida' : (mode === 'due-7' ? 'a vencer em 7 dias' : 'a vencer em 15 dias');
    return `<div class="pcp-deadline-empty">Nenhuma BSP ${emptyText} para o PCP acompanhar.</div>`;
  }
  return `
    <div class="pcp-deadline-client-list" role="list">
      ${groups.map((group, index) => {
        const summary = getPcpDeadlineGroupSummary(group, mode);
        const bodyId = `pcp-deadline-client-body-${index}`;
        return `
          <section class="pcp-deadline-client-block ${index === 0 ? 'is-expanded' : ''} ${summary.urgencyClass}" data-pcp-deadline-client-block="${escapeHtml(group.key)}" role="listitem">
            <button type="button" class="pcp-deadline-client-toggle" data-pcp-deadline-toggle="${escapeHtml(group.key)}" aria-expanded="${index === 0 ? 'true' : 'false'}" aria-controls="${escapeHtml(bodyId)}">
              <span class="pcp-deadline-client-main">
                <span class="pcp-deadline-client-chevron" aria-hidden="true">›</span>
                <span>
                  <strong>${escapeHtml(group.label)}</strong>
                  <small>${escapeHtml(summary.pmSummary)} • ${formatNumber(summary.units)} unidade${summary.units === 1 ? '' : 's'}</small>
                </span>
              </span>
              <span class="pcp-deadline-client-metrics" aria-label="Resumo do cliente">
                <span class="pcp-deadline-pill pcp-deadline-pill--total">${formatNumber(summary.total)} BSP${summary.total === 1 ? '' : 's'}</span>
                ${summary.overdue ? `<span class="pcp-deadline-pill pcp-deadline-pill--overdue">${formatNumber(summary.overdue)} vencida${summary.overdue === 1 ? '' : 's'}</span>` : ''}
                ${summary.dueToday ? `<span class="pcp-deadline-pill pcp-deadline-pill--today">${formatNumber(summary.dueToday)} hoje</span>` : ''}
                ${summary.nextSeven && mode !== 'overdue' ? `<span class="pcp-deadline-pill pcp-deadline-pill--due">${formatNumber(summary.nextSeven)} em 7 dias</span>` : ''}
                <span class="pcp-deadline-pill pcp-deadline-pill--date">${escapeHtml(summary.earliestText)} • ${escapeHtml(summary.earliestDate)}</span>
              </span>
            </button>
            <div class="pcp-deadline-client-body ${index === 0 ? '' : 'hidden'}" id="${escapeHtml(bodyId)}" data-pcp-deadline-client-body="${escapeHtml(group.key)}">
              <div class="modal-table-wrap pcp-deadline-table-wrap">
                <table class="modal-table pcp-deadline-table">
                  <thead>
                    <tr>
                      <th>BSP / Projeto</th>
                      <th>Unidade</th>
                      <th>Término planejado</th>
                      <th>Situação</th>
                      <th>PM responsável</th>
                      <th>Etapa atual</th>
                      <th>Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${group.items.map(({ project, diffDays, deadlineDate }) => {
                      const statusPresentation = getProjectStatusPresentation(project);
                      return `
                        <tr data-pcp-deadline-project-id="${escapeHtml(project.rowId)}">
                          <td>${escapeHtml(project.projectDisplay || project.project || '—')}</td>
                          <td>${escapeHtml(project.vessel || project.unit || '—')}</td>
                          <td>${escapeHtml(formatPcpDeadlineDate(deadlineDate))}</td>
                          <td><span class="pcp-deadline-chip ${diffDays < 0 ? 'pcp-deadline-chip--overdue' : 'pcp-deadline-chip--due'}">${escapeHtml(getPcpDeadlineStatusText(diffDays))}</span></td>
                          <td>${escapeHtml(getPcpDeadlinePm(project))}</td>
                          <td>${escapeHtml(getProjectCurrentStageDisplay(project))}</td>
                          <td><span class="cell-status cell-status--${escapeHtml(statusPresentation.state || 'in_progress')}">${escapeHtml(statusPresentation.text || '—')}</span></td>
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        `;
      }).join('')}
    </div>
  `;
}

function openPcpDeadlinesModal(mode = 'overdue') {
  if (!isPcpStageUser()) return;
  const modal = document.getElementById('pcp-deadlines-modal');
  const titleEl = document.getElementById('pcp-deadlines-title');
  const subtitleEl = document.getElementById('pcp-deadlines-subtitle');
  const contentEl = document.getElementById('pcp-deadlines-content');
  if (!modal || !contentEl) return;
  const isOverdue = mode === 'overdue';
  const isDue7 = mode === 'due-7';
  const total = getPcpDeadlineProjects(mode).length;
  const clients = getPcpDeadlineClientGroups(mode).length;
  if (titleEl) titleEl.textContent = isOverdue ? 'BSPs vencidas' : (isDue7 ? 'BSPs a vencer em até 7 dias' : 'BSPs a vencer em 15 dias');
  if (subtitleEl) subtitleEl.textContent = `${formatNumber(total)} BSP(s) em ${formatNumber(clients)} cliente(s). Clique no cliente para expandir; na BSP, 1 clique seleciona e 2 cliques abrem o detalhamento.`;
  contentEl.innerHTML = renderPcpDeadlineModalContent(mode);
  modal.dataset.mode = mode;
  modal.classList.remove('hidden');
  modal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
}

function closePcpDeadlinesModal() {
  const modal = document.getElementById('pcp-deadlines-modal');
  if (!modal) return;
  modal.classList.add('hidden');
  modal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
}

function renderStats() {
  const statsSource = getStatsProjectsSource();
  const stats = buildClientStats(statsSource);
  // v37.68: Total de Projetos = quantidade real de projetos raiz visíveis na tabela/exportação.
  // Os demais cards são recortes operacionais e não devem reduzir o total.
  stats.totalProjects = Array.isArray(statsSource) ? statsSource.length : Number(stats.totalProjects || 0);
  state.visibleStats = stats;
  const totalFinishedWeight = getTotalFinishedWeightAllProjects();
  const totalWeightRaw = Number(stats.totalWeightKg || 0);
  const totalWeldedWeight = Number(stats.totalWeldedWeightKg || 0);
  // O card de pendente precisa bater com os próprios valores exibidos nos cards.
  // Ex.: se Programado mostra 154.931 kg e Soldado mostra 104.385 kg,
  // Pendente deve mostrar 50.546 kg, e não variar por diferença de decimais ocultos.
  const totalWeightForCard = Math.round(totalWeightRaw);
  const totalWeldedForCard = Math.round(totalWeldedWeight);
  const totalBacklogWelding = Math.max(0, totalWeightForCard - totalWeldedForCard);
  document.getElementById("stat-projects").textContent = formatNumber(stats.totalProjects);
  const startedProjectsEl = document.getElementById("stat-started-projects");
  if (startedProjectsEl) startedProjectsEl.textContent = formatNumber(stats.startedProjects || 0);
  document.getElementById("stat-spools").textContent = `${formatNumber(totalWeldedWeight, 0)} kg`;
  document.getElementById("stat-total-weight").textContent = `${formatNumber(stats.totalWeightKg, 0)} kg`;
  const backlogWeldingEl = document.getElementById("stat-backlog-welding");
  if (backlogWeldingEl) backlogWeldingEl.textContent = `${formatNumber(totalBacklogWelding, 0)} kg`;

  const currentWeekEl = document.getElementById("stat-current-week");
  if (currentWeekEl) {
    currentWeekEl.textContent = `Total enviado ${formatNumber(totalFinishedWeight, 0)} kg`;
  }

  const setTags = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.textContent = `Tags ${formatNumber(value ?? 0)}`;
  };

  setTags("stat-started-tags", stats.startedTags || 0);

  document.getElementById("stat-not-started").textContent = formatNumber(stats.notStarted);
  setTags("stat-not-started-tags", stats.notStartedTags);

  const notStartedHoldEl = document.getElementById("stat-not-started-hold");
  if (notStartedHoldEl) notStartedHoldEl.textContent = formatNumber(stats.notStartedHold);
  setTags("stat-not-started-hold-tags", stats.notStartedHoldTags);

  document.getElementById("stat-in-progress").textContent = formatNumber(stats.inProgress);
  setTags("stat-in-progress-tags", stats.inProgressTags);

  const inspectionEl = document.getElementById("stat-inspection");
  if (inspectionEl) inspectionEl.textContent = formatNumber(stats.inspectionProjects);
  setTags("stat-inspection-tags", stats.inspectionTags);

  const paintingEl = document.getElementById("stat-painting");
  if (paintingEl) paintingEl.textContent = formatNumber(stats.paintingProjects);
  setTags("stat-painting-tags", stats.paintingTags);

  const awaitingEl = document.getElementById("stat-awaiting-shipment");
  if (awaitingEl) awaitingEl.textContent = formatNumber(stats.awaitingShipment);
  const awaitingTagsEl = document.getElementById("stat-awaiting-tags");
  if (awaitingTagsEl) awaitingTagsEl.textContent = `Aguardando envio • Tags ${formatNumber(stats.awaitingShipmentTags ?? 0)}`;

  const paintingM2El = document.getElementById("stat-painting-m2");
  if (paintingM2El) paintingM2El.textContent = `${formatNumber(stats.totalPaintingM2, 3)} m²`;

  const completedEl = document.getElementById("stat-completed");
  if (completedEl) completedEl.textContent = formatNumber(stats.completed);
  setTags("stat-completed-tags", stats.completedTags);

  renderProjectDrillPanel();
  renderPcpDeadlineCards();
  detectPcpDeadlineReminder();
}


function getMilestoneDateValueFromItem(item, keys) {
  const itemKey = normalizeText(item?.key || '');
  const itemLabel = normalizeText(item?.label || '');
  for (const key of keys) {
    const normalizedKey = normalizeText(key);
    if (itemKey === normalizedKey || itemLabel === normalizedKey) {
      const value = item?.value || item?.display || item?.date || '';
      if (value && String(value).trim() && String(value).trim() !== 'N/A') return String(value).trim();
    }
  }
  return '';
}

function getProjectShipmentDate(project) {
  const keys = [
    'Project Finish Date',
    'Data de envio',
    'Data Envio',
    'Package and Delivered Date',
    'Package Delivered Date',
    'Delivery Date',
    'Shipment Date',
  ];

  const directValues = [
    project?.shipmentDate,
    project?.shippingDate,
    project?.sendDate,
    project?.projectFinishDate,
  ];
  for (const value of directValues) {
    if (value && String(value).trim() && String(value).trim() !== 'N/A') return String(value).trim();
  }

  const stageValues = project?.stageValues || {};
  for (const key of keys) {
    const value = stageValues[key];
    if (value && String(value).trim() && String(value).trim() !== 'N/A') return String(value).trim();
  }

  const milestones = Array.isArray(project?.milestones) ? project.milestones : [];
  for (const item of milestones) {
    const value = getMilestoneDateValueFromItem(item, keys);
    if (value) return value;
  }

  const spoolDates = [];
  for (const spool of Array.isArray(project?.spools) ? project.spools : []) {
    const spoolStageValues = spool?.stageValues || {};
    let value = '';
    for (const key of keys) {
      const candidate = spoolStageValues[key];
      if (candidate && String(candidate).trim() && String(candidate).trim() !== 'N/A') {
        value = String(candidate).trim();
        break;
      }
    }
    if (!value) {
      for (const item of Array.isArray(spool?.milestones) ? spool.milestones : []) {
        value = getMilestoneDateValueFromItem(item, keys);
        if (value) break;
      }
    }
    if (!value) continue;
    const parsed = parseDateObject(value);
    spoolDates.push({ value, time: parsed ? parsed.getTime() : 0 });
  }

  if (spoolDates.length) {
    spoolDates.sort((a, b) => b.time - a.time);
    return spoolDates[0].value;
  }


  return '—';
}

function sanitizeFilenamePart(value) {
  return normalizeText(value || '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || 'todos';
}

function excelXmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, ' ');
}

function excelCell(value, type = 'String') {
  if (type === 'Number') {
    const number = Number(value);
    const safeNumber = Number.isFinite(number) ? number : 0;
    return `<Cell><Data ss:Type="Number">${safeNumber}</Data></Cell>`;
  }
  return `<Cell><Data ss:Type="String">${excelXmlEscape(value == null || value === '' ? '—' : value)}</Data></Cell>`;
}

const CLIENT_TRACKING_REPORT_COLUMNS = [
  { label: 'Primary', width: 118 },
  { label: 'Project', width: 135 },
  { label: 'Client', width: 115 },
  { label: 'Vessel', width: 115 },
  { label: 'Client PO Number', width: 145 },
  { label: 'Priority', width: 70, type: 'number' },
  { label: 'Project Type', width: 110 },
  { label: 'PM', width: 150 },
  { label: 'Drawing', width: 230 },
  { label: 'Line Nº', width: 180 },
  { label: 'Size', width: 80 },
  { label: 'Quantity Spools', width: 105, type: 'number' },
  { label: 'Kilos', width: 95, type: 'number' },
  { label: 'Drawing Execution Advance%', width: 135, type: 'percent' },
  { label: 'Spool Assemble and tack weld', width: 150, type: 'percent' },
  { label: 'Full welding execution', width: 135, type: 'percent' },
  { label: 'Non Destructive Examination (QC)', width: 165, type: 'percent' },
  { label: 'Final Dimensional Inpection/3D (QC)', width: 185, type: 'percent' },
  { label: 'Hydro Test Pressure (QC)', width: 150, type: 'percent' },
  { label: 'HDG / FBE.  (PAINT)', width: 135, type: 'percent-or-text' },
  { label: 'Surface preparation and/or coating', width: 180, type: 'percent' },
  { label: 'Final Inspection', width: 120, type: 'percent' },
  { label: 'Package and Delivered', width: 140, type: 'percent' },
  { label: '% Individual Progress', width: 130, type: 'percent' },
];

const CLIENT_TRACKING_REPORT_STAGE_KEYS = {
  'Drawing Execution Advance%': 'Drawing Execution Advance%',
  'Spool Assemble and tack weld': 'Spool Assemble and tack weld',
  'Full welding execution': 'Full welding execution',
  'Non Destructive Examination (QC)': 'Non Destructive Examination (QC)',
  'Final Dimensional Inpection/3D (QC)': 'Final Dimensional Inpection/3D (QC)',
  'Hydro Test Pressure (QC)': 'Hydro Test Pressure (QC)',
  'HDG / FBE.  (PAINT)': 'HDG / FBE.  (PAINT)',
  'Surface preparation and/or coating': 'Surface preparation and/or coating',
  'Final Inspection': 'Final Inspection',
  'Package and Delivered': 'Package and Delivered',
};

function excelCellWithStyle(value, type = 'String', styleId = '') {
  const styleAttr = styleId ? ` ss:StyleID="${excelXmlEscape(styleId)}"` : '';
  if (value == null || value === '') return `<Cell${styleAttr}/>`;
  if (type === 'Number' || type === 'Percent') {
    const number = Number(value);
    if (!Number.isFinite(number)) return `<Cell${styleAttr}><Data ss:Type="String">${excelXmlEscape(value)}</Data></Cell>`;
    return `<Cell${styleAttr}><Data ss:Type="Number">${number}</Data></Cell>`;
  }
  return `<Cell${styleAttr}><Data ss:Type="String">${excelXmlEscape(value)}</Data></Cell>`;
}

function getClientTrackingReportPo(project) {
  const list = Array.isArray(project?.customerPoList) ? project.customerPoList : [];
  const candidates = [project?.customerPo, ...list, project?.customerPoDisplay]
    .map((value) => String(value || '').trim())
    .filter((value) => value && !/aguardando\s+po/i.test(value));
  return Array.from(new Set(candidates)).join(', ');
}

function clientReportNumberValue(value) {
  if (value == null || value === '' || value === '—') return '';
  let text = String(value).trim().replace(/\s/g, '');
  if (!text) return '';
  const hasComma = text.includes(',');
  const hasDot = text.includes('.');
  if (hasComma && hasDot) {
    text = text.lastIndexOf(',') > text.lastIndexOf('.')
      ? text.replace(/\./g, '').replace(',', '.')
      : text.replace(/,/g, '');
  } else if (hasComma) {
    text = text.replace(',', '.');
  }
  text = text.replace(/[^\d.-]/g, '');
  const number = Number(text);
  return Number.isFinite(number) ? number : '';
}

function clientReportPercentNumber(value) {
  if (value == null || value === '') return '';
  if (String(value).trim().toUpperCase() === 'N/A') return 'N/A';
  const raw = String(value).replace('%', '').replace(',', '.').trim();
  const number = Number(raw);
  if (!Number.isFinite(number)) return '';
  if (number > 1) return Number((number / 100).toFixed(6));
  return number;
}

function getClientTrackingReportStageValue(source, label) {
  const key = CLIENT_TRACKING_REPORT_STAGE_KEYS[label];
  if (!key) return '';
  const stageValues = source?.stageValues || {};
  const direct = stageValues[key];
  if (direct == null || direct === '') return '';
  if (String(direct).trim().toUpperCase() === 'N/A') return 'N/A';
  return clientReportPercentNumber(direct);
}

function getClientTrackingReportProjectText(project, spool = null) {
  if (spool) {
    const ref = String(spool.projectRef || spool.projectDisplay || '').trim();
    if (ref) return ref.replace(/^BSP\s+/i, '');
    return String(project?.projectNumber || project?.projectDisplay || '').replace(/^BSP\s+/i, '').trim();
  }
  return project?.projectDisplay || project?.projectNumber || '—';
}

function buildClientTrackingReportRows(project) {
  if (!project) return [];
  const po = getClientTrackingReportPo(project);
  const baseProjectRow = {
    'Primary': project.primary || project.rowNumber || '',
    'Project': getClientTrackingReportProjectText(project),
    'Client': getProjectClientLabel(project),
    'Vessel': getProjectVesselLabel(project),
    'Client PO Number': po,
    'Priority': clientReportNumberValue(project.priority),
    'Project Type': getProjectTypeLabel(project),
    'PM': project.pm || '',
    'Drawing': project.summaryDrawing || 'ISO',
    'Line Nº': project.lineNumber || '',
    'Size': project.size || '',
    'Quantity Spools': Number(project.quantitySpools || getProjectItemCount(project) || 0),
    'Kilos': Number(project.kilos || 0),
    'Drawing Execution Advance%': getClientTrackingReportStageValue(project, 'Drawing Execution Advance%'),
    'Spool Assemble and tack weld': getClientTrackingReportStageValue(project, 'Spool Assemble and tack weld'),
    'Full welding execution': getClientTrackingReportStageValue(project, 'Full welding execution'),
    'Non Destructive Examination (QC)': getClientTrackingReportStageValue(project, 'Non Destructive Examination (QC)'),
    'Final Dimensional Inpection/3D (QC)': getClientTrackingReportStageValue(project, 'Final Dimensional Inpection/3D (QC)'),
    'Hydro Test Pressure (QC)': getClientTrackingReportStageValue(project, 'Hydro Test Pressure (QC)'),
    'HDG / FBE.  (PAINT)': getClientTrackingReportStageValue(project, 'HDG / FBE.  (PAINT)'),
    'Surface preparation and/or coating': getClientTrackingReportStageValue(project, 'Surface preparation and/or coating'),
    'Final Inspection': getClientTrackingReportStageValue(project, 'Final Inspection'),
    'Package and Delivered': getClientTrackingReportStageValue(project, 'Package and Delivered'),
    '% Individual Progress': clientReportPercentNumber(project.individualProgress || project.overallProgress || 0),
    _summary: true,
  };

  const rows = [baseProjectRow];
  const spools = getDisplaySpoolsForProject(project);
  spools.forEach((spool, index) => {
    rows.push({
      'Primary': spool.primary || `${project.primary || project.rowNumber || ''}-${String(index + 1).padStart(2, '0')}`,
      'Project': getClientTrackingReportProjectText(project, spool),
      'Client': spool.client || getProjectClientLabel(project),
      'Vessel': spool.vessel || getProjectVesselLabel(project),
      'Client PO Number': po,
      'Priority': clientReportNumberValue(spool.priority || project.priority),
      'Project Type': spool.projectType || getProjectTypeLabel(project),
      'PM': spool.pm || project.pm || '',
      'Drawing': spool.drawing || spool.iso || '',
      'Line Nº': spool.lineNumber || '',
      'Size': spool.size || '',
      'Quantity Spools': Number(spool.quantitySpools || 1),
      'Kilos': Number(spool.kilos || 0),
      'Drawing Execution Advance%': getClientTrackingReportStageValue(spool, 'Drawing Execution Advance%'),
      'Spool Assemble and tack weld': getClientTrackingReportStageValue(spool, 'Spool Assemble and tack weld'),
      'Full welding execution': getClientTrackingReportStageValue(spool, 'Full welding execution'),
      'Non Destructive Examination (QC)': getClientTrackingReportStageValue(spool, 'Non Destructive Examination (QC)'),
      'Final Dimensional Inpection/3D (QC)': getClientTrackingReportStageValue(spool, 'Final Dimensional Inpection/3D (QC)'),
      'Hydro Test Pressure (QC)': getClientTrackingReportStageValue(spool, 'Hydro Test Pressure (QC)'),
      'HDG / FBE.  (PAINT)': getClientTrackingReportStageValue(spool, 'HDG / FBE.  (PAINT)'),
      'Surface preparation and/or coating': getClientTrackingReportStageValue(spool, 'Surface preparation and/or coating'),
      'Final Inspection': getClientTrackingReportStageValue(spool, 'Final Inspection'),
      'Package and Delivered': getClientTrackingReportStageValue(spool, 'Package and Delivered'),
      '% Individual Progress': clientReportPercentNumber(spool.individualProgress ?? spool.overallProgress ?? 0),
      _summary: false,
    });
  });
  return rows;
}

function excelCellWithOptions(value, type = 'String', styleId = '', options = {}) {
  const attrs = [];
  if (styleId) attrs.push(`ss:StyleID="${excelXmlEscape(styleId)}"`);
  if (Number.isFinite(Number(options.mergeAcross)) && Number(options.mergeAcross) > 0) attrs.push(`ss:MergeAcross="${Number(options.mergeAcross)}"`);
  if (Number.isFinite(Number(options.index)) && Number(options.index) > 0) attrs.push(`ss:Index="${Number(options.index)}"`);
  const attrText = attrs.length ? ` ${attrs.join(' ')}` : '';
  if (value == null || value === '') return `<Cell${attrText}/>`;
  if (type === 'Number' || type === 'Percent') {
    const number = Number(value);
    if (!Number.isFinite(number)) return `<Cell${attrText}><Data ss:Type="String">${excelXmlEscape(value)}</Data></Cell>`;
    return `<Cell${attrText}><Data ss:Type="Number">${number}</Data></Cell>`;
  }
  return `<Cell${attrText}><Data ss:Type="String">${excelXmlEscape(value)}</Data></Cell>`;
}

function excelMergedCell(value, styleId = 'Section', mergeAcross = 1) {
  return excelCellWithOptions(value, 'String', styleId, { mergeAcross });
}

function excelRow(cells = [], height = null) {
  const heightAttr = Number.isFinite(Number(height)) ? ` ss:Height="${Number(height)}"` : '';
  return `<Row${heightAttr}>${cells.join('')}</Row>`;
}

function excelBlankRow() {
  return '<Row/>';
}

function excelSectionRow(title, subtitle = '', mergeAcross = 9) {
  const text = subtitle ? `${title} | ${subtitle}` : title;
  return excelRow([excelMergedCell(text, 'Section', mergeAcross)], 22);
}

function excelMetricPair(label, value) {
  return [excelCellWithStyle(label, 'String', 'MetricLabel'), excelCellWithStyle(value, 'String', 'MetricValue')];
}

function excelMetricRows(metrics = [], pairsPerRow = 5) {
  const rows = [];
  for (let i = 0; i < metrics.length; i += pairsPerRow) {
    const slice = metrics.slice(i, i + pairsPerRow);
    const cells = [];
    slice.forEach((item) => cells.push(...excelMetricPair(item[0], item[1])));
    while (cells.length < pairsPerRow * 2) cells.push(excelCellWithStyle('', 'String', 'NormalCell'));
    rows.push(excelRow(cells, 24));
  }
  return rows.join('');
}

function excelPercentNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return Number((clampClientPercent(number) / 100).toFixed(6));
}

function excelSignedPercentNumber(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return '';
  return Number((number / 100).toFixed(6));
}

function excelPanelStageStatus(percent) {
  const value = clampClientPercent(percent);
  if (value >= 99.9) return 'Concluído';
  if (value > 0) return 'Em andamento';
  return 'Não iniciado';
}

function buildClientExecutiveDetailRows(project) {
  const spools = Array.isArray(project?.spools) ? [...project.spools] : [];
  return spools.sort(compareClientSpoolsByPriority).map((spool) => {
    const state = getClientSpoolVisualState(spool);
    const statusText = spool?.currentStatus || spool?.stage || uiStateLabel(spool?.uiState);
    return {
      iso: spool.iso || spool.tag || spool.projectRef || '—',
      description: spool.description || '—',
      observations: spool.observations ? String(spool.observations).trim() : '—',
      status: statusText || '—',
      sector: spool.currentSector || spool.operationalSector || '—',
      progress: excelPercentNumber(spool.overallProgress),
      kilos: Number(spool.kilos || 0),
      state,
    };
  });
}

function buildClientExecutivePanelWorkbook(project) {
  const reportRows = buildClientTrackingReportRows(project);
  const safeProjectName = sanitizeWorksheetName((project?.projectNumber || project?.projectDisplay || 'Painel').replace(/^BSP\s+/i, '') || 'Painel');
  const status = getProjectStatusPresentation(project);
  const stages = getClientProductionStages(project);
  const overall = getClientOverallProgress(project);
  const fabrication = getClientFabricationProgress(project);
  const plannedToday = getClientPlannedToday(project);
  const completedTags = getClientCompletedTags(project);
  const totalTags = getProjectItemCount(project);
  const remainingTags = Math.max(0, totalTags - completedTags);
  const weight = Number(project?.kilos || 0);
  const welded = Number(project?.weldedWeightKg || 0);
  const pending = Math.max(0, weight - welded);
  const timeline = getClientStageTimeline(project);
  const attention = getClientAttentionPoints(project);
  const scheduleRows = buildClientExecutiveSchedule(project);
  const curveRows = buildClientSCurveData(project);
  const detailRows = buildClientExecutiveDetailRows(project);
  const startDate = clientFormatDateValue(getClientAnalyticStartDate(project));
  const finishDate = clientFormatDateValue(getClientAnalyticFinishDate(project));
  const shipmentDate = clientFormatDateValue(getProjectShipmentDate(project));
  const deviationPercent = Math.max(0, plannedToday - overall);
  const generatedAt = new Date().toLocaleString('pt-BR');
  const po = getClientTrackingReportPo(project) || '—';

  const panelCols = [120, 145, 120, 145, 120, 145, 120, 145, 120, 145].map((width) => `<Column ss:Width="${width}"/>`).join('');
  const panelRows = [];
  panelRows.push(excelRow([excelMergedCell('STEP OIL & GAS | SPOOL FABRICATION - DASHBOARD', 'PanelTitle', 9)], 30));
  panelRows.push(excelRow([excelMergedCell(`${getClientProjectDisplayCode(project)} • ${getProjectClientLabel(project)} • ${getProjectVesselLabel(project)}`, 'PanelSubtitle', 9)], 24));
  panelRows.push(excelRow([excelMergedCell(`Exportado em ${generatedAt} | Layout completo do painel executivo do cliente`, 'PanelMeta', 9)], 20));
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('Informações principais', 'Tracking + Work in Progress'));
  panelRows.push(excelMetricRows([
    ['Project', getClientTrackingReportProjectText(project)],
    ['Client', getProjectClientLabel(project)],
    ['Vessel', getProjectVesselLabel(project)],
    ['Client PO Number', po],
    ['PM', project?.pm || '—'],
    ['Priority', project?.priority || '—'],
    ['Project Type', getProjectTypeLabel(project) || '—'],
    ['Status', status.text || '—'],
    ['Início', startDate || '—'],
    ['Término planejado', finishDate || '—'],
    ['Planejado hoje', formatPercent(plannedToday)],
    ['Desvio', formatPercent(deviationPercent)],
    ['Envio efetivo', shipmentDate || '—'],
    ['Etapa atual', getProjectCurrentStageDisplay(project) || '—'],
    ['M² programada', formatNumber(project?.m2Painting, 3)],
  ]));
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('KPIs do painel', 'Visão executiva'));
  panelRows.push(excelRow([
    excelCellWithStyle('Indicador', 'String', 'Header'),
    excelCellWithStyle('Valor', 'String', 'Header'),
    excelCellWithStyle('Indicador', 'String', 'Header'),
    excelCellWithStyle('Valor', 'String', 'Header'),
    excelCellWithStyle('Indicador', 'String', 'Header'),
    excelCellWithStyle('Valor', 'String', 'Header'),
    excelCellWithStyle('Indicador', 'String', 'Header'),
    excelCellWithStyle('Valor', 'String', 'Header'),
    excelCellWithStyle('Indicador', 'String', 'Header'),
    excelCellWithStyle('Valor', 'String', 'Header'),
  ], 22));
  panelRows.push(excelMetricRows([
    ['Overall Progress', formatPercent(overall)],
    ['Fabrication Progress', formatPercent(fabrication)],
    ['Peso programado', `${formatNumber(weight, 0)} kg`],
    ['Peso soldado', `${formatNumber(welded, 0)} kg`],
    ['Peso restante', `${formatNumber(pending, 0)} kg`],
    ['Tags totais', formatNumber(totalTags, 0)],
    ['Tags finalizadas', formatNumber(completedTags, 0)],
    ['Tags restantes', formatNumber(remainingTags, 0)],
    ['Progresso planejado hoje', formatPercent(plannedToday)],
    ['Desvio acumulado', formatPercent(deviationPercent)],
  ]));
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('Progress by Production Stage', 'Etapas principais'));
  panelRows.push(excelRow([
    excelCellWithStyle('Etapa', 'String', 'Header'),
    excelCellWithStyle('Progresso', 'String', 'Header'),
    excelCellWithStyle('Status', 'String', 'Header'),
    excelMergedCell('', 'Header', 6),
  ], 22));
  stages.forEach((stage) => {
    panelRows.push(excelRow([
      excelCellWithStyle(stage.label, 'String', 'NormalCell'),
      excelCellWithStyle(excelPercentNumber(stage.percent), 'Percent', 'Percent'),
      excelCellWithStyle(excelPanelStageStatus(stage.percent), 'String', stage.percent >= 99.9 ? 'Good' : stage.percent > 0 ? 'Warning' : 'NormalCell'),
      excelMergedCell('', 'NormalCell', 6),
    ], 22));
  });
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('Timeline operacional', 'Resumo dos passos do schedule'));
  panelRows.push(excelRow([
    excelCellWithStyle('Etapa', 'String', 'Header'),
    excelCellWithStyle('Progresso', 'String', 'Header'),
    excelCellWithStyle('Estado', 'String', 'Header'),
    excelMergedCell('', 'Header', 6),
  ], 22));
  timeline.forEach((item) => {
    panelRows.push(excelRow([
      excelCellWithStyle(item.label, 'String', 'NormalCell'),
      excelCellWithStyle(excelPercentNumber(item.percent), 'Percent', 'Percent'),
      excelCellWithStyle(excelPanelStageStatus(item.percent), 'String', item.percent >= 99.9 ? 'Good' : item.percent > 0 ? 'Warning' : 'NormalCell'),
      excelMergedCell('', 'NormalCell', 6),
    ], 22));
  });
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('S-Curve | Planejado x Realizado', 'Dados que alimentam a curva'));
  panelRows.push(excelRow([
    excelCellWithStyle('Data', 'String', 'Header'),
    excelCellWithStyle('Planejado', 'String', 'Header'),
    excelCellWithStyle('Realizado', 'String', 'Header'),
    excelCellWithStyle('Desvio', 'String', 'Header'),
    excelCellWithStyle('Marco / Observação', 'String', 'Header'),
    excelMergedCell('', 'Header', 4),
  ], 22));
  curveRows.forEach((point) => {
    const planned = Number(point.planned || 0);
    const actual = point.actual == null ? '' : Number(point.actual || 0);
    const deviation = actual === '' ? '' : actual - planned;
    panelRows.push(excelRow([
      excelCellWithStyle(clientFormatDateValue(point.date), 'String', 'NormalCell'),
      excelCellWithStyle(excelPercentNumber(planned), 'Percent', 'Percent'),
      excelCellWithStyle(actual === '' ? '—' : excelPercentNumber(actual), actual === '' ? 'String' : 'Percent', actual === '' ? 'Muted' : 'Percent'),
      excelCellWithStyle(deviation === '' ? '—' : excelSignedPercentNumber(deviation), deviation === '' ? 'String' : 'Percent', deviation > 0 ? 'Good' : deviation < 0 ? 'Warning' : 'Percent'),
      excelCellWithStyle(point.trackingLabel || '', 'String', 'NormalCell'),
      excelMergedCell('', 'NormalCell', 4),
    ], 20));
  });
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('Schedule Executivo da BSP', 'Planejado + datas reais do Tracking quando preenchidas'));
  panelRows.push(excelRow([
    excelCellWithStyle('Etapa', 'String', 'Header'),
    excelCellWithStyle('%', 'String', 'Header'),
    excelCellWithStyle('Prazo médio', 'String', 'Header'),
    excelCellWithStyle('Início', 'String', 'Header'),
    excelCellWithStyle('Término', 'String', 'Header'),
    excelCellWithStyle('Status', 'String', 'Header'),
    excelCellWithStyle('Desvio', 'String', 'Header'),
    excelMergedCell('', 'Header', 2),
  ], 22));
  scheduleRows.forEach((row) => {
    const state = getClientScheduleVisualState(row.progress, row);
    const statusText = state === 'completed' ? 'Concluído' : state === 'in-progress' ? 'Em andamento' : 'Não iniciado';
    panelRows.push(excelRow([
      excelCellWithStyle(row.label || '—', 'String', row.type === 'group' ? 'Summary' : 'NormalCell'),
      excelCellWithStyle(excelPercentNumber(row.progress), 'Percent', 'Percent'),
      excelCellWithStyle(`${formatNumber(row.duration, 0)}d`, 'String', 'NormalCell'),
      excelCellWithStyle(formatClientDateShort(row.start), 'String', 'NormalCell'),
      excelCellWithStyle(formatClientDateShort(row.finish), 'String', 'NormalCell'),
      excelCellWithStyle(statusText, 'String', state === 'completed' ? 'Good' : state === 'in-progress' ? 'Warning' : 'NormalCell'),
      excelCellWithStyle(Number.isFinite(row.deviationDays) && row.deviationDays > 0 ? `+${row.deviationDays}d` : '—', 'String', row.deviationDays > 0 ? 'Warning' : 'NormalCell'),
      excelMergedCell('', 'NormalCell', 2),
    ], 20));
  });
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('S-Curve | Attention Points', 'Análise automática'));
  if (attention.length) {
    attention.forEach((item, index) => panelRows.push(excelRow([
      excelCellWithStyle(index + 1, 'Number', 'Number'),
      excelCellWithOptions(item, 'String', 'NormalCell', { mergeAcross: 8 }),
    ], 22)));
  } else {
    panelRows.push(excelRow([excelMergedCell('Nenhum ponto de atenção automático encontrado.', 'Muted', 9)], 22));
  }
  panelRows.push(excelBlankRow());

  panelRows.push(excelSectionRow('Report Tracking + Work in Progress', 'Tabela consolidada da BSP e spools'));
  panelRows.push(excelRow(CLIENT_TRACKING_REPORT_COLUMNS.map((column) => excelCellWithStyle(column.label, 'String', 'Header')), 26));
  reportRows.forEach((row) => {
    const rowStyle = row._summary ? 'Summary' : 'NormalCell';
    const cells = CLIENT_TRACKING_REPORT_COLUMNS.map((column) => {
      const value = row[column.label];
      if (column.type === 'number') return excelCellWithStyle(value, 'Number', row._summary ? 'SummaryNumber' : 'Number');
      if (column.type === 'percent' || column.type === 'percent-or-text') {
        if (String(value).trim().toUpperCase() === 'N/A') return excelCellWithStyle('N/A', 'String', rowStyle);
        return excelCellWithStyle(value, 'Percent', row._summary ? 'SummaryPercent' : 'Percent');
      }
      return excelCellWithStyle(value, 'String', rowStyle);
    });
    panelRows.push(excelRow(cells, row._summary ? 24 : 21));
  });

  const reportColumnDefs = CLIENT_TRACKING_REPORT_COLUMNS.map((column) => `<Column ss:Width="${column.width || 110}"/>`).join('');
  const reportSheetRows = [];
  reportSheetRows.push(excelRow([excelMergedCell('Report Tracking + Work in Progress', 'PanelTitle', Math.max(1, CLIENT_TRACKING_REPORT_COLUMNS.length - 1))], 28));
  reportSheetRows.push(excelRow([excelMergedCell(`${getClientProjectDisplayCode(project)} | PO ${po}`, 'PanelSubtitle', Math.max(1, CLIENT_TRACKING_REPORT_COLUMNS.length - 1))], 22));
  reportSheetRows.push(excelBlankRow());
  reportSheetRows.push(excelRow(CLIENT_TRACKING_REPORT_COLUMNS.map((column) => excelCellWithStyle(column.label, 'String', 'Header')), 26));
  reportRows.forEach((row) => {
    const rowStyle = row._summary ? 'Summary' : 'NormalCell';
    const cells = CLIENT_TRACKING_REPORT_COLUMNS.map((column) => {
      const value = row[column.label];
      if (column.type === 'number') return excelCellWithStyle(value, 'Number', row._summary ? 'SummaryNumber' : 'Number');
      if (column.type === 'percent' || column.type === 'percent-or-text') {
        if (String(value).trim().toUpperCase() === 'N/A') return excelCellWithStyle('N/A', 'String', rowStyle);
        return excelCellWithStyle(value, 'Percent', row._summary ? 'SummaryPercent' : 'Percent');
      }
      return excelCellWithStyle(value, 'String', rowStyle);
    });
    reportSheetRows.push(excelRow(cells, row._summary ? 24 : 21));
  });

  const detailColumnDefs = [145, 260, 300, 160, 160, 90, 95].map((width) => `<Column ss:Width="${width}"/>`).join('');
  const detailSheetRows = [];
  detailSheetRows.push(excelRow([excelMergedCell('Detalhamento da obra', 'PanelTitle', 6)], 28));
  detailSheetRows.push(excelRow([excelMergedCell(`${getClientProjectDisplayCode(project)} | ${getProjectClientLabel(project)}`, 'PanelSubtitle', 6)], 22));
  detailSheetRows.push(excelBlankRow());
  detailSheetRows.push(excelRow(['Tag/ISO', 'Descrição', 'Observação', 'Status', 'Etapa', '%', 'Peso'].map((label) => excelCellWithStyle(label, 'String', 'Header')), 26));
  if (detailRows.length) {
    detailRows.forEach((row) => {
      detailSheetRows.push(excelRow([
        excelCellWithStyle(row.iso, 'String', 'NormalCell'),
        excelCellWithStyle(row.description, 'String', 'NormalCell'),
        excelCellWithStyle(row.observations, 'String', 'NormalCell'),
        excelCellWithStyle(row.status, 'String', row.state === 'completed' ? 'Good' : row.state === 'in-progress' ? 'Warning' : 'NormalCell'),
        excelCellWithStyle(row.sector, 'String', 'NormalCell'),
        excelCellWithStyle(row.progress, 'Percent', 'Percent'),
        excelCellWithStyle(row.kilos, 'Number', 'Number'),
      ], 22));
    });
  } else {
    detailSheetRows.push(excelRow([excelMergedCell('Nenhuma tag detalhada encontrada para esta BSP.', 'Muted', 6)], 22));
  }

  const styles = `
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10"/></Style>
  <Style ss:ID="PanelTitle"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Bold="1" ss:Color="#FFFFFF" ss:Size="14"/><Interior ss:Color="#0B3A5A" ss:Pattern="Solid"/></Style>
  <Style ss:ID="PanelSubtitle"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Bold="1" ss:Color="#0B3A5A" ss:Size="11"/><Interior ss:Color="#DFF4FF" ss:Pattern="Solid"/></Style>
  <Style ss:ID="PanelMeta"><Alignment ss:Vertical="Center" ss:Horizontal="Center"/><Font ss:Color="#475569" ss:Size="9"/><Interior ss:Color="#F8FAFC" ss:Pattern="Solid"/></Style>
  <Style ss:ID="Section"><Alignment ss:Vertical="Center"/><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#0F766E" ss:Pattern="Solid"/></Style>
  <Style ss:ID="Header"><Alignment ss:Vertical="Center" ss:Horizontal="Center" ss:WrapText="1"/><Font ss:Bold="1" ss:Color="#0F2F46"/><Interior ss:Color="#EAF6FB" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#94A3B8"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  <Style ss:ID="MetricLabel"><Alignment ss:Vertical="Center"/><Font ss:Bold="1" ss:Color="#0F766E" ss:Size="9"/><Interior ss:Color="#F0FDFA" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#CCFBF1"/></Borders></Style>
  <Style ss:ID="MetricValue"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Bold="1" ss:Color="#0F172A"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  <Style ss:ID="NormalCell"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Color="#1E293B"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#EEF2F7"/></Borders></Style>
  <Style ss:ID="Number"><Alignment ss:Vertical="Center"/><NumberFormat ss:Format="0.00"/><Font ss:Color="#1E293B"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  <Style ss:ID="Percent"><Alignment ss:Vertical="Center"/><NumberFormat ss:Format="0.00%"/><Font ss:Color="#1E293B"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  <Style ss:ID="Good"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Bold="1" ss:Color="#065F46"/><Interior ss:Color="#ECFDF5" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#A7F3D0"/></Borders></Style>
  <Style ss:ID="Warning"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Bold="1" ss:Color="#92400E"/><Interior ss:Color="#FFFBEB" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#FDE68A"/></Borders></Style>
  <Style ss:ID="Muted"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Color="#64748B" ss:Italic="1"/><Interior ss:Color="#F8FAFC" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#E2E8F0"/></Borders></Style>
  <Style ss:ID="Summary"><Alignment ss:Vertical="Center" ss:WrapText="1"/><Font ss:Bold="1" ss:Color="#0F172A"/><Interior ss:Color="#F1F5F9" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#94A3B8"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#94A3B8"/></Borders></Style>
  <Style ss:ID="SummaryNumber"><Alignment ss:Vertical="Center"/><NumberFormat ss:Format="0.00"/><Font ss:Bold="1" ss:Color="#0F172A"/><Interior ss:Color="#F1F5F9" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#94A3B8"/></Borders></Style>
  <Style ss:ID="SummaryPercent"><Alignment ss:Vertical="Center"/><NumberFormat ss:Format="0.00%"/><Font ss:Bold="1" ss:Color="#0F172A"/><Interior ss:Color="#F1F5F9" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#94A3B8"/></Borders></Style>
 </Styles>`;

  return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Title>${excelXmlEscape(`${project?.projectDisplay || 'BSP'} Painel Executivo`)}</Title>
  <Author>STEP Dashboard</Author>
  <Created>${new Date().toISOString()}</Created>
 </DocumentProperties>
 ${styles}
 <Worksheet ss:Name="${excelXmlEscape(sanitizeWorksheetName(`Painel ${safeProjectName}`))}">
  <Table>${panelCols}${panelRows.join('')}</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>3</SplitHorizontal><TopRowBottomPane>3</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Report Tracking WIP">
  <Table>${reportColumnDefs}${reportSheetRows.join('')}</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Detalhamento">
  <Table>${detailColumnDefs}${detailSheetRows.join('')}</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>4</SplitHorizontal><TopRowBottomPane>4</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions>
 </Worksheet>
</Workbook>`;
}

function buildClientTrackingReportWorkbook(project) {
  return buildClientExecutivePanelWorkbook(project);
}

function sanitizeWorksheetName(value) {
  return String(value || 'Report')
    .replace(/[\\/?:*\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 31) || 'Report';
}

function downloadClientTrackingReport(project) {
  if (!project) {
    window.alert('Não foi possível localizar a BSP para gerar o report.');
    return;
  }
  const workbook = buildClientTrackingReportWorkbook(project);
  const blob = new Blob(['\ufeff', workbook], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const po = sanitizeFilenamePart(getClientTrackingReportPo(project) || 'sem-po');
  const projectPart = sanitizeFilenamePart(project.projectNumber || project.projectDisplay || 'bsp');
  const filename = `painel_executivo_${projectPart}_PO_${po}_${new Date().toISOString().slice(0, 10)}.xls`;
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(link.href);
    link.remove();
  }, 0);
}

function buildFilteredProjectsExportRows() {
  const projects = Array.isArray(state.filteredProjects) ? state.filteredProjects : [];
  const stageOrder = Array.isArray(state.meta?.stageOrder) ? state.meta.stageOrder : [];
  const stageColumns = stageOrder.map((stage) => ({
    key: stage.key,
    label: stage.label || stage.key,
    type: stage.type || 'text',
  }));

  const baseHeaders = [
    'Projeto',
    'Tipo',
    'Cliente',
    'Unidade / Vessel',
    'PM',
    'Status do projeto',
    'Demanda atual',
    'Setor responsável',
    'Término planejado',
    'Data de envio',
    'Qtd. itens do projeto',
    'Item / ISO / Spool',
    'Descrição do item',
    'Observação',
    '% individual projeto',
    '% geral projeto',
    '% item',
    'Peso total projeto (kg)',
    'Peso soldado projeto (kg)',
    'Peso item (kg)',
    'Peso soldado item (kg)',
    'Área operacional projeto (m²)',
    'Área item (m²)',
    'Semana solda',
    'Etapa item',
    'Setor item',
    'Finalizado?',
  ];
  const headers = [...baseHeaders, ...stageColumns.map((stage) => stage.label)];
  const rows = [];

  for (const project of projects) {
    const spools = getDisplaySpoolsForProject(project);
    const exportItems = spools.length ? spools : [null];
    const statusPresentation = getProjectStatusPresentation(project);
    const projectStage = getProjectCurrentStageDisplay(project);
    const projectSector = sectorLabel(getProjectSectorForScopedView(project)) || sectorLabel(getFlowSectorKey(project?.flow || {})) || '—';

    for (const spool of exportItems) {
      const spoolStage = spool ? getSpoolStageLabel(project, spool) : projectStage;
      const spoolSector = spool
        ? (spool.currentSector || spool.operationalSector || sectorLabel(getSpoolCompetenceSector(project, spool)) || sectorLabel(getFlowSectorKey(spool.flow || {})) || projectSector)
        : projectSector;
      const itemFinished = spool ? Boolean(spool.finished || spool.uiState === 'completed') : Boolean(project.finished || project.uiState === 'completed');
      const baseValues = [
        project.projectDisplay || project.projectNumber || '—',
        getProjectTypeLabel(project),
        getProjectClientLabel(project),
        getProjectVesselLabel(project),
        project.pm || '—',
        statusPresentation.text,
        projectStage,
        projectSector,
        project.plannedFinishDate || '—',
        getProjectShipmentDate(project),
        getProjectItemCount(project),
        spool ? (spool.iso || spool.drawing || '—') : 'Projeto sem itens internos detalhados',
        spool ? (spool.description || '—') : (project.description || '—'),
        spool ? (spool.observations || '—') : (project.observations || '—'),
        formatPercent(project.individualProgress),
        formatPercent(project.overallProgress),
        spool ? formatPercent(spool.individualProgress ?? spool.overallProgress ?? spool.progress ?? '') : '—',
        formatNumber(project.kilos, 2),
        formatNumber(project.weldedWeightKg, 2),
        spool ? formatNumber(spool.kilos, 2) : '—',
        spool ? formatNumber(spool.weldedWeightKg, 2) : '—',
        formatNumber(project.m2Painting, 3),
        spool ? formatNumber(spool.m2Painting, 3) : '—',
        spool ? (spool.weldingWeek || project.weldingWeek || '—') : (project.weldingWeek || '—'),
        spoolStage,
        spoolSector,
        itemFinished ? 'Sim' : 'Não',
      ];
      const stageValues = spool?.stageValues || project.stageValues || {};
      const dynamicValues = stageColumns.map((stage) => {
        const value = stageValues[stage.key];
        if (value == null || value === '') return '—';
        return stage.type === 'percent' ? formatPercent(value) : String(value);
      });
      rows.push([...baseValues, ...dynamicValues]);
    }
  }

  return { headers, rows };
}

function getActiveExportFilterLabel() {
  const pieces = [];
  if (state.projectView === 'mine') pieces.push('meus-projetos');
  if (state.sectorScopedView) pieces.push(`setor-${sectorLabel(getPrimaryUserSector()).toLowerCase()}`);
  if (state.demandFilter) pieces.push(`demanda-${state.demandFilter}`);
  if (state.projectTypeFilter) pieces.push(`tipo-${state.projectTypeFilter}`);
  if (state.weekFilter) pieces.push(`semana-${state.weekFilter}`);
  if (Array.isArray(state.statusFilters) && state.statusFilters.length) pieces.push(`status-${state.statusFilters.join('-')}`);
  if (state.searchQuery) pieces.push(`busca-${state.searchQuery}`);
  return pieces.length ? pieces.map(sanitizeFilenamePart).filter(Boolean).join('_') : 'todos-os-projetos';
}

function updateExportFilteredProjectsButton() {
  if (!exportFilteredProjectsEl) return;
  const count = Array.isArray(state.filteredProjects) ? state.filteredProjects.length : 0;
  exportFilteredProjectsEl.disabled = count <= 0;
  exportFilteredProjectsEl.textContent = count > 0 ? `Baixar Excel (${formatNumber(count)})` : 'Baixar Excel';
  exportFilteredProjectsEl.title = count > 0
    ? `Baixar o detalhamento dos ${formatNumber(count)} projeto(s) filtrado(s)`
    : 'Nenhum projeto filtrado para exportar';
}

function downloadFilteredProjectsExcel() {
  const projects = Array.isArray(state.filteredProjects) ? state.filteredProjects : [];
  if (!projects.length) {
    window.alert('Nenhum projeto filtrado para exportar.');
    return;
  }

  const { headers, rows } = buildFilteredProjectsExportRows();
  const generatedAt = new Date().toLocaleString('pt-BR');
  const title = 'Detalhamento de projetos filtrados';
  const filterLabel = getActiveExportFilterLabel();

  const headerRow = `<Row>${headers.map((header) => excelCell(header)).join('')}</Row>`;
  const dataRows = rows.map((row) => `<Row>${row.map((cell) => excelCell(cell)).join('')}</Row>`).join('');
  const columnDefs = headers.map(() => '<Column ss:AutoFitWidth="1" ss:Width="135"/>').join('');

  const workbook = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Title>${excelXmlEscape(title)}</Title>
  <Author>STEP Dashboard</Author>
  <Created>${new Date().toISOString()}</Created>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="10"/></Style>
  <Style ss:ID="Title"><Font ss:Bold="1" ss:Size="13"/><Interior ss:Color="#D9EAF7" ss:Pattern="Solid"/></Style>
  <Style ss:ID="Header"><Font ss:Bold="1"/><Interior ss:Color="#C6E0B4" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
 </Styles>
 <Worksheet ss:Name="Detalhamento">
  <Table>
   ${columnDefs}
   <Row ss:StyleID="Title"><Cell ss:MergeAcross="5"><Data ss:Type="String">${excelXmlEscape(title)}</Data></Cell></Row>
   <Row><Cell><Data ss:Type="String">Gerado em</Data></Cell><Cell><Data ss:Type="String">${excelXmlEscape(generatedAt)}</Data></Cell></Row>
   <Row><Cell><Data ss:Type="String">Projetos filtrados</Data></Cell><Cell><Data ss:Type="String">${excelXmlEscape(formatNumber(projects.length))}</Data></Cell></Row>
   <Row><Cell><Data ss:Type="String">Linhas detalhadas</Data></Cell><Cell><Data ss:Type="String">${excelXmlEscape(formatNumber(rows.length))}</Data></Cell></Row>
   <Row></Row>
   ${headerRow.replace('<Row>', '<Row ss:StyleID="Header">')}
   ${dataRows}
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>6</SplitHorizontal>
   <TopRowBottomPane>6</TopRowBottomPane>
   <ActivePane>2</ActivePane>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;

  const blob = new Blob(['\ufeff', workbook], { type: 'application/vnd.ms-excel;charset=utf-8;' });
  const link = document.createElement('a');
  const filename = `detalhamento_${filterLabel}_${new Date().toISOString().slice(0, 10)}.xls`;
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  window.setTimeout(() => {
    URL.revokeObjectURL(link.href);
    link.remove();
  }, 0);
}

function renderTable() {
  if (!state.filteredProjects.length) {
    bodyEl.innerHTML = '<tr><td colspan="21" class="loading-cell">Nenhum projeto encontrado para a busca informada.</td></tr>';
    searchCountEl.textContent = "0 resultado(s)";
    updateExportFilteredProjectsButton();
    return;
  }

  searchCountEl.textContent = `${state.filteredProjects.length} resultado(s)`;
  updateExportFilteredProjectsButton();

  bodyEl.innerHTML = state.filteredProjects
    .map((project) => {
      const isActive = project.rowId === state.selectedProjectId;
      const statusPresentation = getProjectStatusPresentation(project);
      const statusText = statusPresentation.text;
      const rowClass = [
        ["completed", "awaiting_shipment"].includes(project.uiState) ? "completed-row" : "",
        project.uiState === "in_progress" ? "in-progress-row" : "",
        project.uiState === "not_started" ? "not-started-row" : "",
        isActive ? "active-row" : "",
      ]
        .filter(Boolean)
        .join(" ");

      const stageMap = project.stageValues || {};
      const completedSymbol = ["completed", "awaiting_shipment"].includes(project.uiState) ? "✓" : "✕";
      const statusState = statusPresentation.state;

      return `
        <tr class="${rowClass}" data-project-id="${project.rowId}">
          <td>${project.projectDisplay || "—"}</td>
          <td><span class="type-pill">${getProjectTypeLabel(project)}</span></td>
          <td class="project-client-cell" title="${escapeHtml(getProjectClientLabel(project))}">${escapeHtml(getProjectClientLabel(project))}</td>
          <td class="project-vessel-cell" title="${escapeHtml(getProjectVesselLabel(project))}">${escapeHtml(getProjectVesselLabel(project))}</td>
          <td>${project.plannedFinishDate || "—"}</td>
          <td>${formatNumber(getProjectItemCount(project))}</td>
          <td>${formatNumber(project.weldedWeightKg, 0)}</td>
          <td>${project.weldingWeek || "—"}</td>
          <td>${formatNumber(project.kilos, 2)}</td>
          <td>${formatNumber(project.m2Painting, 3)}</td>
          <td>
            <span class="stage-pill">
              <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
              <span class="stage-text">${getProjectCurrentStageDisplay(project)}</span>
            </span>
          </td>
          <td>${formatPercent(project.individualProgress)}</td>
          <td>${formatPercent(project.overallProgress)}</td>
          <td><span class="cell-status cell-status--${statusState}">${statusText}</span></td>
          <td>${stageMap["Fabrication Start Date"] || "—"}</td>
          <td>${stageMap["Boilermaker Finish Date"] || "—"}</td>
          <td>${stageMap["Welding Finish Date"] || "—"}</td>
          <td>${stageMap["Inspection Finish Date (QC)"] || "—"}</td>
          <td>${stageMap["TH Finish Date"] || "—"}</td>
          <td>${getProjectShipmentDate(project)}</td>
          <td class="cell-finished cell-finished--${project.finished ? "yes" : "no"}">${completedSymbol}</td>
        </tr>
      `;
    })
    .join("");
}

function renderSelectedProjectCard() {
  if (!detailCardEl) return;

  if (!state.selectedProjectDrawerOpen) {
    if (detailDrawerEl) detailDrawerEl.classList.add("hidden");
    detailCardEl.innerHTML = "";
    return;
  }

  const project = getSelectedProject();
  if (!project) {
    if (detailDrawerEl) detailDrawerEl.classList.add("hidden");
    detailCardEl.innerHTML = "";
    return;
  }

  if (detailDrawerEl) detailDrawerEl.classList.remove("hidden");

  const statusPresentation = getProjectStatusPresentation(project);
  const statusText = statusPresentation.text;
  const matchedSpools = getProjectItemCount(project);

  detailCardEl.innerHTML = `
    <div class="detail-hero compact">
      <div class="detail-project-title">
        <div>
          <p class="detail-project-subtitle">Projeto selecionado</p>
          <h3>${projectDisplayWithClient(project)}</h3>
        </div>
        <span class="badge badge--${statusPresentation.state}">${statusText}</span>
      </div>

      <div class="detail-grid compact-grid">
        <div class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(getProjectItemCount(project))}</strong></div>
        <div class="metric-chip"><span>Tipo</span><strong>${getProjectTypeLabel(project)}</strong></div>
        <div class="metric-chip"><span>Cliente</span><strong>${getProjectClientLabel(project)}</strong></div>
        <div class="metric-chip"><span>Unidade</span><strong>${getProjectVesselLabel(project)}</strong></div>
        <div class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></div>
        <button class="metric-chip metric-chip--button" type="button" id="open-backlog-project">
          <span>Backlog KG</span><strong>${formatNumber(getBacklogKg(project), 0)} kg</strong><small>${formatBacklogItemText(project)}</small>
        </button>
        <div class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></div>
        <div class="metric-chip"><span>Início planejado</span><strong>${project.plannedStartDate || "—"}</strong></div>
        <div class="metric-chip"><span>Término planejado</span><strong>${project.plannedFinishDate || "—"}</strong></div>
        <div class="metric-chip"><span>Data de envio</span><strong>${getProjectShipmentDate(project)}</strong></div>
        <div class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 0)}kg</strong></div>
        <div class="metric-chip"><span>Área operacional</span><strong>${formatNumber(project.m2Painting, 3)}</strong></div>
        <div class="metric-chip"><span>% Individual</span><strong>${formatPercent(project.individualProgress)}</strong></div>
        <div class="metric-chip"><span>% Geral</span><strong>${formatPercent(project.overallProgress)}</strong></div>
        <div class="metric-chip"><span>Itens internos</span><strong>${matchedSpools}</strong></div>
      </div>

      <div class="current-stage-box ${project.currentStageAlert ? "alert" : ""}">
        <div class="current-stage-head">
          <span class="current-stage-label">Etapa atual</span>
          <span class="stage-progress">${formatPercent(project.currentStagePercent)}</span>
        </div>
        <div class="stage-pill">
          <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
          <span class="stage-name">${getProjectCurrentStageDisplay(project)}</span>
        </div>
      </div>

      <div class="detail-actions">
        <button class="primary-button" type="button" id="open-selected-project">Abrir detalhamento completo</button>
        ${canOpenClientBspPanel(project) ? '<button class="ghost-button" type="button" id="open-selected-client-panel">Painel do Cliente</button>' : ''}
      </div>
    </div>
  `;

  const button = document.getElementById("open-selected-project");
  if (button) {
    button.addEventListener("click", () => openProjectModal(project));
  }

  const clientPanelButton = document.getElementById("open-selected-client-panel");
  if (clientPanelButton) {
    clientPanelButton.addEventListener("click", () => {
      if (canManageClientBspPanel(project)) openClientBspExecutiveForPmEdit(project);
      else if (canOpenClientBspPanel(project)) openClientBspExecutive(project);
    });
  }

  const backlogButton = document.getElementById("open-backlog-project");
  if (backlogButton) {
    backlogButton.addEventListener("click", () => openProjectModal(project, { pendingOnly: true }));
  }
}

function renderModal(project) {
  const stageOrder = state.meta?.stageOrder || [];
  const milestoneList = (project.milestones || [])
    .map((item) => `<div class="milestone-chip"><span>${item.key || item.label}</span><strong>${item.value}</strong></div>`)
    .join("");

  const baseSpools = state.modalPendingOnly ? getPendingSpools(project) : (project.spools || []);
  const sourceSpools = getDisplaySpoolsForProject(project, baseSpools);
  const modalIsoSortMode = state.modalIsoSortMode || 'urgency';
  const sortByIsoNatural = (a, b) => String(a?.iso || '').localeCompare(String(b?.iso || ''), 'pt-BR', { numeric: true, sensitivity: 'base' });
  const sortedSpools = [...sourceSpools].sort((a, b) => {
    if (modalIsoSortMode === 'iso') return sortByIsoNatural(a, b);
    const aProgress = Number.isFinite(Number(a?.individualProgress)) ? Number(a.individualProgress) : 999999;
    const bProgress = Number.isFinite(Number(b?.individualProgress)) ? Number(b.individualProgress) : 999999;
    if (aProgress !== bProgress) return aProgress - bProgress;
    return sortByIsoNatural(a, b);
  });
  const spoolRows = sortedSpools
    .map((spool) => {
      const stageColumns = stageOrder
        .map((stage) => {
          const value = spool.stageValues?.[stage.key];
          const formatted = value == null || value === "" ? "—" : stage.type === "percent" ? formatPercent(value) : value;
          const cellClass = tableCellClass(value, stage.type);
          return `<td class="${cellClass}">${formatted}</td>`;
        })
        .join("");

      const observations = spool.observations ? escapeHtml(spool.observations).replace(/\n/g, "<br>") : "—";
      const spoolStatusText = spool.currentStatus || spool.stage || uiStateLabel(spool.uiState);
      const spoolSectorText = spool.currentSector || spool.operationalSector || sectorLabel(getFlowSectorKey(spool.flow || {})) || "—";
      const spoolStatusClass = spool.finished || spool.uiState === "completed" ? "completed" : (spool.uiState === "awaiting_shipment" ? "preparing_shipment" : (spool.uiState || "in_progress"));

      return `
        <tr data-modal-row="true">
          <td>${spool.iso || "—"}</td>
          <td>${spool.description || "—"}</td>
          <td class="modal-observation-cell">${observations}</td>
          <td>${formatNumber(spool.weldedWeightKg, 0)} kg</td>
          <td>${spool.weldingWeek || "—"}</td>
          <td>${formatNumber(spool.kilos, 2)}</td>
          <td>${formatNumber(spool.m2Painting, 3)}</td>
          <td><span class="cell-status cell-status--${spoolStatusClass}">${escapeHtml(spoolStatusText)}</span></td>
          <td class="${percentStateClass(spool.stagePercent)}">${escapeHtml(spoolSectorText)}</td>
          <td class="${percentStateClass(spool.individualProgress)}">${formatPercent(spool.individualProgress)}</td>
          <td class="${percentStateClass(spool.overallProgress)}">${formatPercent(spool.overallProgress)}</td>
          ${stageColumns}
        </tr>
      `;
    })
    .join("");

  const stageHeaders = stageOrder.map((stage) => `<th>${stage.label}</th>`).join("");
  const statusPresentation = getProjectStatusPresentation(project);
  const statusText = statusPresentation.text;

  modalTitleEl.textContent = projectDisplayWithClient(project);
  modalSubtitleEl.textContent = `${statusText} • ${state.modalPendingOnly ? getPendingSpools(project).length : (project.spools?.length || 0)} item(ns) interno(s)`;

  modalContentEl.innerHTML = `
    <section class="modal-summary-grid">
      <article class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(getProjectItemCount(project))}</strong></article>
      <article class="metric-chip"><span>Tipo</span><strong>${getProjectTypeLabel(project)}</strong></article>
      <article class="metric-chip"><span>Cliente</span><strong>${getProjectClientLabel(project)}</strong></article>
      <article class="metric-chip"><span>Unidade</span><strong>${getProjectVesselLabel(project)}</strong></article>
      <article class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></article>
      <article class="metric-chip metric-chip--button" id="modal-open-backlog" role="button" tabindex="0"><span>Backlog KG</span><strong>${formatNumber(getBacklogKg(project), 0)} kg</strong><small>${formatBacklogItemText(project)}</small></article>
      <article class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></article>
      <article class="metric-chip"><span>Início planejado</span><strong>${project.plannedStartDate || "—"}</strong></article>
      <article class="metric-chip"><span>Término planejado</span><strong>${project.plannedFinishDate || "—"}</strong></article>
      <article class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 0)}kg</strong></article>
      <article class="metric-chip"><span>Área operacional total</span><strong>${formatNumber(project.m2Painting, 3)}</strong></article>
      <article class="metric-chip"><span>% Individual</span><strong>${formatPercent(project.individualProgress)}</strong></article>
      <article class="metric-chip"><span>% Geral</span><strong>${formatPercent(project.overallProgress)}</strong></article>
      <article class="metric-chip"><span>Status atual</span><strong>${statusText}</strong></article>
      <article class="metric-chip"><span>Etapa atual</span><strong>${getProjectSectorSummary(project) || simplifyCurrentStage(project)}</strong></article>
    </section>

    <section class="modal-milestones">
      ${milestoneList || '<div class="empty-inline">Nenhum marco de data disponível.</div>'}
    </section>

    ${canOpenClientBspPanel(project) ? `<section class="modal-client-panel-action"><button class="ghost-button ghost-button--compact modal-client-panel-button" type="button" data-open-client-panel="${escapeHtml(project.rowId)}">${canManageClientBspPanel(project) ? 'Painel do Cliente / Editar datas' : 'Painel do Cliente'}</button><span>${canManageClientBspPanel(project) ? 'Abre a visão executiva do cliente já no modo de ajuste.' : 'Abre a visão executiva do cliente somente para consulta.'}</span></section>` : ''}

    ${renderProjectSignals(project)}
    ${renderClientTratativaNotice(project)}
    ${renderClientOnHoldNotice(project)}

    <section class="modal-iso-toolbar" aria-label="Ordenação dos ISOs">
      <div>
        <span>Organizar ISOs</span>
        <strong>${modalIsoSortMode === 'iso' ? 'Por ISO' : 'Por urgência'}</strong>
      </div>
      <div class="modal-iso-toolbar__actions">
        <button type="button" class="modal-sort-button ${modalIsoSortMode === 'urgency' ? 'is-active' : ''}" data-modal-iso-sort="urgency">Urgência</button>
        <button type="button" class="modal-sort-button ${modalIsoSortMode === 'iso' ? 'is-active' : ''}" data-modal-iso-sort="iso">ISO</button>
      </div>
    </section>

    <section class="modal-table-wrap">
      <table class="modal-table">
        <thead>
          <tr>
            <th>ISO</th>
            <th>Descrição</th>
            <th>Observações</th>
            <th>Peso soldado</th>
            <th>Semana finalizado</th>
            <th>Peso</th>
            <th>Área operacional</th>
            <th>Status</th>
            <th>Etapa atual</th>
            <th>% Individual</th>
            <th>% Geral</th>
            ${stageHeaders}
          </tr>
        </thead>
        <tbody>
          ${spoolRows || `<tr><td colspan="999" class="loading-cell">${state.modalPendingOnly ? "Nenhuma peça pendente encontrada." : "Nenhum item interno encontrado."}</td></tr>`}
        </tbody>
      </table>
    </section>
  `;
}

function openProjectModal(project, options = {}) {
  state.selectedProjectId = project.rowId;
  state.modalPendingOnly = Boolean(options.pendingOnly);
  renderTable();
  renderSelectedProjectCard();
  renderModal(project);
  modalEl.classList.remove("hidden");
  modalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeProjectModal() {
  state.modalPendingOnly = false;
  modalEl.classList.add("hidden");
  modalEl.setAttribute("aria-hidden", "true");
  if (alertModalEl.classList.contains("hidden")) {
    document.body.classList.remove("modal-open");
  }
}

function getAlertStorageKey() {
  const userKey = normalizeText(state.user?.username || state.user?.name || "guest") || "guest";
  return `step-alert-popup-state:${userKey}`;
}

function getAlertSignature() {
  return state.meta?.alertSignature || "no-alerts";
}

function getAlertCooldownMs() {
  if (!shouldUseSectorScopedToggle(state.user) && userHasProjectsScope(state.user) && state.projectView === 'mine') {
    return 0;
  }
  return 4 * 60 * 60 * 1000;
}

function getNextProjectAlertWindowTimestamp(now = new Date()) {
  const next = new Date(now);
  next.setSeconds(0, 0);
  const hour = now.getHours();
  if (hour < 9 || (hour === 9 && now.getMinutes() === 0)) {
    next.setHours(9, 0, 0, 0);
    return next.getTime();
  }
  if (hour < 14 || (hour === 14 && now.getMinutes() === 0)) {
    next.setHours(14, 0, 0, 0);
    return next.getTime();
  }
  next.setDate(next.getDate() + 1);
  next.setHours(9, 0, 0, 0);
  return next.getTime();
}

function readSavedAlertState() {
  try {
    const raw = window.localStorage.getItem(getAlertStorageKey());
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function getSuppressedUntil() {
  const saved = readSavedAlertState();
  const localSuppressedUntil = Number(saved?.suppressedUntil || 0);
  const memorySuppressedUntil = Number(state.alertPopupSuppressedUntil || 0);
  return Math.max(localSuppressedUntil, memorySuppressedUntil);
}

function shouldOpenAlertPopup() {
  const visibleAlerts = getVisibleAlertsSource();
  if (!visibleAlerts.length) return false;

  const now = Date.now();
  if (!shouldUseSectorScopedToggle(state.user) && userHasProjectsScope(state.user) && state.projectView === 'mine') {
    const windowOpen = getProjectAlertWindow();
    if (!windowOpen) return false;
  }

  const suppressedUntil = getSuppressedUntil();
  if (suppressedUntil > now) return false;
  return true;
}

function persistAlertDismiss() {
  const now = Date.now();
  let suppressedUntil = now + getAlertCooldownMs();
  if (!shouldUseSectorScopedToggle(state.user) && userHasProjectsScope(state.user) && state.projectView === 'mine') {
    suppressedUntil = getNextProjectAlertWindowTimestamp(new Date(now));
  }
  state.alertPopupSuppressedUntil = suppressedUntil;
  try {
    window.localStorage.setItem(
      getAlertStorageKey(),
      JSON.stringify({
        signature: getAlertSignature(),
        dismissedAt: now,
        suppressedUntil,
      })
    );
  } catch {}
}

function renderAlertBadge() {
  if (!alertBadgeCountEl) return;
  const totalAlerts = getVisibleAlertsSource().length || 0;
  alertBadgeCountEl.textContent = String(totalAlerts);
  if (openAlertsButtonEl) {
    openAlertsButtonEl.disabled = totalAlerts === 0;
    openAlertsButtonEl.classList.toggle("alert-badge--empty", totalAlerts === 0);
    openAlertsButtonEl.title = totalAlerts === 0 ? "Nenhum alerta ativo no momento" : "Clique para abrir os alertas";
  }
}

function renderAlertModal() {
  if (!alertModalContentEl) return;

  const visibleAlerts = getVisibleAlertsSource();
  const mediumCount = visibleAlerts.filter((alert) => getAlertSeverity(alert) === "medium").length;
  const urgentCount = visibleAlerts.filter((alert) => getAlertSeverity(alert) === "urgent").length;
  const sectorButtons = [
    { key: "solda", label: "Solda", match: ["Solda"] },
    { key: "calderaria", label: "Calderaria", match: ["Calderaria"] },
    { key: "inspecao", label: "Qualidade", match: ["Qualidade", "Inspeção"] },
    { key: "pintura", label: "Pintura", match: ["Pintura"] },
    { key: "envio", label: "Logística", match: ["Logística", "Envio", "Pendente de envio"] },
  ];
  const sectorCounts = Object.fromEntries(
    sectorButtons.map((button) => [
      button.key,
      visibleAlerts.filter((alert) => normalizeAlertSectorFilterValue(alert.sector) === button.key).length,
    ])
  );
  const filteredAlerts = getFilteredAlerts();

  const filterBar = `
    <div class="alert-filter-stack">
      <div class="alert-filter-bar">
        <button type="button" class="alert-filter-button ${state.alertFilter === "all" ? "is-active" : ""}" data-alert-filter="all">Tudo <strong>${visibleAlerts.length}</strong></button>
        <button type="button" class="alert-filter-button alert-filter-button--medium ${state.alertFilter === "medium" ? "is-active" : ""}" data-alert-filter="medium">Médio <strong>${mediumCount}</strong></button>
        <button type="button" class="alert-filter-button alert-filter-button--urgent ${state.alertFilter === "urgent" ? "is-active" : ""}" data-alert-filter="urgent">Urgente <strong>${urgentCount}</strong></button>
      </div>
      <div class="alert-filter-bar alert-filter-bar--sector">
        <button type="button" class="alert-filter-button ${state.alertSectorFilter === "all" ? "is-active" : ""}" data-alert-sector="all">Todos os setores <strong>${visibleAlerts.length}</strong></button>
        ${sectorButtons.map((button) => `<button type="button" class="alert-filter-button alert-filter-button--sector ${state.alertSectorFilter === button.key ? "is-active" : ""}" data-alert-sector="${button.key}">${button.label} <strong>${sectorCounts[button.key]}</strong></button>`).join("")}
      </div>
      <div class="alert-toolbar-row">
        <label class="alert-client-search">
          <span>Buscar cliente</span>
          <input type="text" value="${escapeHtml(state.alertClientQuery)}" placeholder="Ex.: Prio" data-alert-client-search="true" autocomplete="off" />
        </label>
        <button type="button" class="ghost-button alert-download-button" data-alert-download-pdf="true">Baixar PDF</button>
      </div>
    </div>
  `;

  if (!visibleAlerts.length) {
    alertModalContentEl.innerHTML = `${filterBar}<div class="alert-empty">Nenhum prazo em alerta no momento.</div>`;
    return;
  }

  if (!filteredAlerts.length) {
    alertModalContentEl.innerHTML = `${filterBar}<div class="alert-empty">Nenhum alerta encontrado para este filtro.</div>`;
    return;
  }

  const items = filteredAlerts
    .map((alert) => {
      const severity = getAlertSeverity(alert);
      const tone = severity === "urgent" ? "overdue" : "conference";
      const severityLabel = severity === "urgent" ? "Urgente" : "Médio";
      const projectLine = [alert.projectDisplay, alert.client].filter(Boolean).join(" ");
      const daysLabel = alert.daysRemaining < 0
        ? `${Math.abs(alert.daysRemaining)} dia(s) em atraso`
        : `${alert.daysRemaining} dia(s) para o término planejado`;
      return `
        <article class="alert-item alert-item--${tone} alert-item--clickable" data-alert-project-id="${alert.projectRowId || ""}" data-alert-project-number="${escapeHtml(alert.projectNumber || "")}">
          <div class="alert-item-head">
            <strong>${escapeHtml(projectLine)}</strong>
            <div class="alert-tag-group">
              <span class="alert-item-tag alert-item-tag--${severity}">${severityLabel}</span>
              <span class="alert-item-tag alert-item-tag--sector">${escapeHtml(alert.sector || "Geral")}</span>
              <span class="alert-item-tag">${escapeHtml(alert.title)}</span>
            </div>
          </div>
          <div class="alert-item-meta">
            <span>Término planejado: <strong>${escapeHtml(alert.plannedFinishDate || "—")}</strong></span>
            <span>${escapeHtml(daysLabel)}</span>
            <span>Pintura: <strong>${formatPercent(alert.coatingPercent)}</strong></span>
            <span>Etapa: <strong>${escapeHtml(alert.currentStage || "—")}</strong></span>
          </div>
          <p>${escapeHtml(alert.message)}</p>
        </article>
      `;
    })
    .join("");

  alertModalContentEl.innerHTML = `${filterBar}<div class="alert-list">${items}</div>`;
}


function findProjectFromAlertElement(element) {
  if (!element) return null;
  const projectId = Number(element.dataset.alertProjectId || 0);
  if (projectId) {
    const direct = state.projects.find((project) => project.rowId === projectId);
    if (direct) return direct;
  }

  const projectNumber = normalizeText(element.dataset.alertProjectNumber || "");
  if (!projectNumber) return null;
  return state.projects.find((project) => normalizeText(project.projectNumber) === projectNumber || normalizeText(project.projectDisplay) === projectNumber) || null;
}

function openAlertModal(force = false, options = {}) {
  if (!alertModalEl) return;

  const manualOpen = Boolean(options.manual);

  // O alerta continua existindo no botão/contador, mas o modal grande não abre sozinho ao carregar/reabrir o link.
  // Isso evita a tela apagada/bloqueada para novos usuários.
  if (!manualOpen) return;

  if (!force && !shouldOpenAlertPopup()) return;
  renderAlertModal();
  alertModalEl.classList.remove("hidden");
  alertModalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeAlertModal() {
  if (!alertModalEl) return;
  persistAlertDismiss();
  alertModalEl.classList.add("hidden");
  alertModalEl.setAttribute("aria-hidden", "true");
  if (modalEl.classList.contains("hidden")) {
    document.body.classList.remove("modal-open");
  }
}


function isPageHidden() {
  return document.visibilityState === 'hidden' || document.hidden === true;
}

