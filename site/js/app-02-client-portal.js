/* STEP Dashboard v37.11 - Portal do cliente, painel executivo, imagens e API cliente. Arquivo gerado a partir da divisão segura do app.js. */
function getClientPortalName(user = state.user) {
  return String(user?.clientName || user?.clientKey || user?.name || 'Cliente').trim() || 'Cliente';
}

function getClientPortalLogo(user = state.user) {
  return String(user?.clientLogoUrl || '').trim() || './assets/step-logo.png';
}


function parseClientPlatformImages(value) {
  if (!value) return {};
  if (typeof value === 'object' && !Array.isArray(value)) {
    return Object.entries(value).reduce((acc, [key, src]) => {
      const cleanKey = String(key || '').trim();
      const cleanSrc = String(src || '').trim();
      if (cleanKey && cleanSrc) acc[cleanKey] = cleanSrc;
      return acc;
    }, {});
  }

  const text = String(value || '').trim();
  if (!text) return {};

  if (text.startsWith('{')) {
    try { return parseClientPlatformImages(JSON.parse(text)); } catch (_) {}
  }

  return text.split(/\n+/).reduce((acc, line) => {
    const raw = String(line || '').trim();
    if (!raw) return acc;
    const separator = raw.includes('=') ? '=' : (raw.includes('|') ? '|' : ':');
    const parts = raw.split(separator);
    const key = String(parts.shift() || '').trim();
    const src = parts.join(separator).trim();
    if (key && src) acc[key] = src;
    return acc;
  }, {});
}

function formatClientPlatformImages(value) {
  const map = parseClientPlatformImages(value);
  return Object.entries(map).map(([key, src]) => `${key}=${src}`).join('\n');
}

function getClientPlatformImages(user = state.user) {
  return parseClientPlatformImages(user?.clientPlatformImages || user?.clientPlatformImagesText || '');
}

function getClientPortalPlatformImage(vesselLabel = '', user = state.user) {
  const images = getClientPlatformImages(user);
  const vesselKey = normalizeText(vesselLabel || '');
  for (const [key, src] of Object.entries(images)) {
    if (normalizeText(key) === vesselKey) return String(src || '').trim();
  }
  // Sem fallback: se a plataforma/vessel não tiver foto própria cadastrada, o card fica sem imagem.
  return '';
}

function ensureClientDashboardEl() {
  let el = document.getElementById('client-dashboard');
  if (el) return el;
  el = document.createElement('section');
  el.id = 'client-dashboard';
  el.className = 'client-dashboard hidden';
  el.innerHTML = `
    <div class="client-hero">
      <div class="client-identity">
        <div class="client-logo-box"><img id="client-dashboard-logo" src="./assets/step-logo.png" alt="Logo do cliente" /></div>
        <div>
          <p class="client-kicker">Portal do Cliente</p>
          <h2 id="client-dashboard-name">Cliente</h2>
          <p id="client-dashboard-meta">Demandas filtradas por empresa</p>
        </div>
      </div>
      <div class="client-hero-actions">
        <span id="client-dashboard-sync">Atualização: --</span>
        <button id="client-dashboard-refresh" class="mini-action-button client-refresh-button" type="button" title="Buscar as informações mais recentes do Smartsheet">Atualizar</button>
        <!-- Ações de macro e API são exibidas em outras áreas do portal. Ocultamos aqui para evitar duplicação.
        <button class="mini-action-button client-macro-button" type="button" data-client-open-macro-dashboard>Visão executiva da carteira</button>
        <button class="mini-action-button client-api-button" type="button" data-client-open-api>Gerar API</button>
        -->
        <!-- Botão para projetos em desenvolvimento (apenas Yinson) -->
        <button id="client-under-dev-button" class="mini-action-button client-under-dev-button" type="button" data-client-open-under-dev>Projetos em desenvolvimento</button>
      </div>
    </div>
    <div class="client-summary-grid">
      <article class="client-summary-card-button" data-client-open-macro-dashboard title="Abrir visão executiva da carteira"><span>BSPs</span><strong id="client-stat-bsps">--</strong><small>abrir visão executiva</small></article>
      <article><span>Tags</span><strong id="client-stat-tags">--</strong></article>
      <article><span>Peso programado</span><strong id="client-stat-weight">--</strong></article>
      <article><span>Peso soldado</span><strong id="client-stat-welded">--</strong></article>
      <article><span>M² programada</span><strong id="client-stat-m2">--</strong></article>
      <article><span>Progresso médio</span><strong id="client-stat-progress">--</strong></article>
    </div>
    <div class="client-search-row">
      <div class="search-box">
        <span class="search-label">Localizar BSP, PO ou Focal Point</span>
        <input id="client-project-search" type="text" placeholder="Digite BSP, PO ou nome do focal point. Ex.: Sergio Ramos" autocomplete="off" />
      </div>
      <button id="client-clear-search" class="ghost-button" type="button">Limpar busca</button>
    </div>
    <div class="client-section-head">
      <div><p class="client-kicker">Vessels / Unidades</p><h3>Carteira por unidade</h3></div>
      <p>Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos e o PDF da unidade.</p>
    </div>
    <div id="client-vessel-grid" class="client-vessel-grid"></div>
    <div id="client-bsp-panel" class="client-bsp-panel hidden">
      <div class="client-section-head client-section-head--compact">
        <div><p class="client-kicker">BSPs</p><h3 id="client-bsp-title">Projetos</h3></div>
        <button id="client-clear-vessel" class="mini-action-button" type="button">Ver todas</button>
      </div>
      <div class="client-bsp-content">
        <div id="client-bsp-table" class="client-table-wrap"></div>
        <div id="client-project-detail" class="client-project-detail hidden"></div>
      </div>
    </div>
  `;
  const anchor = document.querySelector('.summary-row');
  if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(el, anchor.nextSibling);
  else document.querySelector('.page-shell')?.appendChild(el);
  el.addEventListener('click', handleClientDashboardClick);
  el.addEventListener('dblclick', handleClientDashboardDblClick);

  const searchInput = el.querySelector('#client-project-search');
  const clearBtn = el.querySelector('#client-clear-search');
  const refreshBtn = el.querySelector('#client-dashboard-refresh');

  if (refreshBtn) {
    refreshBtn.addEventListener('click', refreshClientPortalData);
  }

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      // Sincroniza com o campo de busca global para manter consistência
      const globalSearch = document.getElementById('project-search');
      if (globalSearch) globalSearch.value = e.target.value;
      
      applyFilter();
      renderClientDashboard();
      renderClientBspPanel(); // v32.5: Atualiza a tabela de BSPs imediatamente
    });
  }

  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      state.searchQuery = '';
      if (searchInput) searchInput.value = '';
      const globalSearch = document.getElementById('project-search');
      if (globalSearch) globalSearch.value = '';
      
      applyFilter();
      renderClientDashboard();
      renderClientBspPanel(); // v32.5: Limpa a filtragem da tabela de BSPs
    });
  }

  return el;
}

function setClientDashboardMode() {
  const enabled = isClientUser();
  document.body.classList.toggle('client-mode', enabled);
  const el = ensureClientDashboardEl();
  el.classList.toggle('hidden', !enabled);
  if (openSectorAlertsEl) openSectorAlertsEl.classList.toggle('hidden', enabled);
  if (openMyProjectSignalsEl && enabled) openMyProjectSignalsEl.classList.add('hidden');
  if (openProjectSignalsEl && enabled) openProjectSignalsEl.classList.add('hidden');
  if (openStageUpdatesEl && enabled) openStageUpdatesEl.classList.add('hidden');
}

async function refreshClientPortalData(event) {
  if (event?.preventDefault) event.preventDefault();
  if (!isClientUser()) return;

  const button = document.getElementById('client-dashboard-refresh');
  const syncEl = document.getElementById('client-dashboard-sync');
  const originalText = button?.textContent || 'Atualizar';

  try {
    if (button) {
      button.disabled = true;
      button.textContent = 'Atualizando...';
      button.classList.add('is-loading');
    }
    if (syncEl) syncEl.textContent = 'Atualizando informações...';

    state.lastProjectsFetchAt = 0;

    // v37.17/v37.20: atualização automática é independente do usuário.
    // O botão do cliente é opcional/manual e passa pelo mesmo lock do Supabase.
    const syncResult = await triggerTrackingCacheSync({ force: true, manual: true, auto: false });
    if (syncResult?.cacheUpdatedAt) {
      state.meta = {
        ...(state.meta || {}),
        persistentCacheUpdatedAt: syncResult.cacheUpdatedAt,
        persistentCacheAgeMs: syncResult.cacheAgeMs ?? 0,
        cacheUpdatedAt: syncResult.cacheUpdatedAt,
        cacheAgeMs: syncResult.cacheAgeMs ?? 0,
      };
      if (syncEl) applyTrackingCacheUpdateLabel(syncEl, state.meta, { prefix: 'Última atualização do cache' });
    }
    if (syncResult?.staleCacheKept && syncEl) {
      const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
      syncEl.textContent = `${label} • ${describeTrackingSyncWarning(syncResult)}`;
    }
    // v37.24: Portal do Cliente não deve exibir popup quando o cache válido foi mantido.
    if (syncResult && syncResult.ok === false) {
      if (hasVisibleOperationalCache() && syncEl) {
        const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
        syncEl.textContent = `${label} • ${describeTrackingSyncWarning(syncResult)}`;
      } else {
        throw new Error(syncResult.error || 'Falha ao sincronizar cache.');
      }
    }

    await loadProjects({
      force: false,
      skipLocalCache: true,
      suppressLoadingState: true,
      requireData: false,
      requireClientPo: false,
      preferServerCache: true,
    });

    applyFilter();
    renderClientDashboard();
    renderClientBspPanel();

    if (syncEl) applyTrackingCacheUpdateLabel(syncEl, state.meta, { prefix: 'Última atualização do cache' });
  } catch (error) {
    console.error('[Portal Cliente] Falha ao atualizar informações:', error);
    if (syncEl && hasVisibleOperationalCache()) {
      const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
      syncEl.textContent = `${label} • cache mantido; tentativa manual não concluída`;
    } else {
      if (syncEl) syncEl.textContent = `Falha ao atualizar: ${error?.message || 'tente novamente'}`;
      window.alert(error?.message || 'Falha ao atualizar as informações do cliente.');
    }
  } finally {
    if (button) {
      button.disabled = false;
      button.textContent = originalText;
      button.classList.remove('is-loading');
    }
  }
}

function getClientVesselGroups(projects = state.projects) {
  const groups = new Map();
  for (const project of Array.isArray(projects) ? projects : []) {
    const label = getProjectVesselLabel(project) || 'Unidade não informada';
    const key = createProjectDrillKey(label);
    if (!groups.has(key)) {
      groups.set(key, { key, label, projects: [], tags: 0, weight: 0, welded: 0, m2: 0, progress: 0 });
    }
    const group = groups.get(key);
    group.projects.push(project);
    group.tags += getProjectItemCount(project);
    group.weight += Number(project.kilos || 0);
    group.welded += Number(project.weldedWeightKg || 0);
    group.m2 += Number(project.m2Painting || 0);
    group.progress += Number(project.overallProgress || 0);
  }
  return Array.from(groups.values()).map((group) => ({
    ...group,
    avgProgress: group.projects.length ? group.progress / group.projects.length : 0,
  })).sort((a, b) => b.projects.length - a.projects.length || a.label.localeCompare(b.label, 'pt-BR'));
}

function renderClientDashboard() {
  setClientDashboardMode();
  if (!isClientUser()) return;
  const el = ensureClientDashboardEl();
  // v32.4: Usa filteredProjects se houver uma busca ativa, caso contrário usa todos os projetos do cliente
  const projects = (state.searchQuery && state.searchQuery.trim()) 
    ? (state.filteredProjects || []) 
    : (state.projects || []);
  const logo = document.getElementById('client-dashboard-logo');
  if (logo) logo.src = getClientPortalLogo();
  const nameEl = document.getElementById('client-dashboard-name');
  if (nameEl) nameEl.textContent = getClientPortalName();
  const metaEl = document.getElementById('client-dashboard-meta');
  if (metaEl) metaEl.textContent = `${formatNumber(projects.length)} BSP(s) vinculada(s) ao cliente`;
  const syncEl = document.getElementById('client-dashboard-sync');
  if (syncEl) applyTrackingCacheUpdateLabel(syncEl, state.meta, { prefix: 'Última atualização do cache' });

  const totals = projects.reduce((acc, project) => {
    acc.bsps += 1;
    acc.tags += getProjectItemCount(project);
    acc.weight += Number(project.kilos || 0);
    acc.welded += Number(project.weldedWeightKg || 0);
    acc.m2 += Number(project.m2Painting || 0);
    acc.progress += Number(project.overallProgress || 0);
    return acc;
  }, { bsps: 0, tags: 0, weight: 0, welded: 0, m2: 0, progress: 0 });
  const setText = (id, text) => { const node = document.getElementById(id); if (node) node.textContent = text; };
  setText('client-stat-bsps', formatNumber(totals.bsps));
  setText('client-stat-tags', formatNumber(totals.tags));
  setText('client-stat-weight', `${formatNumber(totals.weight, 0)} kg`);
  setText('client-stat-welded', `${formatNumber(totals.welded, 0)} kg`);
  setText('client-stat-m2', `${formatNumber(totals.m2, 3)} m²`);
  setText('client-stat-progress', formatPercent(totals.bsps ? totals.progress / totals.bsps : 0));

  // Show or hide the "Projetos em desenvolvimento" button based on client name.
  // Only the Yinson portal should display this button. Other clients (e.g. SBM) must not see
  // projects under development from Yinson. We perform a case-insensitive check
  // against the portal name and toggle the button's visibility accordingly.
  (function toggleUnderDevButton() {
    const btn = document.getElementById('client-under-dev-button');
    if (!btn) return;
    const name = String(getClientPortalName() || '').toLowerCase();
    const shouldShow = name.includes('yinson');
    // Use inline style to hide the element completely rather than relying on CSS classes
    btn.style.display = shouldShow ? '' : 'none';
  })();

  const groups = getClientVesselGroups(projects);
  if (!state.clientPortal.selectedVesselKey && groups.length) state.clientPortal.selectedVesselKey = groups[0].key;
  const grid = document.getElementById('client-vessel-grid');
  if (grid) {
    // v34.1: trava o layout das unidades no próprio componente.
    // Isso evita que regras antigas de auto-fit/auto-fill ou cache visual empilhem 7/8 cards por linha.
    grid.classList.add('client-vessel-grid--locked');
    grid.style.display = 'grid';
    grid.style.gridTemplateColumns = 'repeat(4, minmax(0, 1fr))';
    grid.style.gap = '16px';
    grid.style.alignItems = 'stretch';
    grid.style.width = '100%';

    grid.innerHTML = groups.length ? groups.map((group) => `
      ${(() => {
        const platformImage = getClientPortalPlatformImage(group.label);
        return `
          <button type="button" class="client-vessel-card ${platformImage ? 'has-media' : 'no-media'} ${state.clientPortal.selectedVesselKey === group.key ? 'is-active' : ''}" data-client-vessel="${escapeHtml(group.key)}" title="Clique uma vez para listar as BSPs; dê 2 cliques para abrir os gráficos da unidade">
            <div class="client-vessel-info">
              <span>${escapeHtml(group.label)}</span>
              <strong>${formatNumber(group.projects.length)} BSP(s)</strong>
              <small>${formatNumber(group.tags)} tag(s) • ${formatNumber(group.weight, 0)} kg programado</small>
              <small>${formatNumber(group.welded, 0)} kg soldado • ${formatPercent(group.avgProgress)}</small>
              <small class="client-vessel-action">2 cliques: gráficos / PDF da unidade</small>
            </div>
            ${platformImage ? `<div class="client-vessel-media"><img src="${escapeHtml(platformImage)}" alt="Foto da plataforma ${escapeHtml(group.label)}" /></div>` : ''}
          </button>
        `;
      })()}
    `).join('') : '<div class="client-empty">Nenhuma demanda encontrada para este cliente.</div>';
  }
  renderClientBspPanel();
}

function getClientPortalSourceProjects() {
  // v32.5: Se houver pesquisa ativa, usa os projetos filtrados. Caso contrário, usa todos os projetos do cliente.
  return (state.searchQuery && state.searchQuery.trim())
    ? (state.filteredProjects || [])
    : (state.projects || []);
}

function getClientProjectsByVesselKey(vesselKey, projects = getClientPortalSourceProjects()) {
  const selected = String(vesselKey || '').trim();
  const sourceProjects = Array.isArray(projects) ? projects : [];
  if (!selected) return sourceProjects;
  return sourceProjects.filter((project) => createProjectDrillKey(getProjectVesselLabel(project) || 'Unidade não informada') === selected);
}

function getClientVesselGroupByKey(vesselKey, projects = getClientPortalSourceProjects()) {
  const selected = String(vesselKey || '').trim();
  if (!selected) return null;
  return getClientVesselGroups(projects).find((group) => group.key === selected) || null;
}

function getClientSelectedVesselProjects() {
  return getClientProjectsByVesselKey(state.clientPortal.selectedVesselKey);
}

function openClientUnitExecutiveByKey(unitKey) {
  if (!isClientUser()) return false;
  window.clearTimeout(state.clientPortal.vesselClickTimer);
  state.clientPortal.vesselClickTimer = null;
  const group = getClientVesselGroupByKey(unitKey, getClientPortalSourceProjects());
  if (!group) return false;
  state.clientPortal.selectedVesselKey = group.key;
  state.clientPortal.selectedProjectId = null;
  openClientMacroExecutive(group.projects, { scope: 'unit', unitKey: group.key, unitLabel: group.label });
  renderClientDashboard();
  return true;
}

function renderClientBspPanel() {
  const panel = document.getElementById('client-bsp-panel');
  const table = document.getElementById('client-bsp-table');
  const title = document.getElementById('client-bsp-title');
  if (!panel || !table || !title) return;
  const groups = getClientVesselGroups();
  const activeGroup = groups.find((group) => group.key === state.clientPortal.selectedVesselKey) || groups[0] || null;
  if (!activeGroup) {
    panel.classList.add('hidden');
    renderClientProjectDetail(null);
    return;
  }
  panel.classList.remove('hidden');
  title.textContent = `${activeGroup.label} • ${formatNumber(activeGroup.projects.length)} BSP(s)`;
  const projects = getClientSelectedVesselProjects().sort(compareProjectsByPlannedFinishDate);
  if (!state.clientPortal.selectedProjectId && projects.length) state.clientPortal.selectedProjectId = projects[0].rowId;
  table.innerHTML = `
    <table class="client-bsp-table">
      <thead><tr><th>BSP</th><th>Tags</th><th>Peso</th><th>Soldado</th><th>M²</th><th>Status</th><th>Etapa</th><th>% Geral</th><th>Término</th></tr></thead>
      <tbody>
        ${projects.map((project) => {
          const status = getProjectStatusPresentation(project);
          const selected = String(state.clientPortal.selectedProjectId || '') === String(project.rowId || '');
          return `<tr class="${selected ? 'is-selected' : ''}" data-client-project-id="${escapeHtml(project.rowId)}" title="Clique uma vez para selecionar; dê 2 cliques para abrir a visão executiva">
            <td><strong>${escapeHtml(getClientProjectDisplayCode(project))}</strong></td>
            <td>${formatNumber(getProjectItemCount(project))}</td>
            <td>${formatNumber(project.kilos, 0)} kg</td>
            <td>${formatNumber(project.weldedWeightKg, 0)} kg</td>
            <td>${formatNumber(project.m2Painting, 3)}</td>
            <td><span class="cell-status cell-status--${status.state}">${escapeHtml(status.text)}</span></td>
            <td>${escapeHtml(getProjectCurrentStageDisplay(project))}</td>
            <td>${formatPercent(project.overallProgress)}</td>
            <td>${escapeHtml(project.plannedFinishDate || '—')}</td>
          </tr>`;
        }).join('') || '<tr><td colspan="9" class="loading-cell">Nenhuma BSP nesta unidade.</td></tr>'}
      </tbody>
    </table>
  `;
  const selectedProject = projects.find((project) => String(project.rowId) === String(state.clientPortal.selectedProjectId)) || projects[0] || null;
  renderClientProjectDetail(selectedProject);
}

function getClientSpoolProgress(spool) {
  return clampClientPercent(spool?.overallProgress);
}

function isClientSpoolFinished(spool) {
  const progress = getClientSpoolProgress(spool);
  const statusText = normalizeText(spool?.currentStatus || spool?.stage || uiStateLabel(spool?.uiState));
  const uiState = normalizeCompactText(spool?.uiState || '');
  return progress >= 99.9 || uiState === 'completed' || /finalizado|concluido|completed|finished|enviado|delivered/.test(statusText);
}

function isClientDetailingWaitingStatus(value) {
  const text = normalizeText(value);
  return /(^|\s)ag\.?\s*emissao de detalhamento/.test(text) || /aguardando emissao de detalhamento/.test(text);
}

function isClientFinishedStatusText(value) {
  const text = normalizeText(value);
  return /finalizado|concluido|completed|finished|enviado|delivered/.test(text);
}

function getClientSpoolVisualState(spool) {
  const progress = getClientSpoolProgress(spool);
  const statusText = spool?.currentStatus || spool?.stage || uiStateLabel(spool?.uiState);
  const stageStatus = normalizeCompactText(spool?.stageStatus || spool?.currentStageStatus || spool?.flow?.stageStatus || '');
  if (isClientDetailingWaitingStatus(statusText)) {
    if (isClientFinishedStatusText(statusText)) return 'completed';
    if (stageStatus === 'inprogress' || (progress > 0 && progress < 99.9)) return 'in-progress';
    return 'not-started';
  }
  if (isClientFinishedStatusText(statusText)) return 'completed';
  if (progress <= 0 || stageStatus === 'waiting') return 'not-started';
  return 'in-progress';
}

function getClientStageStripVisualState(value) {
  const percent = clientPercentValue(value);
  if (percent >= 99.9) return 'completed';
  if (percent > 0) return 'in-progress';
  return 'not-started';
}

function renderClientStageStripCard(label, value) {
  const stateName = getClientStageStripVisualState(value);
  const displayValue = value == null || value === '' ? '—' : (String(value).includes('%') ? escapeHtml(value) : formatPercent(value));
  return `<div class="client-stage-card client-stage-card--${stateName}"><span>${escapeHtml(label)}</span><strong>${displayValue}</strong></div>`;
}

function compareClientSpoolsByPriority(a, b) {
  const aFinished = isClientSpoolFinished(a);
  const bFinished = isClientSpoolFinished(b);
  if (aFinished !== bFinished) return aFinished ? 1 : -1;
  const progressDiff = getClientSpoolProgress(a) - getClientSpoolProgress(b);
  if (Math.abs(progressDiff) > 0.001) return progressDiff;
  return String(a?.iso || '').localeCompare(String(b?.iso || ''), 'pt-BR', { numeric: true, sensitivity: 'base' });
}

function getClientSpoolPanelKey(spool, index = 0) {
  const candidates = [spool?.iso, spool?.drawing, spool?.description, spool?.lineNumber, spool?.primary]
    .map((value) => String(value || '').trim())
    .filter(Boolean);
  const raw = candidates.join('|') || `spool-${index}`;
  return normalizeCompactText(raw) || `spool-${index}`;
}

function findClientSpoolByPanelKey(project, key) {
  const wanted = String(key || '').trim();
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (!wanted || !spools.length) return null;
  return spools.find((spool, index) => getClientSpoolPanelKey(spool, index) === wanted) || null;
}

function formatClientStandalonePanelValue(value) {
  if (value == null || value === '') return '—';
  return String(value);
}

function formatClientStandalonePanelPercent(value) {
  if (value == null || value === '') return '—';
  if (String(value).trim().toUpperCase() === 'N/A') return 'N/A';
  return String(value).includes('%') ? String(value) : formatPercent(clientPercentValue(value));
}

function getClientStandaloneStageRows(project, spool = null) {
  const projectStageValues = project?.stageValues && typeof project.stageValues === 'object' ? project.stageValues : {};
  const spoolStageValues = spool?.stageValues && typeof spool.stageValues === 'object' ? spool.stageValues : {};
  const source = Object.keys(spoolStageValues).length ? spoolStageValues : projectStageValues;
  const keys = getClientDetailStageKeys(project);
  return keys.map((key) => {
    const label = (state.meta?.stageOrder || []).find((stage) => stage.key === key)?.label || key;
    return { label, value: source[key] };
  });
}

function buildClientStandalonePanelHtml(project, spool = null) {
  const isSpool = Boolean(spool);
  const status = getProjectStatusPresentation(project);
  const visualState = isSpool ? getClientSpoolVisualState(spool) : getClientProjectVisualState(project);
  const title = isSpool ? (spool.iso || spool.drawing || 'Tag / ISO') : getClientProjectDisplayCode(project);
  const subtitle = isSpool ? `${getClientProjectDisplayCode(project)} • ${getProjectVesselLabel(project)}` : `${getProjectClientLabel(project)} • ${getProjectVesselLabel(project)}`;
  const progress = isSpool ? getClientSpoolProgress(spool) : getClientOverallProgress(project);
  const stageRows = getClientStandaloneStageRows(project, spool);
  const po = getClientTrackingReportPo(project) || '—';
  const description = isSpool ? (spool.description || spool.drawing || '—') : (project.summaryDrawing || project.projectDisplay || '—');
  const observation = isSpool
    ? (spool.observations || spool.OBSERVATIONS || spool.observation || spool.note || spool.notes || '—')
    : getProjectObservationContexts(project).map((item) => `${item.source}: ${item.text}`).join(' | ') || '—';
  const currentStatus = isSpool ? (spool.currentStatus || spool.stage || uiStateLabel(spool.uiState)) : status.text;
  const currentStage = isSpool ? (spool.currentSector || spool.operationalSector || spool.flow?.sector || '—') : getProjectCurrentStageDisplay(project);
  const kilos = isSpool ? Number(spool.kilos || 0) : Number(project.kilos || 0);
  const m2 = isSpool ? Number(spool.m2Painting || 0) : Number(project.m2Painting || 0);
  const isoOverride = isSpool ? (getClientIsoDateOverride(project, spool) || {}) : {};
  const finishedDate = isSpool
    ? (isoOverride.finishDate || spool.plannedFinishDate || spool.finishDate || spool.deliveryDate || project.plannedFinishDate)
    : (project.plannedFinishDate || getClientAnalyticFinishDate(project));
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const siblingRows = !isSpool ? spools.slice().sort(compareClientSpoolsByPriority).map((item, index) => {
    const itemState = getClientSpoolVisualState(item);
    return `<tr><td>${escapeHtml(item.iso || item.drawing || `Tag ${index + 1}`)}</td><td>${escapeHtml(item.description || '—')}</td><td><span class="chip chip-${itemState}">${escapeHtml(item.currentStatus || item.stage || uiStateLabel(item.uiState))}</span></td><td>${formatPercent(getClientSpoolProgress(item))}</td><td>${formatNumber(item.kilos || 0, 2)} kg</td></tr>`;
  }).join('') : '';
  const drawingDetailNumber = isSpool ? String(spool?.iso || spool?.drawing || spool?.projectRef || title || '').trim() : '';
  const drawingDetailNumberJson = clientSafeJsonForScript(drawingDetailNumber);
  const isoScheduleRows = isSpool ? buildClientIsoExecutiveSchedule(project, spool) : [];
  const isoScheduleModel = isoScheduleRows.map((row) => ({
    type: row.type,
    key: row.key,
    label: row.label,
    progress: Number(row.progress || 0),
    duration: Math.max(1, Number(row.duration || 1)),
    start: clientDateInputValue(row.start),
    finish: clientDateInputValue(row.finish),
  }));
  const isoPanelKey = isSpool ? getClientSpoolPanelKey(spool) : '';
  const canEditIsoDates = isSpool && canManageClientBspPanel(project);
  const showIsoScheduleSection = isSpool && canEditIsoDates;
  const projectOverride = getClientBspOverride(project) || {};
  const isoEditorPayload = isSpool ? {
    canEdit: canEditIsoDates,
    project: {
      projectRowId: getClientBspOverrideProjectRowId(project),
      projectNumber: getClientBspOverrideProjectNumber(project),
      projectDisplay: project.projectDisplay || project.projectNumber || '',
      clientName: getProjectClientLabel(project),
      vessel: getProjectVesselLabel(project),
      pm: project.pm || '',
    },
    iso: {
      key: isoPanelKey,
      iso: spool?.iso || '',
      drawing: spool?.drawing || '',
      description: spool?.description || '',
    },
    existingCustomFields: projectOverride.customFields || {},
    initialOverride: isoOverride,
    scheduleRows: isoScheduleModel,
  } : null;

  return `<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)} • Painel Individual STEP</title>
<style>
  :root { color-scheme: light; font-family: Inter, Segoe UI, Arial, sans-serif; --blue:#083a63; --teal:#008a75; --soft:#eef8fb; --line:#c9dde8; --text:#12324a; }
  * { box-sizing: border-box; }
  body { margin: 0; background: linear-gradient(135deg, #f8fcff 0%, #eef8fb 100%); color: var(--text); }
  .page { max-width: 1280px; margin: 0 auto; padding: 28px; }
  .hero { display: grid; grid-template-columns: 1fr auto; gap: 18px; align-items: center; background: #fff; border: 1px solid var(--line); border-radius: 24px; padding: 24px; box-shadow: 0 18px 44px rgba(8,58,99,.11); }
  .kicker { margin: 0 0 8px; color: var(--teal); font-weight: 900; letter-spacing: .14em; text-transform: uppercase; font-size: 12px; }
  h1 { margin: 0; font-size: clamp(26px, 4vw, 46px); color: var(--blue); line-height: 1.05; }
  .subtitle { margin: 10px 0 0; color: #49677b; font-weight: 700; }
  .logo { height: 54px; max-width: 130px; object-fit: contain; }
  .toolbar { display:flex; gap:10px; justify-content:flex-end; margin-top:14px; }
  button { border:0; border-radius:999px; padding:11px 16px; font-weight:900; cursor:pointer; background:var(--blue); color:#fff; }
  button.secondary { background:#e8f3f7; color:var(--blue); }
  .iso-schedule-head { display:flex; justify-content:space-between; gap:14px; align-items:flex-start; flex-wrap:wrap; }
  .iso-schedule-head p { margin:4px 0 0; color:#49677b; font-weight:700; }
  .iso-schedule-actions { display:flex; gap:8px; flex-wrap:wrap; align-items:center; justify-content:flex-end; }
  .iso-schedule-editor { margin:14px 0; padding:14px; border:1px dashed var(--line); border-radius:18px; background:#f8fcff; }
  .iso-schedule-editor.hidden { display:none; }
  .iso-schedule-editor-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; }
  .iso-schedule-editor label { display:grid; gap:6px; color:#49677b; font-weight:900; font-size:12px; text-transform:uppercase; letter-spacing:.08em; }
  .iso-schedule-editor input { width:100%; border:1px solid var(--line); border-radius:12px; padding:11px 12px; font:inherit; color:var(--blue); background:#fff; }
  .iso-schedule-editor-actions { display:flex; gap:8px; margin-top:12px; flex-wrap:wrap; }
  .iso-schedule-status { margin-top:8px; color:#49677b; font-weight:800; }
  .iso-schedule-status.error { color:#991b1b; }
  .iso-manual-grid { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:12px; margin:12px 0; }
  .iso-manual-card { border:1px solid var(--line); border-radius:16px; padding:13px; background:#f4fbfe; }
  .iso-manual-card span { display:block; color:#638094; font-size:10px; font-weight:900; letter-spacing:.10em; text-transform:uppercase; margin-bottom:5px; }
  .iso-manual-card strong { color:var(--blue); font-size:17px; }
  .grid { display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:14px; margin-top:18px; }
  .card { background:#fff; border:1px solid var(--line); border-radius:20px; padding:16px; box-shadow:0 12px 30px rgba(8,58,99,.08); min-width:0; }
  .card span { display:block; color:#638094; font-size:11px; font-weight:900; letter-spacing:.10em; text-transform:uppercase; margin-bottom:6px; }
  .card strong { display:block; font-size:20px; color:var(--blue); overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .card.full { grid-column: 1 / -1; }
  .progress-wrap { display:grid; grid-template-columns: 180px 1fr; gap:18px; align-items:center; margin-top:18px; }
  .gauge { width:160px; aspect-ratio:1; border-radius:50%; background:conic-gradient(var(--teal) ${clampClientPercent(progress)}%, #dce7ee 0); display:grid; place-items:center; position:relative; margin:auto; }
  .gauge:after { content:""; position:absolute; inset:18px; border-radius:50%; background:#fff; }
  .gauge strong { position:relative; z-index:1; color:var(--blue); font-size:28px; }
  .section { margin-top:18px; background:#fff; border:1px solid var(--line); border-radius:22px; padding:18px; box-shadow:0 12px 30px rgba(8,58,99,.08); }
  .section h2 { margin:0 0 14px; color:var(--blue); font-size:20px; }
  table { width:100%; border-collapse:separate; border-spacing:0; overflow:hidden; border-radius:14px; border:1px solid var(--line); }
  th, td { padding:12px 14px; text-align:left; border-bottom:1px solid #e2edf2; vertical-align:top; }
  th { background:var(--soft); color:var(--blue); font-size:12px; letter-spacing:.08em; text-transform:uppercase; }
  tr:last-child td { border-bottom:0; }
  .chip { display:inline-flex; border-radius:999px; padding:6px 10px; font-weight:900; font-size:12px; background:#e8f3f7; color:var(--blue); }
  .chip-completed { background:#dcfce7; color:#166534; }
  .chip-in-progress { background:#fef3c7; color:#92400e; }
  .chip-not-started { background:#e2e8f0; color:#475569; }
  .client-spool-chip, .client-spool-progress { display:inline-flex; align-items:center; border-radius:999px; padding:6px 10px; font-weight:900; font-size:12px; background:#e8f3f7; color:var(--blue); }
  .client-spool-chip--completed, .client-spool-progress--completed { background:#dcfce7; color:#166534; }
  .client-spool-chip--in-progress, .client-spool-progress--in-progress { background:#fef3c7; color:#92400e; }
  .client-spool-chip--not-started, .client-spool-progress--not-started { background:#e2e8f0; color:#475569; }
  .client-schedule-child { display:inline-block; padding-left:18px; color:#49677b; }
  .client-planned-date { display:block; margin-top:3px; color:#638094; font-size:11px; font-weight:800; }
  .client-schedule-row--group td { background:#f8fcff; font-weight:900; }
  .note { color:#49677b; line-height:1.5; white-space:pre-wrap; }
  .drawing-detail-status { color:#49677b; font-weight:800; }
  .drawing-detail-status.error { color:#991b1b; }
  .drawing-detail-grid { display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:12px; margin:12px 0 16px; }
  .drawing-detail-card { background:#f4fbfe; border:1px solid var(--line); border-radius:16px; padding:13px; min-width:0; }
  .drawing-detail-card span { display:block; color:#638094; font-size:10px; font-weight:900; letter-spacing:.10em; text-transform:uppercase; margin-bottom:5px; }
  .drawing-detail-card strong { display:block; color:var(--blue); font-size:16px; line-height:1.2; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  .drawing-revisions { display:grid; gap:12px; margin-top:14px; }
  .drawing-revision { border:1px solid var(--line); border-radius:16px; overflow:hidden; background:#fff; }
  .drawing-revision h3 { margin:0; padding:12px 14px; background:#eef8fb; color:var(--blue); font-size:15px; }
  .drawing-revision table { border:0; border-radius:0; }
  .drawing-revision td:first-child { width:36%; color:#49677b; font-weight:900; }
  @media (max-width: 850px) { .hero, .progress-wrap { grid-template-columns:1fr; } .grid, .drawing-detail-grid { grid-template-columns:1fr 1fr; } .page { padding:14px; } }
  @media print { .toolbar { display:none; } body { background:#fff; } .page { max-width:none; } }
</style>
</head>
<body>
  <main class="page">
    <section class="hero">
      <div>
        <p class="kicker">${isSpool ? 'Painel Individual da Tag / ISO' : 'Painel Individual da Obra'}</p>
        <h1>${escapeHtml(title)}</h1>
        <p class="subtitle">${escapeHtml(subtitle)}</p>
      </div>
      <div>
        <img class="logo" src="${escapeHtml(new URL('./assets/step-logo.png', window.location.href).href)}" alt="STEP">
        <div class="toolbar"><button onclick="window.print()">Imprimir / PDF</button><button class="secondary" onclick="window.close()">Fechar</button></div>
      </div>
    </section>

    <section class="progress-wrap">
      <div class="gauge"><strong>${formatPercent(progress)}</strong></div>
      <div class="grid">
        <article class="card"><span>BSP</span><strong>${escapeHtml(getClientProjectDisplayCode(project))}</strong></article>
        <article class="card"><span>PO</span><strong>${escapeHtml(po)}</strong></article>
        <article class="card"><span>Cliente</span><strong>${escapeHtml(getProjectClientLabel(project))}</strong></article>
        <article class="card"><span>Unidade</span><strong>${escapeHtml(getProjectVesselLabel(project))}</strong></article>
        <article class="card"><span>Status</span><strong>${escapeHtml(currentStatus || '—')}</strong></article>
        <article class="card"><span>Etapa</span><strong>${escapeHtml(currentStage || '—')}</strong></article>
        <article class="card"><span>Peso</span><strong>${formatNumber(kilos, 2)} kg</strong></article>
        <article class="card"><span>M²</span><strong>${formatNumber(m2, 3)}</strong></article>
        <article class="card"><span>Descrição</span><strong>${escapeHtml(description)}</strong></article>
        <article class="card"><span>Término</span><strong id="client-standalone-finish-display">${escapeHtml(clientFormatDateValue(finishedDate) || '—')}</strong></article>
      </div>
    </section>

    ${showIsoScheduleSection ? `<section class="section" id="iso-schedule-section">
      <div class="iso-schedule-head">
        <div>
          <h2>Cronograma da ISO</h2>
          <p>Datas específicas desta ISO. Quando o PM informar início e término manual, o cronograma da ISO passa a usar essas datas.</p>
        </div>
        ${canEditIsoDates ? `<div class="iso-schedule-actions"><button type="button" id="iso-schedule-edit-button">Editar início/fim da ISO</button></div>` : ''}
      </div>
      <div class="iso-manual-grid">
        <article class="iso-manual-card"><span>Início manual da ISO</span><strong id="iso-manual-start-display">${escapeHtml(clientFormatDateValue(isoOverride.startDate) || '—')}</strong></article>
        <article class="iso-manual-card"><span>Término manual da ISO</span><strong id="iso-manual-finish-display">${escapeHtml(clientFormatDateValue(isoOverride.finishDate) || '—')}</strong></article>
      </div>
      ${canEditIsoDates ? `<form id="iso-schedule-editor-form" class="iso-schedule-editor hidden">
        <div class="iso-schedule-editor-grid">
          <label><span>Data de início da ISO</span><input type="date" id="iso-schedule-start-input" value="${escapeHtml(clientDateInputValue(isoOverride.startDate))}" /></label>
          <label><span>Data de término da ISO</span><input type="date" id="iso-schedule-finish-input" value="${escapeHtml(clientDateInputValue(isoOverride.finishDate))}" /></label>
        </div>
        <div class="iso-schedule-editor-actions">
          <button type="submit">Salvar datas da ISO</button>
          <button class="secondary" type="button" id="iso-schedule-cancel-button">Cancelar</button>
        </div>
        <div id="iso-schedule-save-status" class="iso-schedule-status"></div>
      </form>` : ''}
      <div id="iso-schedule-table-content">${renderClientExecutiveScheduleRows(isoScheduleRows, 'Cronograma não disponível para esta ISO.')}</div>
    </section>
    <script>
      (function () {
        var payload = ${clientSafeJsonForScript(isoEditorPayload)};
        if (!payload) return;
        var editButton = document.getElementById('iso-schedule-edit-button');
        var form = document.getElementById('iso-schedule-editor-form');
        var cancelButton = document.getElementById('iso-schedule-cancel-button');
        var startInput = document.getElementById('iso-schedule-start-input');
        var finishInput = document.getElementById('iso-schedule-finish-input');
        var statusEl = document.getElementById('iso-schedule-save-status');
        var startDisplay = document.getElementById('iso-manual-start-display');
        var finishDisplay = document.getElementById('iso-manual-finish-display');
        var tableTarget = document.getElementById('iso-schedule-table-content');
        var finishHeaderDisplay = document.getElementById('client-standalone-finish-display');
        function esc(value) {
          return String(value == null || value === '' ? '—' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
        }
        function parseDate(value) {
          var match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
          if (!match) return null;
          return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
        }
        function ymd(date) {
          if (!(date instanceof Date) || isNaN(date)) return '';
          return date.getUTCFullYear() + '-' + String(date.getUTCMonth() + 1).padStart(2, '0') + '-' + String(date.getUTCDate()).padStart(2, '0');
        }
        function br(value) {
          var d = parseDate(value);
          if (!d) return '—';
          return String(d.getUTCDate()).padStart(2, '0') + '/' + String(d.getUTCMonth() + 1).padStart(2, '0') + '/' + d.getUTCFullYear();
        }
        function isBiz(date) {
          var day = date.getUTCDay();
          return day !== 0 && day !== 6;
        }
        function addBiz(date, days) {
          var next = new Date(date.getTime());
          var remaining = Math.max(0, Number(days) || 0);
          while (remaining > 0) {
            next.setUTCDate(next.getUTCDate() + 1);
            if (isBiz(next)) remaining -= 1;
          }
          return next;
        }
        function bizDays(a, b) {
          var start = parseDate(a);
          var finish = parseDate(b);
          if (!start || !finish) return 0;
          if (start > finish) { var tmp = start; start = finish; finish = tmp; }
          var count = 0;
          var cursor = new Date(start.getTime());
          while (cursor <= finish) {
            if (isBiz(cursor)) count += 1;
            cursor.setUTCDate(cursor.getUTCDate() + 1);
          }
          return Math.max(1, count);
        }
        function scaleDurations(items, total) {
          var safeItems = (items || []).map(function (item) { return Object.assign({}, item, { duration: Math.max(1, Number(item.duration || 1)) }); });
          var safeTotal = Math.max(safeItems.length || 1, Number(total) || 1);
          var baseTotal = safeItems.reduce(function (sum, item) { return sum + item.duration; }, 0) || safeItems.length || 1;
          var scaled = safeItems.map(function (item) {
            return Object.assign({}, item, { duration: Math.max(1, Math.round((item.duration / baseTotal) * safeTotal)) });
          });
          var diff = safeTotal - scaled.reduce(function (sum, item) { return sum + item.duration; }, 0);
          var i = 0;
          while (diff !== 0 && scaled.length) {
            var idx = i % scaled.length;
            if (diff > 0) { scaled[idx].duration += 1; diff -= 1; }
            else if (scaled[idx].duration > 1) { scaled[idx].duration -= 1; diff += 1; }
            i += 1;
            if (i > 500) break;
          }
          return scaled;
        }
        function percent(value) {
          var n = Number(value || 0);
          if (!isFinite(n)) return '—';
          return n.toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%';
        }
        function visualState(progress) {
          var n = Number(progress || 0);
          if (n >= 99.9) return 'completed';
          if (n > 0) return 'in-progress';
          return 'not-started';
        }
        function buildRowsForPeriod(startValue, finishValue) {
          var start = parseDate(startValue);
          var finish = parseDate(finishValue);
          var base = Array.isArray(payload.scheduleRows) ? payload.scheduleRows : [];
          if (!start || !finish || !base.length) return base;
          if (start > finish) { var tmp = start; start = finish; finish = tmp; }
          var total = bizDays(ymd(start), ymd(finish));
          var groups = base.filter(function (row) { return row.type === 'group'; });
          var scaledGroups = scaleDurations(groups, total);
          var rows = [];
          var cursor = new Date(start.getTime());
          scaledGroups.forEach(function (group, groupIndex) {
            var originalGroupIndex = base.findIndex(function (row) { return row.type === 'group' && row.key === group.key; });
            var nextGroupIndex = base.findIndex(function (row, idx) { return idx > originalGroupIndex && row.type === 'group'; });
            if (nextGroupIndex < 0) nextGroupIndex = base.length;
            var children = base.slice(originalGroupIndex + 1, nextGroupIndex).filter(function (row) { return row.type === 'child'; });
            var groupStart = new Date(cursor.getTime());
            var groupFinish = addBiz(groupStart, Math.max(0, group.duration - 1));
            if (groupIndex === scaledGroups.length - 1 || groupFinish > finish) groupFinish = new Date(finish.getTime());
            rows.push(Object.assign({}, group, { start: ymd(groupStart), finish: ymd(groupFinish) }));
            var childCursor = new Date(groupStart.getTime());
            scaleDurations(children, group.duration).forEach(function (child, childIndex, arr) {
              var childStart = new Date(childCursor.getTime());
              var childFinish = addBiz(childStart, Math.max(0, child.duration - 1));
              if (childIndex === arr.length - 1 || childFinish > groupFinish) childFinish = new Date(groupFinish.getTime());
              rows.push(Object.assign({}, child, { start: ymd(childStart), finish: ymd(childFinish) }));
              childCursor = addBiz(childFinish, 1);
            });
            cursor = addBiz(groupFinish, 1);
            if (cursor > finish) cursor = new Date(finish.getTime());
          });
          return rows;
        }
        function renderRows(rows) {
          if (!tableTarget) return;
          if (!rows || !rows.length) {
            tableTarget.innerHTML = '<div class="client-empty-state">Cronograma não disponível para esta ISO.</div>';
            return;
          }
          tableTarget.innerHTML = '<div class="client-table-wrap client-table-wrap--compact client-exec-schedule-table"><table class="client-bsp-table client-bsp-table--schedule"><thead><tr><th>Etapa</th><th>%</th><th>Prazo médio</th><th>Início</th><th>Término</th><th>Status</th></tr></thead><tbody>' + rows.map(function (row) {
            var state = visualState(row.progress);
            var label = row.type === 'group' ? '<strong>' + esc(row.label) + '</strong>' : '<span class="client-schedule-child">' + esc(row.label) + '</span>';
            var status = state === 'completed' ? 'Concluído' : state === 'in-progress' ? 'Em andamento' : 'Não iniciado';
            return '<tr class="client-schedule-row client-schedule-row--' + state + ' client-schedule-row--' + esc(row.type) + '"><td>' + label + '</td><td><span class="client-spool-progress client-spool-progress--' + state + '">' + percent(row.progress) + '</span></td><td>' + esc(row.duration || 1) + 'd</td><td>' + br(row.start) + '</td><td>' + br(row.finish) + '</td><td><span class="client-spool-chip client-spool-chip--' + state + '">' + status + '</span></td></tr>';
          }).join('') + '</tbody></table></div>';
        }
        function updateManualDisplays(startValue, finishValue) {
          if (startDisplay) startDisplay.textContent = br(startValue);
          if (finishDisplay) finishDisplay.textContent = br(finishValue);
          if (finishHeaderDisplay && finishValue) finishHeaderDisplay.textContent = br(finishValue);
          if (startValue && finishValue) renderRows(buildRowsForPeriod(startValue, finishValue));
        }
        if (editButton && form) {
          editButton.addEventListener('click', function () { form.classList.remove('hidden'); });
        }
        if (cancelButton && form) {
          cancelButton.addEventListener('click', function () { form.classList.add('hidden'); });
        }
        if (payload.initialOverride) updateManualDisplays(payload.initialOverride.startDate || '', payload.initialOverride.finishDate || '');
        if (form && payload.canEdit) {
          form.addEventListener('submit', function (event) {
            event.preventDefault();
            var startValue = startInput ? startInput.value : '';
            var finishValue = finishInput ? finishInput.value : '';
            if (startValue && finishValue && parseDate(startValue) > parseDate(finishValue)) {
              if (statusEl) { statusEl.classList.add('error'); statusEl.textContent = 'A data de início não pode ser maior que a data de término.'; }
              return;
            }
            if (statusEl) { statusEl.classList.remove('error'); statusEl.textContent = 'Salvando datas da ISO...'; }
            var customFields = Object.assign({}, payload.existingCustomFields || {});
            var map = customFields.__isoDateOverrides && typeof customFields.__isoDateOverrides === 'object' ? Object.assign({}, customFields.__isoDateOverrides) : {};
            map[payload.iso.key] = Object.assign({}, map[payload.iso.key] || {}, {
              startDate: startValue,
              finishDate: finishValue,
              iso: payload.iso.iso || '',
              drawing: payload.iso.drawing || '',
              description: payload.iso.description || '',
              updatedAt: new Date().toISOString()
            });
            customFields.__isoDateOverrides = map;
            fetch('/api/client-bsp-overrides', {
              method: 'POST',
              credentials: 'same-origin',
              cache: 'no-store',
              headers: { 'content-type': 'application/json' },
              body: JSON.stringify(Object.assign({}, payload.project, { mergeCustomFields: true, customFields: customFields }))
            })
              .then(function (response) { return response.json().catch(function () { return null; }).then(function (data) { return { response: response, data: data }; }); })
              .then(function (result) {
                if (!result.response.ok || !result.data || !result.data.ok) throw new Error((result.data && result.data.error) || 'Falha ao salvar as datas da ISO.');
                payload.existingCustomFields = result.data.override && result.data.override.customFields ? result.data.override.customFields : customFields;
                updateManualDisplays(startValue, finishValue);
                if (statusEl) { statusEl.classList.remove('error'); statusEl.textContent = 'Datas da ISO salvas com sucesso.'; }
                if (form) form.classList.add('hidden');
              })
              .catch(function (error) {
                if (statusEl) { statusEl.classList.add('error'); statusEl.textContent = error && error.message ? error.message : 'Falha ao salvar as datas da ISO.'; }
              });
          });
        }
      })();
    </script>` : ''}

    <section class="section">
      <h2>Evolução por etapa</h2>
      <table><thead><tr><th>Etapa</th><th>Avanço</th></tr></thead><tbody>
        ${stageRows.map((row) => `<tr><td>${escapeHtml(row.label)}</td><td><span class="chip">${escapeHtml(formatClientStandalonePanelPercent(row.value))}</span></td></tr>`).join('')}
      </tbody></table>
    </section>

    <section class="section">
      <h2>Observações</h2>
      <div class="note">${escapeHtml(formatClientStandalonePanelValue(observation))}</div>
    </section>

    ${isSpool ? `<section class="section" id="drawing-control-section" data-drawing-number="${escapeHtml(drawingDetailNumber)}">
      <h2>Detalhamento de Drawing</h2>
      <div id="drawing-control-content" class="drawing-detail-status">Carregando datas e revisões do Drawing Documentation Control...</div>
    </section>
    <script>
      (function () {
        var drawingNumber = ${drawingDetailNumberJson};
        var content = document.getElementById('drawing-control-content');
        if (!content || !drawingNumber) return;
        function esc(value) {
          return String(value == null || value === '' ? '—' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
        }
        function row(label, value) {
          return '<tr><td>' + esc(label) + '</td><td>' + esc(value || '—') + '</td></tr>';
        }
        function card(label, value) {
          return '<article class="drawing-detail-card"><span>' + esc(label) + '</span><strong>' + esc(value || '—') + '</strong></article>';
        }
        function revisionBlock(title, rows) {
          return '<article class="drawing-revision"><h3>' + esc(title) + '</h3><table><tbody>' + rows.join('') + '</tbody></table></article>';
        }
        function render(detail) {
          var revA = detail.revA || {};
          var html = '';
          html += '<div class="drawing-detail-grid">';
          html += card('Drawing Number', detail.drawingNumber);
          html += card('Status atual', detail.currentDrawingStatus);
          html += card('Início desenho', detail.drawingStartDate);
          html += card('Final desenho', detail.drawingFinishDate || detail.approvalDate);
          html += card('Unidade', detail.unit);
          html += card('Quantidade/Peso', detail.quantity == null ? '' : detail.quantity);
          html += card('BSP', detail.bspKey || detail.projectNumber);
          html += card('Título', detail.title);
          html += '</div>';
          html += '<div class="drawing-revisions">';
          html += revisionBlock('Revisão A', [
            row('Drawings start date', revA.startDate),
            row('Reviewer approval', revA.reviewerApproval),
            row('Approver approval', revA.approverApproval),
            row('Internally approved & sent to PM', revA.internallySentToPmDate),
            row('PM approval', revA.pmApproval),
            row('Approval Date', revA.approvalDate),
            row('Status da aprovação', revA.approved ? 'Approved / Finalizado' : 'Em andamento ou pendente')
          ]);
          (detail.revisions || []).forEach(function (revision) {
            html += revisionBlock('Revisão ' + revision.revision, [
              row('Last client comments received via PM', revision.clientCommentsDate),
              row('Last revision Update start', revision.updateStartDate),
              row('Reviewer Update approval', revision.reviewerApproval),
              row('Approver Update approval', revision.approverApproval),
              row('Internally approved & sent to PM', revision.internallySentToPmDate),
              row('PM approval', revision.pmApproval)
            ]);
          });
          html += '</div>';
          content.classList.remove('error');
          content.innerHTML = html;
        }
        fetch('/api/drawing-detail?drawingNumber=' + encodeURIComponent(drawingNumber) + '&force=1', { credentials: 'same-origin', cache: 'no-store' })
          .then(function (response) {
            return response.text().then(function (text) {
              var data = null;
              var body = String(text || '').trim();
              if (!body) {
                throw new Error('API de Drawing retornou resposta vazia. Status HTTP: ' + response.status + '. Faça novo deploy com Clear cache e confirme se a função drawing-detail foi enviada.');
              }
              try {
                data = JSON.parse(body);
              } catch (parseError) {
                throw new Error('API de Drawing retornou resposta não JSON. Status HTTP: ' + response.status + '. Retorno: ' + body.slice(0, 220));
              }
              if (!response.ok || !data.ok) {
                var extra = '';
                if (data.totalIndexed != null) extra += ' Itens indexados: ' + data.totalIndexed + '.';
                if (Array.isArray(data.closest) && data.closest.length) extra += ' Próximos encontrados: ' + data.closest.map(function (x) { return x.drawingNumber; }).join(', ') + '.';
                throw new Error((data.error || 'Detalhamento não encontrado.') + extra);
              }
              return data;
            });
          })
          .then(function (data) { render(data.detail || {}); })
          .catch(function (error) {
            content.classList.add('error');
            content.textContent = 'Não foi possível carregar o detalhamento de Drawing para ' + drawingNumber + ': ' + (error && error.message ? error.message : error);
          });
      })();
    </script>` : ''}

    ${!isSpool ? `<section class="section"><h2>Tags / ISOs da obra</h2><table><thead><tr><th>Tag/ISO</th><th>Descrição</th><th>Status</th><th>%</th><th>Peso</th></tr></thead><tbody>${siblingRows || '<tr><td colspan="5">Nenhuma tag detalhada encontrada.</td></tr>'}</tbody></table></section>` : ''}
  </main>
</body>
</html>`;
}

function openClientSpoolIndividualPanel(project, spool = null) {
  if (!project) return;
  const newWindow = window.open('', '_blank');
  if (!newWindow) {
    window.alert('O navegador bloqueou a nova aba. Libere pop-ups para abrir o painel individual.');
    return;
  }
  try { newWindow.opener = null; } catch (_) {}
  const html = buildClientStandalonePanelHtml(project, spool);
  newWindow.document.open();
  newWindow.document.write(html);
  newWindow.document.close();
}

function renderClientSpoolRows(spools, limit = 120, project = null) {
  const items = (Array.isArray(spools) ? [...spools] : []).sort(compareClientSpoolsByPriority).slice(0, limit);
  if (!items.length) return '<tr><td colspan="7" class="loading-cell">Nenhuma tag detalhada encontrada para esta BSP.</td></tr>';
  const projectId = project?.rowId || state.clientBspOverrides.activeExecutiveProjectId || state.clientPortal.selectedProjectId || '';
  return items.map((spool, index) => {
    const state = getClientSpoolVisualState(spool);
    const statusText = spool?.currentStatus || spool?.stage || uiStateLabel(spool?.uiState);
    const observationText = spool?.observations ? String(spool.observations).trim() : '—';
    const key = getClientSpoolPanelKey(spool, index);
    return `<tr class="client-spool-row client-spool-row--${state} client-spool-row--clickable" data-client-spool-panel="${escapeHtml(key)}" data-client-spool-project-id="${escapeHtml(projectId)}" title="Clique para abrir o painel individual e o detalhamento de Drawing em uma nova aba"><td><strong>${escapeHtml(spool.iso || '—')}</strong></td><td>${escapeHtml(spool.description || '—')}</td><td>${escapeHtml(observationText)}</td><td><span class="client-spool-chip client-spool-chip--${state}">${escapeHtml(statusText)}</span></td><td>${escapeHtml(spool.currentSector || spool.operationalSector || '—')}</td><td><span class="client-spool-progress client-spool-progress--${state}">${formatPercent(spool.overallProgress)}</span></td><td>${formatNumber(spool.kilos, 2)} kg</td></tr>`;
  }).join('');
}

function renderClientProjectDetail(project) {
  const detail = document.getElementById('client-project-detail');
  if (!detail) return;
  if (!project) {
    detail.classList.add('hidden');
    detail.innerHTML = '';
    return;
  }
  detail.classList.remove('hidden');
  const status = getProjectStatusPresentation(project);
  const stageValues = project.stageValues || {};
  const spools = Array.isArray(project.spools) ? project.spools : [];
  detail.innerHTML = `
    <div class="client-detail-head">
      <div><p class="client-kicker">Detalhamento da BSP</p><h3>${escapeHtml(getClientProjectDisplayCode(project))}</h3><p>${escapeHtml(getProjectVesselLabel(project))} • ${escapeHtml(getProjectClientLabel(project))}</p></div>
      <button class="primary-button" type="button" data-client-open-analytics="${escapeHtml(project.rowId)}">Abrir visão executiva</button>
    </div>
    <div class="client-summary-grid client-summary-grid--detail">
      <article><span>Tags</span><strong>${formatNumber(getProjectItemCount(project))}</strong></article>
      <article><span>Peso programado</span><strong>${formatNumber(project.kilos, 0)} kg</strong></article>
      <article><span>Peso soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></article>
      <article><span>Área operacional</span><strong>${formatNumber(project.m2Painting, 3)} m²</strong></article>
      <article><span>Status</span><strong>${escapeHtml(status.text)}</strong></article>
      <article><span>Progresso geral</span><strong>${formatPercent(project.overallProgress)}</strong></article>
    </div>
    ${renderClientTratativaNotice(project)}
    ${renderClientOnHoldNotice(project)}

    <div class="client-stage-strip">
      ${getClientDetailStageKeys(project).map((key) => {
        const label = (state.meta?.stageOrder || []).find((stage) => stage.key === key)?.label || key;
        const value = stageValues[key];
        return renderClientStageStripCard(label, value);
      }).join('')}
    </div>
    <div class="client-table-wrap client-table-wrap--compact">
      <table class="client-bsp-table"><thead><tr><th>Tag/ISO</th><th>Descrição</th><th>Observação</th><th>Status</th><th>Etapa</th><th>%</th><th>Peso</th></tr></thead><tbody>
        ${renderClientSpoolRows(spools, 80, project)}
      </tbody></table>
    </div>
  `;
}


function clampClientPercent(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.min(100, num));
}

function clientFormatDateValue(value) {
  const parsed = parseDateObject(value);
  const date = parsed || (value instanceof Date && !Number.isNaN(value.getTime()) ? value : null);
  if (date) return date.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
  return value ? String(value) : '';
}

function clientPercentValue(value) {
  if (value == null || value === '' || value === 'N/A') return 0;
  if (typeof value === 'number') {
    if (value >= 0 && value <= 1) return clampClientPercent(value * 100);
    return clampClientPercent(value);
  }
  const raw = String(value || '').trim();
  if (!raw) return 0;
  const cleaned = raw.replace('%', '').replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const parsed = Number(cleaned);
  if (!Number.isFinite(parsed)) return 0;
  if (parsed >= 0 && parsed <= 1 && !raw.includes('%')) return clampClientPercent(parsed * 100);
  return clampClientPercent(parsed);
}

function getClientStageValue(project, keys) {
  const stageValues = project?.stageValues || {};
  for (const key of keys) {
    const value = stageValues[key];
    if (value != null && value !== '' && value !== 'N/A') return clientPercentValue(value);
  }
  return 0;
}

function isClientSpoolMaterial(project) {
  const explicitType = normalizeText([project?.projectType, project?.type, project?.project_type].filter(Boolean).join(' '));
  if (explicitType) {
    if (explicitType.includes('spool')) return true;
    if (explicitType.includes('support') || explicitType.includes('suporte') || explicitType === 'sup' || explicitType.includes('structure') || explicitType.includes('estrutura') || explicitType.includes('frame')) return false;
  }

  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const evidence = normalizeText([
    project?.summaryDrawing,
    project?.drawing,
    ...spools.flatMap((spool) => [spool?.iso, spool?.drawing, spool?.description]),
  ].filter(Boolean).join(' '));

  if (/\bspl\b/.test(evidence) || evidence.includes('spool')) return true;
  if (/\bsup\b/.test(evidence) || evidence.includes('support') || evidence.includes('suporte') || evidence.includes('structure') || evidence.includes('estrutura') || evidence.includes('frame')) return false;
  return false;
}


function clientProjectHasNoHydroObservation(project) {
  const chunks = [];
  const collect = (value) => {
    if (value == null) return;
    const text = String(value).trim();
    if (text) chunks.push(text);
  };
  collect(project?.observations);
  collect(project?.observation);
  collect(project?.OBSERVATIONS);
  collect(project?.comments);
  collect(project?.notes);
  if (project?.stageValues && typeof project.stageValues === 'object') {
    collect(project.stageValues.OBSERVATIONS);
    collect(project.stageValues.Observations);
    collect(project.stageValues.observations);
    collect(project.stageValues['Observation']);
    collect(project.stageValues['OBSERVATION']);
    collect(project.stageValues['Comments']);
  }
  for (const milestone of Array.isArray(project?.milestones) ? project.milestones : []) {
    collect(milestone?.key);
    collect(milestone?.label);
    collect(milestone?.value);
  }
  for (const spool of Array.isArray(project?.spools) ? project.spools : []) {
    collect(spool?.observations);
    collect(spool?.observation);
    collect(spool?.OBSERVATIONS);
    collect(spool?.comments);
    collect(spool?.notes);
    collect(spool?.description);
    if (spool?.stageValues && typeof spool.stageValues === 'object') {
      collect(spool.stageValues.OBSERVATIONS);
      collect(spool.stageValues.Observations);
      collect(spool.stageValues.observations);
      collect(spool.stageValues['Observation']);
      collect(spool.stageValues['OBSERVATION']);
      collect(spool.stageValues['Comments']);
    }
    for (const milestone of Array.isArray(spool?.milestones) ? spool.milestones : []) {
      collect(milestone?.key);
      collect(milestone?.label);
      collect(milestone?.value);
    }
  }
  const text = normalizeText(chunks.join(' | '));
  if (!text) return false;
  return /solda\s+(de|em)\s+campo/.test(text)
    || /soldagem\s+(de|em)\s+campo/.test(text)
    || /field\s*weld(ing)?/.test(text)
    || /\bf\.?w\.?\b/.test(text)
    || /sem\s+(th|hydro|teste\s+hidrostatico)/.test(text)
    || /nao\s+(requer|aplica|necessita)\s+(th|hydro|teste\s+hidrostatico)/.test(text)
    || /th\s+(nao\s+aplicavel|n\/a|na|dispensado)/.test(text);
}

function shouldClientShowHydro(project) {
  return isClientSpoolMaterial(project) && !clientProjectHasNoHydroObservation(project);
}

function normalizeClientTrackingKey(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, ' ').trim();
}

function clientTrackingKeyMatches(sourceKey, wantedKeys) {
  const source = normalizeClientTrackingKey(sourceKey);
  if (!source) return false;
  return wantedKeys.some((key) => {
    const wanted = normalizeClientTrackingKey(key);
    return wanted && (source === wanted || source.includes(wanted) || wanted.includes(source));
  });
}

function clientReadDateCandidatesFromSource(source, wantedKeys, candidates) {
  if (!source || typeof source !== 'object') return;
  for (const key of wantedKeys) {
    const direct = source[key];
    const parsedDirect = parseClientSafeDateObject(direct);
    if (parsedDirect) candidates.push(parsedDirect);
  }
  for (const [key, value] of Object.entries(source)) {
    if (!clientTrackingKeyMatches(key, wantedKeys)) continue;
    const parsed = parseClientSafeDateObject(value);
    if (parsed) candidates.push(parsed);
  }
}

function clientReadDateCandidatesFromMilestones(milestones, wantedKeys, candidates) {
  for (const item of Array.isArray(milestones) ? milestones : []) {
    if (!item || typeof item !== 'object') continue;
    if (!clientTrackingKeyMatches([item.key, item.label, item.name].filter(Boolean).join(' '), wantedKeys)) continue;
    const parsed = parseClientSafeDateObject(item.value ?? item.date ?? item.finishDate ?? item.startDate);
    if (parsed) candidates.push(parsed);
  }
}

function getClientTrackingDate(project, wantedKeys, mode = 'last') {
  const candidates = [];
  clientReadDateCandidatesFromSource(project, wantedKeys, candidates);
  clientReadDateCandidatesFromSource(project?.stageValues, wantedKeys, candidates);
  clientReadDateCandidatesFromMilestones(project?.milestones, wantedKeys, candidates);
  for (const spool of Array.isArray(project?.spools) ? project.spools : []) {
    clientReadDateCandidatesFromSource(spool, wantedKeys, candidates);
    clientReadDateCandidatesFromSource(spool?.stageValues, wantedKeys, candidates);
    clientReadDateCandidatesFromMilestones(spool?.milestones, wantedKeys, candidates);
  }
  const unique = Array.from(new Map(candidates.filter(Boolean).map((date) => [date.getTime(), date])).values()).sort((a, b) => a - b);
  if (!unique.length) return null;
  return mode === 'first' ? unique[0] : unique[unique.length - 1];
}

function getClientTrackingDates(project, options = {}) {
  const hydro = shouldClientShowHydro(project);
  const trackingDates = {
    drawingsStart: getClientTrackingDate(project, ['Drawing Start Date', 'Drawings Start Date', 'Engineering Start Date', 'Start Drawing Date'], 'first'),
    drawingsFinish: getClientTrackingDate(project, ['Drawing Execution Advance%', 'Drawing Finish Date', 'Drawings Finish Date', 'Engineering Finish Date', 'Drawing'], 'last'),
    procurementStart: getClientTrackingDate(project, ['Procurement Start Date', 'Procuremnt Start Date', 'Material Procurement Start Date', 'Materials Start Date'], 'first'),
    procurementFinish: getClientTrackingDate(project, ['Procuremnt Status %', 'Procurement Status %', 'Procurement Finish Date', 'Material Separation', 'Material Release to Fabrication', 'Materials Finish Date'], 'last'),
    fabricationStart: getClientTrackingDate(project, ['Fabrication Start Date', 'Fab. Início', 'FAB INICIO', 'Fab Inicio'], 'first'),
    boilermakerFinish: getClientTrackingDate(project, ['Boilermaker Finish Date', 'Caldeiraria', 'Calderaria Finish Date'], 'last'),
    weldingFinish: getClientTrackingDate(project, ['Welding Finish Date', 'Solda', 'Weld Finish Date', 'Full Welding Finish Date'], 'last'),
    inspectionFinish: getClientTrackingDate(project, ['Inspection Finish Date (QC)', 'Final Dimensional Inspection Finish Date', 'Inspection Finish'], 'last'),
    thFinish: hydro ? getClientTrackingDate(project, ['TH Finish Date', 'Hydro Finish Date', 'Hydro Testing Finish Date'], 'last') : null,
    coatingFinish: getClientTrackingDate(project, ['Coating Finish Date', 'Painting Finish Date', 'HDG / FBE DATE RETORNO (PAINT)'], 'last'),
    projectFinish: getClientTrackingDate(project, ['Project Finish Date', 'Data de Envio', 'Shipment Date', 'Delivery Date'], 'last'),
    sources: {},
  };
  if (options.ignoreOverrides) return trackingDates;

  const override = getClientBspOverride(project);
  if (!hasClientBspOverrideContent(override)) return trackingDates;

  const withOverrides = { ...trackingDates, sources: { ...trackingDates.sources } };
  const applyDate = (targetKey, overrideKey) => {
    const value = clientOverrideDateObject(override?.[overrideKey]);
    if (value) {
      withOverrides[targetKey] = value;
      withOverrides.sources[targetKey] = 'PM';
    }
  };
  applyDate('drawingsStart', 'drawingsStartOverride');
  applyDate('drawingsFinish', 'drawingsFinishOverride');
  applyDate('procurementStart', 'procurementStartOverride');
  applyDate('procurementFinish', 'procurementFinishOverride');
  applyDate('fabricationStart', 'fabricationStartOverride');
  applyDate('boilermakerFinish', 'boilermakerFinishOverride');
  applyDate('weldingFinish', 'weldingFinishOverride');
  applyDate('inspectionFinish', 'inspectionFinishOverride');
  if (hydro) applyDate('thFinish', 'thFinishOverride');
  applyDate('coatingFinish', 'coatingFinishOverride');
  applyDate('projectFinish', 'projectFinishOverride');
  return withOverrides;
}

function clientNextBusinessDay(dateValue) {
  const date = parseDateObject(dateValue);
  if (!date) return null;
  return addBusinessDaysUtc(date, 1) || date;
}

function clientLatestDate(...values) {
  const dates = values.map(parseDateObject).filter(Boolean).sort((a, b) => a - b);
  return dates.length ? dates[dates.length - 1] : null;
}


function clientEarliestDate(...values) {
  const dates = values.map(parseDateObject).filter(Boolean).sort((a, b) => a - b);
  return dates.length ? dates[0] : null;
}

function applyClientExecutiveScheduleReplan(project, rows) {
  if (!CLIENT_REPLANNING_UI_ENABLED) return rows;
  const replannedFinish = getClientSCurveReplannedFinishDate(project);
  if (!replannedFinish || !Array.isArray(rows) || !rows.length) return rows;

  const plannedFinish = parseDateObject(getClientSCurvePlannedFinishDate(project)) || parseDateObject(getClientAnalyticFinishDate(project));
  if (!plannedFinish || replannedFinish <= plannedFinish) return rows;

  const today = getCurrentBrazilDate();
  const nextRows = rows.map((row) => ({ ...row }));
  const isChild = (row) => row?.type === 'child';
  const isComplete = (row) => clampClientPercent(row?.progress) >= 99.9;
  const isReached = (row) => {
    const finish = parseDateObject(row?.finish);
    return finish && finish <= today;
  };

  let anchorIndex = nextRows.findIndex((row) => isChild(row) && !isComplete(row) && isReached(row));
  if (anchorIndex < 0) {
    // Se a data replanejada já existe mas o prazo de uma etapa ainda não venceu,
    // planejamos a partir da primeira etapa pendente para manter previsibilidade.
    anchorIndex = nextRows.findIndex((row) => isChild(row) && !isComplete(row));
  }
  if (anchorIndex < 0) return nextRows;

  const anchorRow = nextRows[anchorIndex];
  const anchorFinish = parseDateObject(anchorRow.finish) || plannedFinish;
  const replannedStart = addBusinessDaysUtc(anchorFinish, 1) || anchorFinish;
  if (!replannedStart || replannedStart > replannedFinish) {
    anchorRow.replannedFinish = replannedFinish;
    anchorRow.replannedSource = 'WIP';
    return nextRows;
  }

  const remainingChildIndexes = [];
  for (let i = anchorIndex; i < nextRows.length; i += 1) {
    const row = nextRows[i];
    if (!isChild(row)) continue;
    if (isComplete(row)) continue;
    remainingChildIndexes.push(i);
  }
  if (!remainingChildIndexes.length) return nextRows;

  const replanBusinessDays = Math.max(
    remainingChildIndexes.length,
    countBusinessDaysInclusive(replannedStart, replannedFinish) || remainingChildIndexes.length,
  );
  const scaled = scaleDurationVector(
    remainingChildIndexes.map((index) => ({ key: `row-${index}`, base: Math.max(1, Number(nextRows[index].duration || 1)) })),
    replanBusinessDays,
  );

  let cursor = new Date(replannedStart.getTime());
  for (const index of remainingChildIndexes) {
    const row = nextRows[index];
    const duration = Math.max(1, Number((scaled.find((item) => item.key === `row-${index}`) || {}).duration || row.duration || 1));
    const finish = addBusinessDaysUtc(cursor, duration - 1) || cursor;
    row.replannedStart = new Date(cursor.getTime());
    row.replannedFinish = finish > replannedFinish ? replannedFinish : finish;
    row.replannedSource = 'WIP';
    row.replannedDuration = Math.max(1, countBusinessDaysInclusive(row.replannedStart, row.replannedFinish) || duration);
    cursor = addBusinessDaysUtc(row.replannedFinish, 1) || row.replannedFinish;
    if (cursor > replannedFinish) cursor = new Date(replannedFinish.getTime());
  }

  // Consolida o replanejado nas linhas de grupo conforme os filhos daquele grupo.
  let currentGroupIndex = -1;
  for (let i = 0; i < nextRows.length; i += 1) {
    const row = nextRows[i];
    if (row.type === 'group') {
      currentGroupIndex = i;
      continue;
    }
    if (currentGroupIndex < 0 || !row.replannedFinish) continue;
    const group = nextRows[currentGroupIndex];
    group.replannedStart = clientEarliestDate(group.replannedStart, row.replannedStart);
    group.replannedFinish = clientLatestDate(group.replannedFinish, row.replannedFinish);
    group.replannedSource = 'WIP';
  }

  return nextRows;
}

function clientApplyRowDates(row, start, finish, sourceLabel = '') {
  const next = { ...row };
  const parsedStart = parseDateObject(start);
  const parsedFinish = parseDateObject(finish);
  if (parsedStart) {
    next.plannedStart = next.plannedStart || next.start;
    next.start = parsedStart;
    next.dateSource = sourceLabel || next.dateSource || 'tracking';
  }
  if (parsedFinish) {
    next.plannedFinish = next.plannedFinish || next.finish;
    next.finish = parsedFinish;
    next.dateSource = sourceLabel || next.dateSource || 'tracking';
  }
  if (parseDateObject(next.start) && parseDateObject(next.finish) && parseDateObject(next.start) > parseDateObject(next.finish)) {
    next.start = new Date(parseDateObject(next.finish).getTime());
  }
  next.duration = Math.max(1, countBusinessDaysInclusive(next.start, next.finish) || next.duration || 1);
  if (next.plannedFinish && next.finish) {
    next.deviationDays = Math.round((parseDateObject(next.finish) - parseDateObject(next.plannedFinish)) / 86400000);
  }
  return next;
}

function clientHasExecutedStart(row) {
  return Boolean(row?.plannedStart && parseDateObject(row?.start));
}

function clientHasExecutedFinish(row) {
  return Boolean(row?.plannedFinish && parseDateObject(row?.finish));
}

function clientIsTrackingOrEstimatedRow(row) {
  return Boolean(row?.dateSource || clientHasExecutedStart(row) || clientHasExecutedFinish(row));
}

function clientFillMissingActualRowsBetweenAnchors(rows, childIndexes, startBoundary, finishBoundary) {
  const parsedStart = parseDateObject(startBoundary);
  const parsedFinish = parseDateObject(finishBoundary);
  if (!parsedStart || !parsedFinish || !Array.isArray(childIndexes) || !childIndexes.length) return;

  const orderedStart = parsedStart <= parsedFinish ? parsedStart : parsedFinish;
  const orderedFinish = parsedStart <= parsedFinish ? parsedFinish : parsedStart;
  const totalDays = Math.max(
    childIndexes.length,
    countBusinessDaysInclusive(orderedStart, orderedFinish) || childIndexes.length,
  );
  const scaled = scaleDurationVector(
    childIndexes.map((index) => ({ key: `row-${index}`, base: Math.max(1, Number(rows[index]?.duration || 1)) })),
    totalDays,
  );

  let cursor = new Date(orderedStart.getTime());
  childIndexes.forEach((index, position) => {
    const row = rows[index];
    if (!row || row.type !== 'child') return;
    const isLast = position === childIndexes.length - 1;
    const scaledItem = scaled.find((item) => item.key === `row-${index}`);
    const duration = Math.max(1, Number(scaledItem?.duration || row.duration || 1));
    let finish = addBusinessDaysUtc(cursor, duration - 1) || cursor;
    if (isLast || finish > orderedFinish) finish = new Date(orderedFinish.getTime());

    const filled = clientApplyRowDates(row, cursor, finish, 'Estimado');
    filled.inferredActualDate = true;
    rows[index] = filled;

    cursor = addBusinessDaysUtc(finish, 1) || finish;
    if (cursor > orderedFinish) cursor = new Date(orderedFinish.getTime());
  });
}

function clientInferMissingExecutiveActualDates(rows) {
  if (!Array.isArray(rows) || !rows.length) return rows;
  const nextRows = rows.map((row) => ({ ...row }));
  const groupIndexes = [];
  for (let i = 0; i < nextRows.length; i += 1) {
    if (nextRows[i]?.type === 'group') groupIndexes.push(i);
  }

  for (let g = 0; g < groupIndexes.length; g += 1) {
    const groupIndex = groupIndexes[g];
    const nextGroupIndex = groupIndexes[g + 1] ?? nextRows.length;
    const childIndexes = [];
    for (let i = groupIndex + 1; i < nextGroupIndex; i += 1) {
      if (nextRows[i]?.type === 'child') childIndexes.push(i);
    }
    if (!childIndexes.length) continue;

    const hasAnyActualDate = childIndexes.some((index) => clientIsTrackingOrEstimatedRow(nextRows[index]));
    if (!hasAnyActualDate) continue;

    for (let c = 0; c < childIndexes.length; c += 1) {
      const currentIndex = childIndexes[c];
      const currentRow = nextRows[currentIndex];
      if (clientIsTrackingOrEstimatedRow(currentRow)) continue;

      let previousIndex = null;
      for (let p = c - 1; p >= 0; p -= 1) {
        const candidateIndex = childIndexes[p];
        if (clientHasExecutedFinish(nextRows[candidateIndex])) {
          previousIndex = candidateIndex;
          break;
        }
      }

      let nextIndex = null;
      for (let n = c + 1; n < childIndexes.length; n += 1) {
        const candidateIndex = childIndexes[n];
        if (clientHasExecutedStart(nextRows[candidateIndex]) || clientHasExecutedFinish(nextRows[candidateIndex])) {
          nextIndex = candidateIndex;
          break;
        }
      }

      const missingBlock = [];
      let cursor = c;
      while (cursor < childIndexes.length && !clientIsTrackingOrEstimatedRow(nextRows[childIndexes[cursor]])) {
        missingBlock.push(childIndexes[cursor]);
        cursor += 1;
      }
      if (!missingBlock.length) continue;

      const prevFinish = previousIndex != null ? parseDateObject(nextRows[previousIndex]?.finish) : null;
      let nextStart = null;
      if (nextIndex != null) {
        const nextRow = nextRows[nextIndex];
        nextStart = clientHasExecutedStart(nextRow) ? parseDateObject(nextRow.start) : null;
        if (!nextStart && clientHasExecutedFinish(nextRow)) {
          // Quando a próxima etapa só possui data de término no Tracking, usamos
          // essa data como limite para estimar o bloco anterior, sem alterar o planejado.
          nextStart = parseDateObject(nextRow.finish);
        }
      }

      if (prevFinish && nextStart) {
        clientFillMissingActualRowsBetweenAnchors(nextRows, missingBlock, prevFinish, nextStart);
        c = cursor - 1;
      }
    }

    const childrenWithExecutedDates = childIndexes.map((index) => nextRows[index]).filter(clientIsTrackingOrEstimatedRow);
    if (childrenWithExecutedDates.length) {
      const firstStart = clientEarliestDate(...childrenWithExecutedDates.map((row) => row.start));
      const lastFinish = clientLatestDate(...childrenWithExecutedDates.map((row) => row.finish));
      if (firstStart || lastFinish) {
        const group = nextRows[groupIndex];
        const updatedGroup = clientApplyRowDates(
          group,
          firstStart || group.start,
          lastFinish || group.finish,
          childrenWithExecutedDates.some((row) => row.dateSource === 'Tracking' || row.dateSource === 'PM') ? 'Tracking' : 'Estimado',
        );
        updatedGroup.inferredActualDate = childrenWithExecutedDates.some((row) => row.inferredActualDate);
        nextRows[groupIndex] = updatedGroup;
      }
    }
  }

  return nextRows;
}

function getClientSCurveActualMilestones(project) {
  const real = getClientTrackingDates(project);
  const hydro = shouldClientShowHydro(project);
  // Escala dinâmica para a curva S: ajusta os percentuais conforme o progresso geral atual.
  // Isso evita que uma única ISO com alto progresso cause saltos irreais na curva realizada.
  const currentOverall = getClientOverallProgress(project);
  // Garante fator de escala mínimo 0 para evitar NaN; se currentOverall for 0, escala vira 0.
  const scaleFactor = Number.isFinite(currentOverall) && currentOverall > 0 ? (currentOverall / 100) : 0;
  // A Curva S precisa ser acumulada e sempre crescente.
  // Não usar o percentual atual das etapas antigas aqui, porque uma etapa anterior
  // já concluída em 100% fazia a linha realizada nascer em 100% e depois cair.
  const baseInspectionPercent = hydro ? 72 : 82;
  const points = [
    { date: real.drawingsStart, actual: 0 * scaleFactor, label: 'Drawings Start' },
    { date: real.drawingsFinish, actual: 12 * scaleFactor, label: 'Drawings Finish' },
    { date: real.procurementStart, actual: 15 * scaleFactor, label: 'Procurement Start' },
    { date: real.procurementFinish, actual: 25 * scaleFactor, label: 'Procurement Finish' },
    { date: real.fabricationStart, actual: 30 * scaleFactor, label: 'Fabrication Start' },
    { date: real.boilermakerFinish, actual: 42 * scaleFactor, label: 'Boilermaker Finish' },
    { date: real.weldingFinish, actual: 58 * scaleFactor, label: 'Welding Finish' },
    { date: real.inspectionFinish, actual: baseInspectionPercent * scaleFactor, label: 'Inspection Finish' },
    { date: real.thFinish, actual: 84 * scaleFactor, label: 'TH Finish' },
    { date: real.coatingFinish, actual: 92 * scaleFactor, label: 'Coating Finish' },
    { date: real.projectFinish, actual: 100 * scaleFactor, label: 'Project Finish' },
  ].filter((item) => item.date && (hydro || item.label !== 'TH Finish'));

  const byDate = new Map();
  for (const item of points) {
    const key = item.date.getTime();
    const previous = byDate.get(key);
    if (!previous || item.actual > previous.actual) byDate.set(key, item);
  }

  let runningMax = 0;
  return Array.from(byDate.values())
    .sort((a, b) => a.date - b.date)
    .map((item) => {
      runningMax = Math.max(runningMax, clampClientPercent(item.actual));
      return { ...item, actual: runningMax };
    });
}

function clientInterpolateDate(start, finish, ratio) {
  const a = parseDateObject(start);
  const b = parseDateObject(finish);
  if (!a || !b) return null;
  const safeRatio = Math.max(0, Math.min(1, Number(ratio) || 0));
  const time = a.getTime() + (b.getTime() - a.getTime()) * safeRatio;
  return new Date(time);
}

function normalizeClientSCurveActualMilestones(project, scheduleStart, milestones, scheduleFinish = null) {
  const start = parseDateObject(scheduleStart);
  const finish = parseDateObject(scheduleFinish) || null;
  const list = Array.isArray(milestones) ? milestones.filter((item) => item?.date) : [];
  if (!start || !list.length) return list;

  const real = getClientTrackingDates(project);
  const fabricationStart = parseDateObject(real.fabricationStart) || list[0]?.date;
  const synthetic = [{ date: start, actual: 0, label: 'Início planejado' }];

  // Quando a fabricação real começa muito depois da data inicial da BSP,
  // criamos marcos executivos intermediários para a Curva S nascer no começo
  // do cronograma e subir de forma controlada até o início real da fabricação.
  if (fabricationStart && fabricationStart > start) {
    const preFabricationSpan = fabricationStart.getTime() - start.getTime();
    if (preFabricationSpan > 86400000) {
      synthetic.push(
        { date: clientInterpolateDate(start, fabricationStart, 0.40), actual: 10, label: 'Pré-fabricação | avanço inicial' },
        { date: clientInterpolateDate(start, fabricationStart, 0.70), actual: 20, label: 'Pré-fabricação | avanço intermediário' },
      );
    }
  }

  const plannedDays = finish ? clientDaysBetween(start, finish) : 120;
  const minGapDays = Math.max(1, Math.min(7, Math.ceil(plannedDays * 0.025)));
  const usedTimes = new Set();
  let startCollisionIndex = 0;
  const normalizedItems = [];

  const ordered = synthetic.concat(list)
    .map((item) => ({ ...item, date: parseDateObject(item.date), actual: clampClientPercent(item.actual) }))
    .filter((item) => item.date)
    .sort((a, b) => (a.date - b.date) || (a.actual - b.actual));

  for (const item of ordered) {
    let date = parseDateObject(item.date);
    const actual = clampClientPercent(item.actual);
    if (!date) continue;

    // Se o Tracking traz vários marcos no mesmo dia do início planejado,
    // a linha realizada nascia já em 40%, 60% ou 90% e parecia quebrada.
    // Mantemos o início em 0% e espaçamos visualmente esses marcos para
    // preservar uma curva crescente e fluida, sem alterar os dados originais.
    if (actual > 0 && date.getTime() <= start.getTime()) {
      startCollisionIndex += 1;
      date = addUtcDays(start, startCollisionIndex * minGapDays);
    }

    let guard = 0;
    while (actual > 0 && usedTimes.has(date.getTime()) && guard < 20) {
      date = addUtcDays(date, minGapDays);
      guard += 1;
    }

    const key = date.getTime();
    usedTimes.add(key);
    normalizedItems.push({ ...item, date, actual });
  }

  let runningMax = 0;
  return normalizedItems
    .sort((a, b) => (a.date - b.date) || (a.actual - b.actual))
    .map((item) => {
      runningMax = Math.max(runningMax, clampClientPercent(item.actual));
      return { ...item, actual: runningMax };
    });
}

function clientDateForPlannedPercent(startValue, finishValue, percentValue) {
  const start = parseDateObject(startValue);
  const finish = parseDateObject(finishValue);
  const target = clampClientPercent(percentValue);
  if (!start || !finish || finish <= start) return null;
  let low = 0;
  let high = 1;
  for (let i = 0; i < 32; i += 1) {
    const mid = (low + high) / 2;
    const planned = clientSchedulePlannedPercent(mid);
    if (planned < target) low = mid;
    else high = mid;
  }
  return clientInterpolateDate(start, finish, high);
}

function clientGetCurrentActualPointDate(start, plannedFinish, today, plannedToday, lastMilestone, actualNow) {
  const safeToday = parseDateObject(today);
  const safeStart = parseDateObject(start);
  const safeFinish = parseDateObject(plannedFinish);
  if (!safeToday || !safeStart || !safeFinish || !lastMilestone) return { date: safeToday, label: 'Atual' };

  const duration = Math.max(1, safeFinish - safeStart);
  const elapsedRatio = Math.max(0, Math.min(1, (safeToday - safeStart) / duration));
  const progressJump = clampClientPercent(actualNow) - clampClientPercent(lastMilestone.actual);
  const isCrushedAtStart = elapsedRatio < 0.10 && progressJump > 25 && clampClientPercent(actualNow) > clampClientPercent(plannedToday) + 25;

  if (!isCrushedAtStart) return { date: safeToday, label: 'Atual' };

  const equivalentDate = clientDateForPlannedPercent(safeStart, safeFinish, actualNow);
  const minDate = addUtcDays(parseDateObject(lastMilestone.date) || safeStart, Math.max(2, Math.ceil(clientDaysBetween(safeStart, safeFinish) * 0.04)));
  const date = clientDateMax(equivalentDate, minDate, safeToday) || safeToday;
  return { date, label: 'Atual estimado pelo progresso' };
}

function getClientDetailStageKeys(project) {
  const keys = [
    'Drawing Execution Advance%',
    'Procuremnt Status %',
    'Material Separation',
    'Full welding execution',
    'Non Destructive Examination (QC)',
    'Hydro Test Pressure (QC)',
    'Surface preparation and/or coating',
    'Final Inspection',
    'Package and Delivered',
  ];
  return shouldClientShowHydro(project) ? keys : keys.filter((key) => key !== 'Hydro Test Pressure (QC)');
}

function getClientFabricationStageItems(project) {
  const items = [
    { keys: ['Welding Preparation', 'Spool Assemble and tack weld'], weight: 10 },
    { keys: ['Initial Dimensional Inspection/3D'], weight: 8 },
    { keys: ['Full welding execution'], weight: 25 },
    { keys: ['Non Destructive Examination (QC)'], weight: 12 },
    { keys: ['Final Dimensional Inpection/3D (QC)'], weight: 8 },
    { keys: ['Hydro Test Pressure (QC)'], weight: 7, spoolOnly: true },
    { keys: ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)'], weight: 15 },
  ];
  return items.filter((item) => !item.spoolOnly || shouldClientShowHydro(project));
}

const CLIENT_PRODUCTION_STAGE_EVIDENCE_KEYS = [
  'Drawing Execution Advance%',
  'Drawing',
  'Procuremnt Status %',
  'Procurement Status %',
  'Procurement',
  'Material Separation',
  'Material Release to Fabrication',
  'Welding Preparation',
  'Spool Assemble and tack weld',
  'Initial Dimensional Inspection/3D',
  'Full welding execution',
  'Non Destructive Examination (QC)',
  'Final Dimensional Inpection/3D (QC)',
  'Hydro Test Pressure (QC)',
  'Surface preparation and/or coating',
  'HDG / FBE.  (PAINT)',
  'Final Inspection',
  'Package and Delivered',
];

function hasClientStageEvidence(project, keys = CLIENT_PRODUCTION_STAGE_EVIDENCE_KEYS) {
  const stageValues = project?.stageValues || {};
  return keys.some((key) => {
    const value = stageValues[key];
    return value != null && value !== '' && value !== 'N/A' && value !== 'Não';
  });
}

function getClientStageEvidenceValue(project, keys) {
  const stageValues = project?.stageValues || {};
  for (const key of keys) {
    const value = stageValues[key];
    if (value != null && value !== '' && value !== 'N/A' && value !== 'Não') {
      return { hasEvidence: true, percent: clientPercentValue(value) };
    }
  }
  return { hasEvidence: false, percent: 0 };
}

function getClientProductionStageSnapshots(project) {
  const engineering = getClientStageEvidenceValue(project, ['Drawing Execution Advance%', 'Drawing']);
  const procurementCandidates = [
    getClientStageEvidenceValue(project, ['Procuremnt Status %', 'Procurement Status %', 'Procurement']),
    getClientStageEvidenceValue(project, ['Material Separation']),
    getClientStageEvidenceValue(project, ['Material Release to Fabrication']),
  ];
  const procurement = procurementCandidates.reduce((best, item) => {
    if (!item.hasEvidence) return best;
    if (!best.hasEvidence || item.percent > best.percent) return item;
    return best;
  }, { hasEvidence: false, percent: 0 });
  const fabricationEvidenceKeys = getClientFabricationStageItems(project).flatMap((item) => item.keys);
  const fabrication = {
    hasEvidence: hasClientStageEvidence(project, fabricationEvidenceKeys),
    percent: getClientFabricationProgress(project),
  };
  const packageDelivery = getClientStageEvidenceValue(project, ['Package and Delivered', 'Final Inspection']);
  return [
    { key: 'engineering', label: 'Engineering / Drawing', percent: engineering.percent, weight: 15, hasEvidence: engineering.hasEvidence },
    { key: 'procurement', label: 'Procurement', percent: procurement.percent, weight: 15, hasEvidence: procurement.hasEvidence },
    { key: 'fabrication', label: 'Fabrication', percent: fabrication.percent, weight: 65, hasEvidence: fabrication.hasEvidence },
    { key: 'package', label: 'Package / Delivery', percent: packageDelivery.percent, weight: 5, hasEvidence: packageDelivery.hasEvidence },
  ];
}

function hasClientIncompleteProductionEvidence(project) {
  const stages = getClientProductionStageSnapshots(project);
  return stages.some((stage) => stage.hasEvidence && clampClientPercent(stage.percent) < 99.9);
}

function getClientFirstIncompleteProductionStage(project) {
  return getClientProductionStageSnapshots(project).find((stage) => stage.hasEvidence && clampClientPercent(stage.percent) < 99.9) || null;
}

function getClientStageBasedOverallProgress(project) {
  const stages = getClientProductionStages(project);
  const totalWeight = stages.reduce((sum, stage) => sum + stage.weight, 0) || 100;
  return clampClientPercent(stages.reduce((sum, stage) => sum + stage.percent * stage.weight, 0) / totalWeight);
}

function getClientFabricationProgress(project) {
  const painting = getClientStageValue(project, ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)']);
  if (painting >= 99.9) return 100;

  const stages = getClientFabricationStageItems(project);
  const totalWeight = stages.reduce((sum, item) => sum + item.weight, 0);
  if (!totalWeight) return 0;
  return clampClientPercent(stages.reduce((sum, item) => sum + getClientStageValue(project, item.keys) * item.weight, 0) / totalWeight);
}

function getClientDatabookProgress(project) {
  return getClientStageValue(project, [
    'Databook Issuance',
    'Databook',
    'DATA BOOK',
    'Data Book',
    'DataBook',
  ]);
}

/**
 * Computes the package/delivery progress for a project based on its spools.
 *
 * Historically this function returned the maximum value of the
 * "Package and Delivered" and "Final Inspection" stages recorded on the
 * summary row of a BSP. This caused a single spool with progress in the
 * final stages to mark the entire BSP as completed, even when the vast
 * majority of spools had not progressed beyond unitization. To provide
 * a more representative metric, we now calculate a weighted average of
 * the packaging progress across all spools. Each spool contributes
 * proportionally to its weight (kilos) and spools without a defined
 * weight are treated with weight 1. This ensures the overall delivery
 * progress reflects the state of all spools rather than a single outlier.
 *
 * If no spools are available, the function falls back to the summary
 * stage value for "Package and Delivered".
 *
 * @param {Object} project The project (BSP) object containing spools and stage values.
 * @returns {number} The weighted percentage (0-100) of the packaging stage.
 */
function getClientPackageProgress(project) {
  // If there are spools, compute weighted average of packaging progress across them
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (spools.length > 0) {
    let totalWeight = 0;
    let weightedSum = 0;
    for (const spool of spools) {
      const weight = Number(spool.kilos) || 1;
      totalWeight += weight;
      // Determine spool's packaging progress. We ignore the final inspection
      // when computing delivery progress so that only unitization progress
      // influences the result. Use getClientStageValue on the spool to parse
      // percentage strings (e.g. "95%") into numbers.
      const spoolProgress = getClientStageValue(spool, ['Package and Delivered']);
      weightedSum += spoolProgress * weight;
    }
    return totalWeight > 0 ? (weightedSum / totalWeight) : 0;
  }
  // Fall back to project-level stage value if no spools are available
  return getClientStageValue(project, ['Package and Delivered']);
}

/**
 * Computes the packing progress across all spools based on the
 * "Package and Delivered" stage. Each spool contributes according to its
 * weight (kilos), or 1 if not specified. If no spools are present, the
 * value from the project's summary row is used. This ensures the "Packing"
 * row in the schedule reflects the proportion of ISOs that have reached
 * the unitization/dispatch stage.
 *
 * @param {Object} project The project (BSP) with spools and stage values.
 * @returns {number} The weighted percentage (0-100) of packing progress.
 */
function getClientPackingProgress(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (spools.length > 0) {
    let totalWeight = 0;
    let weightedSum = 0;
    for (const spool of spools) {
      const weight = Number(spool.kilos) || 1;
      totalWeight += weight;
      const progress = getClientStageValue(spool, ['Package and Delivered']);
      weightedSum += progress * weight;
    }
    return totalWeight > 0 ? (weightedSum / totalWeight) : 0;
  }
  return getClientStageValue(project, ['Package and Delivered']);
}

/**
 * Computes the final inspection progress across all spools. Each spool
 * contributes according to its weight (kilos), or 1 if unspecified. When no
 * spools exist, falls back to the project's summary row value. This
 * calculation ensures the "Final Inspection" row of the schedule uses the
 * average progress of all ISOs instead of a single spool's completion.
 *
 * @param {Object} project The project (BSP) with spools and stage values.
 * @returns {number} The weighted percentage (0-100) of final inspection progress.
 */
function getClientFinalInspectionProgress(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (spools.length > 0) {
    let totalWeight = 0;
    let weightedSum = 0;
    for (const spool of spools) {
      const weight = Number(spool.kilos) || 1;
      totalWeight += weight;
      const progress = getClientStageValue(spool, ['Final Inspection']);
      weightedSum += progress * weight;
    }
    return totalWeight > 0 ? (weightedSum / totalWeight) : 0;
  }
  return getClientStageValue(project, ['Final Inspection']);
}

function getClientProductionStages(project) {
  const engineering = getClientStageValue(project, ['Drawing Execution Advance%', 'Drawing']);
  const procurement = Math.max(
    getClientStageValue(project, ['Procuremnt Status %', 'Procurement Status %', 'Procurement']),
    getClientStageValue(project, ['Material Separation']),
    getClientStageValue(project, ['Material Release to Fabrication'])
  );
  const fabrication = getClientFabricationProgress(project);
  const packageDelivery = getClientPackageProgress(project);
  return [
    { key: 'engineering', label: 'Engineering / Drawing', percent: engineering, weight: 15 },
    { key: 'procurement', label: 'Procurement', percent: procurement, weight: 15 },
    { key: 'fabrication', label: 'Fabrication', percent: fabrication, weight: 65 },
    { key: 'package', label: 'Package / Delivery', percent: packageDelivery, weight: 5 },
  ];
}

function getClientOverallProgress(project) {
  const stageBased = getClientStageBasedOverallProgress(project);
  if (hasClientStageEvidence(project)) return stageBased;

  const direct = clientPercentValue(project?.overallProgress);
  if (direct > 0) return direct;
  return stageBased;
}

function getClientAnalyticStartDate(project) {
  const candidates = [
    project?.plannedStartDate,
    project?.startDate,
    project?.stageValues?.['Project Start Date'],
    project?.stageValues?.['Fabrication Start Date'],
    project?.stageValues?.['Drawing Start Date'],
  ];
  for (const value of candidates) {
    const parsed = parseClientSafeDateObject(value);
    if (parsed) return parsed;
  }
  const today = getCurrentBrazilDate();
  const fallback = new Date(today);
  fallback.setUTCDate(today.getUTCDate() - 30);
  return fallback;
}

function getClientAnalyticFinishDate(project) {
  const candidates = [
    project?.plannedFinishDate,
    project?.projectFinishDate,
    project?.stageValues?.['Project Finish Date'],
    getProjectShipmentDate(project),
  ];
  for (const value of candidates) {
    const parsed = parseClientSafeDateObject(value);
    if (parsed) return parsed;
  }
  const start = getClientAnalyticStartDate(project);
  const fallback = new Date(start);
  fallback.setUTCDate(start.getUTCDate() + 120);
  return fallback;
}

function addUtcDays(date, days) {
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + days);
  return next;
}

function clientDaysBetween(start, end) {
  return Math.max(1, Math.round((end - start) / 86400000));
}

function clientSchedulePlannedPercent(progressRatio) {
  const x = Math.max(0, Math.min(1, Number(progressRatio) || 0));
  const segments = [
    { start: 0.00, end: 0.15, base: 0, weight: 15 },
    { start: 0.15, end: 0.30, base: 15, weight: 15 },
    { start: 0.30, end: 0.90, base: 30, weight: 60 },
    { start: 0.90, end: 0.95, base: 90, weight: 5 },
    { start: 0.95, end: 1.00, base: 95, weight: 5 },
  ];
  let total = 0;
  for (const segment of segments) {
    if (x >= segment.end) {
      total = segment.base + segment.weight;
      continue;
    }
    if (x > segment.start) {
      const local = (x - segment.start) / Math.max(0.001, segment.end - segment.start);
      const smooth = local * local * (3 - 2 * local);
      return clampClientPercent(segment.base + smooth * segment.weight);
    }
    return clampClientPercent(total);
  }
  return 100;
}

function getClientPlannedToday(project) {
  const start = getClientAnalyticStartDate(project);
  const finish = getClientAnalyticFinishDate(project);
  const today = getCurrentBrazilDate();
  if (today <= start) return 0;
  if (today >= finish) return 100;
  return clientSchedulePlannedPercent((today - start) / Math.max(1, finish - start));
}


function getClientSCurvePlannedFinishDate(project) {
  const candidates = [
    project?.plannedFinishDate,
    project?.plannedEndDate,
    project?.baselineFinishDate,
    project?.stageValues?.['Planned Finish Date'],
    project?.stageValues?.['Planned Finish'],
    project?.stageValues?.['Baseline Finish Date'],
  ];
  for (const value of candidates) {
    const parsed = parseClientSafeDateObject(value);
    if (parsed) return parsed;
  }
  return getClientAnalyticFinishDate(project);
}

function getProjectReplannedFinishDate(project) {
  if (!CLIENT_REPLANNING_UI_ENABLED) return '';
  const candidates = [
    project?.replannedFinishDate,
    project?.replannedFinish,
    project?.replannedDate,
    project?.deadlineDateAsAgreededWithClient,
    project?.deadlineDateAsAgreedWithClient,
    project?.stageValues?.['Deadline Date as Agreeded with Client*'],
    project?.stageValues?.['Deadline Date as Agreeded with Client'],
    project?.stageValues?.['Deadline Date as Agreed with Client*'],
    project?.stageValues?.['Deadline Date as Agreed with Client'],
    project?.stageValues?.['Data Replanejada'],
    project?.stageValues?.['Replanejado'],
  ];
  for (const value of candidates) {
    const parsed = parseClientSafeDateObject(value);
    if (parsed) return clientFormatDateValue(parsed);
  }
  return '';
}

function getClientSCurveReplannedFinishDate(project) {
  const parsed = parseClientSafeDateObject(getProjectReplannedFinishDate(project));
  return parsed || null;
}

function getClientSCurveReplanInfo(project) {
  const plannedFinish = parseDateObject(getClientSCurvePlannedFinishDate(project));
  const replannedFinish = getClientSCurveReplannedFinishDate(project);
  if (!plannedFinish || !replannedFinish || replannedFinish <= plannedFinish) return null;
  const days = Math.max(1, Math.round((replannedFinish.getTime() - plannedFinish.getTime()) / 86400000));
  return {
    start: plannedFinish,
    end: replannedFinish,
    days,
    label: 'Replanejado',
    tooltip: [
      'Replanejamento',
      `Término planejado: ${clientFormatDateValue(plannedFinish)}`,
      `Término replanejado: ${clientFormatDateValue(replannedFinish)}`,
      `Extensão de prazo: +${days} dia(s)`,
      project?.replannedFinishSource ? `Origem: ${project.replannedFinishSource}` : 'Origem: Work in Progress',
    ].join('\n'),
  };
}

function getClientSCurveShipmentDate(project) {
  const realDates = getClientTrackingDates(project);
  return parseClientSafeDateObject(realDates.projectFinish) || parseClientSafeDateObject(getProjectShipmentDate(project));
}

function getClientSCurveDelayInfo(project) {
  const plannedFinish = parseDateObject(getClientSCurvePlannedFinishDate(project));
  if (!plannedFinish) return null;
  const replannedFinish = getClientSCurveReplannedFinishDate(project);
  const effectiveDeadline = replannedFinish && replannedFinish > plannedFinish ? replannedFinish : plannedFinish;
  const shipmentDate = getClientSCurveShipmentDate(project);
  const today = getCurrentBrazilDate();
  const actualNow = getClientOverallProgress(project);
  let end = null;
  let status = '';
  let endLabel = '';
  let completed = false;

  if (shipmentDate && shipmentDate > effectiveDeadline) {
    end = shipmentDate;
    status = 'Finalizado com atraso';
    endLabel = 'Envio real';
    completed = true;
  } else if (!shipmentDate && actualNow < 100 && today > effectiveDeadline) {
    end = today;
    status = 'Em atraso';
    endLabel = 'Hoje';
  }

  if (!end || end <= effectiveDeadline) return null;
  const days = Math.max(1, Math.round((end.getTime() - effectiveDeadline.getTime()) / 86400000));
  return {
    start: effectiveDeadline,
    end,
    days,
    status,
    endLabel,
    completed,
    tooltip: [
      'Desvio de prazo',
      `Status: ${status}`,
      `${replannedFinish && replannedFinish > plannedFinish ? 'Término replanejado' : 'Término planejado'}: ${clientFormatDateValue(effectiveDeadline)}`,
      `${endLabel}: ${clientFormatDateValue(end)}`,
      `Atraso: +${days} dia(s)`,
    ].join('\n'),
  };
}

function clientDateMax(...values) {
  const dates = values.map(parseDateObject).filter(Boolean).sort((a, b) => a - b);
  return dates.length ? dates[dates.length - 1] : null;
}

function buildClientSCurveData(project) {
  const start = getClientAnalyticStartDate(project);
  let plannedFinish = getClientSCurvePlannedFinishDate(project);
  if (plannedFinish <= start) plannedFinish = addUtcDays(start, 30);
  const today = getCurrentBrazilDate();
  const actualNow = getClientOverallProgress(project);
  const plannedToday = getClientPlannedToday(project);
  const replanInfo = getClientSCurveReplanInfo(project);
  const delayInfo = getClientSCurveDelayInfo(project);
  let trackingMilestones = normalizeClientSCurveActualMilestones(project, start, getClientSCurveActualMilestones(project), plannedFinish);
  const lastTrackingMilestone = trackingMilestones[trackingMilestones.length - 1];
  if (lastTrackingMilestone && today >= lastTrackingMilestone.date && actualNow > lastTrackingMilestone.actual) {
    const currentPoint = clientGetCurrentActualPointDate(start, plannedFinish, today, plannedToday, lastTrackingMilestone, actualNow);
    trackingMilestones = trackingMilestones.concat([{
      date: currentPoint.date || today,
      planned: getClientPlannedToday(project),
      actual: clampClientPercent(actualNow),
      trackingLabel: currentPoint.label || 'Atual',
      label: currentPoint.label || 'Atual',
    }]);
  }

  const lastMilestoneDate = trackingMilestones.length ? trackingMilestones[trackingMilestones.length - 1].date : null;
  const chartFinish = clientDateMax(plannedFinish, replanInfo?.end, delayInfo?.end, lastMilestoneDate) || plannedFinish;
  const plannedDuration = Math.max(1, clientDaysBetween(start, plannedFinish));
  const step = Math.max(1, Math.ceil(plannedDuration / 14));
  const rawPoints = [];

  for (let day = 0; day <= plannedDuration; day += step) {
    const date = addUtcDays(start, day);
    const ratio = day / plannedDuration;
    const planned = clientSchedulePlannedPercent(ratio);
    rawPoints.push({ date, planned, actual: null });
  }
  if (rawPoints[rawPoints.length - 1]?.date < plannedFinish) {
    rawPoints.push({ date: plannedFinish, planned: 100, actual: null, trackingLabel: 'Término planejado' });
  }
  if (replanInfo?.end && replanInfo.end > plannedFinish) {
    rawPoints.push({ date: replanInfo.end, planned: 100, actual: null, trackingLabel: 'Término replanejado', replanInfo });
  }
  if (delayInfo?.end && delayInfo.end > (replanInfo?.end || plannedFinish)) {
    rawPoints.push({ date: delayInfo.end, planned: 100, actual: null, trackingLabel: delayInfo.endLabel || 'Desvio de prazo', delayInfo });
  }
  if (chartFinish > plannedFinish && (!delayInfo || chartFinish.getTime() !== delayInfo.end.getTime()) && (!replanInfo || chartFinish.getTime() !== replanInfo.end.getTime())) {
    rawPoints.push({ date: chartFinish, planned: 100, actual: null });
  }

  for (const milestone of trackingMilestones) {
    const ratio = Math.max(0, Math.min(1, (milestone.date - start) / Math.max(1, plannedFinish - start)));
    rawPoints.push({ date: milestone.date, planned: clientSchedulePlannedPercent(ratio), actual: milestone.actual, trackingLabel: milestone.label });
  }

  const sorted = rawPoints
    .filter((point) => point.date instanceof Date && !Number.isNaN(point.date.getTime()))
    .sort((a, b) => a.date - b.date);

  let lastActual = null;
  return sorted.map((point) => {
    const sameDayMilestones = trackingMilestones.filter((milestone) => milestone.date.getTime() <= point.date.getTime());
    if (sameDayMilestones.length) {
      lastActual = sameDayMilestones[sameDayMilestones.length - 1].actual;
    }
    let actual = null;
    if (trackingMilestones.length) {
      actual = point.date <= today && lastActual != null ? lastActual : null;
    } else if (point.date <= today) {
      if (plannedToday > 0) {
        actual = clampClientPercent((point.planned / plannedToday) * actualNow);
      } else {
        actual = actualNow > 0 ? actualNow : 0;
      }
    }
    return { ...point, actual, replanInfo: point.replanInfo || replanInfo || null, delayInfo: point.delayInfo || delayInfo || null };
  });
}

function clientChartDateBounds(points) {
  const dates = (Array.isArray(points) ? points : [])
    .map((point) => parseDateObject(point?.date))
    .filter(Boolean)
    .sort((a, b) => a - b);
  const start = dates[0] || getCurrentBrazilDate();
  const finish = dates[dates.length - 1] || addUtcDays(start, 1);
  return { start, finish: finish > start ? finish : addUtcDays(start, 1) };
}

function clientChartX(point, points, width, pad) {
  const { start, finish } = clientChartDateBounds(points);
  const date = parseDateObject(point?.date) || start;
  const innerW = width - pad.left - pad.right;
  const ratio = Math.max(0, Math.min(1, (date - start) / Math.max(1, finish - start)));
  return pad.left + ratio * innerW;
}

function clientChartY(value, height, pad) {
  const innerH = height - pad.top - pad.bottom;
  return pad.top + (1 - clampClientPercent(value) / 100) * innerH;
}

function clientSvgPolyline(points, width, height, getValue, domainPoints = points) {
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const usable = (Array.isArray(points) ? points : []).filter((point) => getValue(point) != null);
  if (!usable.length) return '';
  return usable.map((point, index) => {
    const x = clientChartX(point, domainPoints, width, pad);
    const y = clientChartY(getValue(point), height, pad);
    return `${index ? 'L' : 'M'} ${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(' ');
}

function clientActualValueAtDate(points, dateValue) {
  const target = parseDateObject(dateValue);
  if (!target) return null;
  const actualPoints = (Array.isArray(points) ? points : [])
    .map((point) => ({ ...point, date: parseDateObject(point?.date) }))
    .filter((point) => point.date && point.actual != null)
    .sort((a, b) => a.date - b.date);
  if (!actualPoints.length) return null;

  let previous = null;
  let next = null;
  for (const point of actualPoints) {
    if (point.date.getTime() <= target.getTime()) previous = point;
    if (point.date.getTime() >= target.getTime()) {
      next = point;
      break;
    }
  }

  if (previous && previous.date.getTime() === target.getTime()) return clampClientPercent(previous.actual);
  if (next && next.date.getTime() === target.getTime()) return clampClientPercent(next.actual);
  if (previous && next && next.date.getTime() > previous.date.getTime()) {
    const ratio = (target.getTime() - previous.date.getTime()) / Math.max(1, next.date.getTime() - previous.date.getTime());
    return clampClientPercent(previous.actual + (next.actual - previous.actual) * ratio);
  }
  if (previous) return clampClientPercent(previous.actual);
  if (next) return clampClientPercent(next.actual);
  return null;
}

function splitClientSCurveActualSegments(points, delayInfo) {
  const actualPoints = (Array.isArray(points) ? points : [])
    .map((point) => ({ ...point, date: parseDateObject(point?.date) }))
    .filter((point) => point.date && point.actual != null)
    .sort((a, b) => a.date - b.date);
  const delayStart = parseDateObject(delayInfo?.start);
  if (!delayStart || !actualPoints.length) return { normal: actualPoints, delayed: [] };

  const deadlineActual = clientActualValueAtDate(actualPoints, delayStart);
  const hasDeadlinePoint = actualPoints.some((point) => point.date.getTime() === delayStart.getTime());
  const deadlinePoint = deadlineActual == null ? null : {
    date: delayStart,
    planned: 100,
    actual: deadlineActual,
    trackingLabel: 'Início do desvio',
  };

  const normal = actualPoints.filter((point) => point.date.getTime() <= delayStart.getTime());
  const delayed = actualPoints.filter((point) => point.date.getTime() >= delayStart.getTime());

  if (deadlinePoint && !hasDeadlinePoint) {
    normal.push(deadlinePoint);
    delayed.unshift(deadlinePoint);
  } else if (deadlinePoint && !delayed.length) {
    delayed.push(deadlinePoint);
  }

  const dedupe = (list) => {
    const map = new Map();
    for (const point of list.sort((a, b) => (a.date - b.date) || ((a.actual || 0) - (b.actual || 0)))) {
      const key = point.date.toISOString().slice(0, 10);
      const previous = map.get(key);
      if (!previous || (point.actual ?? 0) >= (previous.actual ?? 0)) map.set(key, point);
    }
    return Array.from(map.values()).sort((a, b) => a.date - b.date);
  };

  return {
    normal: dedupe(normal),
    delayed: dedupe(delayed),
  };
}

function buildClientChartHoverTargets(points, width, height) {
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const byDate = new Map();
  for (const point of Array.isArray(points) ? points : []) {
    const date = parseDateObject(point?.date);
    if (!date) continue;
    const key = date.toISOString().slice(0, 10);
    const previous = byDate.get(key) || { date, planned: null, actual: null, labels: [] };
    previous.planned = point.planned != null ? clampClientPercent(point.planned) : previous.planned;
    previous.actual = point.actual != null ? clampClientPercent(point.actual) : previous.actual;
    if (point.trackingLabel && !previous.labels.includes(point.trackingLabel)) previous.labels.push(point.trackingLabel);
    byDate.set(key, previous);
  }
  return Array.from(byDate.values()).sort((a, b) => a.date - b.date).map((point) => {
    const valueForY = point.actual != null ? point.actual : point.planned;
    const x = clientChartX(point, points, width, pad);
    const y = clientChartY(valueForY ?? 0, height, pad);
    const tooltipLines = [
      `Data: ${clientFormatDateValue(point.date)}`,
      `Planejado: ${point.planned == null ? '--' : formatPercent(point.planned)}`,
      `Realizado: ${point.actual == null ? '--' : formatPercent(point.actual)}`,
    ];
    if (point.labels.length) tooltipLines.push(`Marco: ${point.labels.join(' / ')}`);
    const tooltipText = tooltipLines.join('\n');
    return `<circle class="client-chart-hover-target" cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="12" data-client-chart-tooltip="${escapeHtml(tooltipText)}"><title>${escapeHtml(tooltipText)}</title></circle>`;
  }).join('');
}


function renderClientSCurveReplanOverlay(project, points, width, height, pad) {
  const replanInfo = getClientSCurveReplanInfo(project);
  if (!replanInfo) return '';
  const x1 = clientChartX({ date: replanInfo.start }, points, width, pad);
  const x2 = clientChartX({ date: replanInfo.end }, points, width, pad);
  if (!Number.isFinite(x1) || !Number.isFinite(x2) || x2 <= x1) return '';
  const y100 = clientChartY(100, height, pad);
  const bottom = height - pad.bottom;
  const tooltip = escapeHtml(replanInfo.tooltip);
  const labelX = Math.min(width - pad.right - 120, Math.max(pad.left + 6, x1 + 8));
  const labelY = Math.max(pad.top + 14, y100 + 16);
  return `
    <line x1="${x1.toFixed(1)}" y1="${bottom}" x2="${x1.toFixed(1)}" y2="${y100.toFixed(1)}" class="client-chart-replan-boundary" />
    <line x1="${x2.toFixed(1)}" y1="${bottom}" x2="${x2.toFixed(1)}" y2="${y100.toFixed(1)}" class="client-chart-replan-boundary client-chart-replan-boundary--end" />
    <line x1="${x1.toFixed(1)}" y1="${y100.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y100.toFixed(1)}" class="client-chart-replan-line" />
    <text x="${labelX.toFixed(1)}" y="${labelY.toFixed(1)}" class="client-chart-replan-label">Replanejado +${replanInfo.days}d</text>
    <rect x="${x1.toFixed(1)}" y="${pad.top}" width="${(x2 - x1).toFixed(1)}" height="${(bottom - pad.top)}" class="client-chart-hover-target client-chart-replan-hover" data-client-chart-tooltip="${tooltip}"><title>${tooltip}</title></rect>
  `;
}


function renderClientSCurveDelayOverlay(project, points, width, height, pad) {
  const delayInfo = getClientSCurveDelayInfo(project);
  if (!delayInfo) return '';
  const x1 = clientChartX({ date: delayInfo.start }, points, width, pad);
  const x2 = clientChartX({ date: delayInfo.end }, points, width, pad);
  if (!Number.isFinite(x1) || !Number.isFinite(x2) || x2 <= x1) return '';
  const top = pad.top;
  const bottom = height - pad.bottom;
  const y100 = clientChartY(100, height, pad);
  const tooltip = escapeHtml(delayInfo.tooltip);
  const labelX = Math.min(width - pad.right - 90, Math.max(pad.left + 6, x1 + 8));
  const labelY = Math.max(pad.top + 12, y100 - 10);
  return `
    <rect x="${x1.toFixed(1)}" y="${top}" width="${(x2 - x1).toFixed(1)}" height="${(bottom - top)}" class="client-chart-delay-band" />
    <line x1="${x1.toFixed(1)}" y1="${top}" x2="${x1.toFixed(1)}" y2="${bottom}" class="client-chart-delay-boundary" />
    <line x1="${x2.toFixed(1)}" y1="${top}" x2="${x2.toFixed(1)}" y2="${bottom}" class="client-chart-delay-boundary client-chart-delay-boundary--end" />
    <line x1="${x1.toFixed(1)}" y1="${y100.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y100.toFixed(1)}" class="client-chart-delay-line" />
    <text x="${labelX.toFixed(1)}" y="${labelY.toFixed(1)}" class="client-chart-delay-label">+${delayInfo.days}d</text>
    <rect x="${x1.toFixed(1)}" y="${top}" width="${(x2 - x1).toFixed(1)}" height="${(bottom - top)}" class="client-chart-hover-target client-chart-delay-hover" data-client-chart-tooltip="${tooltip}"><title>${tooltip}</title></rect>
  `;
}

function getClientSCurveSvgWidth(points) {
  const baseWidth = 760;
  if (!Array.isArray(points) || points.length < 2) return baseWidth;
  const { start, finish } = clientChartDateBounds(points);
  const totalDays = clientDaysBetween(start, finish);
  // Até aproximadamente 70 dias o gráfico cabe no card.
  // Acima disso, o SVG ganha largura extra e o card mostra barra de rolagem,
  // mantendo a Curva S fluida sem achatar prazos longos.
  const expandedWidth = baseWidth + Math.max(0, totalDays - 70) * 8;
  return Math.min(2200, Math.max(baseWidth, Math.round(expandedWidth)));
}

function wrapClientSCurveSvg(svgMarkup, width) {
  return `
    <div class="client-scurve-scroll" role="region" aria-label="Curva S com rolagem horizontal" tabindex="0">
      <div class="client-scurve-canvas" style="min-width:${Number(width) || 760}px">
        ${svgMarkup}
      </div>
    </div>
  `;
}

function renderClientSCurveSvg(project) {
  const points = buildClientSCurveData(project);
  const width = getClientSCurveSvgWidth(points);
  const height = 260;
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const plannedPath = clientSvgPolyline(points, width, height, (point) => point.planned);
  // v36.43: curva macro/carteira/unidade não possui uma única BSP para calcular desvio.
  // A versão anterior chamava getClientSCurveDelayInfo(project), mas `project` não existe
  // neste escopo e quebrava a abertura da visão executiva da carteira/unidade.
  const delayInfo = null;
  const actualSegments = splitClientSCurveActualSegments(points, delayInfo);
  const actualPath = clientSvgPolyline(actualSegments.normal, width, height, (point) => point.actual, points);
  const delayActualPath = clientSvgPolyline(actualSegments.delayed, width, height, (point) => point.actual, points);
  const grid = [0, 25, 50, 75, 100].map((value) => {
    const y = pad.top + (1 - value / 100) * innerH;
    return `<line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" class="client-chart-grid" /><text x="8" y="${y + 4}" class="client-chart-label">${value}%</text>`;
  }).join('');
  const first = points[0]?.date ? clientFormatDateValue(points[0].date) : '';
  const mid = points[Math.floor(points.length / 2)]?.date ? clientFormatDateValue(points[Math.floor(points.length / 2)].date) : '';
  const last = points[points.length - 1]?.date ? clientFormatDateValue(points[points.length - 1].date) : '';
  const actualCircle = (() => {
    const lastActual = points.filter((point) => point.actual != null).pop();
    if (!lastActual) return '';
    const x = clientChartX(lastActual, points, width, pad);
    const y = clientChartY(lastActual.actual, height, pad);
    const isDelayed = delayInfo?.start && parseDateObject(lastActual.date) && parseDateObject(lastActual.date) >= parseDateObject(delayInfo.start);
    return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="5" class="${isDelayed ? 'client-chart-dot client-chart-dot--delay' : 'client-chart-dot'}" />`;
  })();
  const replanOverlay = renderClientSCurveReplanOverlay(project, points, width, height, pad);
  const delayOverlay = renderClientSCurveDelayOverlay(project, points, width, height, pad);
  const hoverTargets = buildClientChartHoverTargets(points, width, height);
  return wrapClientSCurveSvg(`
    <svg class="client-scurve-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Curva S planejado versus realizado">
      ${grid}
      ${replanOverlay}
      ${delayOverlay}
      <line x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <path d="${plannedPath}" class="client-chart-planned" />
      ${actualPath ? `<path d="${actualPath}" class="client-chart-actual" />` : ''}
      ${delayActualPath ? `<path d="${delayActualPath}" class="client-chart-actual-delay" />` : ''}
      ${actualCircle}
      ${hoverTargets}
      <text x="${pad.left}" y="${height - 12}" class="client-chart-date">${escapeHtml(first)}</text>
      <text x="${pad.left + innerW / 2 - 38}" y="${height - 12}" class="client-chart-date">${escapeHtml(mid)}</text>
      <text x="${width - pad.right - 72}" y="${height - 12}" class="client-chart-date">${escapeHtml(last)}</text>
    </svg>
  `, width);
}

function isClientDeviationActive(deadlineDate) {
  const deadline = parseDateObject(deadlineDate);
  if (!deadline) return true;
  const today = new Date();
  const utcToday = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate()));
  return utcToday.getTime() > deadline.getTime();
}

function getClientVisibleDeviationPercent(actualPercent, plannedPercent, deadlineDate = null) {
  if (!isClientDeviationActive(deadlineDate)) return 0;
  return Math.max(0, clampClientPercent(plannedPercent) - clampClientPercent(actualPercent));
}

function renderClientGauge(percent, label, plannedPercent = null, options = {}) {
  const p = clampClientPercent(percent);
  const plannedBase = plannedPercent == null ? p : clampClientPercent(plannedPercent);
  const deviationActive = isClientDeviationActive(options?.deadlineDate || null);
  const planned = deviationActive ? Math.max(p, plannedBase) : p;
  const deviation = deviationActive ? Math.max(0, planned - p) : 0;
  const deliveryDate = options?.deliveryDate ? clientFormatDateValue(options.deliveryDate) : '';
  const noteText = options?.note ? String(options.note) : '';
  const ringStyle = deviationActive
    ? `background: conic-gradient(#0b9b7a 0 ${p}%, #efc14f ${p}% ${planned}%, #d4dce2 ${planned}% 100%)`
    : `background: conic-gradient(#0b9b7a 0 ${p}%, #d4dce2 ${p}% 100%)`;
  return `
    <div class="client-exec-gauge" style="--p:${p}">
      <div class="client-exec-gauge-ring" style="${ringStyle}"></div>
      <div class="client-exec-gauge-center">
        <strong>${formatPercent(p)}</strong>
        <span>${escapeHtml(label)}</span>
      </div>
      <div class="client-exec-gauge-meta">
        <small><i class="actual"></i>Realizado ${formatPercent(p)}</small>
        <small><i class="deviation"></i>Desvio ${formatPercent(deviation)}</small>
        ${noteText ? `<small><i class="planned"></i>${escapeHtml(noteText)}</small>` : ''}
        ${deliveryDate ? `<small><i class="delivery"></i>Envio: ${escapeHtml(deliveryDate)}</small>` : ''}
      </div>
    </div>
  `;
}

function getClientCompletedTags(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (!spools.length) return 0;
  return spools.filter((spool) => !hasClientIncompleteProductionEvidence(spool) && (clientPercentValue(spool?.overallProgress) >= 100 || String(spool?.uiState || '').toLowerCase() === 'completed')).length;
}

function getClientAttentionPoints(project) {
  const actual = getClientOverallProgress(project);
  const planned = getClientPlannedToday(project);
  const stages = getClientProductionStages(project);
  const finish = getClientAnalyticFinishDate(project);
  const today = getCurrentBrazilDate();
  const daysToFinish = Math.round((finish - today) / 86400000);
  const points = [];
  if (planned - actual >= 10) points.push(`Realizado ${formatPercent(actual)} contra planejado ${formatPercent(planned)}: desvio de ${formatPercent(planned - actual)}.`);
  if (getClientFabricationProgress(project) < 25 && planned > 35) points.push('Fabricação abaixo do ritmo planejado para a data atual.');
  const painting = getClientStageValue(project, ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)']);
  if (getClientFabricationProgress(project) >= 60 && painting <= 0) points.push('Pintura ainda não iniciada após avanço relevante da fabricação.');
  const delivery = getClientPackageProgress(project);
  if (daysToFinish <= 21 && daysToFinish >= 0 && actual < 80) points.push(`Término previsto em ${daysToFinish} dia(s) com progresso abaixo de 80%.`);
  const blocked = stages.find((stage) => stage.percent <= 0);
  if (blocked && planned > 20) points.push(`${blocked.label} sem progresso registrado.`);
  if (!points.length) points.push('Sem ponto crítico automático identificado neste momento.');
  return points;
}

function countBusinessDaysInclusive(startValue, endValue) {
  const start = parseDateObject(startValue);
  const end = parseDateObject(endValue);
  if (!start || !end) return 0;
  let current = new Date(start.getTime());
  let finish = new Date(end.getTime());
  if (current > finish) [current, finish] = [finish, current];
  let total = 0;
  while (current <= finish) {
    const day = current.getUTCDay();
    if (day !== 0 && day !== 6) total += 1;
    current.setUTCDate(current.getUTCDate() + 1);
  }
  return total;
}

function addBusinessDaysUtc(dateValue, amount) {
  const date = parseDateObject(dateValue);
  if (!date) return null;
  const result = new Date(date.getTime());
  let remaining = Number(amount || 0);
  const direction = remaining >= 0 ? 1 : -1;
  remaining = Math.abs(remaining);
  while (remaining > 0) {
    result.setUTCDate(result.getUTCDate() + direction);
    const day = result.getUTCDay();
    if (day !== 0 && day !== 6) remaining -= 1;
  }
  return result;
}

function formatClientDateShort(value) {
  return clientFormatDateValue(value) || '—';
}

function scaleDurationVector(baseItems, targetTotal) {
  const totalBase = baseItems.reduce((sum, item) => sum + Number(item.base || 0), 0) || 1;
  const safeTarget = Math.max(baseItems.length, Number(targetTotal || 0));
  const raw = baseItems.map((item) => ({ ...item, raw: (Number(item.base || 0) / totalBase) * safeTarget }));
  let rows = raw.map((item) => ({ ...item, duration: Math.max(1, Math.floor(item.raw)), fraction: item.raw - Math.floor(item.raw) }));
  let currentTotal = rows.reduce((sum, item) => sum + item.duration, 0);
  if (currentTotal < safeTarget) {
    rows.sort((a, b) => b.fraction - a.fraction);
    for (let i = 0; currentTotal < safeTarget && rows.length; i = (i + 1) % rows.length) {
      rows[i].duration += 1;
      currentTotal += 1;
    }
  } else if (currentTotal > safeTarget) {
    rows.sort((a, b) => a.fraction - b.fraction);
    for (let i = 0; currentTotal > safeTarget && rows.length; i = (i + 1) % rows.length) {
      if (rows[i].duration > 1) {
        rows[i].duration -= 1;
        currentTotal -= 1;
      }
    }
  }
  return baseItems.map((item) => rows.find((row) => row.key === item.key) || { ...item, duration: Math.max(1, Number(item.base || 1)) });
}

function buildClientExecutiveSchedule(project, options = {}) {
  const start = parseDateObject(options.start) || getClientAnalyticStartDate(project);
  const finish = parseDateObject(options.finish) || getClientAnalyticFinishDate(project) || getProjectShipmentDate(project);
  const totalBusinessDays = Math.max(5, countBusinessDaysInclusive(start, finish) || 116);
  const realDates = options.realDates || getClientTrackingDates(project);
  const hasRealFabricationDates = Boolean(realDates.drawingsStart || realDates.drawingsFinish || realDates.procurementStart || realDates.procurementFinish || realDates.fabricationStart || realDates.boilermakerFinish || realDates.weldingFinish || realDates.inspectionFinish || realDates.thFinish || realDates.coatingFinish);

  const groupTemplate = [
    { key: 'engineering', label: 'ENGINEERING', base: 15, percent: getClientStageValue(project, ['Drawing Execution Advance%', 'Drawing']) },
    { key: 'procurement', label: 'MATERIAL PROCUREMENT', base: 15, percent: Math.max(getClientStageValue(project, ['Procuremnt Status %', 'Procurement Status %', 'Procurement']), getClientStageValue(project, ['Material Separation']), getClientStageValue(project, ['Material Release to Fabrication'])) },
    { key: 'fabrication', label: 'FABRICATION', base: 81, percent: getClientFabricationProgress(project) },
    { key: 'delivery', label: 'DELIVERY', base: 2, percent: getClientPackageProgress(project) },
  ];
  const groupDurations = scaleDurationVector(groupTemplate, totalBusinessDays);

  const fabricationTemplate = [
    { key: 'fitup', label: 'Fit up', base: 15, percent: Math.max(getClientStageValue(project, ['Welding Preparation', 'Spool Assemble and tack weld']), 0) },
    { key: 'initial-dimensional', label: 'Initial Dimensional Inspection', base: 10, percent: getClientStageValue(project, ['Initial Dimensional Inspection/3D']) },
    { key: 'weld', label: 'Weld', base: 10, percent: getClientStageValue(project, ['Full welding execution']) },
    { key: 'nde', label: 'Non Destructive Examination', base: 8, percent: getClientStageValue(project, ['Non Destructive Examination (QC)']) },
    { key: 'final-dimensional', label: 'Final Dimensional Inspection', base: 8, percent: getClientStageValue(project, ['Final Dimensional Inpection/3D (QC)']) },
    { key: 'hydro', label: 'Hydro Testing', base: 6, percent: getClientStageValue(project, ['Hydro Test Pressure (QC)']), spoolOnly: true },
    { key: 'painting', label: 'Painting', base: 14, percent: getClientStageValue(project, ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)']) },
  ].filter((item) => !item.spoolOnly || shouldClientShowHydro(project));
  const logisticsTemplate = [
    // "Packing" reflects the unitization and dispatch stage (Package and Delivered).
    // "Packing" reflects the weighted average of unitization/dispatch across spools.
    { key: 'packing', label: 'Packing', base: 1, percent: getClientPackingProgress(project) },
    // "Final Inspection" reflects the weighted average of final inspection progress across spools.
    { key: 'final-inspection', label: 'Final Inspection', base: 1, percent: getClientFinalInspectionProgress(project) },
    // "Delivery" row reflects the aggregated packaging progress across spools.
    { key: 'delivery', label: 'Delivery', base: 1, percent: getClientPackageProgress(project) },
  ];
  const fabricationGroup = groupDurations.find((item) => item.key === 'fabrication');
  const deliveryGroup = groupDurations.find((item) => item.key === 'delivery');
  const fabricationDurations = scaleDurationVector(fabricationTemplate, fabricationGroup?.duration || 81);
  const logisticsDurations = scaleDurationVector(logisticsTemplate, deliveryGroup?.duration || 3);
  const scheduleFallbackPercents = getClientScheduleFallbackPercents(project);

  const rows = [];
  let cursor = parseDateObject(start) || parseDateObject(getProjectShipmentDate(project)) || new Date();
  for (const group of groupDurations) {
    const groupStart = new Date(cursor.getTime());
    const groupFinish = addBusinessDaysUtc(groupStart, Math.max(0, (group.duration || 1) - 1)) || groupStart;
    rows.push({ type: 'group', key: group.key, label: group.label, progress: group.percent, duration: group.duration, start: groupStart, finish: groupFinish });
    let children = [];
    if (group.key === 'engineering') {
      children = [{ key: 'drawings', label: 'Drawings', progress: group.percent, duration: group.duration }];
    } else if (group.key === 'procurement') {
      children = [{ key: 'materials', label: 'Materials for Application Acquisition', progress: group.percent, duration: group.duration }];
    } else if (group.key === 'fabrication') {
      children = fabricationDurations.map((item) => ({ key: item.key, label: item.label, progress: item.percent, duration: item.duration }));
    } else if (group.key === 'delivery') {
      children = logisticsDurations.map((item) => ({ key: item.key, label: item.label, progress: item.percent, duration: item.duration }));
    }

    let childCursor = new Date(groupStart.getTime());
    for (const child of children) {
      const childStart = new Date(childCursor.getTime());
      const childFinish = addBusinessDaysUtc(childStart, Math.max(0, (child.duration || 1) - 1)) || childStart;
      rows.push({ type: 'child', key: child.key, label: child.label, progress: child.progress, duration: child.duration, start: childStart, finish: childFinish });
      childCursor = addBusinessDaysUtc(childFinish, 1) || childFinish;
    }

    cursor = addBusinessDaysUtc(groupFinish, 1) || groupFinish;
  }

  if (!hasRealFabricationDates) return options.skipReplan ? rows : applyClientExecutiveScheduleReplan(project, rows);

  const rowsWithTracking = rows.map((row) => {
    if (row.key === 'engineering' || row.key === 'drawings') {
      if (realDates.drawingsStart || realDates.drawingsFinish) {
        return clientApplyRowDates(row, realDates.drawingsStart, realDates.drawingsFinish, realDates.sources?.drawingsStart || realDates.sources?.drawingsFinish || 'PM');
      }
      return row;
    }
    if (row.key === 'procurement' || row.key === 'materials') {
      if (realDates.procurementStart || realDates.procurementFinish) {
        return clientApplyRowDates(row, realDates.procurementStart, realDates.procurementFinish, realDates.sources?.procurementStart || realDates.sources?.procurementFinish || 'PM');
      }
      return row;
    }
    if (row.key === 'fabrication') {
      const fabFinish = clientLatestDate(realDates.coatingFinish, realDates.thFinish, realDates.inspectionFinish, realDates.weldingFinish, realDates.boilermakerFinish);
      return clientApplyRowDates(row, realDates.fabricationStart, fabFinish, realDates.sources?.fabricationStart || realDates.sources?.coatingFinish || realDates.sources?.thFinish || realDates.sources?.inspectionFinish || realDates.sources?.weldingFinish || realDates.sources?.boilermakerFinish || 'Tracking');
    }
    if (row.key === 'fitup') return clientApplyRowDates(row, realDates.fabricationStart, realDates.boilermakerFinish, realDates.sources?.fabricationStart || realDates.sources?.boilermakerFinish || 'Tracking');
    if (row.key === 'weld') return clientApplyRowDates(row, realDates.boilermakerFinish ? clientNextBusinessDay(realDates.boilermakerFinish) : null, realDates.weldingFinish, realDates.sources?.weldingFinish || 'Tracking');
    if (row.key === 'final-dimensional') return clientApplyRowDates(row, realDates.weldingFinish ? clientNextBusinessDay(realDates.weldingFinish) : null, realDates.inspectionFinish, realDates.sources?.inspectionFinish || 'Tracking');
    if (row.key === 'hydro' && realDates.thFinish) return clientApplyRowDates(row, realDates.inspectionFinish ? clientNextBusinessDay(realDates.inspectionFinish) : null, realDates.thFinish, realDates.sources?.thFinish || 'Tracking');
    if (row.key === 'painting') {
      const paintStartBase = shouldClientShowHydro(project) ? realDates.thFinish : realDates.inspectionFinish;
      if (realDates.coatingFinish) return clientApplyRowDates(row, paintStartBase ? clientNextBusinessDay(paintStartBase) : null, realDates.coatingFinish, realDates.sources?.coatingFinish || 'Tracking');
      return row;
    }
    if (row.key === 'delivery' && row.type === 'group') return clientApplyRowDates(row, null, realDates.projectFinish, realDates.sources?.projectFinish || 'Tracking');
    return row;
  });

  const rowsWithEstimatedActuals = clientInferMissingExecutiveActualDates(rowsWithTracking).map((row) => {
    const next = { ...row };
    if (next.key === 'fitup' && clampClientPercent(next.progress) <= 0) {
      if (scheduleFallbackPercents.fitup > 0) next.progress = scheduleFallbackPercents.fitup;
      else if (clientHasExecutedStart(next) || clientHasExecutedFinish(next)) next.progress = 1;
    }
    if (next.key === 'fabrication' && clampClientPercent(next.progress) <= 0) {
      if (scheduleFallbackPercents.fabrication > 0) next.progress = scheduleFallbackPercents.fabrication;
      else if (clientHasExecutedStart(next) || clientHasExecutedFinish(next)) next.progress = 1;
    }
    return next;
  });

  const fitupRow = rowsWithEstimatedActuals.find((row) => row.key === 'fitup');
  const fabricationRow = rowsWithEstimatedActuals.find((row) => row.key === 'fabrication');
  if (fabricationRow) {
    const childProgressMax = rowsWithEstimatedActuals
      .filter((row) => row.type === 'child' && ['fitup', 'initial-dimensional', 'weld', 'nde', 'final-dimensional', 'hydro', 'painting'].includes(row.key))
      .reduce((max, row) => Math.max(max, clampClientPercent(row.progress)), 0);
    if (childProgressMax > clampClientPercent(fabricationRow.progress)) fabricationRow.progress = childProgressMax;
    if (clampClientPercent(fabricationRow.progress) <= 0 && (clientHasExecutedStart(fabricationRow) || (fitupRow && clampClientPercent(fitupRow.progress) > 0))) fabricationRow.progress = Math.max(1, clampClientPercent(fitupRow?.progress || 0));
  }

  return options.skipReplan ? rowsWithEstimatedActuals : applyClientExecutiveScheduleReplan(project, rowsWithEstimatedActuals);
}

function getClientScheduleFallbackPercents(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const total = spools.length || 0;
  if (!total) return { fitup: 0, fabrication: 0 };

  const reachedFitupOrBeyond = spools.filter((spool) => {
    const directText = normalizeStageWorkspaceText([
      spool?.currentStatus,
      spool?.stage,
      spool?.currentSector,
      spool?.operationalSector,
      spool?.flow?.status,
      spool?.flow?.sector,
      spool?.etapaAtual,
    ].filter(Boolean).join(' '));
    if (directText.includes('fit up') || directText.includes('spool assemble') || directText.includes('tack weld') || directText.includes('welding preparation')) return true;
    const sector = getSpoolCompetenceSector(project, spool);
    return ['calderaria', 'solda', 'inspecao', 'pintura', 'pendente_envio'].includes(sector) || isClientSpoolFinished(spool);
  }).length;

  const percent = clampClientPercent((reachedFitupOrBeyond / total) * 100);
  return { fitup: percent, fabrication: percent };
}

function getClientScheduleVisualState(progress, row = null) {
  const value = clampClientPercent(progress);
  if (value >= 99.9) return 'completed';
  if (value > 0) return 'in-progress';
  if (row && (clientHasExecutedStart(row) || clientHasExecutedFinish(row))) return 'in-progress';
  return 'not-started';
}

function renderClientExecutiveScheduleRows(rows, emptyText = 'Schedule não disponível para esta BSP.') {
  if (!rows.length) return `<div class="client-empty-state">${escapeHtml(emptyText)}</div>`;
  return `
    <div class="client-table-wrap client-table-wrap--compact client-exec-schedule-table">
      <table class="client-bsp-table client-bsp-table--schedule">
        <thead><tr><th>Etapa</th><th>%</th><th>Prazo médio</th><th>Início</th><th>Término</th><th>Status</th></tr></thead>
        <tbody>
          ${rows.map((row) => {
            const state = getClientScheduleVisualState(row.progress, row);
            const label = row.type === 'group' ? `<strong>${escapeHtml(row.label)}</strong>` : `<span class="client-schedule-child">${escapeHtml(row.label)}</span>`;
            const startCell = `${formatClientDateShort(row.start)}${row.plannedStart ? `<small class="client-planned-date">Plan.: ${formatClientDateShort(row.plannedStart)}</small>` : ''}`;
            const finishCell = `${formatClientDateShort(row.finish)}${row.plannedFinish ? `<small class="client-planned-date">Plan.: ${formatClientDateShort(row.plannedFinish)}</small>` : ''}`;
            const deviationText = Number.isFinite(row.deviationDays) && row.deviationDays > 0 ? ` • +${row.deviationDays}d` : '';
            const statusText = state === 'completed' ? 'Concluído' : state === 'in-progress' ? 'Em andamento' : 'Não iniciado';
            return `<tr class="client-schedule-row client-schedule-row--${state} client-schedule-row--${row.type}"><td>${label}</td><td><span class="client-spool-progress client-spool-progress--${state}">${formatPercent(row.progress)}</span></td><td>${formatNumber(row.duration, 0)}d</td><td>${startCell}</td><td>${finishCell}</td><td><span class="client-spool-chip client-spool-chip--${state}">${statusText}${deviationText}</span></td></tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderClientExecutiveSchedule(project) {
  return renderClientExecutiveScheduleRows(buildClientExecutiveSchedule(project), 'Schedule não disponível para esta BSP.');
}


function formatClientTrackingReportScreenCell(value, column = {}) {
  if (value == null || value === '') return '—';
  const rawText = String(value).trim();
  if (!rawText) return '—';
  if (column.type === 'percent' || column.type === 'percent-or-text') {
    if (rawText.toUpperCase() === 'N/A') return 'N/A';
    const number = Number(value);
    if (!Number.isFinite(number)) return escapeHtml(rawText);
    return formatPercent(number * 100);
  }
  if (column.type === 'number') {
    const number = Number(value);
    if (!Number.isFinite(number)) return escapeHtml(rawText);
    const fraction = /kilos/i.test(column.label || '') ? 0 : 0;
    return formatNumber(number, fraction);
  }
  return escapeHtml(rawText);
}

function renderClientTrackingReportPreview(project) {
  const rows = buildClientTrackingReportRows(project);
  if (!rows.length) return '<section class="client-exec-card client-exec-report-preview"><div class="client-empty-state">Report não disponível para esta BSP.</div></section>';
  const summary = rows[0] || {};
  const columns = CLIENT_TRACKING_REPORT_COLUMNS;
  const po = summary['Client PO Number'] || getClientTrackingReportPo(project) || '—';
  const progress = formatClientTrackingReportScreenCell(summary['% Individual Progress'], { type: 'percent' });
  const spoolCount = summary['Quantity Spools'] || getProjectItemCount(project) || 0;
  const kilos = summary['Kilos'] || project?.kilos || 0;
  return `
    <section class="client-exec-card client-exec-report-preview">
      <div class="client-exec-card-head client-exec-card-head--report">
        <div>
          <h3>Report do Cliente</h3>
          <span>Dados do Tracking + Work in Progress exibidos dentro da visão principal</span>
        </div>
        <button class="client-exec-pdf-button client-exec-report-button" type="button" data-client-download-report="${escapeHtml(project?.rowId || '')}">Baixar Excel do Cronograma</button>
      </div>
      <div class="client-report-summary-strip">
        <article><span>Project</span><strong>${escapeHtml(getClientTrackingReportProjectText(project))}</strong></article>
        <article><span>Client PO Number</span><strong>${escapeHtml(po)}</strong></article>
        <article><span>PM</span><strong>${escapeHtml(summary.PM || project?.pm || '—')}</strong></article>
        <article><span>Project Type</span><strong>${escapeHtml(summary['Project Type'] || getProjectTypeLabel(project) || '—')}</strong></article>
        <article><span>Quantity Spools</span><strong>${formatNumber(spoolCount, 0)}</strong></article>
        <article><span>Kilos</span><strong>${formatNumber(kilos, 0)}</strong></article>
        <article><span>% Individual Progress</span><strong>${escapeHtml(progress)}</strong></article>
      </div>
      <div class="client-table-wrap client-report-table-wrap">
        <table class="client-bsp-table client-tracking-report-table">
          <thead><tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join('')}</tr></thead>
          <tbody>
            ${rows.map((row) => `<tr class="${row._summary ? 'client-report-summary-row' : ''}">${columns.map((column) => `<td>${formatClientTrackingReportScreenCell(row[column.label], column)}</td>`).join('')}</tr>`).join('')}
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function getClientStageTimeline(project) {
  const stages = getClientProductionStages(project);
  return stages.map((stage) => ({
    ...stage,
    state: stage.percent >= 100 ? 'done' : stage.percent > 0 ? 'active' : 'future',
  }));
}


function isClientBspOverridesMissingTableMessage(message = '') {
  return /client_bsp_overrides|Tabela client_bsp_overrides/i.test(String(message || ''));
}

function getClientSafeOverrideFeedback(message = '') {
  return isClientBspOverridesMissingTableMessage(message) ? '' : String(message || '');
}


function renderClientBspOverrideNotice(project) {
  const override = getClientBspOverride(project);
  if (!hasClientBspOverrideContent(override)) return '';
  const custom = clientOverrideCustomFieldsArray(override);
  const updated = override.updatedAt ? new Date(override.updatedAt).toLocaleString('pt-BR') : '';
  const meta = [override.updatedByName || override.updatedBy, updated].filter(Boolean).join(' • ');
  return `
    <section class="client-pm-override-notice">
      <div>
        <span class="client-pm-override-badge">Informação executiva PM</span>
        <h3>${escapeHtml(override.executiveStatus || 'Ajuste executivo aplicado')}</h3>
        ${override.executiveNote ? `<p>${escapeHtml(override.executiveNote)}</p>` : ''}
        ${override.delayReason ? `<small>Motivo / desvio: ${escapeHtml(override.delayReason)}</small>` : ''}
        ${custom.length ? `<div class="client-pm-custom-fields">${custom.map((item) => `<span><strong>${escapeHtml(item.label || 'Campo')}:</strong> ${escapeHtml(item.value || '—')}</span>`).join('')}</div>` : ''}
      </div>
      ${meta ? `<small class="client-pm-override-meta">Última atualização: ${escapeHtml(meta)}</small>` : ''}
    </section>
  `;
}

function renderClientBspOverrideEditor(project) {
  if (!canManageClientBspPanel(project)) return '';
  if (String(state.clientBspOverrides.editingProjectId || '') !== String(project?.rowId || '')) return '';
  const override = getClientBspOverride(project) || {};
  const tracking = getClientTrackingDates(project, { ignoreOverrides: true });
  const field = (name, label, trackingValue, overrideValue) => `
    <label class="client-pm-date-field">
      <span>${escapeHtml(label)}</span>
      <input type="date" name="${escapeHtml(name)}" value="${escapeHtml(clientDateInputValue(overrideValue))}" />
      <small>Tracking: ${escapeHtml(formatClientDateShort(trackingValue))}</small>
    </label>
  `;
  const custom = clientOverrideCustomFieldsArray(override)[0] || { label: '', value: '' };
  return `
    <section class="client-pm-editor" data-client-pm-editor="${escapeHtml(project.rowId)}">
      <div class="client-exec-card-head">
        <h3>Editar datas e informações do cliente</h3>
        <span>Preencha somente o que deseja ajustar. Em branco, o painel continua usando Tracking/planejado. Os dados ficam salvos no Supabase e não alteram o Tracking/Smartsheet.</span>
      </div>
      <form class="client-pm-editor-form" data-client-bsp-override-form="${escapeHtml(project.rowId)}">
        <div class="client-pm-date-grid">
          ${field('drawingsStartOverride', 'Drawings Start Date', tracking.drawingsStart, override.drawingsStartOverride)}
          ${field('drawingsFinishOverride', 'Drawings Finish Date', tracking.drawingsFinish, override.drawingsFinishOverride)}
          ${field('procurementStartOverride', 'Procurement Start Date', tracking.procurementStart, override.procurementStartOverride)}
          ${field('procurementFinishOverride', 'Procurement Finish Date', tracking.procurementFinish, override.procurementFinishOverride)}
          ${field('fabricationStartOverride', 'Fabrication Start Date', tracking.fabricationStart, override.fabricationStartOverride)}
          ${field('boilermakerFinishOverride', 'Boilermaker Finish Date', tracking.boilermakerFinish, override.boilermakerFinishOverride)}
          ${field('weldingFinishOverride', 'Welding Finish Date', tracking.weldingFinish, override.weldingFinishOverride)}
          ${field('inspectionFinishOverride', 'Inspection Finish Date (QC)', tracking.inspectionFinish, override.inspectionFinishOverride)}
          ${shouldClientShowHydro(project) ? field('thFinishOverride', 'TH Finish Date', tracking.thFinish, override.thFinishOverride) : ''}
          ${field('coatingFinishOverride', 'Coating Finish Date', tracking.coatingFinish, override.coatingFinishOverride)}
          ${field('projectFinishOverride', 'Project Finish / Envio', tracking.projectFinish, override.projectFinishOverride)}
        </div>
        <div class="client-pm-text-grid">
          <label><span>Status executivo</span><input type="text" name="executiveStatus" maxlength="160" value="${escapeHtml(override.executiveStatus || '')}" placeholder="Ex.: Em tratativa / Prazo revisado" /></label>
          <label><span>Motivo do desvio</span><input type="text" name="delayReason" maxlength="220" value="${escapeHtml(override.delayReason || '')}" placeholder="Ex.: Aguardando liberação do cliente" /></label>
          <label class="client-pm-editor-wide"><span>Observação para o cliente</span><textarea name="executiveNote" rows="3" maxlength="1000" placeholder="Resumo executivo visível no painel do cliente">${escapeHtml(override.executiveNote || '')}</textarea></label>
          <label><span>Campo adicional</span><input type="text" name="customFieldLabel" maxlength="80" value="${escapeHtml(custom.label || '')}" placeholder="Ex.: Pendência" /></label>
          <label><span>Valor do campo adicional</span><input type="text" name="customFieldValue" maxlength="220" value="${escapeHtml(custom.value || '')}" placeholder="Ex.: PO em revisão" /></label>
        </div>
        <div id="client-pm-editor-feedback" class="client-pm-editor-feedback">${escapeHtml(getClientSafeOverrideFeedback(state.clientBspOverrides.feedback || ''))}</div>
        <div class="client-pm-editor-actions">
          <button class="primary-button" type="submit">Salvar ajustes</button>
          <button class="ghost-button" type="button" data-client-bsp-edit-cancel>Cancelar</button>
          ${override.id ? '<button class="ghost-button ghost-button--danger" type="button" data-client-bsp-clear-override>Limpar ajuste</button>' : ''}
        </div>
      </form>
    </section>
  `;
}

function collectClientBspOverrideForm(form, project = null) {
  const fd = new FormData(form);
  const customLabel = String(fd.get('customFieldLabel') || '').trim();
  const customValue = String(fd.get('customFieldValue') || '').trim();
  const val = (key) => String(fd.get(key) || '').trim();
  const existingOverride = project ? getClientBspOverride(project) : null;
  const existingCustom = existingOverride?.customFields && typeof existingOverride.customFields === 'object' ? existingOverride.customFields : {};
  const customFields = {};
  if (existingCustom.__isoDateOverrides && typeof existingCustom.__isoDateOverrides === 'object') {
    customFields.__isoDateOverrides = existingCustom.__isoDateOverrides;
  }
  customFields.__drawingsStartOverride = val('drawingsStartOverride');
  customFields.__drawingsFinishOverride = val('drawingsFinishOverride');
  customFields.__procurementStartOverride = val('procurementStartOverride');
  customFields.__procurementFinishOverride = val('procurementFinishOverride');
  if (customLabel || customValue) customFields[customLabel || 'Campo adicional'] = customValue;
  return {
    fabricationStartOverride: val('fabricationStartOverride'),
    boilermakerFinishOverride: val('boilermakerFinishOverride'),
    weldingFinishOverride: val('weldingFinishOverride'),
    inspectionFinishOverride: val('inspectionFinishOverride'),
    thFinishOverride: val('thFinishOverride'),
    coatingFinishOverride: val('coatingFinishOverride'),
    projectFinishOverride: val('projectFinishOverride'),
    executiveStatus: val('executiveStatus'),
    executiveNote: val('executiveNote'),
    delayReason: val('delayReason'),
    customFields,
  };
}

async function saveClientBspOverride(project, form) {
  if (!project || !form || !canManageClientBspPanel(project)) return;
  state.clientBspOverrides.feedback = 'Salvando ajustes executivos...';
  const feedbackEl = document.getElementById('client-pm-editor-feedback');
  if (feedbackEl) feedbackEl.textContent = state.clientBspOverrides.feedback;
  try {
    const response = await fetch('/api/client-bsp-overrides', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        projectRowId: getClientBspOverrideProjectRowId(project),
        projectNumber: getClientBspOverrideProjectNumber(project),
        projectDisplay: project.projectDisplay || '',
        clientName: getProjectClientLabel(project),
        vessel: getProjectVesselLabel(project),
        pm: project.pm || '',
        ...collectClientBspOverrideForm(form, project),
      }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao salvar ajustes.');
    const next = normalizeClientBspOverride(data.override);
    if (!next) throw new Error('Ajuste salvo, mas a API não retornou os dados atualizados. Atualize a tela para conferir.');
    const items = state.clientBspOverrides.items.filter((item) => String(item.projectRowId) !== String(next.projectRowId));
    items.unshift(next);
    setClientBspOverrides(items);
    state.clientBspOverrides.feedback = 'Ajustes salvos com sucesso.';
    state.clientBspOverrides.editingProjectId = null;
    openClientBspExecutive(project);
    renderClientDashboard();
    if (modalEl && !modalEl.classList.contains('hidden')) renderModal(project);
  } catch (error) {
    state.clientBspOverrides.feedback = getClientSafeOverrideFeedback(error?.message || 'Falha ao salvar ajustes.');
    if (feedbackEl) feedbackEl.textContent = state.clientBspOverrides.feedback;
  }
}

async function clearClientBspOverride(project) {
  const override = getClientBspOverride(project);
  if (!project || !override?.id || !canManageClientBspPanel(project)) return;
  state.clientBspOverrides.feedback = 'Removendo ajuste executivo...';
  const feedbackEl = document.getElementById('client-pm-editor-feedback');
  if (feedbackEl) feedbackEl.textContent = state.clientBspOverrides.feedback;
  try {
    const response = await fetch(`/api/client-bsp-overrides?id=${encodeURIComponent(override.id)}&projectRowId=${encodeURIComponent(getClientBspOverrideProjectRowId(project))}`, {
      method: 'DELETE',
      credentials: 'same-origin',
      cache: 'no-store',
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao remover ajuste.');
    setClientBspOverrides(state.clientBspOverrides.items.filter((item) => String(item.id) !== String(override.id)));
    state.clientBspOverrides.feedback = '';
    state.clientBspOverrides.editingProjectId = null;
    openClientBspExecutive(project);
    renderClientDashboard();
    if (modalEl && !modalEl.classList.contains('hidden')) renderModal(project);
  } catch (error) {
    state.clientBspOverrides.feedback = getClientSafeOverrideFeedback(error?.message || 'Falha ao remover ajuste.');
    if (feedbackEl) feedbackEl.textContent = state.clientBspOverrides.feedback;
  }
}

function ensureClientBspExecutiveModalEl() {
  let modal = document.getElementById('client-bsp-executive-modal');
  if (modal) return modal;
  modal = document.createElement('div');
  modal.id = 'client-bsp-executive-modal';
  modal.className = 'client-exec-modal hidden';
  modal.innerHTML = `
    <div class="client-exec-backdrop" data-client-exec-close></div>
    <section class="client-exec-shell" role="dialog" aria-modal="true" aria-label="Visão Executiva da BSP">
      <button type="button" class="client-exec-close" data-client-exec-close>×</button>
      <div id="client-bsp-executive-content"></div>
    </section>
  `;
  document.body.appendChild(modal);
  modal.addEventListener('click', (event) => {
    const spoolPanelRow = event.target.closest('[data-client-spool-panel]');
    if (spoolPanelRow) {
      event.preventDefault();
      event.stopPropagation();
      const projectId = spoolPanelRow.dataset.clientSpoolProjectId || state.clientBspOverrides.activeExecutiveProjectId;
      const project = state.projects.find((item) => String(item.rowId) === String(projectId));
      const spool = findClientSpoolByPanelKey(project, spoolPanelRow.dataset.clientSpoolPanel || '');
      if (project && spool) openClientSpoolIndividualPanel(project, spool);
      return;
    }
    const projectPanelRow = event.target.closest('[data-client-project-panel]');
    if (projectPanelRow) {
      event.preventDefault();
      event.stopPropagation();
      const project = state.projects.find((item) => String(item.rowId) === String(projectPanelRow.dataset.clientProjectPanel));
      if (project) openClientSpoolIndividualPanel(project, null);
      return;
    }
    const pdfButton = event.target.closest('[data-client-download-pdf]');
    if (pdfButton) {
      event.preventDefault();
      event.stopPropagation();
      handleClientExecutivePdfDownload(pdfButton);
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
    const editButton = event.target.closest('[data-client-bsp-edit]');
    if (editButton) {
      const project = state.projects.find((item) => String(item.rowId) === String(editButton.dataset.clientBspEdit));
      if (project && canManageClientBspPanel(project)) {
        openClientBspExecutiveForPmEdit(project);
      }
      return;
    }
    if (event.target.closest('[data-client-bsp-edit-cancel]')) {
      const project = state.projects.find((item) => String(item.rowId) === String(state.clientBspOverrides.activeExecutiveProjectId));
      state.clientBspOverrides.editingProjectId = null;
      if (project) openClientBspExecutive(project);
      return;
    }
    if (event.target.closest('[data-client-bsp-clear-override]')) {
      const project = state.projects.find((item) => String(item.rowId) === String(state.clientBspOverrides.activeExecutiveProjectId));
      if (project) clearClientBspOverride(project);
      return;
    }
    const uploadImageButton = event.target.closest('[data-client-bsp-upload-image]');
    if (uploadImageButton) {
      event.preventDefault();
      event.stopPropagation();
      const rowId = uploadImageButton.dataset.clientBspUploadImage;
      const project = state.projects.find((item) => String(item.rowId) === String(rowId));
      if (!project || !canManageClientBspPanel(project)) {
        alert('A importação de imagem é restrita ao administrador e ao PM responsável pela BSP.');
        return;
      }
      openClientBspUploadImageModal(rowId);
      return;
    }
    // Novo botão para carregar imagens da BSP. Ao clicar, abre um modal
    // interno com as imagens associadas à BSP. Não abre uma nova guia,
    // pois o carregamento das imagens é assíncrono e ocorre somente
    // quando o usuário clicar no botão. O rowId é passado para a
    // função que monta e exibe o modal.
    const imagesButton = event.target.closest('[data-client-bsp-images]');
    if (imagesButton) {
      event.preventDefault();
      event.stopPropagation();
      const rowId = imagesButton.dataset.clientBspImages;
      // Prepare row numbers and row IDs of spools for this BSP.  The
      // attachments may be located on child spool rows as well, so we
      // include them when present.  If the sheet ID of the project is
      // available via project.sheetId (not currently defined), it can
      // be passed to the modal as well.
      let spoolRowNumbers = [];
      let spoolRowIds = [];
      let sheetIdForProject = null;
      try {
        const proj = state.projects && state.projects.find((item) => String(item.rowId) === String(rowId));
        if (proj) {
          if (Array.isArray(proj.spools)) {
            for (const sp of proj.spools) {
              if (sp && sp.rowNumber != null) spoolRowNumbers.push(sp.rowNumber);
              if (sp && sp.rowId != null) spoolRowIds.push(sp.rowId);
            }
          }
          // If the project object ever includes a sheetId field in the
          // future, forward it to the modal.  This remains null for now.
          if (proj.sheetId) sheetIdForProject = proj.sheetId;
        }
      } catch (err) {
        console.error(err);
      }
      openClientBspImagesModal(rowId, spoolRowNumbers, spoolRowIds, sheetIdForProject);
      return;
    }
    if (event.target.closest('[data-client-exec-close]')) closeClientBspExecutive();
  });
  modal.addEventListener('submit', (event) => {
    const form = event.target.closest('[data-client-bsp-override-form]');
    if (!form) return;
    event.preventDefault();
    const project = state.projects.find((item) => String(item.rowId) === String(form.dataset.clientBspOverrideForm));
    if (project) saveClientBspOverride(project, form);
  });
  return modal;
}

function closeClientBspExecutive() {
  const modal = document.getElementById('client-bsp-executive-modal');
  if (modal) modal.classList.add('hidden');
  document.body.classList.remove('client-exec-open');
  state.clientBspOverrides.activeExecutiveProjectId = null;
  state.clientBspOverrides.editingProjectId = null;
}

function getClientMacroProjects(projects = state.projects) {
  return (Array.isArray(projects) ? projects : []).filter(Boolean);
}

function getClientMacroTotals(projects = state.projects) {
  const list = getClientMacroProjects(projects);
  return list.reduce((acc, project) => {
    const totalTags = getProjectItemCount(project);
    const completedTags = getClientCompletedTags(project);
    const weight = Number(project.kilos || 0);
    const welded = Number(project.weldedWeightKg || 0);
    acc.bsps += 1;
    acc.tags += totalTags;
    acc.completedTags += completedTags;
    acc.remainingTags += Math.max(0, totalTags - completedTags);
    acc.weight += weight;
    acc.welded += welded;
    acc.pending += Math.max(0, weight - welded);
    acc.m2 += Number(project.m2Painting || 0);
    acc.progress += getClientOverallProgress(project);
    acc.fabrication += getClientFabricationProgress(project);
    acc.planned += getClientPlannedToday(project);
    return acc;
  }, { bsps: 0, tags: 0, completedTags: 0, remainingTags: 0, weight: 0, welded: 0, pending: 0, m2: 0, progress: 0, fabrication: 0, planned: 0 });
}

function getClientMacroAverage(value, count) {
  return count ? clampClientPercent(value / count) : 0;
}

function getClientMacroDateRange(projects = state.projects) {
  const list = getClientMacroProjects(projects);
  const starts = [];
  const finishes = [];
  for (const project of list) {
    const start = getClientAnalyticStartDate(project);
    const finish = getClientAnalyticFinishDate(project);
    if (start) starts.push(start);
    if (finish) finishes.push(finish);
  }
  const start = starts.length ? new Date(Math.min(...starts.map((date) => date.getTime()))) : getCurrentBrazilDate();
  const finish = finishes.length ? new Date(Math.max(...finishes.map((date) => date.getTime()))) : addUtcDays(start, 120);
  return { start, finish: finish <= start ? addUtcDays(start, 30) : finish };
}

function getClientMacroProductionStages(projects = state.projects) {
  const list = getClientMacroProjects(projects);
  const template = [
    { key: 'engineering', label: 'Engineering / Drawing', weight: 15 },
    { key: 'procurement', label: 'Procurement', weight: 15 },
    { key: 'fabrication', label: 'Fabrication', weight: 65 },
    { key: 'package', label: 'Package / Delivery', weight: 5 },
  ];
  if (!list.length) return template.map((stage) => ({ ...stage, percent: 0 }));
  return template.map((stage) => {
    const total = list.reduce((sum, project) => {
      const match = getClientProductionStages(project).find((item) => item.key === stage.key);
      return sum + Number(match?.percent || 0);
    }, 0);
    return { ...stage, percent: clampClientPercent(total / list.length) };
  });
}

function getClientMacroPlannedToday(projects = state.projects) {
  const totals = getClientMacroTotals(projects);
  return getClientMacroAverage(totals.planned, totals.bsps);
}

function getClientMacroOverallProgress(projects = state.projects) {
  const totals = getClientMacroTotals(projects);
  return getClientMacroAverage(totals.progress, totals.bsps);
}

function getClientMacroFabricationProgress(projects = state.projects) {
  const totals = getClientMacroTotals(projects);
  return getClientMacroAverage(totals.fabrication, totals.bsps);
}

function buildClientMacroSCurveData(projects = state.projects) {
  const { start, finish } = getClientMacroDateRange(projects);
  const duration = clientDaysBetween(start, finish);
  const step = Math.max(1, Math.ceil(duration / 14));
  const today = getCurrentBrazilDate();
  const actualNow = getClientMacroOverallProgress(projects);
  const plannedToday = getClientMacroPlannedToday(projects);
  const points = [];
  for (let day = 0; day <= duration; day += step) {
    const date = addUtcDays(start, day);
    const ratio = day / duration;
    const planned = clientSchedulePlannedPercent(ratio);
    let actual = null;
    if (date <= today) {
      actual = plannedToday > 0 ? clampClientPercent((planned / plannedToday) * actualNow) : actualNow;
    }
    points.push({ date, planned, actual });
  }
  if (points[points.length - 1]?.date < finish) points.push({ date: finish, planned: 100, actual: finish <= today ? actualNow : null });
  return points;
}

function renderClientMacroSCurveSvg(projects = state.projects) {
  const points = buildClientMacroSCurveData(projects);
  const width = getClientSCurveSvgWidth(points);
  const height = 260;
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const plannedPath = clientSvgPolyline(points, width, height, (point) => point.planned);
  // v36.43: curva macro/carteira/unidade não possui uma única BSP para calcular desvio.
  // A versão anterior chamava getClientSCurveDelayInfo(project), mas `project` não existe
  // neste escopo e quebrava a abertura da visão executiva da carteira/unidade.
  const delayInfo = null;
  const actualSegments = splitClientSCurveActualSegments(points, delayInfo);
  const actualPath = clientSvgPolyline(actualSegments.normal, width, height, (point) => point.actual, points);
  const delayActualPath = clientSvgPolyline(actualSegments.delayed, width, height, (point) => point.actual, points);
  const grid = [0, 25, 50, 75, 100].map((value) => {
    const y = pad.top + (1 - value / 100) * innerH;
    return `<line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" class="client-chart-grid" /><text x="8" y="${y + 4}" class="client-chart-label">${value}%</text>`;
  }).join('');
  const first = points[0]?.date ? clientFormatDateValue(points[0].date) : '';
  const mid = points[Math.floor(points.length / 2)]?.date ? clientFormatDateValue(points[Math.floor(points.length / 2)].date) : '';
  const last = points[points.length - 1]?.date ? clientFormatDateValue(points[points.length - 1].date) : '';
  const actualCircle = (() => {
    const lastActual = points.filter((point) => point.actual != null).pop();
    if (!lastActual) return '';
    const x = clientChartX(lastActual, points, width, pad);
    const y = clientChartY(lastActual.actual, height, pad);
    const isDelayed = delayInfo?.start && parseDateObject(lastActual.date) && parseDateObject(lastActual.date) >= parseDateObject(delayInfo.start);
    return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="5" class="${isDelayed ? 'client-chart-dot client-chart-dot--delay' : 'client-chart-dot'}" />`;
  })();
  const hoverTargets = buildClientChartHoverTargets(points, width, height);
  return wrapClientSCurveSvg(`
    <svg class="client-scurve-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Curva S macro planejado versus realizado">
      ${grid}
      <line x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <path d="${plannedPath}" class="client-chart-planned" />
      ${actualPath ? `<path d="${actualPath}" class="client-chart-actual" />` : ''}
      ${delayActualPath ? `<path d="${delayActualPath}" class="client-chart-actual-delay" />` : ''}
      ${actualCircle}
      ${hoverTargets}
      <text x="${pad.left}" y="${height - 12}" class="client-chart-date">${escapeHtml(first)}</text>
      <text x="${pad.left + innerW / 2 - 38}" y="${height - 12}" class="client-chart-date">${escapeHtml(mid)}</text>
      <text x="${width - pad.right - 72}" y="${height - 12}" class="client-chart-date">${escapeHtml(last)}</text>
    </svg>
  `, width);
}

function getClientMacroAttentionPoints(projects = state.projects) {
  const list = getClientMacroProjects(projects);
  const actual = getClientMacroOverallProgress(list);
  const planned = getClientMacroPlannedToday(list);
  const points = [];
  if (planned - actual >= 10) points.push(`Carteira realizada em ${formatPercent(actual)} contra planejado ${formatPercent(planned)}: desvio macro de ${formatPercent(planned - actual)}.`);
  const notStarted = list.filter((project) => getClientOverallProgress(project) <= 0).length;
  const finished = list.filter((project) => getClientOverallProgress(project) >= 99.9 || isProjectFinishedForTotal(project)).length;
  const delayed = list.filter((project) => getClientPlannedToday(project) - getClientOverallProgress(project) >= 10).length;
  if (notStarted) points.push(`${formatNumber(notStarted)} BSP(s) ainda sem avanço registrado.`);
  if (delayed) points.push(`${formatNumber(delayed)} BSP(s) com desvio maior ou igual a 10% em relação ao planejado.`);
  if (finished) points.push(`${formatNumber(finished)} BSP(s) finalizada(s) na carteira.`);
  if (!points.length) points.push('Carteira sem ponto crítico automático identificado neste momento.');
  return points;
}

function getClientProjectVisualState(project) {
  const progress = getClientOverallProgress(project);
  if (hasClientIncompleteProductionEvidence(project)) return progress <= 0 ? 'not-started' : 'in-progress';
  if (progress >= 99.9 || isProjectFinishedForTotal(project)) return 'completed';
  if (progress <= 0) return 'not-started';
  return 'in-progress';
}

function renderClientMacroProjectRows(projects = state.projects) {
  const list = getClientMacroProjects(projects)
    .sort((a, b) => {
      const aDone = getClientProjectVisualState(a) === 'completed';
      const bDone = getClientProjectVisualState(b) === 'completed';
      if (aDone !== bDone) return aDone ? 1 : -1;
      return getClientOverallProgress(a) - getClientOverallProgress(b);
    })
    .slice(0, 180);
  if (!list.length) return '<tr><td colspan="8" class="loading-cell">Nenhuma BSP encontrada para este cliente.</td></tr>';
  return list.map((project) => {
    const status = getProjectStatusPresentation(project);
    const visualState = getClientProjectVisualState(project);
    return `<tr class="client-spool-row client-spool-row--${visualState} client-spool-row--clickable" data-client-project-panel="${escapeHtml(project.rowId || '')}" title="Clique para abrir o painel individual da obra em uma nova aba"><td><strong>${escapeHtml(getClientProjectDisplayCode(project))}</strong></td><td>${escapeHtml(getProjectVesselLabel(project) || '—')}</td><td>${formatNumber(getProjectItemCount(project))}</td><td>${formatNumber(project.kilos, 0)} kg</td><td>${formatNumber(project.weldedWeightKg, 0)} kg</td><td><span class="client-spool-chip client-spool-chip--${visualState}">${escapeHtml(status.text)}</span></td><td><span class="client-spool-progress client-spool-progress--${visualState}">${formatPercent(getClientOverallProgress(project))}</span></td><td>${escapeHtml(project.plannedFinishDate || '—')}</td></tr>`;
  }).join('');
}


function sanitizeClientReportFilenamePart(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  return text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^-+|-+$/g, '')
    .slice(0, 90);
}

function getClientReportPoLabel(project) {
  const values = [];
  if (Array.isArray(project?.customerPoList)) values.push(...project.customerPoList);
  values.push(project?.customerPo, project?.customerPoDisplay, project?.clientDisplayCode);
  const found = [];
  for (const value of values) {
    const raw = String(value || '').trim();
    if (!raw || /aguardando\s+po/i.test(raw)) continue;
    const explicitPo = raw.match(/(?:^|[-–—]\s*)POs?\s+(.+)$/i);
    const cleaned = (explicitPo ? explicitPo[1] : raw)
      .replace(/^POs?\s*/i, '')
      .replace(/\bPOs?\b/ig, '')
      .trim();
    const parts = cleaned.split(/[\/;,]+/).map((part) => part.trim()).filter(Boolean);
    for (const part of parts) {
      const safe = sanitizeClientReportFilenamePart(part);
      if (safe && !found.includes(safe)) found.push(safe);
    }
  }
  return found.slice(0, 3).join('_');
}

function buildClientExecutivePdfFileName(project = null) {
  const po = getClientReportPoLabel(project);
  if (po) return `relatorio_PO_${po}.pdf`;
  return 'relatorio.pdf';
}

function buildClientUnitExecutivePdfFileName(unitLabel = '') {
  const client = sanitizeClientReportFilenamePart(getClientPortalName());
  const unit = sanitizeClientReportFilenamePart(unitLabel || 'unidade');
  if (client && unit) return `relatorio_${client}_UNIDADE_${unit}.pdf`;
  if (unit) return `relatorio_UNIDADE_${unit}.pdf`;
  return 'relatorio.pdf';
}

function drawClientPdfFooter(doc) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8);
  doc.setTextColor(86, 107, 124);
  doc.text(`STEP • Página ${doc.internal.getNumberOfPages()}`, pageWidth - 14, pageHeight - 7, { align: 'right' });
}

async function drawClientPdfHeader(doc, title, subtitle, metaLine = '') {
  const pageWidth = doc.internal.pageSize.getWidth();
  const generatedAt = new Date().toLocaleString('pt-BR');
  const logoDataUrl = await loadImageAsDataUrl('./assets/step-logo.png');
  doc.setFillColor(238, 248, 255);
  doc.rect(0, 0, pageWidth, 34, 'F');
  if (logoDataUrl) {
    try { doc.addImage(logoDataUrl, 'PNG', 12, 8, 34, 11); } catch (error) { console.warn('Não foi possível renderizar a logo no PDF.', error); }
  }
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.setTextColor(11, 55, 97);
  doc.text(title, 52, 13);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(61, 100, 127);
  doc.text(subtitle, 52, 20);
  if (metaLine) doc.text(metaLine, 52, 27);
  doc.text(`Gerado em: ${generatedAt}`, pageWidth - 14, 13, { align: 'right' });
  drawClientPdfFooter(doc);
}

function drawClientPdfKpi(doc, x, y, w, h, label, value) {
  doc.setDrawColor(193, 216, 232);
  doc.setFillColor(255, 255, 255);
  doc.roundedRect(x, y, w, h, 2, 2, 'FD');
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7.5);
  doc.setTextColor(95, 127, 149);
  doc.text(String(label || ''), x + 3, y + 5);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.setTextColor(8, 46, 81);
  const lines = doc.splitTextToSize(String(value || '—'), w - 6);
  doc.text(lines.slice(0, 2), x + 3, y + 12);
}

function drawClientPdfKpiGrid(doc, kpis, yStart, options = {}) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = options.margin || 12;
  const gap = options.gap || 4;
  const columns = options.columns || 4;
  const boxH = options.boxH || 18;
  const boxW = (pageWidth - margin * 2 - gap * (columns - 1)) / columns;
  kpis.forEach((item, index) => {
    const col = index % columns;
    const row = Math.floor(index / columns);
    drawClientPdfKpi(doc, margin + col * (boxW + gap), yStart + row * (boxH + gap), boxW, boxH, item[0], item[1]);
  });
  return yStart + Math.ceil(kpis.length / columns) * boxH + Math.max(0, Math.ceil(kpis.length / columns) - 1) * gap;
}

function drawClientPdfSectionTitle(doc, title, x, y) {
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(11, 55, 97);
  doc.text(String(title || ''), x, y);
}

function drawClientPdfProgressBar(doc, x, y, w, label, percent, tone = 'normal') {
  const p = clampClientPercent(percent);
  const fill = tone === 'planned' ? [239, 193, 79] : [11, 155, 122];
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(11, 55, 97);
  doc.text(String(label || '—'), x, y);
  doc.setDrawColor(210, 220, 226);
  doc.setFillColor(232, 239, 244);
  doc.roundedRect(x + 55, y - 4, w, 4, 1.5, 1.5, 'FD');
  doc.setFillColor(...fill);
  doc.roundedRect(x + 55, y - 4, Math.max(1, (w * p) / 100), 4, 1.5, 1.5, 'F');
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.text(formatPercent(p), x + 57 + w, y);
}

function drawClientPdfSCurve(doc, points, x, y, width, height, title) {
  const safePoints = Array.isArray(points) ? points : [];
  const padLeft = 12;
  const padBottom = 12;
  const chartX = x + padLeft;
  const chartY = y + 8;
  const chartW = width - padLeft - 4;
  const chartH = height - padBottom - 12;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(10);
  doc.setTextColor(11, 55, 97);
  doc.text(title, x, y);
  doc.setDrawColor(214, 224, 232);
  doc.setLineWidth(0.1);
  [0, 25, 50, 75, 100].forEach((value) => {
    const gy = chartY + (1 - value / 100) * chartH;
    doc.line(chartX, gy, chartX + chartW, gy);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(95, 127, 149);
    doc.text(`${value}%`, x, gy + 2);
  });
  doc.setDrawColor(110, 135, 153);
  doc.line(chartX, chartY, chartX, chartY + chartH);
  doc.line(chartX, chartY + chartH, chartX + chartW, chartY + chartH);

  const toXY = (point, index) => {
    const px = chartX + (index / Math.max(1, safePoints.length - 1)) * chartW;
    const value = clampClientPercent(point);
    const py = chartY + (1 - value / 100) * chartH;
    return { px, py };
  };
  const drawSeries = (getter, color, dashed = false) => {
    const series = safePoints.map((point, index) => ({ value: getter(point), index })).filter((item) => item.value != null);
    if (series.length < 2) return;
    doc.setDrawColor(...color);
    doc.setLineWidth(0.7);
    if (dashed && doc.setLineDashPattern) doc.setLineDashPattern([2, 1.5], 0);
    for (let i = 1; i < series.length; i += 1) {
      const a = toXY(series[i - 1].value, series[i - 1].index);
      const b = toXY(series[i].value, series[i].index);
      doc.line(a.px, a.py, b.px, b.py);
    }
    if (doc.setLineDashPattern) doc.setLineDashPattern([], 0);
  };
  drawSeries((point) => point.planned, [239, 193, 79], true);
  drawSeries((point) => point.actual, [11, 155, 122], false);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(7);
  doc.setTextColor(61, 100, 127);
  const first = safePoints[0]?.date ? clientFormatDateValue(safePoints[0].date) : '';
  const last = safePoints[safePoints.length - 1]?.date ? clientFormatDateValue(safePoints[safePoints.length - 1].date) : '';
  if (first) doc.text(first, chartX, y + height - 2);
  if (last) doc.text(last, chartX + chartW, y + height - 2, { align: 'right' });
  doc.setFillColor(239, 193, 79);
  doc.circle(x + width - 44, y + 2, 1.3, 'F');
  doc.text('Planejado', x + width - 40, y + 3);
  doc.setFillColor(11, 155, 122);
  doc.circle(x + width - 20, y + 2, 1.3, 'F');
  doc.text('Realizado', x + width - 16, y + 3);
}

function clientReportStageLabel(key) {
  return (state.meta?.stageOrder || []).find((stage) => stage.key === key)?.label || key;
}

function handleClientExecutivePdfDownload(button) {
  const type = button?.dataset?.clientReportType || 'project';
  const projectId = button?.dataset?.clientReportProjectId || '';
  const unitKey = button?.dataset?.clientReportUnitKey || '';
  const originalText = button.textContent;
  button.disabled = true;
  button.textContent = 'Gerando PDF...';
  let task;
  if (type === 'macro') {
    task = downloadClientMacroExecutivePdf();
  } else if (type === 'unit') {
    task = downloadClientUnitExecutivePdf(unitKey);
  } else {
    task = downloadClientBspExecutivePdf(state.projects.find((item) => String(item.rowId) === String(projectId)));
  }
  Promise.resolve(task)
    .catch((error) => {
      console.error('Falha ao gerar PDF executivo.', error);
      window.alert(error?.message || 'Não foi possível gerar o PDF. Atualize a página e tente novamente.');
    })
    .finally(() => {
      button.disabled = false;
      button.textContent = originalText || 'Baixar PDF';
    });
}

async function downloadClientBspExecutivePdf(project) {
  if (!project) throw new Error('Nenhuma BSP selecionada para gerar o relatório.');
  const jsPdfApi = window.jspdf?.jsPDF;
  if (!jsPdfApi) throw new Error('A biblioteca de PDF não foi carregada.');
  const doc = new jsPdfApi({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const contentX = 12;
  const contentW = pageWidth - 24;
  const status = getProjectStatusPresentation(project);
  const stages = getClientProductionStages(project);
  const overall = getClientOverallProgress(project);
  const fabrication = getClientFabricationProgress(project);
  const plannedToday = getClientPlannedToday(project);
  const totalTags = getProjectItemCount(project);
  const completedTags = getClientCompletedTags(project);
  const remainingTags = Math.max(0, totalTags - completedTags);
  const weight = Number(project.kilos || 0);
  const welded = Number(project.weldedWeightKg || 0);
  const pending = Math.max(0, weight - welded);
  const projectCode = getClientProjectDisplayCode(project);
  const subtitle = `${getProjectClientLabel(project)} • ${getProjectVesselLabel(project)} • ${status.text}`;
  const metaLine = `BSP/PO: ${projectCode}`;

  await drawClientPdfHeader(doc, 'Relatório Operacional STEP', subtitle, metaLine);

  const kpis = [
    ['Progresso geral', formatPercent(overall)],
    ['Fabricação', formatPercent(fabrication)],
    ['Planejado hoje', formatPercent(plannedToday)],
    ['Peso programado', `${formatNumber(weight, 0)} kg`],
    ['Peso soldado', `${formatNumber(welded, 0)} kg`],
    ['Peso restante', `${formatNumber(pending, 0)} kg`],
    ['Tags totais', formatNumber(totalTags)],
    ['Tags restantes', formatNumber(remainingTags)],
  ];
  const afterKpisY = drawClientPdfKpiGrid(doc, kpis, 42, { columns: 4, boxH: 18 });

  let y = afterKpisY + 12;
  drawClientPdfSectionTitle(doc, 'Indicadores dos gráficos', contentX, y);
  y += 10;
  drawClientPdfProgressBar(doc, contentX, y, 112, 'Overall Progress', overall);
  y += 9;
  drawClientPdfProgressBar(doc, contentX, y, 112, 'Fabrication Progress', fabrication);
  y += 9;
  drawClientPdfProgressBar(doc, contentX, y, 112, 'Planejado hoje', plannedToday, 'planned');
  y += 13;
  stages.forEach((stage) => {
    drawClientPdfProgressBar(doc, contentX, y, 112, stage.label, stage.percent);
    y += 8;
  });

  const curveY = Math.max(y + 10, 188);
  drawClientPdfSCurve(doc, buildClientSCurveData(project), contentX, curveY, contentW, 82, 'Curva S | Planejado x Realizado');

  doc.addPage('a4', 'portrait');
  await drawClientPdfHeader(doc, 'Relatório Operacional STEP', subtitle, metaLine);
  drawClientPdfSectionTitle(doc, 'Processos da BSP por etapa', contentX, 42);
  const detailStageKeys = getClientDetailStageKeys(project);
  const stageRows = detailStageKeys.map((key) => [clientReportStageLabel(key), formatPercent(getClientStageValue(project, [key]))]);
  doc.autoTable({
    startY: 48,
    head: [['Processo', '%']],
    body: stageRows,
    tableWidth: contentW,
    styles: { font: 'helvetica', fontSize: 8, cellPadding: 2, lineColor: [220, 228, 236], lineWidth: 0.1, overflow: 'linebreak' },
    headStyles: { fillColor: [22, 83, 126], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 150 }, 1: { cellWidth: 36, halign: 'center' } },
    margin: { left: contentX, right: contentX },
    didDrawPage: () => drawClientPdfFooter(doc),
  });

  const scheduleStartY = Math.min((doc.lastAutoTable?.finalY || 104) + 12, 132);
  drawClientPdfSectionTitle(doc, 'Schedule Executivo', contentX, scheduleStartY);
  const scheduleRows = buildClientExecutiveSchedule(project).map((row) => [
    row.type === 'group' ? row.label : `  ${row.label}`,
    row.type === 'group' ? 'Grupo' : 'Etapa',
    formatPercent(row.progress),
    String(row.duration || '—'),
    formatClientDateShort(row.start),
    formatClientDateShort(row.finish),
  ]);
  doc.autoTable({
    startY: scheduleStartY + 6,
    head: [['Schedule Executivo', 'Tipo', '%', 'Dias úteis', 'Início', 'Fim']],
    body: scheduleRows,
    tableWidth: contentW,
    margin: { left: contentX, right: contentX },
    styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 1.7, lineColor: [220, 228, 236], lineWidth: 0.1, overflow: 'linebreak' },
    headStyles: { fillColor: [22, 83, 126], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 60 }, 1: { cellWidth: 22 }, 2: { cellWidth: 20, halign: 'center' }, 3: { cellWidth: 28, halign: 'center' }, 4: { cellWidth: 28 }, 5: { cellWidth: 28 } },
    didDrawPage: () => drawClientPdfFooter(doc),
  });

  doc.addPage('a4', 'portrait');
  await drawClientPdfHeader(doc, 'Relatório Operacional STEP', subtitle, metaLine);
  drawClientPdfSectionTitle(doc, 'Detalhamento das Tags / ISOs', contentX, 42);
  const spools = Array.isArray(project.spools) ? project.spools : [];
  const spoolRows = spools.slice(0, 180).map((spool) => [
    String(spool.iso || '—'),
    String(spool.description || '—'),
    String(spool.currentStatus || spool.stage || uiStateLabel(spool.uiState) || '—'),
    String(spool.currentSector || spool.operationalSector || '—'),
    formatPercent(spool.overallProgress),
    `${formatNumber(spool.kilos, 2)} kg`,
  ]);
  doc.autoTable({
    startY: 48,
    head: [['Tag/ISO', 'Descrição', 'Status', 'Etapa', '%', 'Peso']],
    body: spoolRows.length ? spoolRows : [['—', 'Nenhuma tag detalhada encontrada para esta BSP.', '—', '—', '—', '—']],
    styles: { font: 'helvetica', fontSize: 6.8, cellPadding: 1.35, overflow: 'linebreak', valign: 'middle', lineColor: [220, 228, 236], lineWidth: 0.1 },
    headStyles: { fillColor: [22, 83, 126], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 42 }, 1: { cellWidth: 40 }, 2: { cellWidth: 36 }, 3: { cellWidth: 30 }, 4: { cellWidth: 16, halign: 'center' }, 5: { cellWidth: 22, halign: 'right' } },
    margin: { left: contentX, right: contentX },
    didDrawPage: () => drawClientPdfFooter(doc),
  });

  doc.save(buildClientExecutivePdfFileName(project));
}

async function downloadClientUnitExecutivePdf(unitKey = '') {
  const group = getClientVesselGroupByKey(unitKey, state.projects);
  if (!group) throw new Error('Nenhuma unidade encontrada para gerar o relatório.');
  return downloadClientMacroExecutivePdf(group.projects, { scope: 'unit', unitKey: group.key, unitLabel: group.label });
}

async function downloadClientMacroExecutivePdf(projects = state.projects, options = {}) {
  const jsPdfApi = window.jspdf?.jsPDF;
  if (!jsPdfApi) throw new Error('A biblioteca de PDF não foi carregada.');
  const list = getClientMacroProjects(projects);
  const unitLabel = String(options.unitLabel || '').trim();
  const isUnitScope = options.scope === 'unit' && unitLabel;
  const doc = new jsPdfApi({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const contentX = 12;
  const contentW = pageWidth - 24;
  const totals = getClientMacroTotals(list);
  const overall = getClientMacroOverallProgress(list);
  const fabrication = getClientMacroFabricationProgress(list);
  const plannedToday = getClientMacroPlannedToday(list);
  const stages = getClientMacroProductionStages(list);
  const range = getClientMacroDateRange(list);
  const reportSubtitle = isUnitScope ? `${getClientPortalName()} • Unidade ${unitLabel}` : getClientPortalName();
  const reportScopeLine = isUnitScope
    ? `Relatório por unidade • ${formatNumber(totals.bsps)} BSP(s) • ${formatNumber(totals.tags)} tag(s)`
    : `Carteira do cliente • ${formatNumber(totals.bsps)} BSP(s) • ${formatNumber(totals.tags)} tag(s)`;
  await drawClientPdfHeader(doc, 'Relatório Operacional STEP', reportSubtitle, reportScopeLine);
  const kpis = [
    ['BSPs', formatNumber(totals.bsps)],
    ['Tags totais', formatNumber(totals.tags)],
    ['Tags restantes', formatNumber(totals.remainingTags)],
    ['Progresso geral', formatPercent(overall)],
    ['Fabricação', formatPercent(fabrication)],
    ['Planejado hoje', formatPercent(plannedToday)],
    ['Peso programado', `${formatNumber(totals.weight, 0)} kg`],
    ['M² programada', formatNumber(totals.m2, 3)],
  ];
  const afterKpisY = drawClientPdfKpiGrid(doc, kpis, 42, { columns: 4, boxH: 18 });

  let y = afterKpisY + 12;
  drawClientPdfSectionTitle(doc, 'Indicadores dos gráficos', contentX, y);
  y += 10;
  drawClientPdfProgressBar(doc, contentX, y, 112, 'Overall Progress', overall);
  y += 10;
  drawClientPdfProgressBar(doc, contentX, y, 112, 'Fabrication Progress', fabrication);
  y += 12;
  stages.forEach((stage) => {
    drawClientPdfProgressBar(doc, contentX, y, 112, stage.label, stage.percent);
    y += 9;
  });
  const curveY = Math.max(y + 10, 188);
  drawClientPdfSCurve(doc, buildClientMacroSCurveData(list), contentX, curveY, contentW, 82, isUnitScope ? 'Curva S | Unidade' : 'Curva S | Carteira do Cliente');
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(61, 100, 127);
  doc.text(`Início macro: ${clientFormatDateValue(range.start) || '—'}  •  Término macro: ${clientFormatDateValue(range.finish) || '—'}`, contentX, 278);

  doc.addPage('a4', 'portrait');
  await drawClientPdfHeader(doc, 'Relatório Operacional STEP', reportSubtitle, isUnitScope ? `Detalhamento da unidade ${unitLabel}` : 'Detalhamento macro das BSPs');
  drawClientPdfSectionTitle(doc, isUnitScope ? `Detalhamento da unidade ${unitLabel}` : 'Detalhamento macro das BSPs', contentX, 42);
  const projectRows = getClientMacroProjects(list).map((project) => {
    const status = getProjectStatusPresentation(project);
    return [
      getClientProjectDisplayCode(project),
      getProjectVesselLabel(project) || '—',
      formatNumber(getProjectItemCount(project)),
      `${formatNumber(project.kilos, 0)} kg`,
      `${formatNumber(project.weldedWeightKg, 0)} kg`,
      status.text,
      formatPercent(getClientOverallProgress(project)),
      String(project.plannedFinishDate || '—'),
    ];
  });
  doc.autoTable({
    startY: 48,
    head: [['BSP / PO', 'Unidade', 'Tags', 'Peso', 'Soldado', 'Status', '% Geral', 'Término']],
    body: projectRows.length ? projectRows : [['—', '—', '—', '—', '—', 'Nenhuma BSP encontrada para este cliente.', '—', '—']],
    styles: { font: 'helvetica', fontSize: 6.6, cellPadding: 1.35, overflow: 'linebreak', valign: 'middle', lineColor: [220, 228, 236], lineWidth: 0.1 },
    headStyles: { fillColor: [22, 83, 126], textColor: [255, 255, 255], fontStyle: 'bold' },
    columnStyles: { 0: { cellWidth: 42 }, 1: { cellWidth: 24 }, 2: { cellWidth: 13, halign: 'center' }, 3: { cellWidth: 20, halign: 'right' }, 4: { cellWidth: 20, halign: 'right' }, 5: { cellWidth: 32 }, 6: { cellWidth: 16, halign: 'center' }, 7: { cellWidth: 22 } },
    margin: { left: contentX, right: contentX },
    didDrawPage: () => drawClientPdfFooter(doc),
  });
  doc.save(isUnitScope ? buildClientUnitExecutivePdfFileName(unitLabel) : 'relatorio.pdf');
}

function openClientMacroExecutive(projects = state.projects, options = {}) {
  if (!isClientUser()) return;
  const list = getClientMacroProjects(projects);
  const modal = ensureClientBspExecutiveModalEl();
  const content = modal.querySelector('#client-bsp-executive-content');
  if (!content) return;
  const unitLabel = String(options.unitLabel || '').trim();
  const unitKey = String(options.unitKey || '').trim();
  const isUnitScope = options.scope === 'unit' && unitLabel;
  const reportTitle = isUnitScope ? unitLabel : getClientPortalName();
  const reportKicker = isUnitScope ? 'Visão Executiva da Unidade' : 'Visão Executiva da Carteira';
  const reportIntro = isUnitScope
    ? `${getClientPortalName()} • Unidade ${unitLabel} • ${formatNumber(list.length)} BSP(s)`
    : `Carteira completa do cliente • ${formatNumber(list.length)} BSP(s)`;
  const reportButtonAttrs = isUnitScope
    ? `data-client-report-type="unit" data-client-report-unit-key="${escapeHtml(unitKey)}"`
    : 'data-client-report-type="macro"';
  const reportButtonText = isUnitScope ? 'Baixar PDF da unidade' : 'Baixar PDF';

  const totals = getClientMacroTotals(list);
  const overall = getClientMacroOverallProgress(list);
  const fabrication = getClientMacroFabricationProgress(list);
  const plannedToday = getClientMacroPlannedToday(list);
  const stages = getClientMacroProductionStages(list);
  const attention = getClientMacroAttentionPoints(list);
  const range = getClientMacroDateRange(list);
  const deviationPercent = getClientVisibleDeviationPercent(overall, plannedToday, range.finish);
  const vesselGroups = getClientVesselGroups(list);
  const finishedBsps = list.filter((project) => getClientOverallProgress(project) >= 99.9 || isProjectFinishedForTotal(project)).length;
  const pendingBsps = Math.max(0, totals.bsps - finishedBsps);
  const timeline = stages.map((stage) => ({ ...stage, state: stage.percent >= 100 ? 'done' : stage.percent > 0 ? 'active' : 'future' }));

  content.innerHTML = `
    <header class="client-exec-header client-exec-header--macro">
      <div>
        <p class="client-kicker">${escapeHtml(reportKicker)}</p>
        <h2>${escapeHtml(reportTitle)}</h2>
        <p>${escapeHtml(reportIntro)} • ${formatNumber(totals.tags)} tag(s)</p>
        <div class="client-exec-header-actions"><button class="client-exec-pdf-button" type="button" data-client-download-pdf ${reportButtonAttrs}>${escapeHtml(reportButtonText)}</button></div>
      </div>
      <div class="client-exec-dates">
        <span>Início macro: <strong>${escapeHtml(clientFormatDateValue(range.start) || '—')}</strong></span>
        <span>Término macro: <strong>${escapeHtml(clientFormatDateValue(range.finish) || '—')}</strong></span>
        <span>Planejado hoje: <strong>${formatPercent(plannedToday)}</strong></span>
        <span>Desvio: <strong>${formatPercent(deviationPercent)}</strong></span>
      </div>
    </header>

    <div class="client-exec-kpis">
      <article><span>BSPs</span><strong>${formatNumber(totals.bsps)}</strong></article>
      <article><span>Tags totais</span><strong>${formatNumber(totals.tags)}</strong></article>
      <article><span>Tags restantes</span><strong>${formatNumber(totals.remainingTags)}</strong></article>
      <article><span>Peso programado</span><strong>${formatNumber(totals.weight, 0)} kg</strong></article>
      <article><span>Peso soldado</span><strong>${formatNumber(totals.welded, 0)} kg</strong></article>
      <article><span>Peso restante</span><strong>${formatNumber(totals.pending, 0)} kg</strong></article>
      <article><span>M² programada</span><strong>${formatNumber(totals.m2, 3)}</strong></article>
      <article><span>BSPs pendentes</span><strong>${formatNumber(pendingBsps)}</strong></article>
    </div>

    <div class="client-exec-grid client-exec-grid--top">
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Overall Progress</h3><span>${isUnitScope ? 'Unidade consolidada + desvio' : 'Carteira geral + desvio'}</span></div>
        ${renderClientGauge(overall, isUnitScope ? 'unidade' : 'carteira', plannedToday, { note: 'Meta macro até hoje', deadlineDate: range.finish })}
      </section>
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Fabrication Progress</h3><span>${isUnitScope ? 'Fabricação ponderada da unidade' : 'Fabricação ponderada da carteira'}</span></div>
        ${renderClientGauge(fabrication, 'fabricação', plannedToday, { note: 'Meta macro até hoje', deadlineDate: range.finish })}
      </section>
      <section class="client-exec-card client-exec-card--bars">
        <div class="client-exec-card-head"><h3>Progress by Production Stage</h3><span>${isUnitScope ? 'Etapas principais da unidade' : 'Etapas principais da carteira'}</span></div>
        <div class="client-exec-bars">
          ${stages.map((stage) => `<div class="client-exec-bar-row"><span>${escapeHtml(stage.label)}</span><div><i style="width:${clampClientPercent(stage.percent)}%"></i></div><strong>${formatPercent(stage.percent)}</strong></div>`).join('')}
        </div>
      </section>
    </div>

    <div class="client-exec-grid client-exec-grid--main">
      <section class="client-exec-card client-exec-card--curve">
        <div class="client-exec-card-head"><h3>${isUnitScope ? 'Curva S | Unidade' : 'Curva S | Carteira do Cliente'}</h3><span>Planejado x realizado consolidado das BSPs</span></div>
        <div class="client-exec-legend"><span><i class="planned"></i> Planejado</span><span><i class="actual"></i> Realizado</span></div>
        ${renderClientMacroSCurveSvg(list)}
      </section>
      <aside class="client-exec-side">
        <section class="client-exec-card">
          <div class="client-exec-card-head"><h3>Resumo BSPs</h3><span>${isUnitScope ? 'Unidade consolidada' : 'Carteira consolidada'}</span></div>
          <div class="client-exec-mini-table">
            <div><span>Total</span><strong>${formatNumber(totals.bsps)}</strong></div>
            <div><span>Finalizadas</span><strong>${formatNumber(finishedBsps)}</strong></div>
            <div><span>Pendentes</span><strong>${formatNumber(pendingBsps)}</strong></div>
          </div>
        </section>
        <section class="client-exec-card">
          <div class="client-exec-card-head"><h3>${isUnitScope ? 'Unidade' : 'Unidades'}</h3><span>${isUnitScope ? 'Escopo do relatório' : 'BSPs por vessel'}</span></div>
          <div class="client-exec-mini-table">
            ${isUnitScope
              ? `<div><span>${escapeHtml(unitLabel)}</span><strong>${formatNumber(totals.bsps)} BSP(s)</strong></div>`
              : (vesselGroups.slice(0, 8).map((group) => `<div><span>${escapeHtml(group.label)}</span><strong>${formatNumber(group.projects.length)} BSP(s)</strong></div>`).join('') || '<div><span>Sem unidade</span><strong>—</strong></div>')}
          </div>
        </section>
      </aside>
    </div>

    <section class="client-exec-card">
      <div class="client-exec-card-head"><h3>Timeline macro</h3><span>${isUnitScope ? 'Resumo das etapas da unidade' : 'Resumo das etapas da carteira'}</span></div>
      <div class="client-exec-timeline">
        ${timeline.map((item) => `<div class="client-exec-step is-${item.state}"><span></span><strong>${escapeHtml(item.label)}</strong><small>${formatPercent(item.percent)}</small></div>`).join('')}
      </div>
    </section>

    ${'' /* v37.73: Attention Points oculto no painel do cliente */}

    <section class="client-exec-card client-exec-process-detail">
      <div class="client-exec-card-head"><h3>${isUnitScope ? 'Detalhamento da unidade' : 'Detalhamento macro das BSPs'}</h3><span>Menor progresso primeiro; finalizadas no final</span></div>
      <div class="client-table-wrap client-table-wrap--compact client-exec-process-table">
        <table class="client-bsp-table"><thead><tr><th>BSP / PO</th><th>Unidade</th><th>Tags</th><th>Peso</th><th>Soldado</th><th>Status</th><th>% Geral</th><th>Término</th></tr></thead><tbody>
          ${renderClientMacroProjectRows(list)}
        </tbody></table>
      </div>
    </section>
  `;

  modal.classList.remove('hidden');
  document.body.classList.add('client-exec-open');
  if (options.scrollToEditor && canManageClientBspPanel(project)) {
    window.setTimeout(() => {
      const editor = modal.querySelector(`[data-client-pm-editor="${CSS.escape(String(project.rowId || ''))}"]`);
      if (editor) editor.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 80);
  }
}

function openClientBspExecutive(project, options = {}) {
  if (!project || !canOpenClientBspPanel(project)) return;
  state.clientBspOverrides.activeExecutiveProjectId = project.rowId;
  if (!options.keepEditing && String(state.clientBspOverrides.editingProjectId || '') !== String(project.rowId || '')) state.clientBspOverrides.editingProjectId = null;
  const modal = ensureClientBspExecutiveModalEl();
  const content = modal.querySelector('#client-bsp-executive-content');
  if (!content) return;

  const status = getProjectStatusPresentation(project);
  const stages = getClientProductionStages(project);
  const overall = getClientOverallProgress(project);
  const fabrication = getClientFabricationProgress(project);
  const plannedToday = getClientPlannedToday(project);
  const completedTags = getClientCompletedTags(project);
  const totalTags = getProjectItemCount(project);
  const remainingTags = Math.max(0, totalTags - completedTags);
  const weight = Number(project.kilos || 0);
  const welded = Number(project.weldedWeightKg || 0);
  const pending = Math.max(0, weight - welded);
  const timeline = getClientStageTimeline(project);
  const attention = getClientAttentionPoints(project);
  const startDate = clientFormatDateValue(getClientAnalyticStartDate(project));
  const finishDate = clientFormatDateValue(getClientAnalyticFinishDate(project));
  const shipmentDate = clientFormatDateValue(getProjectShipmentDate(project));
  const deviationPercent = getClientVisibleDeviationPercent(overall, plannedToday, finishDate);
  const stageValues = project.stageValues || {};
  const spools = Array.isArray(project.spools) ? project.spools : [];
  const detailStageKeys = getClientDetailStageKeys(project);

  content.innerHTML = `
    <header class="client-exec-header">
      <div>
        <p class="client-kicker">Visão Executiva da BSP</p>
        <h2>${escapeHtml(getClientProjectDisplayCode(project))}</h2>
        <p>${escapeHtml(getProjectClientLabel(project))} • ${escapeHtml(getProjectVesselLabel(project))} • <span class="cell-status cell-status--${status.state}">${escapeHtml(status.text)}</span></p>
        <div class="client-exec-header-actions">
          <button class="client-exec-pdf-button" type="button" data-client-download-pdf data-client-report-type="project" data-client-report-project-id="${escapeHtml(project.rowId)}">Baixar PDF</button>
          <button class="client-exec-pdf-button client-exec-report-button" type="button" data-client-download-report="${escapeHtml(project.rowId)}">Baixar Excel do Cronograma</button>
          ${canManageClientBspPanel(project) ? `<button class="client-exec-pdf-button client-exec-edit-button" type="button" data-client-bsp-edit="${escapeHtml(project.rowId)}">Editar datas / informações</button>` : ''}
          <!-- Botões de imagens da BSP: cliente apenas visualiza; importação fica restrita a admin/PM -->
          <button class="client-exec-pdf-button client-exec-images-button" type="button" data-client-bsp-images="${escapeHtml(project.rowId)}">Imagens</button>
          ${canManageClientBspPanel(project) ? `<button class="client-exec-pdf-button client-exec-upload-image-button" type="button" data-client-bsp-upload-image="${escapeHtml(project.rowId)}">Importar imagem</button>` : ''}
        </div>
      </div>
      <div class="client-exec-dates">
        <span>Início: <strong>${escapeHtml(startDate || '—')}</strong></span>
        <span>Término planejado: <strong>${escapeHtml(finishDate || '—')}</strong></span>
        <span>Planejado hoje: <strong>${formatPercent(plannedToday)}</strong></span>
        <span>Desvio: <strong>${formatPercent(deviationPercent)}</strong></span>
        <span>Envio efetivo: <strong>${escapeHtml(shipmentDate || '—')}</strong></span>
      </div>
    </header>

    ${renderClientBspOverrideEditor(project)}
    ${renderClientBspOverrideNotice(project)}
    ${renderClientOnHoldNotice(project)}
    ${renderClientTratativaNotice(project)}

    <div class="client-exec-kpis">
      <article><span>Progresso geral</span><strong>${formatPercent(overall)}</strong></article>
      <article><span>Peso programado</span><strong>${formatNumber(weight, 0)} kg</strong></article>
      <article><span>Peso soldado</span><strong>${formatNumber(welded, 0)} kg</strong></article>
      <article><span>Peso restante</span><strong>${formatNumber(pending, 0)} kg</strong></article>
      <article><span>Tags totais</span><strong>${formatNumber(totalTags)}</strong></article>
      <article><span>Tags restantes</span><strong>${formatNumber(remainingTags)}</strong></article>
      <article><span>M² programada</span><strong>${formatNumber(project.m2Painting, 3)}</strong></article>
      <article><span>Etapa atual</span><strong>${escapeHtml(getProjectCurrentStageDisplay(project))}</strong></article>
    </div>

    <div class="client-exec-grid client-exec-grid--top">
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Overall Progress</h3><span>Concluído x restante + desvio</span></div>
        ${renderClientGauge(overall, 'concluído', plannedToday, { deliveryDate: shipmentDate, note: 'Meta até hoje', deadlineDate: finishDate })}
      </section>
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Fabrication Progress</h3><span>Fabricação ponderada + desvio</span></div>
        ${renderClientGauge(fabrication, 'fabricação', plannedToday, { note: 'Meta até hoje', deadlineDate: finishDate })}
      </section>
      <section class="client-exec-card client-exec-card--bars">
        <div class="client-exec-card-head"><h3>Progress by Production Stage</h3><span>Etapas principais</span></div>
        <div class="client-exec-bars">
          ${stages.map((stage) => `<div class="client-exec-bar-row"><span>${escapeHtml(stage.label)}</span><div><i style="width:${clampClientPercent(stage.percent)}%"></i></div><strong>${formatPercent(stage.percent)}</strong></div>`).join('')}
        </div>
      </section>
    </div>

    <div class="client-exec-grid client-exec-grid--main">
      <section class="client-exec-card client-exec-card--curve">
        <div class="client-exec-card-head"><h3>Curva S | Planejado x Realizado</h3><span>Baseada na data inicial e final da BSP</span></div>
        <div class="client-exec-legend"><span><i class="planned"></i> Planejado</span><span><i class="actual"></i> Realizado</span><span><i class="delay"></i> Desvio</span></div>
        ${renderClientSCurveSvg(project)}
      </section>
      <aside class="client-exec-side">
        <section class="client-exec-card">
          <div class="client-exec-card-head"><h3>Spool / Tags</h3><span>Resumo da BSP</span></div>
          <div class="client-exec-mini-table">
            <div><span>Total</span><strong>${formatNumber(totalTags)}</strong></div>
            <div><span>Finalizadas</span><strong>${formatNumber(completedTags)}</strong></div>
            <div><span>Restantes</span><strong>${formatNumber(remainingTags)}</strong></div>
          </div>
        </section>
        <section class="client-exec-card">
          <div class="client-exec-card-head"><h3>Weight</h3><span>Programado x soldado</span></div>
          <div class="client-weight-bars">
            <div><span>Soldado</span><div><i style="width:${weight ? clampClientPercent((welded / weight) * 100) : 0}%"></i></div><strong>${formatNumber(welded, 0)} kg</strong></div>
            <div><span>Restante</span><div><i class="remaining" style="width:${weight ? clampClientPercent((pending / weight) * 100) : 0}%"></i></div><strong>${formatNumber(pending, 0)} kg</strong></div>
          </div>
        </section>
      </aside>
    </div>

    <section class="client-exec-card">
      <div class="client-exec-card-head"><h3>Timeline operacional</h3><span>Resumo dos passos do schedule</span></div>
      <div class="client-exec-timeline">
        ${timeline.map((item) => `<div class="client-exec-step is-${item.state}"><span></span><strong>${escapeHtml(item.label)}</strong><small>${formatPercent(item.percent)}</small></div>`).join('')}
      </div>
    </section>

    <section class="client-exec-card client-exec-schedule-card">
      <div class="client-exec-card-head"><h3>Schedule Executivo da BSP</h3><span>Planejado + datas reais do Tracking quando preenchidas</span></div>
      ${renderClientExecutiveSchedule(project)}
    </section>

    ${renderClientTrackingReportPreview(project)}

    ${'' /* v37.73: Attention Points oculto no painel do cliente */}

    <section class="client-exec-card client-exec-process-detail">
      <div class="client-exec-card-head"><h3>Detalhamento da obra</h3><span>Processos da BSP e evolução por etapa</span></div>
      <div class="client-stage-strip client-stage-strip--executive">
        ${detailStageKeys.map((key) => {
          const label = (state.meta?.stageOrder || []).find((stage) => stage.key === key)?.label || key;
          const value = stageValues[key];
          return renderClientStageStripCard(label, value);
        }).join('')}
      </div>
      <div class="client-table-wrap client-table-wrap--compact client-exec-process-table">
        <table class="client-bsp-table"><thead><tr><th>Tag/ISO</th><th>Descrição</th><th>Observação</th><th>Status</th><th>Etapa</th><th>%</th><th>Peso</th></tr></thead><tbody>
          ${renderClientSpoolRows(spools, 120, project)}
        </tbody></table>
      </div>
    </section>
  `;

  modal.classList.remove('hidden');
  document.body.classList.add('client-exec-open');
}


function ensureClientApiModal() {
  let modal = document.getElementById('client-api-modal');
  if (modal) return modal;
  modal = document.createElement('div');
  modal.id = 'client-api-modal';
  modal.className = 'client-api-modal hidden';
  modal.innerHTML = `
    <div class="client-api-dialog" role="dialog" aria-modal="true" aria-labelledby="client-api-title">
      <div class="client-api-head">
        <div>
          <p class="client-kicker">Integração</p>
          <h2 id="client-api-title">API do Portal do Cliente</h2>
          <p>Gere uma chave para consumir, de forma controlada, as informações que já aparecem neste painel.</p>
        </div>
        <button class="client-api-close" type="button" data-client-api-close aria-label="Fechar">×</button>
      </div>
      <div class="client-api-body">
        <section class="client-api-card">
          <div class="client-api-card-head">
            <div>
              <h3>Nova chave</h3>
              <p>O token completo aparece somente uma vez. Copie e guarde em local seguro.</p>
            </div>
          </div>
          <div class="client-api-form-row">
            <label>
              <span>Nome da integração</span>
              <input id="client-api-key-name" type="text" value="API do Portal do Cliente" maxlength="120" />
            </label>
            <button class="mini-action-button" type="button" data-client-api-create>Criar API</button>
          </div>
          <div id="client-api-new-token" class="client-api-token-box hidden"></div>
        </section>
        <section class="client-api-card">
          <div class="client-api-card-head">
            <div>
              <h3>Como consumir</h3>
              <p>Use o endpoint abaixo com o header Authorization.</p>
            </div>
          </div>
          <pre class="client-api-code" id="client-api-example"></pre>
          <div class="client-api-help-grid">
            <span><strong>Resumo:</strong> /api/client-data?format=summary</span>
            <span><strong>Completo:</strong> /api/client-data?format=full</span>
            <span><strong>Com spools:</strong> /api/client-data?includeSpools=1</span>
            <span><strong>Filtro:</strong> /api/client-data?unit=FORTE ou ?bsp=25-1165-38</span>
          </div>
        </section>
        <section class="client-api-card">
          <div class="client-api-card-head">
            <div>
              <h3>Chaves criadas</h3>
              <p>Revogue imediatamente qualquer chave que não estiver mais em uso.</p>
            </div>
            <button class="mini-action-button" type="button" data-client-api-refresh>Atualizar</button>
          </div>
          <div id="client-api-feedback" class="client-api-feedback"></div>
          <div id="client-api-keys-list" class="client-api-keys-list"></div>
        </section>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
  modal.addEventListener('click', handleClientApiModalClick);
  return modal;
}

function getClientApiEndpoint() {
  return `${window.location.origin}/api/client-data`;
}

function renderClientApiExample(token = '<SUA_API_KEY>') {
  const exampleEl = document.getElementById('client-api-example');
  if (!exampleEl) return;
  const endpoint = getClientApiEndpoint();
  exampleEl.textContent = `curl -H "Authorization: Bearer ${token}" \\\n  "${endpoint}?format=summary"`;
}

function renderClientApiKeys() {
  const listEl = document.getElementById('client-api-keys-list');
  const feedbackEl = document.getElementById('client-api-feedback');
  const tokenEl = document.getElementById('client-api-new-token');
  if (feedbackEl) feedbackEl.textContent = state.clientApi.feedback || '';
  if (tokenEl) {
    if (state.clientApi.newToken) {
      tokenEl.classList.remove('hidden');
      tokenEl.innerHTML = `
        <span>Copie sua chave agora:</span>
        <code>${escapeHtml(state.clientApi.newToken)}</code>
        <div class="client-api-token-actions">
          <button class="mini-action-button" type="button" data-client-api-copy-token>Copiar token</button>
        </div>
      `;
    } else {
      tokenEl.classList.add('hidden');
      tokenEl.innerHTML = '';
    }
  }
  if (!listEl) return;
  const rawToken = state.clientApi.newToken || '';
  const rawTokenId = state.clientApi.newTokenKeyId || '';
  const tokenNotice = rawToken ? `
    <article class="client-api-key-item client-api-key-item--new-token">
      <div>
        <strong>Token criado agora</strong>
        <span>Copie este token completo. Ele não será exibido novamente depois que fechar esta janela.</span>
        <code>${escapeHtml(rawToken)}</code>
        <small>Use em Authorization: Bearer ${escapeHtml(rawToken)}</small>
      </div>
      <button class="mini-action-button" type="button" data-client-api-copy-token>Copiar token</button>
    </article>
  ` : '';
  if (state.clientApi.loading) {
    listEl.innerHTML = `${tokenNotice}<div class="client-api-empty">Carregando chaves...</div>`;
    return;
  }
  const keys = Array.isArray(state.clientApi.keys) ? state.clientApi.keys : [];
  if (!keys.length) {
    listEl.innerHTML = `${tokenNotice}<div class="client-api-empty">Nenhuma API criada ainda.</div>`;
    return;
  }
  listEl.innerHTML = tokenNotice + keys.map((key) => {
    const status = key.active === false ? 'Revogada' : 'Ativa';
    const created = key.createdAt ? new Date(key.createdAt).toLocaleString('pt-BR') : '--';
    const used = key.lastUsedAt ? new Date(key.lastUsedAt).toLocaleString('pt-BR') : 'Nunca usada';
    const isNewKey = rawToken && rawTokenId && String(key.id || '') === String(rawTokenId);
    return `
      <article class="client-api-key-item ${key.active === false ? 'is-revoked' : ''} ${isNewKey ? 'is-new-key' : ''}">
        <div>
          <strong>${escapeHtml(key.name || 'API do cliente')}</strong>
          <span>${escapeHtml(key.tokenPreview || 'step_••••')}</span>
          <small>Criada: ${escapeHtml(created)} • Último uso: ${escapeHtml(used)} • ${escapeHtml(status)}</small>
        </div>
        <div class="client-api-key-actions">
          ${isNewKey ? '<button class="mini-action-button" type="button" data-client-api-copy-token>Copiar token</button>' : ''}
          ${key.active === false ? '' : `<button class="mini-action-button mini-action-button--danger" type="button" data-client-api-revoke="${escapeHtml(key.id)}">Revogar</button>`}
        </div>
      </article>
    `;
  }).join('');
}

async function loadClientApiKeys() {
  state.clientApi.loading = true;
  state.clientApi.feedback = '';
  renderClientApiKeys();
  try {
    const response = await fetch('/api/client-api-keys', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar chaves de API.');
    state.clientApi.keys = Array.isArray(data.keys) ? data.keys : [];
  } catch (error) {
    state.clientApi.feedback = error?.message || 'Falha ao carregar chaves de API.';
  } finally {
    state.clientApi.loading = false;
    renderClientApiKeys();
  }
}

function openClientApiModal() {
  if (!isClientUser()) return;
  const modal = ensureClientApiModal();
  state.clientApi.newToken = '';
  state.clientApi.newTokenKeyId = '';
  state.clientApi.feedback = '';
  modal.classList.remove('hidden');
  document.body.classList.add('client-api-open');
  renderClientApiExample();
  renderClientApiKeys();
  loadClientApiKeys();
}

/**
 * Abre uma nova guia para listar projetos em desenvolvimento (pré‑BSP) do cliente.
 * Atualmente o recurso foi criado apenas para o cliente Yinson e usa a página
 * client-under-development.html localizada em /site.
 */
function openClientUnderDevPage() {
  if (!isClientUser()) return;
  // A página nova é aberta em outra guia para que o usuário não perca o contexto do painel.
  // Abre a página estática em uma nova aba. Usamos um caminho relativo sem './' para garantir
  // que o arquivo raiz do site seja resolvido corretamente em ambientes Netlify/SPA.
  window.open('client-under-development.html', '_blank');
}

/**
 * Garante que o modal de imagens da BSP exista no DOM e retorna sua referência.
 * Este modal é criado dinamicamente apenas na primeira vez que for necessário.
 * Ele contém uma tabela para listar os anexos de imagem e um viewer para
 * visualizar cada imagem individualmente. Fechar o modal remove a classe
 * `client-bsp-images-open` do body para restaurar a rolagem.
 */
function ensureClientBspImagesModalEl() {
  let modal = document.getElementById('client-bsp-images-modal');
  if (modal) return modal;
  modal = document.createElement('div');
  modal.id = 'client-bsp-images-modal';
  modal.className = 'client-bsp-images-modal';
  modal.innerHTML = `
    <div class="client-bsp-images-dialog" role="dialog" aria-modal="true" aria-label="Imagens da BSP">
      <div class="client-bsp-images-header">
        <h2>Imagens da BSP</h2>
        <button type="button" class="client-bsp-images-close" aria-label="Fechar">×</button>
      </div>
      <div class="client-bsp-images-progress-container" hidden>
        <div class="client-bsp-images-progress-bar"></div>
      </div>
      <div class="client-bsp-images-content">
        <table class="client-bsp-images-table">
          <thead>
            <tr>
              <th>Arquivo</th>
              <th>Ação</th>
            </tr>
          </thead>
          <tbody id="client-bsp-images-body">
            <tr><td colspan="2" class="client-bsp-images-placeholder">Carregando imagens…</td></tr>
          </tbody>
        </table>
      </div>
      <!-- Viewer para imagem selecionada, com navegação -->
      <div id="client-bsp-img-viewer" class="client-bsp-img-viewer">
        <button type="button" class="client-bsp-image-prev" aria-label="Imagem anterior">‹</button>
        <img id="client-bsp-img" src="" alt="" />
        <button type="button" class="client-bsp-image-next" aria-label="Próxima imagem">›</button>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
  // Close handlers
  const closeBtn = modal.querySelector('.client-bsp-images-close');
  closeBtn.addEventListener('click', () => closeClientBspImagesModal());
  // Fechar clicando fora do diálogo
  modal.addEventListener('click', (ev) => {
    if (ev.target === modal) closeClientBspImagesModal();
  });
  // Fechar viewer quando clicado nele
  const viewer = modal.querySelector('#client-bsp-img-viewer');
  viewer.addEventListener('click', () => {
    viewer.classList.remove('active');
    const imgEl = modal.querySelector('#client-bsp-img');
    if (imgEl) imgEl.src = '';
  });
  // Navigation buttons inside the viewer. Use event.stopPropagation() to
  // prevent closing the viewer when clicking on navigation controls.
  const prevBtn = viewer.querySelector('.client-bsp-image-prev');
  const nextBtn = viewer.querySelector('.client-bsp-image-next');
  if (prevBtn) {
    prevBtn.addEventListener('click', (ev) => {
      ev.stopPropagation();
      showPrevClientBspImage();
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener('click', (ev) => {
      ev.stopPropagation();
      showNextClientBspImage();
    });
  }
  return modal;
}

/**
 * Abre o modal de imagens da BSP e carrega os anexos de imagem via API. O
 * carregamento é assíncrono: exibe uma barra de progresso enquanto busca os
 * dados. Após carregar, popula a tabela com as imagens disponíveis. Cada
 * imagem possui um botão "Ver" que abre o arquivo no viewer interno. Caso
 * não haja imagens, exibe uma mensagem no lugar da tabela. Erros durante
 * a busca também são tratados com mensagem.
 *
 * @param {string} rowId O identificador da linha/BSP no Smartsheet.
 */
/**
 * Abre o modal de imagens da BSP e carrega os anexos de imagem. Além do
 * identificador da linha da BSP (`rowId`), aceita listas opcionais de
 * números de linha (`spoolRowNumbers`) e identificadores de linha
 * (`spoolRowIds`) para incluir anexos dos spools. Também pode receber o
 * identificador de planilha (`sheetId`) para forçar a busca em uma
 * planilha específica. Ao montar a URL de chamada, adiciona apenas os
 * parâmetros que possuem valores.
 *
 * @param {string|number} rowId - ID da linha da BSP
 * @param {number[]} [spoolRowNumbers=[]] - Lista de números de linha das linhas de spool
 * @param {string[]} [spoolRowIds=[]] - Lista de IDs das linhas de spool
 * @param {string|null} [sheetId=null] - ID da planilha Smartsheet (opcional)
 */
function openClientBspImagesModal(rowId, spoolRowNumbers = [], spoolRowIds = [], sheetId = null) {
  const modal = ensureClientBspImagesModalEl();
  modal.classList.add('active');
  document.body.classList.add('client-bsp-images-open');
  const progressContainer = modal.querySelector('.client-bsp-images-progress-container');
  const progressBar = modal.querySelector('.client-bsp-images-progress-bar');
  const tbody = modal.querySelector('#client-bsp-images-body');
  // Reset viewer state
  const viewer = modal.querySelector('#client-bsp-img-viewer');
  if (viewer) viewer.classList.remove('active');
  const imgEl = modal.querySelector('#client-bsp-img');
  if (imgEl) imgEl.src = '';
  // Show loading placeholder and progress
  tbody.innerHTML = '<tr><td colspan="2" class="client-bsp-images-placeholder">Carregando imagens…</td></tr>';
  progressContainer.hidden = false;
  progressBar.style.width = '0%';
  // Simple progress bar that animates up to 90% until the fetch completes
  let percent = 0;
  const timer = setInterval(() => {
    percent = Math.min(percent + Math.random() * 10, 90);
    progressBar.style.width = percent + '%';
  }, 200);
  // Build query string for API call. Executive BSP images must be pulled
  // from the tracking sheet only (3612820139992964). The Yinson development
  // page uses its own endpoint and is not affected by this request.
  let url = `/api/client-bsp-images?rowId=${encodeURIComponent(rowId)}`;
  if (Array.isArray(spoolRowNumbers) && spoolRowNumbers.length) {
    const numbersParam = spoolRowNumbers.map((n) => String(n).trim()).filter(Boolean).join(',');
    if (numbersParam) url += `&rowNumbers=${encodeURIComponent(numbersParam)}`;
  }
  // Prefer rowNumbers for child rows to keep the URL shorter and avoid invalid rowId batches.
  // Use rowIds only when rowNumbers are not available.
  if ((!Array.isArray(spoolRowNumbers) || !spoolRowNumbers.length) && Array.isArray(spoolRowIds) && spoolRowIds.length) {
    const idsParam = spoolRowIds.map((id) => String(id).trim()).filter(Boolean).join(',');
    if (idsParam) url += `&rowIds=${encodeURIComponent(idsParam)}`;
  }
  const executiveSheetId = sheetId || '3612820139992964';
  url += `&sheetId=${encodeURIComponent(executiveSheetId)}`;
  // Fetch images from Netlify function
  fetch(url, { cache: 'no-store' })
    .then((res) => {
      if (!res.ok) throw new Error('Erro ao carregar imagens');
      return res.json();
    })
    .then((json) => {
      if (json?.error) {
        throw new Error(json.error);
      }
      const images = Array.isArray(json?.images) ? json.images : [];
      tbody.innerHTML = '';
      // Guardar a lista de imagens no modal para navegação subsequente
      modal._bspImages = images;
      modal._currentImageIndex = 0;
      images.forEach((img, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(String(img.name || ''))}</td>
          <td><button type="button" class="client-bsp-image-view-link" data-attachment-id="${escapeHtml(String(img.id))}" data-attachment-name="${escapeHtml(String(img.name || 'Imagem'))}" data-sheet-id="${escapeHtml(String(img.sheetId || ''))}" data-image-index="${idx}">Ver</button></td>
        `;
        tbody.appendChild(tr);
      });
      if (!images.length) {
        tbody.innerHTML = '<tr><td colspan="2" class="client-bsp-images-placeholder">Nenhuma imagem encontrada.</td></tr>';
      }
      bindClientBspImageLinks();
    })
    .catch((err) => {
      console.error(err);
      const message = err?.message ? `Falha ao carregar imagens: ${escapeHtml(err.message)}` : 'Falha ao carregar imagens.';
      tbody.innerHTML = `<tr><td colspan="2" class="client-bsp-images-placeholder">${message}</td></tr>`;
    })
    .finally(() => {
      clearInterval(timer);
      progressBar.style.width = '100%';
      setTimeout(() => { progressContainer.hidden = true; }, 400);
    });
}

/**
 * Percorre todos os botões da tabela de imagens no modal e associa um
 * manipulador de clique. Ao clicar, o viewer interno será aberto com a
 * imagem correspondente. O viewer utiliza o endpoint proxy
 * `/api/client-under-dev-image?id=` para buscar a imagem de forma segura.
 */
function bindClientBspImageLinks() {
  const modal = document.getElementById('client-bsp-images-modal');
  if (!modal) return;
  const buttons = modal.querySelectorAll('.client-bsp-image-view-link');
  buttons.forEach((button) => {
    button.addEventListener('click', (ev) => {
      ev.preventDefault();
      const attId = button.getAttribute('data-attachment-id');
      const name = button.getAttribute('data-attachment-name') || 'Imagem';
      const viewer = modal.querySelector('#client-bsp-img-viewer');
      const imgEl = modal.querySelector('#client-bsp-img');
      // Ao clicar em uma imagem, utilize o índice armazenado para abrir no viewer
      const indexAttr = button.getAttribute('data-image-index');
      const idx = indexAttr != null ? Number(indexAttr) : 0;
      showClientBspImageAtIndex(idx);
    });
  });
}

/**
 * Fecha o modal de imagens da BSP, restaurando a rolagem do corpo e
 * ocultando o conteúdo. A limpeza do viewer é feita na função que
 * encerra a visualização individual.
 */
function closeClientBspImagesModal() {
  const modal = document.getElementById('client-bsp-images-modal');
  if (modal) modal.classList.remove('active');
  document.body.classList.remove('client-bsp-images-open');
}

/**
 * Exibe a imagem no viewer conforme o índice fornecido. Mantém o índice
 * atual no modal para permitir navegação com os botões Anterior/Próximo.
 *
 * @param {number} index Índice da imagem a ser exibida
 */
function showClientBspImageAtIndex(index) {
  const modal = document.getElementById('client-bsp-images-modal');
  if (!modal) return;
  const images = modal._bspImages || [];
  if (!images.length) return;
  // Ajustar o índice para ficar dentro dos limites e permitir ciclo infinito
  const adjIndex = ((Number(index) % images.length) + images.length) % images.length;
  modal._currentImageIndex = adjIndex;
  const imgData = images[adjIndex];
  const viewer = modal.querySelector('#client-bsp-img-viewer');
  const imgEl = modal.querySelector('#client-bsp-img');
  if (imgEl && imgData) {
    const attId = imgData.id;
    const sheetId = imgData.sheetId || '3612820139992964';
    // Use the temporary URL returned by the BSP image endpoint whenever available.
    // This supports both row attachments and Smartsheet cell images. If no URL was
    // returned, fall back to the attachment proxy.
    const directUrl = imgData.url && String(imgData.url).trim();
    const baseUrl = `/api/client-under-dev-image?id=${encodeURIComponent(attId)}`;
    const proxyUrl = sheetId ? `${baseUrl}&sheetId=${encodeURIComponent(sheetId)}` : baseUrl;
    imgEl.src = directUrl || proxyUrl;
    imgEl.alt = imgData.name || 'Imagem';
    if (viewer) viewer.classList.add('active');
  }
}

/**
 * Avança para a próxima imagem no viewer.
 */
function showNextClientBspImage() {
  const modal = document.getElementById('client-bsp-images-modal');
  if (!modal) return;
  const idx = modal._currentImageIndex != null ? Number(modal._currentImageIndex) : 0;
  showClientBspImageAtIndex(idx + 1);
}

/**
 * Retorna para a imagem anterior no viewer.
 */
function showPrevClientBspImage() {
  const modal = document.getElementById('client-bsp-images-modal');
  if (!modal) return;
  const idx = modal._currentImageIndex != null ? Number(modal._currentImageIndex) : 0;
  showClientBspImageAtIndex(idx - 1);
}


/**
 * Modal sob demanda para importar imagens diretamente para a linha da BSP no Smartsheet.
 * O modal só é criado/carregado quando o usuário clica em "Importar imagem", mantendo
 * o painel leve no carregamento inicial.
 */
function ensureClientBspUploadImageModalEl() {
  let modal = document.getElementById('client-bsp-upload-image-modal');
  if (modal) return modal;
  modal = document.createElement('div');
  modal.id = 'client-bsp-upload-image-modal';
  modal.className = 'client-bsp-upload-image-modal';
  modal.innerHTML = `
    <div class="client-bsp-upload-image-dialog" role="dialog" aria-modal="true" aria-label="Importar imagem da BSP">
      <div class="client-bsp-upload-image-header">
        <div>
          <p class="client-kicker">Smartsheet Tracking</p>
          <h2>Importar imagem da BSP</h2>
        </div>
        <button type="button" class="client-bsp-upload-image-close" aria-label="Fechar">×</button>
      </div>
      <div class="client-bsp-upload-image-body">
        <p class="client-bsp-upload-image-note">Selecione uma ou mais imagens. O sistema comprime automaticamente antes de enviar para manter o painel leve.</p>
        <input id="client-bsp-upload-image-input" type="file" accept="image/*" multiple />
        <div id="client-bsp-upload-image-list" class="client-bsp-upload-image-list"></div>
        <p id="client-bsp-upload-image-status" class="client-bsp-upload-image-status"></p>
      </div>
      <div class="client-bsp-upload-image-actions">
        <button type="button" class="client-exec-pdf-button" data-client-bsp-upload-image-cancel>Cancelar</button>
        <button type="button" class="client-exec-pdf-button client-exec-upload-image-submit" data-client-bsp-upload-image-submit>Enviar para Smartsheet</button>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
  modal.addEventListener('click', (ev) => {
    if (ev.target === modal || ev.target.closest('.client-bsp-upload-image-close') || ev.target.closest('[data-client-bsp-upload-image-cancel]')) {
      closeClientBspUploadImageModal();
    }
    const submit = ev.target.closest('[data-client-bsp-upload-image-submit]');
    if (submit) {
      const rowId = modal.dataset.rowId;
      uploadClientBspSelectedImages(rowId);
    }
  });
  const input = modal.querySelector('#client-bsp-upload-image-input');
  input.addEventListener('change', () => renderClientBspUploadSelection());
  return modal;
}

function openClientBspUploadImageModal(rowId) {
  if (!rowId) return;
  const modal = ensureClientBspUploadImageModalEl();
  modal.dataset.rowId = rowId;
  modal.dataset.sheetId = '3612820139992964';
  modal.classList.add('active');
  document.body.classList.add('client-bsp-upload-image-open');
  const input = modal.querySelector('#client-bsp-upload-image-input');
  const list = modal.querySelector('#client-bsp-upload-image-list');
  const status = modal.querySelector('#client-bsp-upload-image-status');
  if (input) input.value = '';
  if (list) list.innerHTML = '';
  if (status) status.textContent = '';
}

function closeClientBspUploadImageModal() {
  const modal = document.getElementById('client-bsp-upload-image-modal');
  if (modal) modal.classList.remove('active');
  document.body.classList.remove('client-bsp-upload-image-open');
}

function renderClientBspUploadSelection() {
  const modal = document.getElementById('client-bsp-upload-image-modal');
  if (!modal) return;
  const input = modal.querySelector('#client-bsp-upload-image-input');
  const list = modal.querySelector('#client-bsp-upload-image-list');
  const files = Array.from(input?.files || []);
  if (!list) return;
  if (!files.length) {
    list.innerHTML = '';
    return;
  }
  list.innerHTML = files.map((file) => `
    <div class="client-bsp-upload-image-item">
      <span>${escapeHtml(file.name)}</span>
      <small>${Math.round(file.size / 1024)} KB</small>
    </div>
  `).join('');
}

function readClientBspFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(reader.error || new Error('Falha ao ler imagem.'));
    reader.readAsDataURL(file);
  });
}

function loadClientBspImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('Não foi possível carregar a imagem para compressão.'));
    img.src = dataUrl;
  });
}

async function prepareClientBspImageForUpload(file) {
  const originalDataUrl = await readClientBspFileAsDataUrl(file);
  // GIF/HEIC e formatos que o canvas pode não processar seguem sem conversão.
  if (/image\/gif/i.test(file.type) || /heic|heif/i.test(file.type || file.name)) {
    return {
      fileName: file.name,
      mimeType: file.type || 'image/jpeg',
      base64: originalDataUrl.split(',')[1] || originalDataUrl,
    };
  }
  try {
    const img = await loadClientBspImage(originalDataUrl);
    const maxSide = 1800;
    const ratio = Math.min(1, maxSide / Math.max(img.width || maxSide, img.height || maxSide));
    const width = Math.max(1, Math.round((img.width || maxSide) * ratio));
    const height = Math.max(1, Math.round((img.height || maxSide) * ratio));
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, width, height);
    const outType = 'image/jpeg';
    const compressed = canvas.toDataURL(outType, 0.82);
    const baseName = String(file.name || 'imagem.jpg').replace(/\.[^.]+$/, '');
    return {
      fileName: `${baseName}.jpg`,
      mimeType: outType,
      base64: compressed.split(',')[1] || compressed,
    };
  } catch (err) {
    return {
      fileName: file.name,
      mimeType: file.type || 'image/jpeg',
      base64: originalDataUrl.split(',')[1] || originalDataUrl,
    };
  }
}

async function uploadClientBspSelectedImages(rowId) {
  const modal = document.getElementById('client-bsp-upload-image-modal');
  if (!modal || !rowId) return;
  const input = modal.querySelector('#client-bsp-upload-image-input');
  const status = modal.querySelector('#client-bsp-upload-image-status');
  const submit = modal.querySelector('[data-client-bsp-upload-image-submit]');
  const files = Array.from(input?.files || []);
  if (!files.length) {
    if (status) status.textContent = 'Selecione pelo menos uma imagem.';
    return;
  }
  if (submit) submit.disabled = true;
  const sheetId = modal.dataset.sheetId || '3612820139992964';
  let okCount = 0;
  try {
    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      if (status) status.textContent = `Preparando ${i + 1}/${files.length}: ${file.name}`;
      const prepared = await prepareClientBspImageForUpload(file);
      if (status) status.textContent = `Enviando ${i + 1}/${files.length}: ${prepared.fileName}`;
      const res = await fetch('/api/client-bsp-upload-image', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          rowId,
          sheetId,
          fileName: prepared.fileName,
          mimeType: prepared.mimeType,
          base64: prepared.base64,
        }),
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok || !json.ok) {
        throw new Error(json.error || `Falha ao enviar ${prepared.fileName}.`);
      }
      okCount += 1;
    }
    if (status) status.textContent = `${okCount} imagem(ns) importada(s) com sucesso. Atualizando galeria...`;
    closeClientBspUploadImageModal();
    const proj = state.projects && state.projects.find((item) => String(item.rowId) === String(rowId));
    const rowNumbers = [];
    const rowIds = [];
    if (proj && Array.isArray(proj.spools)) {
      for (const sp of proj.spools) {
        if (sp && sp.rowNumber != null) rowNumbers.push(sp.rowNumber);
        if (sp && sp.rowId != null) rowIds.push(sp.rowId);
      }
    }
    openClientBspImagesModal(rowId, rowNumbers, rowIds, sheetId);
  } catch (err) {
    if (status) status.textContent = err.message || 'Falha ao importar imagem.';
  } finally {
    if (submit) submit.disabled = false;
  }
}

function closeClientApiModal() {
  const modal = document.getElementById('client-api-modal');
  if (modal) modal.classList.add('hidden');
  document.body.classList.remove('client-api-open');
  state.clientApi.newToken = '';
  state.clientApi.newTokenKeyId = '';
}

async function createClientApiKeyFromModal() {
  const input = document.getElementById('client-api-key-name');
  const name = input?.value || 'API do Portal do Cliente';
  state.clientApi.loading = true;
  state.clientApi.feedback = 'Criando chave...';
  state.clientApi.newToken = '';
  state.clientApi.newTokenKeyId = '';
  renderClientApiKeys();
  try {
    const response = await fetch('/api/client-api-keys', {
      method: 'POST',
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao criar API.');
    state.clientApi.newToken = data.key?.token || '';
    state.clientApi.newTokenKeyId = data.key?.id || '';
    state.clientApi.feedback = 'API criada com sucesso. Copie o token agora.';
    renderClientApiExample(state.clientApi.newToken || '<SUA_API_KEY>');
    await loadClientApiKeys();
    state.clientApi.feedback = 'API criada com sucesso. Copie o token agora.';
    renderClientApiKeys();
    setTimeout(() => {
      const tokenBox = document.getElementById('client-api-new-token') || document.querySelector('.client-api-key-item--new-token');
      tokenBox?.scrollIntoView?.({ behavior: 'smooth', block: 'center' });
    }, 50);
  } catch (error) {
    state.clientApi.feedback = error?.message || 'Falha ao criar API.';
    renderClientApiKeys();
  } finally {
    state.clientApi.loading = false;
  }
}

async function revokeClientApiKeyFromModal(id) {
  if (!id) return;
  state.clientApi.feedback = 'Revogando chave...';
  renderClientApiKeys();
  try {
    const response = await fetch('/api/client-api-keys', {
      method: 'DELETE',
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao revogar API.');
    state.clientApi.feedback = 'Chave revogada.';
    await loadClientApiKeys();
  } catch (error) {
    state.clientApi.feedback = error?.message || 'Falha ao revogar API.';
    renderClientApiKeys();
  }
}

function copyTextToClipboard(value) {
  const text = String(value || '');
  if (!text) return Promise.reject(new Error('Nada para copiar.'));
  if (navigator.clipboard?.writeText && window.isSecureContext !== false) {
    return navigator.clipboard.writeText(text);
  }
  return new Promise((resolve, reject) => {
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', 'readonly');
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      textarea.style.top = '0';
      document.body.appendChild(textarea);
      textarea.select();
      textarea.setSelectionRange(0, textarea.value.length);
      const ok = document.execCommand('copy');
      document.body.removeChild(textarea);
      ok ? resolve() : reject(new Error('Cópia bloqueada pelo navegador.'));
    } catch (error) {
      reject(error);
    }
  });
}

function copyClientApiToken() {
  const token = state.clientApi.newToken || '';
  if (!token) {
    state.clientApi.feedback = 'O token completo só aparece no momento da criação. Crie uma nova API caso não tenha copiado.';
    renderClientApiKeys();
    return;
  }
  copyTextToClipboard(token).then(() => {
    state.clientApi.feedback = 'Token copiado.';
    renderClientApiKeys();
  }).catch(() => {
    state.clientApi.feedback = 'Não consegui copiar automaticamente. Selecione o token completo exibido e copie manualmente.';
    renderClientApiKeys();
  });
}

function handleClientApiModalClick(event) {
  if (event.target.closest('[data-client-api-close]') || event.target.id === 'client-api-modal') {
    closeClientApiModal();
    return;
  }
  if (event.target.closest('[data-client-api-create]')) {
    createClientApiKeyFromModal();
    return;
  }
  if (event.target.closest('[data-client-api-refresh]')) {
    loadClientApiKeys();
    return;
  }
  if (event.target.closest('[data-client-api-copy-token]')) {
    copyClientApiToken();
    return;
  }
  const revokeButton = event.target.closest('[data-client-api-revoke]');
  if (revokeButton) {
    revokeClientApiKeyFromModal(revokeButton.dataset.clientApiRevoke || '');
  }
}

