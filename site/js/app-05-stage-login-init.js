/* STEP Dashboard v37.11 - Login/logout, apontamentos, validação PCP e inicialização. Arquivo gerado a partir da divisão segura do app.js. */
function startPostSessionBackgroundLoads(options = {}) {
  if (!state.user) return Promise.resolve([]);
  const autoOpenStageValidation = Boolean(options.autoOpenStageValidation);
  const tasks = [
    syncPushSubscription(false).catch(() => {}),
    loadManualAlerts().catch((error) => console.warn('Falha ao carregar alertas:', error?.message || error)),
    loadAlertResponses().catch((error) => console.warn('Falha ao carregar respostas:', error?.message || error)),
    loadStageUpdates().catch((error) => console.warn('Falha ao carregar apontamentos:', error?.message || error)),
  ];

  if (state.user?.role === "admin") {
    tasks.push(loadAdminData().catch((error) => console.warn('Falha ao carregar dados administrativos:', error?.message || error)));
  }

  return Promise.allSettled(tasks).then((results) => {
    if (autoOpenStageValidation && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      renderStageUpdatesModal();
    }
    return results;
  });
}

async function handleLoginSubmit(event) {
  event.preventDefault();
  if (!loginFeedbackEl) return;
  loginFeedbackEl.textContent = "Validando acesso...";
  startLoginProgress({
    initialPercent: 6,
    title: 'Validando acesso...',
    message: 'Estamos conferindo suas credenciais e preparando sua sessão.',
    detail: 'Autenticação em andamento.',
  });
  try {
    const response = await fetch("/api/auth-login", {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        username: String(loginUsernameEl.value || "").trim(),
        password: String(loginPasswordEl.value || "").trim(),
        operationRegion: 'BR',
        siteKey: 'BR',
      }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao entrar.");
    }
    setLoginProgress(20, {
      title: 'Acesso validado.',
      message: 'Estamos conectando seu usuário ao ambiente correto.',
      detail: 'Sessão validada.',
    });
    state.user = data.user;
    resetClientBspOverridesState();
    if (shouldUseSectorScopedToggle(state.user)) {
      state.sectorScopedView = loadSectorScopedViewPreference(state.user);
      state.alertSectorFilter = state.sectorScopedView ? normalizeAlertSectorFilterValue(getPrimaryUserSector(state.user)) || 'all' : 'all';
    }
    updateSessionUi();
    closeLoginModal();
    setProjectsLoadingState('Acesso validado. Carregando painel...');
    setLoginProgress(34, {
      title: 'Carregando BSPs...',
      message: 'Estamos carregando as BSPs e organizando os projetos por cliente e unidade.',
      detail: 'BSPs em processamento.',
    });

    const sessionPromise = bootstrapSession();
    const projectsPromise = loadProjects({ preferServerCache: true, requireData: false }).then((result) => {
      setLoginProgress(72, {
        title: 'Atualizando indicadores...',
        message: 'Estamos calculando pesos, status, pendências e alertas operacionais.',
        detail: 'Indicadores sendo preparados.',
      });
      return result;
    });
    const authenticated = await sessionPromise;
    if (!authenticated) {
      failLoginProgress('Sua sessão não pôde ser confirmada. Faça login novamente.');
      return;
    }
    setLoginProgress(55, {
      title: 'Montando Portal do Cliente...',
      message: 'Estamos abrindo as BSPs primeiro. As POs e referências continuam atualizando em segundo plano.',
      detail: 'BSPs e dashboard em processamento.',
    });
    await projectsPromise;
    setLoginProgress(88, {
      title: 'Definindo dashboards...',
      message: 'Estamos definindo os dashboards e montando a visualização final.',
      detail: 'Dashboard quase pronto.',
    });

    await ensureDashboardDataReadyBeforeRelease({ maxAttempts: 2, retryDelayMs: 350 });

    syncStageDraftsForCurrentSector();
    const autoOpenStageValidation = shouldOpenStageValidationWorkspaceFromUrl() && canValidateStageWorkspace();
    if (autoOpenStageValidation) {
      openStageUpdatesModal({ loading: true });
    }
    startPostSessionBackgroundLoads({ autoOpenStageValidation });
    startPresenceHeartbeat();
    startPolling();
    await completeLoginProgress();
  } catch (error) {
    loginFeedbackEl.textContent = error.message || "Falha ao autenticar.";
    failLoginProgress(error.message || "Falha ao autenticar.");
  }
}

function clearClientVisibleCookies() {
  try {
    const rawCookies = String(document.cookie || '').split(';').map((item) => item.trim()).filter(Boolean);
    if (!rawCookies.length) return;

    const hostname = String(window.location.hostname || '').trim();
    const hostParts = hostname.split('.').filter(Boolean);
    const domainCandidates = new Set(['']);
    if (hostname && !/^localhost$/i.test(hostname) && !/^\d{1,3}(\.\d{1,3}){3}$/.test(hostname)) {
      domainCandidates.add(hostname);
      domainCandidates.add(`.${hostname}`);
      for (let index = 1; index < hostParts.length - 1; index += 1) {
        const domain = hostParts.slice(index).join('.');
        if (domain) {
          domainCandidates.add(domain);
          domainCandidates.add(`.${domain}`);
        }
      }
    }

    const pathCandidates = new Set(['/']);
    const pathParts = String(window.location.pathname || '/').split('/').filter(Boolean);
    let currentPath = '';
    pathParts.forEach((part) => {
      currentPath += `/${part}`;
      pathCandidates.add(currentPath);
      pathCandidates.add(`${currentPath}/`);
    });

    rawCookies.forEach((cookie) => {
      const separatorIndex = cookie.indexOf('=');
      const name = separatorIndex >= 0 ? cookie.slice(0, separatorIndex).trim() : cookie.trim();
      if (!name) return;
      domainCandidates.forEach((domain) => {
        pathCandidates.forEach((path) => {
          const domainPart = domain ? `; domain=${domain}` : '';
          document.cookie = `${name}=; expires=Thu, 01 Jan 1970 00:00:00 GMT; max-age=0; path=${path}${domainPart}; SameSite=Lax`;
        });
      });
    });
  } catch (error) {
    console.warn('Falha ao limpar cookies visíveis do navegador.', error);
  }
}

function shouldPreserveLocalStorageKeyOnLogout(key) {
  const name = String(key || '');

  // Mantém somente o cache operacional versionado e escopado por usuário/perfil/cliente.
  // Esse cache é o que evita o próximo login ficar preso esperando Smartsheet/Supabase ao vivo.
  // Dados de sessão, rascunhos, alertas, filtros e usuários locais continuam sendo apagados.
  if (name.startsWith(`${PROJECTS_CACHE_KEY}:`)) return true;

  // Mantém o escopo fixo BR do painel. Não contém dados de usuário e evita reconfiguração desnecessária.
  if (name === 'step_operation_region') return true;

  return false;
}

function clearLocalStorageForFastLogout() {
  try {
    if (!window.localStorage) return;
    Object.keys(window.localStorage).forEach((key) => {
      if (!shouldPreserveLocalStorageKeyOnLogout(key)) {
        window.localStorage.removeItem(key);
      }
    });
  } catch (error) {
    console.warn('Falha ao limpar localStorage seletivo.', error);
  }
}

async function cleanupLogoutRuntimeWithoutBreakingFastBoot() {
  try { window.clearInterval(state.pollTimer); } catch (_) {}
  try { stopPresenceHeartbeat(); } catch (_) {}
  try { window.sessionStorage?.clear(); } catch (error) { console.warn('Falha ao limpar sessionStorage.', error); }
  clearLocalStorageForFastLogout();
  clearClientVisibleCookies();

  // Importante: não remove Service Worker nem Cache Storage do app.
  // Apagar esses caches deixava o próximo login lento e sem base local caso a API demorasse.
  // A atualização limpa visual é feita pelo cache-buster da URL e pelos fetches /api em no-store.
}

function reloadAfterLogoutCleanup() {
  try {
    const target = new URL(window.location.origin + window.location.pathname);
    target.searchParams.set('logout', '1');
    target.searchParams.set('clean', String(Date.now()));
    window.location.replace(target.toString());
  } catch (_) {
    window.location.href = `/?logout=1&clean=${Date.now()}`;
  }
}

async function handleLogout() {
  const previousLogoutLabel = logoutButtonEl?.textContent || 'Sair';
  if (logoutButtonEl) {
    logoutButtonEl.disabled = true;
    logoutButtonEl.textContent = 'Saindo...';
  }

  try {
    await fetch('/api/auth-logout', {
      method: 'POST',
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Cache-Control': 'no-store' },
    });
  } catch (error) {
    console.warn('Falha ao avisar o servidor sobre logout. A limpeza local será executada mesmo assim.', error);
  }

  state.user = null;
  resetClientBspOverridesState();
  state.loadingProjectsRequest = null;
  state.lastProjectsFetchAt = 0;
  state.lastManualAlertsFetchAt = 0;
  state.lastAlertResponsesFetchAt = 0;
  state.lastStageUpdatesFetchAt = 0;
  state.lastAdminDataFetchAt = 0;
  state.manualAlerts = [];
  state.projectSignals = [];
  state.alertResponses = [];
  state.stageUpdates = [];
  state.stageDrafts = {};
  state.stageBatchValidationMode = false;
  state.stageSelectedIds = [];
  state.stageDatePendencies = [];
  state.stageDatePendingLoaded = false;
  state.stageDatePendingLoading = false;
  state.stageTrackingSubmitting = false;
  state.stageDateSelectedIds = [];
  state.attentionPopupQueue = [];
  state.attentionPopupCurrent = null;
  state.incomingAlertState = { manual: { initialized: false, ids: [] }, projectSignals: { initialized: false, ids: [] }, automatic: { initialized: false, ids: [] }, stageUpdates: { initialized: false, ids: [] } };

  await cleanupLogoutRuntimeWithoutBreakingFastBoot();
  resetDashboardForLoggedOutState();
  updateSessionUi();
  if (logoutButtonEl) {
    logoutButtonEl.textContent = previousLogoutLabel;
    logoutButtonEl.disabled = false;
  }
  reloadAfterLogoutCleanup();
}




function waitStageUpdatesRetry(ms) {
  return new Promise((resolve) => setTimeout(resolve, Math.max(0, Number(ms) || 0)));
}

async function fetchStageUpdatesPayloadWithRetry(attempts = 3) {
  let lastError = null;
  const maxAttempts = Math.max(1, Number(attempts) || 3);
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await fetch('/api/stage-updates', { credentials: 'same-origin', cache: 'no-store' });
      const raw = await response.text().catch(() => '');
      let data = null;
      try { data = raw ? JSON.parse(raw) : null; } catch (_) { data = null; }

      if (response.ok && data?.ok) return data;

      const detail = data?.error
        || data?.warning
        || (raw && !raw.trim().startsWith('<') ? raw.slice(0, 180) : '')
        || 'Falha ao carregar apontamentos setoriais.';
      const error = new Error(response.status === 401
        ? 'Sessão expirada. Faça login novamente para carregar os apontamentos.'
        : detail);
      error.status = response.status;
      error.offline = Boolean(data?.offline);
      error.payload = data;
      throw error;
    } catch (error) {
      lastError = error;
      const status = Number(error?.status || 0);
      const canRetry = error?.offline || !status || status >= 500;
      if (!canRetry || attempt >= maxAttempts) break;
      await waitStageUpdatesRetry(500 + attempt * 650);
    }
  }
  throw lastError || new Error('Falha ao carregar apontamentos setoriais.');
}


function mergeStageUpdatesById(incoming = []) {
  const current = Array.isArray(state.stageUpdates) ? state.stageUpdates : [];
  const incomingList = Array.isArray(incoming) ? incoming : [];
  if (!incomingList.length) return current;
  const incomingById = new Map(incomingList.map((item) => [String(item?.id || ''), item]).filter(([id]) => id));
  const merged = current.map((item) => incomingById.has(String(item?.id || '')) ? { ...item, ...incomingById.get(String(item?.id || '')) } : item);
  for (const item of incomingList) {
    const id = String(item?.id || '');
    if (id && !current.some((row) => String(row?.id || '') === id)) merged.unshift(item);
  }
  return merged;
}

function shouldValidateStageTrackingItem(item) {
  if (!item || !isPendingStageStatus(item.status) || isReviewStageStatus(item.status)) return false;
  if (item.trackingMatched === true) return false;
  const status = String(item.trackingStatus || '').trim().toLowerCase();
  return !status || ['checking', 'pending_check', 'timeout', 'not_found', 'waiting'].includes(status);
}

async function validateStageTrackingInBackground() {
  if (!state.user || !canValidateStageWorkspace()) return;
  if (state.stageTrackingValidationRunning) return;
  const now = Date.now();
  if (state.lastStageTrackingValidationAt && now - state.lastStageTrackingValidationAt < STAGE_TRACKING_VALIDATION_COOLDOWN_MS) return;
  state.lastStageTrackingValidationAt = now;
  const pending = (Array.isArray(state.stageUpdates) ? state.stageUpdates : [])
    .filter(shouldValidateStageTrackingItem)
    .slice(0, 25);
  if (!pending.length) return;

  state.stageTrackingValidationRunning = true;
  const previousWarning = state.stageUpdatesWarning || '';
  if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
    state.stageUpdatesWarning = 'Validando Tracking em segundo plano. A tela permanece aberta mesmo se o Smartsheet demorar.';
    renderStageUpdatesModal();
  }

  try {
    const ids = pending.map((item) => encodeURIComponent(String(item.id || ''))).join(',');
    const response = await fetch(`/api/stage-updates?mode=tracking-check&ids=${ids}`, { credentials: 'same-origin', cache: 'no-store' });
    const raw = await response.text().catch(() => '');
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch (_) { data = null; }
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || data?.warning || raw || 'Falha ao validar Tracking em segundo plano.');
    }
    const incomingUpdates = Array.isArray(data.updates) ? data.updates : [];
    state.stageUpdates = mergeStageUpdatesById(incomingUpdates);
    const processedIds = new Set(incomingUpdates
      .filter((item) => isStageUpdateEffectivelyProcessed(item))
      .map((item) => String(item?.id || ''))
      .filter(Boolean));
    if (processedIds.size) {
      setStageSelection((state.stageSelectedIds || []).filter((id) => !processedIds.has(String(id || ''))));
    }
    const autoResolvedCount = Number(data.autoResolvedCount || 0);
    state.stageUpdatesWarning = data.warning || (autoResolvedCount ? `${autoResolvedCount} apontamento(s) já conferido(s) no Tracking foram removidos da fila PCP.` : '');
  } catch (error) {
    state.stageUpdatesWarning = error.message || previousWarning || 'O Smartsheet demorou para validar o Tracking. A lista permanece disponível.';
  } finally {
    state.stageTrackingValidationRunning = false;
    state.lastStageUpdatesFetchAt = Date.now();
    if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      renderStageUpdatesModal();
    }
  }
}

async function loadStageUpdates(options = {}) {
  if (!state.user) {
    state.stageUpdates = [];
    return;
  }
  if (!options.force && !isStageUpdatesWorkspaceOpen()) {
    return;
  }
  if (options.background && shouldSkipBackgroundRequest(options)) return;
  const now = Date.now();
  if (!options.force && options.background && now - state.lastStageUpdatesFetchAt < ALERTS_REFRESH_MS) return;
  try {
    const data = await fetchStageUpdatesPayloadWithRetry(options.force ? 3 : 2);
    state.lastStageUpdatesFetchAt = Date.now();
    state.stageUpdates = Array.isArray(data.updates) ? data.updates : [];
    if (data.warning && stageUpdatesContentEl && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      state.stageUpdatesWarning = data.warning;
    } else {
      state.stageUpdatesWarning = '';
    }
    if (normalizeSectorValue(state.user?.sector) === 'pcp') {
      const pendingForPcp = state.stageUpdates.filter((item) => isPendingStageStatus(item?.status));
      syncIncomingAlerts('stageUpdates', pendingForPcp);
    }
    if (data?.trackingValidationDeferred && canValidateStageWorkspace()) {
      window.setTimeout(() => validateStageTrackingInBackground().catch(() => {}), 250);
    }
  } catch (error) {
    // v35.7: não bloquear o apontamento do setor quando a consulta do histórico/validação falhar.
    // O setor ainda precisa conseguir abrir a tela, buscar a BSP e enviar o apontamento.
    if (!Array.isArray(state.stageUpdates) || !state.stageUpdates.length || !options.background) {
      state.stageUpdates = [];
    }
    state.stageUpdatesWarning = error.message || 'Histórico de apontamentos indisponível no momento; você ainda pode lançar novos apontamentos.';
    state.lastStageUpdatesFetchAt = Date.now();
    if (stageUpdatesContentEl && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      renderStageUpdatesModal();
      const warning = document.createElement('div');
      warning.className = 'stage-inline-warning';
      warning.textContent = state.stageUpdatesWarning;
      stageUpdatesContentEl.prepend(warning);
    }
    return;
  }
}

function renderStageSectorWorkspace() {
  if (!stageUpdatesContentEl) return;
  const sector = getStageWorkspaceSector();
  const stageLabel = getStageWorkspaceLabel(sector);
  const isPcpPointing = isPcpStageUser() && state.stagePcpPointingMode;
  if (isPcpPointing && !sector) {
    stageUpdatesContentEl.innerHTML = `
      <div class="stage-workspace-shell">
        ${state.stageUpdatesWarning ? `<div class="stage-inline-warning">${escapeHtml(state.stageUpdatesWarning)}</div>` : ''}
        <section class="admin-card admin-card--wide">
          <div class="admin-card-head"><h4>Modo PCP de apontamento</h4></div>
          <label class="stack-field">
            <span>Apontar como setor</span>
            <select data-pcp-stage-sector-switch="true"><option value="">Selecione o setor</option>${getStageSectorOptionsHtml(state.pcpStageSelectedSector)}</select>
          </label>
          <div class="stage-row-actions"><button class="ghost-button" type="button" data-stage-back-validation="true">Voltar para Validação PCP</button></div>
        </section>
      </div>`;
    return;
  }
  const matchedProjects = stageWorkspaceSearchProjects();
  const blockedInfo = getStageWorkspaceBlockedInfo();
  const myUpdates = getMyStageUpdates();
  const resolvedMine = myUpdates.filter((item) => isResolvedStageStatus(item.status)).slice(0, 10);
  const draftEntries = getStageDraftEntries(sector);
  const readyDrafts = getReadyStageDraftEntries(sector);
  stageUpdatesContentEl.innerHTML = `
    <div class="stage-workspace-shell">
      ${state.stageUpdatesWarning ? `<div class="stage-inline-warning">${escapeHtml(state.stageUpdatesWarning)}</div>` : ''}
      ${isPcpPointing ? `
      <section class="admin-card admin-card--wide stage-pcp-pointing-card">
        <div class="admin-card-head">
          <h4>Modo PCP de apontamento</h4>
          <span class="stage-badge stage-badge--sector">Apontando como ${escapeHtml(stageLabel)}</span>
        </div>
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Apontar como setor</span>
            <select data-pcp-stage-sector-switch="true">${getStageSectorOptionsHtml(sector)}</select>
          </label>
          <div class="stage-row-actions">
            <button class="ghost-button" type="button" data-stage-back-validation="true">Voltar para Validação PCP</button>
          </div>
        </div>
        <div class="stage-validation-note">O PCP está lançando apontamentos em nome do setor selecionado. A lista continua respeitando a competência da etapa atual do spool.</div>
      </section>` : ''}
      <section class="admin-card admin-card--wide">
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Buscar BSP / cliente</span>
            <input type="text" data-stage-search="true" value="${escapeHtml(state.stageUpdatesSearchQuery || '')}" placeholder="Ex.: BSP 25-732-03 ou BSP2573203" autocomplete="off" inputmode="search" />
          </label>
          <div class="stage-muted">${isPcpPointing ? 'Fila selecionada pelo PCP' : 'Etapa atual do seu login'}: <strong>${escapeHtml(stageLabel)}</strong></div>
        </div>
        <div class="stage-bulk-bar">
          <div class="stage-muted">Rascunhos salvos: <strong>${draftEntries.length}</strong> • Prontos para envio: <strong>${readyDrafts.length}</strong></div>
          <div class="stage-row-actions">
            <button class="ghost-button" type="button" data-stage-clear-drafts="true" ${draftEntries.length ? '' : 'disabled'}>Limpar rascunhos</button>
            <button class="primary-button" type="button" data-stage-bulk-send="true" ${(readyDrafts.length && !state.stageBulkSubmitting) ? '' : 'disabled'}>${state.stageBulkSubmitting ? 'Enviando lote...' : 'Enviar em massa'}</button>
          </div>
        </div>
      </section>
      <div class="stage-two-col">
        <section class="admin-card admin-card--wide">
          <div class="admin-card-head"><h4>Lançar avanço da etapa</h4></div>
          ${matchedProjects.length ? `<div class="stage-project-list">${matchedProjects.map((project) => {
            const spools = Array.isArray(project.spools) ? project.spools : [];
            return `
              <article class="stage-project-card">
                <div class="stage-project-head">
                  <div>
                    <strong>${escapeHtml(project.projectDisplay || project.projectNumber || 'Projeto')}</strong>
                    <div class="stage-update-meta">
                      <span class="stage-badge">${escapeHtml(project.client || 'Sem cliente')}</span>
                      <span class="stage-badge">${spools.length} liberado(s)${Number(project.stageWorkspaceTotalSpools || spools.length) > spools.length ? ` de ${Number(project.stageWorkspaceTotalSpools || spools.length)} spool(s)` : ''}</span>
                    </div>
                  </div>
                </div>
                <div class="table-shell">
                  <table class="stage-inline-table">
                    <thead><tr><th>Spool</th><th>Descrição</th><th>Etapa atual</th><th>Andamento</th><th>Data conclusão</th><th>Obs.</th><th>Ação</th></tr></thead>
                    <tbody>
                      ${spools.map((spool) => {
                        const projectRowId = project.rowId || project.rowNumber;
                        const pending = getPendingStageUpdate(projectRowId, spool.iso, sector);
                        const lastResolved = getLatestResolvedStageUpdate(projectRowId, spool.iso, sector);
                        const submitKey = `${String(projectRowId || '').trim()}::${String(spool.iso || '').trim().toLowerCase()}::${String(sector || '').trim().toLowerCase()}`;
                        const isSubmitting = Boolean(state.stageSubmittingKeys?.[submitKey]);
                        const draft = getStageDraft(projectRowId, spool.iso, sector) || {};
                        return `
                          <tr>
                            <td>${escapeHtml(spool.iso || '—')}</td>
                            <td>${escapeHtml(spool.description || '—')}</td>
                            <td><span class="stage-badge stage-badge--sector">${escapeHtml(getSpoolStageLabel(project, spool))}</span></td>
                            <td>
                              <div class="stage-row-form" data-stage-update-form="true" data-project-row-id="${escapeHtml(String(projectRowId || ''))}" data-project-number="${escapeHtml(project.projectNumber || '')}" data-project-display="${escapeHtml(project.projectDisplay || project.projectNumber || '')}" data-client="${escapeHtml(project.client || '')}" data-spool-iso="${escapeHtml(spool.iso || '')}" data-spool-description="${escapeHtml(spool.description || spool.drawing || '')}" data-spool-stage="${escapeHtml(getSpoolStageLabel(project, spool))}" data-stage-sector="${escapeHtml(sector || '')}">
                                <select name="progress" data-stage-progress="true" ${pending || isSubmitting ? 'disabled' : ''}>
                                  <option value="">Selecione</option>
                                  ${STAGE_PROGRESS_OPTIONS.map((value) => `<option value="${value}" ${Number(draft.progress || 0) === Number(value) ? 'selected' : ''}>${value}%</option>`).join('')}
                                </select>
                              </div>
                            </td>
                            <td><input type="date" name="completionDate" value="${escapeHtml(draft.completionDate || '')}" ${pending || isSubmitting ? 'disabled' : ''} /></td>
                            <td><textarea name="note" rows="2" placeholder="Observação opcional" ${pending || isSubmitting ? 'disabled' : ''}>${escapeHtml(draft.note || '')}</textarea></td>
                            <td>
                              ${pending
                                ? `<span class="stage-badge ${isReviewStageStatus(pending.status) ? 'stage-badge--review' : 'stage-badge--sent'}">${escapeHtml(stageUpdateActionLabel(pending.status))}</span>`
                                : isSubmitting
                                  ? `<button class="primary-button" type="button" disabled>Enviando...</button>`
                                  : `<div class="stage-row-actions"><button class="primary-button" type="button" data-stage-send="true">Enviar</button><button class="ghost-button stage-review-button" type="button" data-stage-review="true">Revisão</button></div>`}
                              ${lastResolved ? `<div class="stage-muted">Último ${escapeHtml(isReviewStageStatus(lastResolved.status) ? 'retorno de revisão' : 'avanço concluído')}: ${escapeHtml(formatStageDate(lastResolved.resolvedAt))}</div>` : ''}
                            </td>
                          </tr>`;
                      }).join('')}
                    </tbody>
                  </table>
                </div>
              </article>`;
          }).join('')}</div>` : `<div class="empty-state">${blockedInfo.count ? `Esta BSP existe, mas nenhum spool está liberado para o setor ${escapeHtml(stageLabel)}. ${blockedInfo.first ? `Etapa atual encontrada: ${escapeHtml(blockedInfo.first.stage)} • Responsável: ${escapeHtml(sectorLabel(blockedInfo.first.sector))}.` : ''}` : 'Pesquise a BSP para visualizar somente os spools liberados para a sua etapa.'}</div>`}
        </section>
        <section class="admin-card admin-card--wide">
          <div class="admin-card-head"><h4>Meus lançamentos</h4></div>
          <div class="stage-history-shell">
            <div class="stage-update-list">
              ${myUpdates.length ? myUpdates.map((item) => `
                <article class="stage-update-card">
                  <div class="stage-update-head">
                    <div>
                      <strong>${escapeHtml(item.projectDisplay || item.projectNumber || 'Projeto')} • ${escapeHtml(item.spoolIso || 'Spool')}</strong>
                      <div class="stage-update-meta">
                        <span class="stage-badge stage-badge--sector">${escapeHtml(sectorLabel(item.sector))}</span>
                        <span class="stage-badge ${isResolvedStageStatus(item.status) ? (isReviewStageStatus(item.status) ? 'stage-badge--review-resolved' : 'stage-badge--resolved') : (isReviewStageStatus(item.status) ? 'stage-badge--review' : 'stage-badge--sent')}">${escapeHtml(isResolvedStageStatus(item.status) ? stageUpdateResolveLabel(item.status) : stageUpdateActionLabel(item.status))}</span>
                        <span class="stage-badge">${escapeHtml(String(item.progress || 0))}%</span>
                        ${stageTrackingBadgeHtml(item)}
                      </div>
                    </div>
                  </div>
                  <p>${escapeHtml(item.note || 'Sem observação.')}</p>
                  <div class="stage-muted">Enviado em: ${escapeHtml(formatStageDate(item.createdAt))}</div>
                  ${isResolvedStageStatus(item.status) ? `<div class="stage-muted">${escapeHtml(stageUpdateResolveLabel(item.status))} por ${escapeHtml(item.resolvedByName || item.resolvedBy || 'PCP')} em ${escapeHtml(formatStageDate(item.resolvedAt))}</div>${item.resolutionNote ? `<div class="response-bubble response-bubble--admin"><strong>${escapeHtml(isReviewStageStatus(item.status) ? 'Tratativa PCP' : 'Fechamento PCP')}</strong><p>${escapeHtml(item.resolutionNote)}</p></div>` : ''}` : ''}
                </article>`).join('') : `<div class="empty-state">Nenhum apontamento enviado ainda.</div>`}
            </div>
          </div>
        </section>
      </div>
    </div>`;
}

function shouldOpenStageValidationWorkspaceFromUrl() {
  try {
    const params = new URLSearchParams(window.location.search || '');
    return params.get('stageWorkspace') === '1' || window.location.hash === '#stage-validation';
  } catch {
    return false;
  }
}

function openStageValidationWorkspaceInline() {
  if (!canValidateStageWorkspace()) return;

  state.stageUpdatesSearchQuery = '';
  state.stagePcpPointingMode = false;
  syncStageDraftsForCurrentSector();

  // Mantém a Validação PCP na aba atual.
  // Antes era usado window.open(...), o que criava uma nova página a cada clique.
  try {
    if (!shouldOpenStageValidationWorkspaceFromUrl()) {
      const url = new URL(window.location.href);
      url.searchParams.set('stageWorkspace', '1');
      url.hash = 'stage-validation';
      window.history.replaceState({}, '', url.toString());
    }
  } catch {}

  openStageUpdatesModal({ loading: true });

  loadStageUpdates()
    .then(() => {
      if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
        renderStageUpdatesModal();
      }
    })
    .catch((error) => {
      if (stageUpdatesContentEl && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
        stageUpdatesContentEl.innerHTML = `<div class="empty-state">${escapeHtml(error?.message || 'Falha ao carregar apontamentos setoriais.')}</div>`;
      }
    });
}

// Mantido como alias para evitar referência quebrada em versões antigas do HTML/cache.
function openStageValidationInNewTab() {
  openStageValidationWorkspaceInline();
}

function getFilteredStageUpdatesForValidation() {
  const query = String(state.stageUpdatesSearchQuery || '').trim();
  const all = Array.isArray(state.stageUpdates) ? state.stageUpdates : [];
  return all.filter((item) => {
    if (!query) return true;
    return matchesFlexibleSearch([
      item.projectNumber,
      item.projectDisplay,
      item.client,
      item.spoolIso,
      item.spoolDescription,
      item.sector,
      sectorLabel(item.sector),
      item.createdByName,
      item.createdBy,
    ], query);
  }).sort((a,b)=> new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
}

function isAdvanceStageUpdate(item) {
  return !isReviewStageStatus(item?.status);
}

function isStageUpdateSelectableForTracking(item) {
  return Boolean(item && isPendingStageStatus(item.status) && isAdvanceStageUpdate(item));
}

function getSelectedVisibleStageIds(items = []) {
  const visibleIds = new Set(items.filter(isStageUpdateSelectableForTracking).map((item) => String(item.id || '')));
  return (Array.isArray(state.stageSelectedIds) ? state.stageSelectedIds : []).filter((id) => visibleIds.has(String(id)));
}

function setStageSelection(ids = []) {
  state.stageSelectedIds = Array.from(new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean)));
}

function setStageDateSelection(ids = []) {
  state.stageDateSelectedIds = Array.from(new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean)));
}

function trackingActionLabel(item, rewrite = false) {
  if (rewrite) return 'Regravar Tracking';
  const info = getStageTrackingInfo(item);
  if (info.current != null && info.current > Number(item?.progress || 0)) return 'Confirmar avanço superior';
  return 'Atualizar Tracking';
}

function stageTrackingMessageFromResult(result) {
  const rows = Number(result?.rowCount || 0);
  const column = result?.progressColumn ? ` • ${result.progressColumn}` : '';
  const message = result?.message || 'Tracking processado.';
  return `${message}${rows ? ` (${rows} linha(s) localizada(s)${column})` : ''}`;
}

function renderStageValidationPendingTable(pending = []) {
  const selectable = pending.filter(isStageUpdateSelectableForTracking);
  const selected = new Set(getSelectedVisibleStageIds(pending));
  const allSelected = selectable.length > 0 && selectable.every((item) => selected.has(String(item.id || '')));
  if (!pending.length) return '<div class="empty-state">Nenhum apontamento pendente no momento.</div>';
  return `
    <div class="table-shell stage-validation-table-shell">
      <table class="stage-inline-table stage-validation-table">
        <thead>
          <tr>
            <th class="stage-check-cell"><input type="checkbox" data-stage-master-check="true" ${allSelected ? 'checked' : ''} ${selectable.length ? '' : 'disabled'} aria-label="Selecionar todos os apontamentos visíveis" /></th>
            <th>BSP / Spool</th>
            <th>Setor</th>
            <th>Avanço</th>
            <th>Tracking</th>
            <th>Enviado por</th>
            <th>Observação</th>
            <th>Ação</th>
          </tr>
        </thead>
        <tbody>
          ${pending.map((item) => {
            const id = String(item.id || '');
            const selectableItem = isStageUpdateSelectableForTracking(item);
            const info = getStageTrackingInfo(item);
            const canConcludeOk = isReviewStageStatus(item.status) || (selectableItem && info.matched);
            const rewriteButton = selectableItem && info.matched
              ? `<button class="ghost-button" type="button" data-stage-tracking-rewrite="${escapeHtml(id)}">${escapeHtml(trackingActionLabel(item, true))}</button>`
              : '';
            const updateButton = selectableItem && !info.matched
              ? `<button class="primary-button" type="button" data-stage-tracking-update="${escapeHtml(id)}">${escapeHtml(trackingActionLabel(item, false))}</button>`
              : '';
            const concludeButton = canConcludeOk
              ? `<button class="ghost-button" type="button" data-stage-conclude="${escapeHtml(id)}">${escapeHtml(isReviewStageStatus(item.status) ? 'Concluir revisão' : 'Concluir OK')}</button>`
              : '';
            const deleteButton = isPendingStageStatus(item.status)
              ? `<button class="ghost-button stage-danger-button" type="button" data-stage-delete="${escapeHtml(id)}">Remover pendência</button>`
              : '';
            return `
              <tr data-stage-row-id="${escapeHtml(id)}" class="${info.matched ? 'stage-row--ok' : ''}">
                <td class="stage-check-cell"><input type="checkbox" data-stage-item-check="${escapeHtml(id)}" ${selected.has(id) ? 'checked' : ''} ${selectableItem ? '' : 'disabled'} aria-label="Selecionar apontamento" /></td>
                <td><strong>${escapeHtml(item.projectDisplay || item.projectNumber || 'Projeto')}</strong><br><span class="stage-muted">${escapeHtml(item.spoolIso || 'Spool')}</span></td>
                <td>${escapeHtml(sectorLabel(item.sector))}<br><span class="stage-badge ${isReviewStageStatus(item.status) ? 'stage-badge--review' : 'stage-badge--pending'}">${escapeHtml(stageUpdatePendingLabel(item.status))}</span></td>
                <td><strong>${escapeHtml(String(item.progress || 0))}%</strong>${Number(item.progress || 0) === 100 ? `<br><span class="stage-muted">Data: ${escapeHtml(item.completionDate || 'data atual')}</span>` : ''}</td>
                <td>${stageTrackingBadgeHtml(item)}</td>
                <td>${escapeHtml(item.createdByName || item.createdBy || 'Usuário')}<br><span class="stage-muted">${escapeHtml(formatStageDate(item.createdAt))}</span></td>
                <td>${escapeHtml(item.note || '—')}</td>
                <td><div class="stage-row-actions stage-row-actions--stack">${updateButton}${rewriteButton}${concludeButton || `<span class="stage-muted">Atualize o Tracking primeiro</span>`}${deleteButton}</div></td>
              </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;
}

function renderStageHistoryList(history = []) {
  if (!history.length) return '<div class="empty-state">Nenhum histórico validado encontrado.</div>';
  return `<div class="stage-update-list stage-history-list">${history.map((item) => `
    <article class="stage-update-card">
      <div class="stage-update-head">
        <div>
          <strong>${escapeHtml(item.projectDisplay || item.projectNumber || 'Projeto')} • ${escapeHtml(item.spoolIso || 'Spool')}</strong>
          <div class="stage-update-meta">
            <span class="stage-badge stage-badge--sector">${escapeHtml(sectorLabel(item.sector))}</span>
            <span class="stage-badge ${isReviewStageStatus(item.status) ? 'stage-badge--review-resolved' : 'stage-badge--resolved'}">${escapeHtml(stageUpdateResolveLabel(item.status))}</span>
            <span class="stage-badge">${escapeHtml(String(item.progress || 0))}%</span>
          </div>
        </div>
      </div>
      <div class="stage-muted">Informado por ${escapeHtml(item.createdByName || item.createdBy || 'Usuário')} • ${escapeHtml(isReviewStageStatus(item.status) ? 'revisão tratada' : 'validado')} por ${escapeHtml(item.resolvedByName || item.resolvedBy || 'PCP')} em ${escapeHtml(formatStageDate(item.resolvedAt))}</div>
      ${item.resolutionNote ? `<div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(item.resolutionNote)}</p></div>` : ''}
    </article>`).join('')}</div>`;
}

function renderStageDatePendencies() {
  if (state.stageDatePendingLoading) return '<div class="empty-state">Carregando pendências de datas do histórico...</div>';
  if (!state.stageDatePendingLoaded) return '<div class="empty-state">Clique em “Pendências de datas do histórico” para verificar somente apontamentos 100% já validados pelo app.</div>';
  const items = Array.isArray(state.stageDatePendencies) ? state.stageDatePendencies : [];
  if (!items.length) return '<div class="empty-state">Nenhuma pendência de data encontrada no histórico validado do app.</div>';
  const selected = new Set(state.stageDateSelectedIds || []);
  const allSelected = items.length > 0 && items.every((item) => selected.has(String(item.id || '')));
  return `
    <div class="table-shell stage-validation-table-shell">
      <table class="stage-inline-table stage-validation-table">
        <thead>
          <tr>
            <th class="stage-check-cell"><input type="checkbox" data-stage-date-master-check="true" ${allSelected ? 'checked' : ''} aria-label="Selecionar todas as pendências visíveis" /></th>
            <th>BSP / Spool</th>
            <th>Processo</th>
            <th>Data faltante</th>
            <th>Data aplicada</th>
            <th>Linhas</th>
            <th>Ação</th>
          </tr>
        </thead>
        <tbody>
          ${items.map((item) => {
            const id = String(item.id || '');
            return `
              <tr>
                <td class="stage-check-cell"><input type="checkbox" data-stage-date-item-check="${escapeHtml(id)}" ${selected.has(id) ? 'checked' : ''} /></td>
                <td><strong>${escapeHtml(item.projectDisplay || item.projectNumber || 'Projeto')}</strong><br><span class="stage-muted">${escapeHtml(item.spoolIso || 'Spool')}</span></td>
                <td>${escapeHtml(item.process || '—')}${item.needsPaintingNextSteps ? '<br><span class="stage-badge stage-badge--tracking-waiting">Pintura 100% + próximas etapas 25%</span>' : ''}</td>
                <td>${escapeHtml(item.missingDateColumn || '—')}</td>
                <td>${escapeHtml(item.applyDate || 'data atual')}</td>
                <td>${escapeHtml(String(item.affectedRows || item.rowCount || 0))}/${escapeHtml(String(item.rowCount || 0))}</td>
                <td><button class="primary-button" type="button" data-stage-date-fix="${escapeHtml(id)}">Corrigir</button></td>
              </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>`;
}

function renderStageValidationWorkspace() {
  if (!stageUpdatesContentEl) return;
  const filtered = getFilteredStageUpdatesForValidation();
  // Na Validação PCP, a fila deve mostrar apenas apontamentos realmente pendentes.
  // Quando a conferência do Tracking já encontrou o mesmo avanço aplicado
  // (trackingMatched/trackingStatus='matched'), o registro não deve continuar
  // aparecendo como pendência visual. Ele permanece armazenado, mas sai da fila
  // operacional para não confundir o PCP com itens já atualizados.
  const pending = filtered.filter((item) => isPendingStageStatus(item.status) && !isStageUpdateEffectivelyProcessed(item));
  const history = filtered.filter((item) => isResolvedStageStatus(item.status)).slice(0, 80);
  const selectable = pending.filter(isStageUpdateSelectableForTracking);
  const selectedIds = getSelectedVisibleStageIds(pending);
  const selectedDateIds = (state.stageDateSelectedIds || []).filter((id) => (state.stageDatePendencies || []).some((item) => String(item.id) === String(id)));
  state.stageSelectedIds = selectedIds;
  state.stageDateSelectedIds = selectedDateIds;
  const submitting = Boolean(state.stageTrackingSubmitting);

  stageUpdatesContentEl.innerHTML = `
    <div class="stage-workspace-shell stage-validation-workspace" id="stage-validation">
      ${state.stageUpdatesWarning ? `<div class="stage-inline-warning">${escapeHtml(state.stageUpdatesWarning)}</div>` : ''}
      <section class="admin-card admin-card--wide stage-validation-header-card">
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Buscar BSP / spool / setor</span>
            <input type="text" data-stage-search="true" value="${escapeHtml(state.stageUpdatesSearchQuery || '')}" placeholder="Ex.: BSP 25-732-03 ou BSP2573203" autocomplete="off" inputmode="search" />
          </label>
          <div class="stage-row-actions">
            <div class="stage-muted">Pendentes: <strong>${pending.length}</strong> • Selecionados: <strong>${selectedIds.length}</strong> • Histórico: <strong>${history.length}</strong></div>
            <button class="ghost-button" type="button" data-stage-load-date-pending="true" ${state.stageDatePendingLoading ? 'disabled' : ''}>${state.stageDatePendingLoading ? 'Verificando...' : 'Pendências de datas do histórico'}</button>
            <button class="primary-button" type="button" data-stage-tracking-bulk="true" ${selectedIds.length && !submitting ? '' : 'disabled'}>${submitting ? 'Atualizando...' : 'Atualizar/Regravar selecionados'}</button>
            <button class="ghost-button" type="button" data-stage-conclude-bulk-ok="true" ${pending.length && !submitting ? '' : 'disabled'}>Concluir lote OK</button>
          </div>
        </div>
        <div class="stage-validation-note">
          A atualização usa a API do Smartsheet com percentuais numéricos: 25% = 0.25, 50% = 0.5, 75% = 0.75 e 100% = 1. O apontamento só sai dos pendentes após confirmação do Tracking.
        </div>
      </section>

      ${isPcpStageUser() ? `
      <section class="admin-card admin-card--wide stage-pcp-pointing-card">
        <div class="admin-card-head">
          <h4>Apontamento pelo PCP</h4>
          <div class="stage-muted">Use quando o PCP precisar apontar demandas de setores como Solda, Pintura, Produção ou Logística.</div>
        </div>
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Apontar como setor</span>
            <select data-pcp-stage-sector-select="true">${getStageSectorOptionsHtml(ensurePcpStageSectorDefault())}</select>
          </label>
          <div class="stage-row-actions">
            <button class="primary-button" type="button" data-stage-open-pcp-pointing="true">Abrir fila para apontamento</button>
          </div>
        </div>
        <div class="stage-validation-note">Após selecionar o setor, o app mostrará somente os spools liberados para aquela competência. O apontamento será registrado pelo usuário PCP em nome do setor escolhido.</div>
      </section>` : ''}

      <section class="admin-card admin-card--wide">
        <div class="admin-card-head">
          <h4>Validação PCP dos apontamentos</h4>
          <div class="stage-muted">Itens elegíveis para lote: ${selectable.length}</div>
        </div>
        ${renderStageValidationPendingTable(pending)}
      </section>

      <section class="admin-card admin-card--wide">
        <div class="admin-card-head">
          <h4>Pendências de datas do histórico</h4>
          <div class="stage-row-actions">
            <span class="stage-muted">Selecionadas: <strong>${selectedDateIds.length}</strong></span>
            <button class="primary-button" type="button" data-stage-date-bulk="true" ${selectedDateIds.length && !submitting ? '' : 'disabled'}>Corrigir selecionadas</button>
            <button class="ghost-button" type="button" data-stage-date-fix-all="true" ${(state.stageDatePendencies || []).length && !submitting ? '' : 'disabled'}>Corrigir em massa</button>
          </div>
        </div>
        ${renderStageDatePendencies()}
      </section>

      <section class="admin-card admin-card--wide">
        <div class="admin-card-head"><h4>Histórico validado</h4></div>
        ${renderStageHistoryList(history)}
      </section>
    </div>`;
}


function renderStageUpdatesModal() {
  if (!stageUpdatesContentEl) return;
  if (isPcpStageUser() && state.stagePcpPointingMode) {
    renderStageSectorWorkspace();
    return;
  }
  if (canValidateStageWorkspace()) {
    renderStageValidationWorkspace();
    return;
  }
  renderStageSectorWorkspace();
}

function openStageUpdatesModal(options = {}) {
  if (!stageUpdatesModalEl) return;
  const titleEl = document.getElementById('stage-updates-title');
  const subtitleEl = document.getElementById('stage-updates-subtitle');
  const isPcpPointing = isPcpStageUser() && state.stagePcpPointingMode;
  if (titleEl) titleEl.textContent = isPcpPointing
    ? `Apontamento PCP • ${getStageWorkspaceLabel()}`
    : (canValidateStageWorkspace() ? 'Validação PCP dos apontamentos' : `Apontamentos da etapa • ${getStageWorkspaceLabel()}`);
  if (subtitleEl) subtitleEl.textContent = isPcpPointing
    ? 'PCP apontando em nome do setor selecionado, respeitando a competência da etapa atual.'
    : (canValidateStageWorkspace()
      ? 'Conclua os registros validados para que saiam da fila e permaneçam no histórico.'
      : 'Cada setor informa somente a sua própria etapa por spool. O PCP valida e mantém o histórico.');

  if (options.loading && stageUpdatesContentEl) {
    stageUpdatesContentEl.innerHTML = `
      <div class="stage-workspace-shell">
        <section class="admin-card admin-card--wide">
          <div class="empty-state">Carregando apontamentos... a tela já está aberta e os dados serão filtrados automaticamente.</div>
        </section>
      </div>`;
  } else {
    renderStageUpdatesModal();
  }

  stageUpdatesModalEl.classList.remove('hidden');
  stageUpdatesModalEl.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
}

function closeStageUpdatesModal() {
  if (!stageUpdatesModalEl) return;
  stageUpdatesModalEl.classList.add('hidden');
  stageUpdatesModalEl.setAttribute('aria-hidden', 'true');
  if (
    modalEl.classList.contains('hidden') &&
    alertModalEl.classList.contains('hidden') &&
    sectorAlertsModalEl.classList.contains('hidden') &&
    stageUpdatesModalEl.classList.contains('hidden') &&
    adminModalEl.classList.contains('hidden') &&
    loginModalEl.classList.contains('hidden')
  ) {
    document.body.classList.remove('modal-open');
  }
}

function getStageSubmitKey(projectRowId, spoolIso, sector = '') {
  return `${String(projectRowId || '').trim()}::${String(spoolIso || '').trim().toLowerCase()}::${String(sector || '').trim().toLowerCase()}`;
}

function setStageSubmitting(projectRowId, spoolIso, sector, value) {
  const key = getStageSubmitKey(projectRowId, spoolIso, sector);
  state.stageSubmittingKeys = { ...(state.stageSubmittingKeys || {}) };
  if (value) state.stageSubmittingKeys[key] = true;
  else delete state.stageSubmittingKeys[key];
}

function persistStageDraftFromRow(rowEl) {
  const formEl = rowEl?.querySelector('[data-stage-update-form="true"]');
  if (!formEl) return;
  const projectRowId = String(formEl.dataset.projectRowId || '').trim();
  const spoolIso = String(formEl.dataset.spoolIso || '').trim();
  const sector = String(formEl.dataset.stageSector || getStageWorkspaceSector() || '').trim();
  if (!projectRowId || !spoolIso) return;
  const progress = String(formEl.querySelector('[name="progress"]')?.value || '').trim();
  const completionDate = String(rowEl.querySelector('[name="completionDate"]')?.value || '').trim();
  const note = String(rowEl.querySelector('[name="note"]')?.value || '').trim();
  const metadata = {
    projectNumber: String(formEl.dataset.projectNumber || '').trim(),
    projectDisplay: String(formEl.dataset.projectDisplay || formEl.dataset.projectNumber || '').trim(),
    client: String(formEl.dataset.client || '').trim(),
    spoolDescription: String(formEl.dataset.spoolDescription || '').trim(),
    spoolStage: String(formEl.dataset.spoolStage || '').trim(),
  };
  if (!progress && !completionDate && !note) {
    removeStageDraft(projectRowId, spoolIso, sector);
    return;
  }
  upsertStageDraft(projectRowId, spoolIso, sector, { progress, completionDate, note, ...metadata });
}

async function handleStageWorkspaceBulkSubmit() {
  const sector = getStageWorkspaceSector();
  const items = getReadyStageDraftEntries(sector).map((item) => ({
    projectRowId: item.projectRowId,
    projectNumber: item.projectNumber || '',
    projectDisplay: item.projectDisplay || item.projectNumber || '',
    client: item.client || '',
    spoolIso: item.spoolIso,
    spoolDescription: item.spoolDescription || '',
    spoolStage: item.spoolStage || '',
    sector: item.sector,
    progress: Number(item.progress || 0),
    completionDate: item.completionDate || '',
    note: item.note || '',
    actionType: item.actionType === 'review' ? 'review' : 'advance',
  }));
  if (!items.length || state.stageBulkSubmitting) return;
  state.stageBulkSubmitting = true;
  renderStageUpdatesModal();
  try {
    const response = await fetch('/api/stage-updates', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ items, sector }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao enviar lote de apontamentos.');
    const created = Array.isArray(data.updates) ? data.updates : [];
    created.forEach((item) => {
      removeStageDraft(item.projectRowId, item.spoolIso, item.sector || sector);
      state.stageUpdates = [item, ...(Array.isArray(state.stageUpdates) ? state.stageUpdates : [])];
    });
    state.stageBulkSubmitting = false;
    renderStageUpdatesModal();
    if (Array.isArray(data.errors) && data.errors.length) {
      window.alert(`Lote enviado parcialmente. Sucesso: ${created.length}. Pendências: ${data.errors.length}.`);
    }
    loadStageUpdates().then(() => renderStageUpdatesModal()).catch(() => {});
  } catch (error) {
    state.stageBulkSubmitting = false;
    renderStageUpdatesModal();
    window.alert(error.message || 'Falha ao enviar lote de apontamentos.');
  }
}

async function loadStageHistoryDatePendencies() {
  if (!canValidateStageWorkspace() || state.stageDatePendingLoading) return;
  state.stageDatePendingLoading = true;
  renderStageUpdatesModal();
  try {
    const response = await fetch('/api/stage-updates?mode=history-date-pending', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar pendências de datas do histórico.');
    state.stageDatePendencies = Array.isArray(data.pendencies) ? data.pendencies : [];
    state.stageDatePendingLoaded = true;
    setStageDateSelection([]);
  } catch (error) {
    window.alert(error.message || 'Falha ao carregar pendências de datas do histórico.');
  } finally {
    state.stageDatePendingLoading = false;
    renderStageUpdatesModal();
  }
}


function getStageUpdatesByIds(ids = []) {
  const idSet = new Set((Array.isArray(ids) ? ids : [ids]).map((id) => String(id || '').trim()).filter(Boolean));
  return (Array.isArray(state.stageUpdates) ? state.stageUpdates : []).filter((item) => idSet.has(String(item?.id || '')));
}

function isStageUpdateEffectivelyProcessed(item) {
  if (!item) return true;
  if (isResolvedStageStatus(item.status)) return true;
  if (item.trackingMatched === true) return true;
  if (String(item.trackingStatus || '').toLowerCase() === 'matched') return true;
  return false;
}

function areStageUpdateIdsProcessed(ids = []) {
  const cleanIds = (Array.isArray(ids) ? ids : [ids]).map((id) => String(id || '').trim()).filter(Boolean);
  if (!cleanIds.length) return false;
  const current = getStageUpdatesByIds(cleanIds);
  if (!current.length) return true;
  const currentById = new Map(current.map((item) => [String(item?.id || ''), item]));
  return cleanIds.every((id) => isStageUpdateEffectivelyProcessed(currentById.get(String(id))));
}

async function refreshStageUpdatesAfterAction(ids = [], options = {}) {
  try {
    await loadStageUpdates({ force: true });
    if (areStageUpdateIdsProcessed(ids)) {
      if (options.clearSelection !== false) {
        const processed = new Set((Array.isArray(ids) ? ids : [ids]).map((id) => String(id || '').trim()).filter(Boolean));
        setStageSelection((state.stageSelectedIds || []).filter((id) => !processed.has(String(id || ''))));
      }
      renderStageUpdatesModal();
      return true;
    }
  } catch (_) {}
  return false;
}

function isProbablyAlreadyProcessedStageError(error) {
  const message = String(error?.message || error || '').toLowerCase();
  return message.includes('não encontrado')
    || message.includes('nao encontrado')
    || message.includes('nenhum apontamento elegível')
    || message.includes('nenhum apontamento elegivel')
    || message.includes('já foi processado')
    || message.includes('ja foi processado')
    || message.includes('já estava')
    || message.includes('ja estava');
}

async function sendStageTrackingUpdate(ids = [], options = {}) {
  const cleanIds = Array.from(new Set((Array.isArray(ids) ? ids : [ids]).map((id) => String(id || '').trim()).filter(Boolean)));
  if (!cleanIds.length || state.stageTrackingSubmitting) return;
  state.stageTrackingSubmitting = true;
  renderStageUpdatesModal();
  try {
    const action = options.dateOnly ? 'fix-history-dates' : 'update-tracking';
    const response = await fetch('/api/stage-updates', {
      method: 'PUT',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action,
        ids: cleanIds,
        forceRewrite: Boolean(options.forceRewrite),
      }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao atualizar o Tracking.');

    const results = Array.isArray(data?.tracking?.results) ? data.tracking.results : [];
    const successResults = results.filter((item) => item?.success);
    const successIds = new Set(successResults.map((item) => String(item?.id || '')).filter(Boolean));
    const errorCount = Array.isArray(data?.errors) ? data.errors.length : 0;

    if (options.dateOnly) {
      // Remove da tela imediatamente o que já foi corrigido com sucesso, antes de exibir qualquer aviso.
      state.stageDatePendencies = (Array.isArray(state.stageDatePendencies) ? state.stageDatePendencies : [])
        .filter((item) => !successIds.has(String(item?.id || '')));
      setStageDateSelection((state.stageDateSelectedIds || []).filter((id) => !successIds.has(String(id || ''))));
      renderStageUpdatesModal();

      const messages = [];
      if (successResults.length) {
        messages.push(`${successResults.length} pendência(s) de data corrigida(s) no Tracking.`);
      }
      if (errorCount) {
        messages.push(`Pendências não processadas: ${errorCount}.`);
      }
      if (messages.length) window.alert(messages.join('\n'));

      await loadStageHistoryDatePendencies();
    } else {
      setStageSelection((state.stageSelectedIds || []).filter((id) => !successIds.has(String(id || ''))));
      renderStageUpdatesModal();

      const messages = [];
      if (successResults.length === 1) {
        messages.push(stageTrackingMessageFromResult(successResults[0]));
      } else if (successResults.length > 1) {
        const totalRows = successResults.reduce((sum, item) => sum + Number(item?.rowCount || 0), 0);
        messages.push(`Tracking atualizado em ${successResults.length} apontamento(s).${totalRows ? ` ${totalRows} linha(s) localizada(s).` : ''}`);
      }
      if (errorCount) {
        messages.push(`Pendências não processadas: ${errorCount}.`);
      }
      if (messages.length) window.alert(messages.join('\n'));
    }

    await loadStageUpdates();
    renderStageUpdatesModal();
  } catch (error) {
    const processedAfterRefresh = await refreshStageUpdatesAfterAction(cleanIds);
    if (processedAfterRefresh) {
      window.alert('Tracking atualizado/conferido. A tela foi sincronizada novamente.');
      return;
    }
    window.alert(error.message || 'Falha ao atualizar o Tracking.');
  } finally {
    state.stageTrackingSubmitting = false;
    renderStageUpdatesModal();
  }
}


function getVisiblePendingValidationIds(onlySelected = false) {
  const pending = getFilteredStageUpdatesForValidation().filter((item) => isPendingStageStatus(item.status));
  if (onlySelected && (state.stageSelectedIds || []).length) return getSelectedVisibleStageIds(pending);
  return pending.map((item) => String(item.id || '')).filter(Boolean);
}


async function deleteStageUpdatePending(id) {
  const cleanId = String(id || '').trim();
  if (!cleanId || state.stageTrackingSubmitting) return;
  const item = (Array.isArray(state.stageUpdates) ? state.stageUpdates : []).find((entry) => String(entry.id || '') === cleanId);
  const label = item ? `${item.projectDisplay || item.projectNumber || 'BSP'} • ${item.spoolIso || 'Spool'}` : 'este apontamento';
  const confirmed = window.confirm(`Remover ${label} da fila de Validação PCP?

Use esta opção somente para apontamento lançado por engano ou spool inexistente no Tracking.`);
  if (!confirmed) return;
  state.stageTrackingSubmitting = true;
  renderStageUpdatesModal();
  try {
    const response = await fetch('/api/stage-updates', {
      method: 'DELETE',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids: [cleanId] }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao remover apontamento.');
    state.stageUpdates = (Array.isArray(state.stageUpdates) ? state.stageUpdates : []).filter((entry) => String(entry.id || '') !== cleanId);
    setStageSelection((state.stageSelectedIds || []).filter((entryId) => String(entryId) !== cleanId));
    await loadStageUpdates();
    renderStageUpdatesModal();
  } catch (error) {
    window.alert(error.message || 'Falha ao remover apontamento.');
  } finally {
    state.stageTrackingSubmitting = false;
    renderStageUpdatesModal();
  }
}

async function concludeStageUpdatesBulkOk() {
  const selected = getVisiblePendingValidationIds(true);
  const ids = selected.length ? selected : getVisiblePendingValidationIds(false);
  await concludeStageUpdatesBulk(ids);
}

async function concludeStageUpdatesBulk(ids = []) {
  const cleanIds = Array.from(new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '').trim()).filter(Boolean)));
  if (!cleanIds.length) return;
  const resolutionInput = window.prompt('Observação de validação do PCP para o lote (opcional):', '');
  if (resolutionInput === null) return;
  const resolutionNote = String(resolutionInput || '').trim();
  try {
    const response = await fetch('/api/stage-updates', {
      method: 'PATCH',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids: cleanIds, resolutionNote }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || (!data?.ok && !data?.partial)) throw new Error(data?.error || 'Falha ao concluir lote de apontamentos.');
    if (Array.isArray(data?.errors) && data.errors.length) {
      window.alert(`Alguns itens não foram concluídos porque ainda precisam de atualização no Tracking: ${data.errors.length}`);
    }
    setStageSelection([]);
    await loadStageUpdates();
    renderStageUpdatesModal();
  } catch (error) {
    const processedAfterRefresh = await refreshStageUpdatesAfterAction(cleanIds, { clearSelection: true });
    if (processedAfterRefresh || isProbablyAlreadyProcessedStageError(error)) {
      await loadStageUpdates({ force: true }).catch(() => {});
      setStageSelection([]);
      renderStageUpdatesModal();
      window.alert('Os apontamentos já estavam concluídos ou foram sincronizados novamente.');
      return;
    }
    window.alert(error.message || 'Falha ao concluir lote de apontamentos.');
  }
}

async function handleStageWorkspaceSubmit(formEl, actionType = 'advance') {
  const rowEl = formEl?.closest('tr');
  const projectRowId = String(formEl?.dataset?.projectRowId || '').trim();
  const spoolIso = String(formEl?.dataset?.spoolIso || '').trim();
  const sector = String(formEl?.dataset?.stageSector || getStageWorkspaceSector() || '').trim();
  const progress = String(formEl?.querySelector('[name="progress"]')?.value || '').trim();
  const completionDate = String(rowEl?.querySelector('[name="completionDate"]')?.value || '').trim();
  const note = String(rowEl?.querySelector('[name="note"]')?.value || '').trim();
  const metadata = {
    projectNumber: String(formEl?.dataset?.projectNumber || '').trim(),
    projectDisplay: String(formEl?.dataset?.projectDisplay || formEl?.dataset?.projectNumber || '').trim(),
    client: String(formEl?.dataset?.client || '').trim(),
    spoolDescription: String(formEl?.dataset?.spoolDescription || '').trim(),
    spoolStage: String(formEl?.dataset?.spoolStage || '').trim(),
  };
  upsertStageDraft(projectRowId, spoolIso, sector, { progress, completionDate, note, actionType, ...metadata });
  if (!projectRowId || !spoolIso || !progress) {
    window.alert('Preencha o avanço do spool antes de enviar.');
    return;
  }
  const submitKey = getStageSubmitKey(projectRowId, spoolIso, sector);
  if (state.stageSubmittingKeys?.[submitKey]) {
    return;
  }
  setStageSubmitting(projectRowId, spoolIso, sector, true);
  renderStageUpdatesModal();
  try {
    const response = await fetch('/api/stage-updates', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ projectRowId, spoolIso, progress, completionDate, note, sector, actionType, ...metadata }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao enviar apontamento.');

    const newUpdate = data?.update || {
      projectRowId: Number(projectRowId || 0),
      spoolIso,
      sector,
      progress: Number(progress || 0),
      completionDate,
      note,
      status: actionType === 'review' ? 'pending_review' : 'pending_advance',
      createdAt: new Date().toISOString(),
    };
    state.stageUpdates = [newUpdate, ...(Array.isArray(state.stageUpdates) ? state.stageUpdates : [])];
    removeStageDraft(projectRowId, spoolIso, sector);
    setStageSubmitting(projectRowId, spoolIso, sector, false);
    renderStageUpdatesModal();
    loadStageUpdates().then(() => {
      renderStageUpdatesModal();
    }).catch(() => {});
  } catch (error) {
    setStageSubmitting(projectRowId, spoolIso, sector, false);
    renderStageUpdatesModal();
    window.alert(error.message || 'Falha ao enviar apontamento.');
  }
}

async function concludeStageUpdate(id) {
  if (!id) return;
  const update = (Array.isArray(state.stageUpdates) ? state.stageUpdates : []).find((item) => String(item.id) === String(id));
  const resolutionPrompt = isReviewStageStatus(update?.status)
    ? 'Observação da tratativa da revisão (opcional):'
    : 'Observação de validação do PCP (opcional):';
  const resolutionInput = window.prompt(resolutionPrompt, '');
  if (resolutionInput === null) {
    return;
  }
  const resolutionNote = String(resolutionInput || '').trim();
  try {
    const response = await fetch('/api/stage-updates', {
      method: 'PATCH',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, resolutionNote }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || (!data?.ok && !data?.partial)) throw new Error(data?.error || 'Falha ao concluir apontamento.');
    if (Array.isArray(data?.errors) && data.errors.length) {
      window.alert('Este apontamento ainda precisa de atualização no Tracking antes de concluir.');
    }
    await loadStageUpdates();
    renderStageUpdatesModal();
  } catch (error) {
    const processedAfterRefresh = await refreshStageUpdatesAfterAction([id], { clearSelection: true });
    if (processedAfterRefresh || isProbablyAlreadyProcessedStageError(error)) {
      await loadStageUpdates({ force: true }).catch(() => {});
      renderStageUpdatesModal();
      window.alert('Este apontamento já estava concluído ou foi sincronizado novamente.');
      return;
    }
    window.alert(error.message || 'Falha ao concluir apontamento.');
  }
}
async function handleAdminUserSubmit(event) {
  event.preventDefault();
  const editingId = adminUserIdEl?.value || "";
  adminUserFeedbackEl.textContent = editingId ? "Salvando usuário..." : "Criando usuário...";
  try {
    const payload = {
      userId: editingId,
      name: document.getElementById("admin-user-name").value,
      username: String(document.getElementById("admin-user-username").value || "").trim(),
      password: String(document.getElementById("admin-user-password").value || "").trim(),
      role: document.getElementById("admin-user-role").value,
      operationRegion: 'BR',
      siteKey: 'BR',
      portalSite: 'BR',
      sector: document.getElementById("admin-user-role").value === 'client' ? 'all' : document.getElementById("admin-user-sector").value,
      alertSectors: document.getElementById("admin-user-role").value === 'client' ? [] : getSelectedAdminAlertSectors(),
      projectPmAliases: adminUserFormHasProjectsScope() ? getAdminProjectPmAliases() : [],
      qualityCompetencies: adminUserFormHasQualityScope() ? getAdminQualityCompetencies() : [],
      canViewClientPanel: adminUserCanViewClientPanelEl?.checked === true,
      clientKey: adminUserClientKeyEl?.value || '',
      clientName: adminUserClientNameEl?.value || '',
      clientLogoUrl: adminUserClientLogoUrlEl?.value || '',
      clientPlatformImageUrl: '',
      clientPlatformImages: parseClientPlatformImages(adminUserClientPlatformImagesEl?.value || ''),
    };
    const response = await fetch("/api/admin-users", {
      method: editingId ? "PUT" : "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || (editingId ? "Falha ao editar usuário." : "Falha ao criar usuário."));
    }
    const savedUser = data.user || {
      id: editingId || `local-${Date.now()}`,
      name: payload.name,
      username: payload.username,
      role: payload.role,
      sector: payload.role === "admin" ? "all" : payload.sector,
      alertSectors: payload.role === "admin" ? [] : payload.alertSectors,
      projectPmAliases: payload.role === "admin" ? [] : payload.projectPmAliases,
      qualityCompetencies: payload.role === "admin" ? [] : payload.qualityCompetencies,
      canViewClientPanel: payload.canViewClientPanel === true,
      clientKey: payload.clientKey,
      clientName: payload.clientName,
      clientLogoUrl: payload.clientLogoUrl,
      clientPlatformImageUrl: payload.clientPlatformImageUrl,
      clientPlatformImages: payload.clientPlatformImages,
      active: true,
      createdAt: new Date().toISOString(),
    };
    upsertLocalUser(savedUser);
    resetAdminUserForm();
    adminUserFeedbackEl.textContent = state.githubSyncEnabled
      ? (editingId ? "Usuário atualizado e salvo no GitHub." : "Usuário criado e salvo no GitHub.")
      : (editingId ? "Usuário atualizado localmente. Para enviar ao GitHub, configure as variáveis GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify e clique em 'Subir pro GitHub'." : "Usuário criado localmente. Para enviar ao GitHub, configure as variáveis GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify e clique em 'Subir pro GitHub'.");
    await loadAdminData();
  } catch (error) {
    adminUserFeedbackEl.textContent = error.message || (editingId ? "Falha ao editar usuário." : "Falha ao criar usuário.");
  }
}

async function updateUserRole(userId, role) {
  try {
    const response = await fetch("/api/admin-users", {
      method: "PATCH",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ userId, role }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao atualizar perfil.");
    }
    await loadAdminData();
  } catch (error) {
    adminUserFeedbackEl.textContent = error.message || "Falha ao atualizar perfil.";
  }
}

async function handleAdminAlertSubmit(event) {
  event.preventDefault();
  adminAlertFeedbackEl.textContent = "Enviando alerta...";
  try {
    const payload = {
      sector: document.getElementById("admin-alert-sector").value,
      title: document.getElementById("admin-alert-title").value,
      message: document.getElementById("admin-alert-message").value,
      priority: document.getElementById("admin-alert-priority").value,
      requiresAck: document.getElementById("admin-alert-requires-ack").checked,
    };
    const response = await fetch("/api/sector-alerts", {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao criar alerta operacional.");
    }
    adminAlertFormEl.reset();
    document.getElementById("admin-alert-requires-ack").checked = true;
    adminAlertFeedbackEl.textContent = "Alerta operacional enviado com sucesso.";
    await loadManualAlerts();
    await loadAdminData();
  } catch (error) {
    adminAlertFeedbackEl.textContent = error.message || "Falha ao criar alerta operacional.";
  }
}

async function acknowledgeManualAlert(alertId) {
  try {
    const response = await fetch("/api/sector-alerts", {
      method: "PATCH",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ alertId }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao confirmar leitura.");
    }
    await loadManualAlerts();
  } catch (error) {
    console.warn(error);
  }
}

async function init() {
  if (alertModalEl) {
    alertModalEl.classList.add("hidden");
    alertModalEl.setAttribute("aria-hidden", "true");
  }
  document.body.classList.remove("modal-open");
  updateConnectionStatus();
  window.addEventListener("online", updateConnectionStatus);
  window.addEventListener("offline", updateConnectionStatus);
  setupInstallExperience();
  registerServiceWorker();
  startClocks();
  bindEvents();
  setupLoginPasswordToggle();
  setupAdminPasswordToggle();
  resetAdminUserForm();
  // v36.85: o login precisa ser a primeira tela.
  // Não aquecemos projetos antes de confirmar sessão para evitar demora/flash do painel antes do login.
  if (!state.user) openLoginModal("Faça login para visualizar o painel.");
  const authenticated = await bootstrapSession();
  if (authenticated) {
    prewarmProjectsApi();
    const autoOpenStageValidation = shouldOpenStageValidationWorkspaceFromUrl() && canValidateStageWorkspace();
    if (autoOpenStageValidation) {
      state.stageUpdatesSearchQuery = '';
      openStageUpdatesModal({ loading: true });
    }
    setProjectsLoadingState('Carregando painel...');
    try {
      await loadProjects({
        preferServerCache: true,
        // v37.17: login deve abrir pelo cache Supabase sem forçar Smartsheet.
        // Carteira vazia legítima não pode travar cliente com escopo exato.
        requireData: false,
        requireClientPo: false,
      });
    } catch (error) {
      console.warn('[Init] Dados iniciais ainda não liberados:', error?.message || error);
      if (isClientUser()) {
        setProjectsLoadingState('Carregando dados do cliente...');
      } else {
        setProjectsLoadingState(error?.message || 'Falha ao carregar painel.');
      }
    }

    // v36.55: quando o usuário já está autenticado e abre uma guia nova/anônima,
    // o init não passava pela validação do login. Se o endpoint respondesse vazio
    // por cache/snapshot temporário, o Portal do Cliente ficava liberado com "--".
    // Agora o cliente só é liberado quando existe carteira válida e PO carregada.
    if (isClientUser()) {
      try {
        await ensureDashboardDataReadyBeforeRelease({ maxAttempts: 6, retryDelayMs: 850 });
      } catch (error) {
        console.warn('[Init] Portal do Cliente ainda sem dados válidos:', error?.message || error);
        setProjectsLoadingState(error?.message || 'Os dados do cliente ainda não carregaram. Clique em Atualizar agora em alguns segundos.');
      }
    }

    syncStageDraftsForCurrentSector();
    startPostSessionBackgroundLoads({ autoOpenStageValidation });
    startPolling();
  }
}

init();
