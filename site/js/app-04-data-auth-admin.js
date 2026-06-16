/* STEP Dashboard v37.65 - Watchdog do cache: fallback automático se o agendador Netlify não tocar em 15 min. */
function isStageUpdatesWorkspaceOpen() {
  return Boolean(stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden'));
}

function shouldSkipBackgroundRequest(options = {}) {
  return !options.force && isPageHidden();
}

let trackingCacheSyncPromise = null;
let trackingCacheSyncCooldownUntil = 0;

// v37.65: watchdog do cache.
// Se a rotina agendada do Netlify não rodar/registrar timestamp, o próprio painel logado
// aciona uma sincronização leve para tocar o updated_at no Supabase. Isso mantém a operação
// com fallback seguro sem depender exclusivamente do cron externo.
let trackingCacheWatchdogPromise = null;
let trackingCacheWatchdogLastRunAt = 0;
const TRACKING_CACHE_WATCHDOG_GRACE_MS = 75 * 1000;
const TRACKING_CACHE_WATCHDOG_COOLDOWN_MS = 4 * 60 * 1000;

function getTrackingCacheRefreshAfterMs(meta = state.meta) {
  const value = Number(meta?.persistentCacheAutoRefreshAfterMs || meta?.cacheAutoRefreshAfterMs || 15 * 60 * 1000);
  return Number.isFinite(value) && value > 0 ? value : 15 * 60 * 1000;
}

function isTrackingCacheStaleForWatchdog(meta = state.meta) {
  if (!meta) return false;
  const ageMs = getTrackingCacheAgeMs(meta);
  if (ageMs == null) return false;
  return ageMs >= getTrackingCacheRefreshAfterMs(meta) + TRACKING_CACHE_WATCHDOG_GRACE_MS;
}

function maybeTriggerTrackingCacheWatchdog(reason = 'watchdog') {
  if (!state.user || isPageHidden()) return Promise.resolve({ ok: true, skipped: true, reason: 'not-visible-or-no-user' });
  if (!isTrackingCacheStaleForWatchdog(state.meta)) return Promise.resolve({ ok: true, skipped: true, reason: 'cache-not-stale' });
  if (trackingCacheWatchdogPromise) return trackingCacheWatchdogPromise;

  const now = Date.now();
  if (trackingCacheWatchdogLastRunAt && now - trackingCacheWatchdogLastRunAt < TRACKING_CACHE_WATCHDOG_COOLDOWN_MS) {
    return Promise.resolve({ ok: true, skipped: true, reason: 'watchdog-cooldown' });
  }

  trackingCacheWatchdogLastRunAt = now;
  const previousCacheTimeMs = getCurrentTrackingCacheTimestampMs();

  trackingCacheWatchdogPromise = (async () => {
    try {
      if (lastSyncEl && state.meta) {
        const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
        lastSyncEl.textContent = `${label} • ativando fallback automático`;
      }

      const syncResult = await triggerTrackingCacheSync({ force: false, manual: false, auto: true, reason });

      if (syncResult?.ok && !syncResult?.skipped && !syncResult?.staleCacheKept) {
        await delay(1800);
        await loadProjects({
          force: false,
          background: true,
          skipLocalCache: true,
          suppressLoadingState: true,
          preferServerCache: true,
        });
      } else if (lastSyncEl && state.meta && syncResult?.reason) {
        const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
        lastSyncEl.textContent = `${label} • fallback: ${describeTrackingSyncWarning(syncResult)}`;
      }

      return syncResult;
    } catch (error) {
      console.warn('[tracking-cache-watchdog] Falha no fallback automático:', error?.message || error);
      return { ok: false, skipped: true, reason: 'watchdog-error', error: error?.message || String(error) };
    } finally {
      window.setTimeout(() => {
        trackingCacheWatchdogPromise = null;
      }, 1500);
    }
  })();

  return trackingCacheWatchdogPromise;
}

function hasVisibleOperationalCache() {
  return Boolean(
    state?.meta?.persistentCacheUpdatedAt ||
    state?.meta?.cacheUpdatedAt ||
    state?.meta?.lastSync ||
    (Array.isArray(state?.projects) && state.projects.length > 0)
  );
}

function normalizeTrackingSyncResult(data, fallbackReason = 'sync-not-completed') {
  const payload = data && typeof data === 'object'
    ? data
    : { ok: false, skipped: true, reason: fallbackReason };

  // v37.24: se já existe cache operacional visível, uma falha de checagem/smartsheet
  // não deve virar popup bloqueante. O painel mantém a base atual e informa inline.
  if (payload.manual === true && payload.ok === false) return payload;

  if (payload.ok === false && hasVisibleOperationalCache()) {
    return {
      ...payload,
      ok: true,
      synced: false,
      skipped: true,
      staleCacheKept: true,
      reason: payload.reason || fallbackReason,
      warning: payload.error || payload.warning || 'Cache atual mantido; rotina tentará novamente.',
    };
  }
  return payload;
}

function getTrackingCacheTimestampMs(value) {
  const time = value ? new Date(value).getTime() : 0;
  return Number.isFinite(time) ? time : 0;
}

function getCurrentTrackingCacheTimestampMs() {
  return Math.max(
    getTrackingCacheTimestampMs(state?.meta?.persistentCacheUpdatedAt),
    getTrackingCacheTimestampMs(state?.meta?.cacheUpdatedAt),
    getTrackingCacheTimestampMs(state?.meta?.lastSync)
  );
}

function assertManualTrackingSyncCompleted(syncResult, previousCacheTimeMs) {
  const synced = Boolean(syncResult?.ok && syncResult?.synced && !syncResult?.skipped && !syncResult?.staleCacheKept);
  const cacheUpdatedAt = syncResult?.cacheUpdatedAt || syncResult?.lastSync || '';
  const nextCacheTimeMs = getTrackingCacheTimestampMs(cacheUpdatedAt);
  if (!synced || !nextCacheTimeMs || nextCacheTimeMs <= Number(previousCacheTimeMs || 0)) {
    throw new Error(syncResult?.warning || syncResult?.error || 'Atualizacao do Tracking nao foi concluida. Cache antigo mantido.');
  }
  return cacheUpdatedAt;
}

function delay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function waitForManualTrackingCacheUpdate(previousCacheTimeMs, options = {}) {
  const timeoutMs = Math.max(30 * 1000, Number(options.timeoutMs || 8 * 60 * 1000));
  const intervalMs = Math.max(2500, Number(options.intervalMs || 5000));
  const startedAt = Date.now();
  let attempt = 0;
  let lastError = null;

  while (Date.now() - startedAt < timeoutMs) {
    attempt += 1;
    if (refreshProjectsButtonEl) {
      refreshProjectsButtonEl.disabled = true;
      refreshProjectsButtonEl.textContent = `Atualizando cache... ${attempt}`;
    }
    if (lastSyncEl) {
      const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Ultima atualizacao do cache' });
      lastSyncEl.textContent = `${label} • atualizando Tracking, aguarde`;
    }
    await delay(attempt === 1 ? 2500 : intervalMs);
    try {
      await loadProjects({
        force: false,
        skipLocalCache: true,
        suppressLoadingState: true,
        preferServerCache: true,
      });
      const currentTimeMs = getCurrentTrackingCacheTimestampMs();
      if (currentTimeMs > Number(previousCacheTimeMs || 0)) {
        return state.meta?.persistentCacheUpdatedAt || state.meta?.cacheUpdatedAt || state.meta?.lastSync || new Date(currentTimeMs).toISOString();
      }
    } catch (error) {
      lastError = error;
      console.warn('[tracking-sync] Aguardando cache atualizado:', error?.message || error);
    }
  }

  throw new Error(lastError?.message || 'Atualizacao ainda nao confirmada pelo Supabase. Cache antigo mantido.');
}

function describeTrackingSyncWarning(syncResult) {
  if (!syncResult) return '';
  const reason = String(syncResult.reason || '').toLowerCase();
  if (syncResult.warning) return String(syncResult.warning);
  if (reason.includes('lock')) return 'Atualização já está em andamento; cache atual mantido.';
  if (reason.includes('version-check')) return 'Smartsheet não respondeu à checagem rápida; cache atual mantido.';
  if (reason.includes('full-refresh')) return 'Smartsheet não concluiu a atualização completa; cache atual mantido.';
  if (reason.includes('request-failed')) return 'Não foi possível chamar a sincronização agora; cache atual mantido.';
  if (reason.includes('cache-too-fresh')) return 'Cache recente; nova atualização manual ainda não é necessária.';
  return 'Cache atual mantido; rotina tentará novamente.';
}

function triggerTrackingCacheSync(options = {}) {
  if (!state.user) return Promise.resolve({ ok: false, skipped: true, reason: 'not-authenticated' });

  const force = Boolean(options.force);
  const auto = options.auto !== false;
  const now = Date.now();
  if (!force && auto && trackingCacheSyncCooldownUntil && now < trackingCacheSyncCooldownUntil) {
    return Promise.resolve({ ok: true, skipped: true, reason: 'frontend-cooldown' });
  }
  if (trackingCacheSyncPromise) return trackingCacheSyncPromise;

  const manual = Boolean(options.manual);
  const useBackgroundSync = Boolean(force && manual);
  const url = useBackgroundSync
    ? `/api/sync-tracking-cache-background?force=1&manual=1&t=${Date.now()}`
    : force
    ? `/api/sync-tracking-cache?force=1${manual ? '&manual=1' : ''}`
    : `/api/sync-tracking-cache?mode=auto${manual ? '&manual=1' : ''}`;
  trackingCacheSyncPromise = fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    cache: 'no-store',
  })
    .then((response) => {
      if (useBackgroundSync && response.status === 202) {
        return { ok: true, manual: true, background: true, accepted: true, synced: false, skipped: false, reason: 'background-sync-started' };
      }
      return response.json().catch(() => {
        if (useBackgroundSync && response.ok) return { ok: true, manual: true, background: true, accepted: true, synced: false, skipped: false, reason: 'background-sync-started' };
        return null;
      });
    })
    .then((data) => {
      trackingCacheSyncCooldownUntil = Date.now() + 60 * 1000;
      const normalized = normalizeTrackingSyncResult(data, 'empty-response');
      if (!normalized?.ok || normalized?.staleCacheKept) {
        console.warn('[tracking-sync] Sincronização controlada não concluída:', normalized?.warning || normalized?.error || normalized);
      }
      return normalized;
    })
    .catch((error) => {
      console.warn('[tracking-sync] Falha ao chamar sincronização controlada:', error?.message || error);
      return normalizeTrackingSyncResult({
        ok: false,
        skipped: true,
        reason: 'request-failed',
        error: error?.message || 'Falha ao chamar sincronização controlada.',
      }, 'request-failed');
    })
    .finally(() => {
      window.setTimeout(() => {
        trackingCacheSyncPromise = null;
      }, 1500);
    });

  return trackingCacheSyncPromise;
}

function setProjectsLoadingState(message = 'Carregando dados operacionais...') {
  if (!state.user) return;
  if (bodyEl && !state.projects.length) {
    bodyEl.innerHTML = `<tr><td colspan="21" class="loading-cell">${escapeHtml(message)}</td></tr>`;
  }
  if (detailCardEl && !state.projects.length) {
    detailCardEl.innerHTML = `<div class="detail-placeholder">${escapeHtml(message)}</div>`;
  }
  if (searchCountEl && !state.projects.length) searchCountEl.textContent = 'Carregando...';
  if (lastSyncEl) lastSyncEl.textContent = message;
}

function revalidateProjectsInBackground(force = false) {
  if (!state.user || state.loadingProjectsRequest) return Promise.resolve();

  // v37.17: a atualização pesada fica independente do usuário via Scheduled Function.
  // Em background, o navegador apenas relê a última base do Supabase/Function cache.
  // Smartsheet só é chamado pela rotina agendada ou pelo botão manual de atualização.
  return loadProjects({
    force: false,
    background: true,
    skipLocalCache: true,
    suppressLoadingState: true,
    preferServerCache: true,
  }).catch((error) => {
    console.warn('Falha ao revalidar projetos em background:', error?.message || error);
  });
}

function getProjectsCacheKey(user = state.user) {
  const role = String(user?.role || 'guest').trim().toLowerCase();
  const username = normalizeText(user?.username || user?.name || 'guest').replace(/[^a-z0-9]+/g, '_') || 'guest';
  // Para usuários cliente, usar clientKey ou clientName como parte da chave.
  // Se ambos estiverem vazios, usar o ID do usuário como fallback para evitar cache compartilhado.
  let client = normalizeText(user?.clientKey || user?.clientName || '').replace(/[^a-z0-9]+/g, '_');
  if (!client && user?.id) {
    client = normalizeText(user.id).replace(/[^a-z0-9]+/g, '_').substring(0, 16);
  }
  if (!client && user?.sub) {
    client = normalizeText(user.sub).replace(/[^a-z0-9]+/g, '_').substring(0, 16);
  }
  return `${PROJECTS_CACHE_KEY}:${role}:${username}:${client}`;
}

function readProjectsCache() {
  try {
    const raw = window.localStorage.getItem(getProjectsCacheKey());
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || !parsed.payload) return null;
    return parsed;
  } catch {
    return null;
  }
}

function hasResolvedCustomerPo(project) {
  const values = [
    project?.customerPo,
    project?.customerPoDisplay,
    ...(Array.isArray(project?.customerPoList) ? project.customerPoList : []),
  ];
  return values.some((value) => {
    const text = String(value || '').trim();
    return text && !/aguardando\s+po/i.test(text);
  });
}

function isClientProjectsPayloadPoReady(payload) {
  if (!isClientUser()) return true;
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  if (!projects.length) return false;

  // Quando a base complementar de PO respondeu, o payload está pronto mesmo que alguma BSP isolada
  // permaneça sem PO por falta de correspondência no Smartsheet.
  if (payload?.meta?.wipStepPoAvailable === true) return true;

  // Fallbacks/snapshots antigos costumam vir com todas as BSPs como "Aguardando PO".
  // Eles não podem ser tratados como carregamento final no Portal do Cliente.
  if (projects.some(hasResolvedCustomerPo)) return true;

  const servedFromFallback = Boolean(payload?.meta?.servedFromDiskFallback || payload?.meta?.source === 'disk-snapshot' || payload?.meta?.cacheReason === 'disk-fallback');
  if (servedFromFallback) return false;

  const allWaitingPo = projects.every((project) => /aguardando\s+po/i.test(String(getClientProjectDisplayCode(project) || '')));
  return !allWaitingPo;
}

function writeProjectsCache(payload) {
  try {
    const meta = payload?.meta || {};
    if (meta.servedFromBundledFallback || String(meta.cacheReason || '').includes('bundled-fallback')) {
      console.warn('[Cache] Não salvou fallback empacotado no navegador para evitar base operacional antiga');
      return;
    }
    if (isClientUser() && !isClientProjectsPayloadPoReady(payload)) {
      console.warn('[Cache] Não salvou cache do Portal do Cliente sem PO carregada');
      return;
    }
    window.localStorage.setItem(getProjectsCacheKey(), JSON.stringify({
      savedAt: Date.now(),
      payload,
    }));
  } catch {
    // Cache local é apenas otimização. Se o navegador bloquear, o app continua funcionando.
  }
}

function clearProjectsCache() {
  try {
    window.localStorage.removeItem(getProjectsCacheKey());
    Object.keys(window.localStorage || {}).forEach((key) => {
      if (String(key).startsWith(PROJECTS_CACHE_KEY + ':')) window.localStorage.removeItem(key);
    });
  } catch {}
}

function isProjectsCacheFresh(cacheEntry) {
  const savedAt = Number(cacheEntry?.savedAt || 0);
  return savedAt > 0 && Date.now() - savedAt <= PROJECTS_CACHE_TTL_MS;
}

function shouldIgnoreCachedProjectsPayload(cacheEntry) {
  if (!cacheEntry?.payload) return true;
  const cachedProjects = Array.isArray(cacheEntry.payload.projects) ? cacheEntry.payload.projects : [];
  const meta = cacheEntry.payload.meta || {};
  const isOldStaticSnapshot = Boolean(
    meta.servedFromStaticFallback
    || meta.servedFromErrorFallback
    || meta.servedFromDiskFallback
    || meta.servedFromDiskFastCache
    || meta.servedFromBundledFallback
    || meta.fallbackMayBeIncomplete
    || meta.cacheReason === 'disk-fast-boot'
    || meta.cacheReason === 'client-disk-fast-boot'
    || String(meta.cacheReason || '').includes('bundled-fallback')
    || meta.source === 'static-fallback'
    || meta.source === 'snapshot'
    || meta.source === 'disk-snapshot'
    || String(meta.version || '').toLowerCase().includes('snapshot')
  );
  if (isOldStaticSnapshot) {
    console.warn('[Cache] Ignorando snapshot/fallback local antigo para forçar leitura direta do Smartsheet');
    return true;
  }
  const savedAt = Number(cacheEntry?.savedAt || 0);
  if (savedAt > 0 && Date.now() - savedAt > PROJECTS_CACHE_TTL_MS) {
    console.warn('[Cache] Ignorando cache local vencido para não exibir base operacional antiga');
    return true;
  }
  // Evita reaproveitar cache vazio criado por versão anterior do Portal do Cliente.
  // Cache vazio de cliente travava os cards em "--" até expirar.
  if (isClientUser() && cachedProjects.length === 0) {
    console.warn('[Cache] Ignorando cache vazio para usuário cliente');
    return true;
  }
  // Validação adicional: se o cache tem meta.clientPortal=true mas nenhum projeto, é inválido.
  if (isClientUser() && cacheEntry.payload.meta?.clientPortal && cachedProjects.length === 0) {
    console.warn('[Cache] Ignorando cache com clientPortal=true mas sem projetos');
    return true;
  }
  // Evita reaproveitar snapshot antigo/fallback em que todas as BSPs aparecem como "Aguardando PO".
  // Esse era o motivo de abrir em guia anônima sem PO e só corrigir depois de vários F5.
  if (isClientUser() && !isClientProjectsPayloadPoReady(cacheEntry.payload)) {
    console.warn('[Cache] Ignorando cache do Portal do Cliente sem PO carregada');
    return true;
  }
  return false;
}

function applyProjectsPayload(data, options = {}) {
  // v37.04: se o backend reidratou os dados do cliente pelo Supabase,
  // sincroniza também o estado do frontend para evitar Portal do Cliente ficar com nome/logo padrão.
  if (isClientUser() && data?.meta) {
    state.user = {
      ...state.user,
      clientName: state.user?.clientName || data.meta.clientName || '',
      clientKey: state.user?.clientKey || data.meta.clientKey || '',
      clientLogoUrl: state.user?.clientLogoUrl || data.meta.clientLogoUrl || '',
    };
  }
  state.projects = enrichProjects(data.projects || []);
  renderAdminProjectPmAliasOptions();
  renderProjectViewTabs();
  state.stats = data.stats || null;
  state.meta = data.meta || null;
  state.alerts = data.alerts || [];
  state.projectsLoadedFromCache = Boolean(options.fromCache);
  buildDemandOptions();
  buildProjectTypeOptions();
  buildWeekOptions();

  if (!state.selectedProjectId && state.projects.length) {
    state.selectedProjectId = state.projects[0].rowId;
  }

  applyFilter();
  renderStats();
  renderTable();
  renderSelectedProjectCard();
  renderAlertBadge();
  updateMeta();
  renderAlertModal();
  renderClientDashboard();
  scheduleClientBspOverridesLoad();
  if (state.user && sectorAlertsModalEl && !sectorAlertsModalEl.classList.contains('hidden')) {
    renderManualAlerts();
  }
}

function updateMeta() {
  if (!state.meta) return;
  sheetNameEl.textContent = state.meta.clientPortal ? 'Base operacional' : (state.meta.sheetName || "Smartsheet");
  applyTrackingCacheUpdateLabel(lastSyncEl, state.meta, { prefix: 'Última atualização do cache' });
  footerVersionEl.textContent = (state.meta.clientPortal || isClientUser()) ? `Versão dos dados: ${state.meta.version || '--'}` : `Versão da sheet: ${state.meta.version}`;
}


function prewarmProjectsApi() {
  if (state.user) return Promise.resolve(null);
  if (projectsWarmupPromise) return projectsWarmupPromise;

  window.clearTimeout(projectsWarmupResetTimer);
  projectsWarmupPromise = fetch('/api/projects?warmup=1', {
    credentials: 'same-origin',
    cache: 'no-store',
  })
    .then((response) => response.json().catch(() => null))
    .catch((error) => {
      console.warn('Pré-aquecimento dos dados operacionais não concluído:', error?.message || error);
      return null;
    })
    .finally(() => {
      // Mantém a referência por pouco tempo para evitar chamadas repetidas enquanto o usuário faz login.
      projectsWarmupResetTimer = window.setTimeout(() => {
        projectsWarmupPromise = null;
      }, 60000);
    });

  return projectsWarmupPromise;
}

async function waitForProjectsWarmup(maxWaitMs = 2000) {
  const warmupPromise = projectsWarmupPromise || prewarmProjectsApi();
  if (!warmupPromise) return null;

  let timeoutId = null;
  const timeoutPromise = new Promise((resolve) => {
    timeoutId = window.setTimeout(() => resolve(null), maxWaitMs);
  });

  try {
    return await Promise.race([warmupPromise, timeoutPromise]);
  } finally {
    if (timeoutId) window.clearTimeout(timeoutId);
  }
}

async function loadProjects(options = {}) {
  const force = Boolean(options.force);
  const background = Boolean(options.background);
  const skipLocalCache = Boolean(options.skipLocalCache);
  const suppressLoadingState = Boolean(options.suppressLoadingState);

  if (!state.user) {
    resetDashboardForLoggedOutState();
    return;
  }

  if (background && shouldSkipBackgroundRequest(options)) return;

  const cached = skipLocalCache ? null : readProjectsCache();
  const cacheFresh = isProjectsCacheFresh(cached);
  const shouldUseCache = !force && cached?.payload && cacheFresh && !shouldIgnoreCachedProjectsPayload(cached);
  if (shouldUseCache) {
    applyProjectsPayload(cached.payload, { fromCache: true });
    const fresh = isProjectsCacheFresh(cached);
    if (lastSyncEl && state.meta) {
      const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
      lastSyncEl.textContent = `${label} • exibindo cache local${fresh ? '' : ' enquanto atualiza'}`;
    }

    if (!force && !background && !state.loadingProjectsRequest) {
      // Stale-while-revalidate: a tela aparece imediatamente e a API atualiza em background.
      window.setTimeout(() => revalidateProjectsInBackground(false), 0);
    }

    // Em navegações/login, não bloqueia a thread aguardando a API quando já existe cache aproveitável.
    if (!force && !background) {
      if (fresh) state.lastProjectsFetchAt = Date.now();
      return {
        ok: true,
        fromCache: true,
        revalidating: true,
        projectsCount: Array.isArray(state.projects) ? state.projects.length : 0,
      };
    }

    // Em polling/background, cache fresco evita tráfego; cache vencido segue para a API.
    if (!force && background && fresh) {
      state.lastProjectsFetchAt = Date.now();
      return {
        ok: true,
        fromCache: true,
        revalidating: false,
        projectsCount: Array.isArray(state.projects) ? state.projects.length : 0,
      };
    }
  }
  if (isClientUser() && !shouldUseCache && cached) {
    console.warn('[LoadProjects] Cache rejeitado para usuário cliente, forçando API call');
  }

  if (!force && state.loadingProjectsRequest) {
    return state.loadingProjectsRequest;
  }

  if (!background && !suppressLoadingState && !state.projects.length) {
    setProjectsLoadingState('Carregando dados operacionais...');
  }

  const request = (async () => {
    try {
      if (refreshProjectsButtonEl) {
        refreshProjectsButtonEl.disabled = true;
        refreshProjectsButtonEl.textContent = force ? 'Atualizando...' : 'Sincronizando...';
      }
      const preferServerCache = Boolean(options.preferServerCache || !force);
      const requestUrl = force ? "/api/projects?force=1" : (preferServerCache ? "/api/projects?preferCache=1" : "/api/projects");
      const response = await fetch(requestUrl, { cache: "no-store", credentials: "same-origin" });
      let data = null;

      try {
        data = await response.json();
      } catch (parseError) {
        throw new Error("Falha ao atualizar dados operacionais.");
      }

      if (response.status === 401) {
        state.user = null;
        clearProjectsCache();
        updateSessionUi();
        resetDashboardForLoggedOutState();
        openLoginModal(data?.error || "Faça login para visualizar o painel.");
        return;
      }

      if (!response.ok || !data.ok) {
        throw new Error(data?.error || "Falha ao carregar projetos.");
      }
      const projectsFromApi = Array.isArray(data.projects) ? data.projects : [];
      if (options.requireData && projectsFromApi.length === 0) {
        throw new Error(isClientUser()
          ? 'As BSPs do cliente ainda não foram recebidas. Mantendo carregamento até os dados aparecerem.'
          : 'Os dados operacionais ainda não foram recebidos. Mantendo carregamento até os dados aparecerem.');
      }
      if (isClientUser() && projectsFromApi.length === 0) {
        console.warn('[LoadProjects] Aviso: usuário cliente recebeu 0 projetos da API');
      }
      if (isClientUser() && options.requireClientPo && !isClientProjectsPayloadPoReady(data)) {
        throw new Error('As BSPs carregaram, mas as POs ainda não foram recebidas. Buscando a base completa novamente.');
      }
      state.lastProjectsFetchAt = Date.now();
      writeProjectsCache(data);
      applyProjectsPayload(data, { fromCache: false });
      if (!force && !background) {
        window.setTimeout(() => maybeTriggerTrackingCacheWatchdog('load-projects').catch(() => {}), 1200);
      }
      if (options.requireFreshCacheAfter) {
        const requiredTimeMs = getTrackingCacheTimestampMs(options.requireFreshCacheAfter);
        const currentTimeMs = getCurrentTrackingCacheTimestampMs();
        if (!currentTimeMs || currentTimeMs < requiredTimeMs) {
          throw new Error('A API ainda retornou cache antigo. Aguardando gravacao nova do Supabase.');
        }
      }

      // v11 performance: se a API respondeu com snapshot/cache de boot rápido, libera a tela
      // imediatamente e dispara uma revalidação real em segundo plano. Nunca deixa o painel
      // em branco esperando o Smartsheet.
      const meta = data?.meta || {};
      const shouldRevalidateAfterFastBoot = Boolean(
        !force && !background && (
          meta.autoRefreshRecommended ||
          meta.servedFromDiskFastCache ||
          meta.servedFromFastCache ||
          meta.cacheReason === 'disk-fast-boot' ||
          meta.cacheReason === 'client-disk-fast-boot'
        )
      );
      if (shouldRevalidateAfterFastBoot) {
        if (lastSyncEl) {
          lastSyncEl.textContent = state.meta
            ? `${formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última base válida do cache' })} • sincronizando em segundo plano`
            : 'Base segura carregada • sincronizando em segundo plano';
        }
        window.setTimeout(() => revalidateProjectsInBackground(true), 800);
      }

      return {
        ok: true,
        fromCache: Boolean(meta.servedFromDiskFastCache || meta.servedFromFastCache),
        revalidating: shouldRevalidateAfterFastBoot,
        projectsCount: projectsFromApi.length,
      };
    } catch (error) {
      const fallbackMessage = error?.message || "Falha ao atualizar dados operacionais.";

      if (options.requireData) {
        throw error;
      }

      if (state.projects.length) {
        const staleSuffix = state.meta
          ? ` | exibindo ${formatTrackingCacheUpdateLabel(state.meta, { prefix: 'última base válida do cache' })}`
          : "";
        lastSyncEl.textContent = `Conexão instável com os dados operacionais${staleSuffix}`;
        console.warn("Falha temporária ao atualizar projetos:", fallbackMessage);
        return;
      }

      // v11 performance: se a API falhou antes de popular a tela, tenta reaplicar o último
      // cache local mesmo vencido. É melhor manter a última base visível do que deixar o
      // painel zerado enquanto a sincronização falha.
      const staleCache = readProjectsCache();
      const staleProjects = Array.isArray(staleCache?.payload?.projects) ? staleCache.payload.projects : [];
      if (staleProjects.length) {
        applyProjectsPayload(staleCache.payload, { fromCache: true, staleFallback: true });
        if (lastSyncEl) {
          const staleDate = staleCache.payload?.meta?.lastSync || staleCache.savedAt;
          lastSyncEl.textContent = staleDate
            ? `Conexão instável • usando última base válida de ${new Date(staleDate).toLocaleString("pt-BR")} • ${formatDurationPtBr(Date.now() - new Date(staleDate).getTime())}`
            : 'Conexão instável • usando última base válida local';
        }
        // v37.17: não força Smartsheet em recuperação automática de login.
        // Rele apenas cache Supabase/Function; atualização pesada fica no agendador ou botão.
        window.setTimeout(() => revalidateProjectsInBackground(false), 2500);
        return;
      }

      const retryCount = Number(options.retryCount || 0);
      const canRetryCacheRead = !background && !state.projects.length && retryCount < 2;
      if (canRetryCacheRead) {
        const waitMs = 1800 + retryCount * 2500;
        const retryMessage = `Aguardando cache operacional do Supabase. Tentando novamente (${retryCount + 1}/2)...`;
        if (bodyEl) bodyEl.innerHTML = `<tr><td colspan="21" class="loading-cell">${escapeHtml(retryMessage)}</td></tr>`;
        if (detailCardEl) detailCardEl.innerHTML = `<div class="detail-placeholder">${escapeHtml(retryMessage)}</div>`;
        if (lastSyncEl) lastSyncEl.textContent = retryMessage;
        window.setTimeout(() => {
          loadProjects({
            force: false,
            skipLocalCache: true,
            suppressLoadingState: false,
            preferServerCache: true,
            retryCount: retryCount + 1,
          }).catch((retryError) => {
            console.warn('Nova tentativa de leitura do cache falhou:', retryError?.message || retryError);
          });
        }, waitMs);
        return { ok: false, retrying: true, error: fallbackMessage };
      }

      bodyEl.innerHTML = `<tr><td colspan="21" class="loading-cell">${escapeHtml(fallbackMessage)}</td></tr>`;
      detailCardEl.innerHTML = `<div class="detail-placeholder">${escapeHtml(fallbackMessage)}</div>`;
      if (isClientUser()) {
        const syncEl = document.getElementById('client-dashboard-sync');
        const metaEl = document.getElementById('client-dashboard-meta');
        const grid = document.getElementById('client-vessel-grid');
        if (syncEl) syncEl.textContent = 'Carregando dados do cliente...';
        if (metaEl) metaEl.textContent = 'Aguardando retorno da carteira do cliente';
        if (grid && !grid.querySelector('[data-client-vessel]')) {
          grid.innerHTML = `<div class="client-empty">${escapeHtml(fallbackMessage)} Clique em Atualizar em alguns segundos.</div>`;
        }
      }
    } finally {
      state.loadingProjectsRequest = null;
      if (refreshProjectsButtonEl) {
        refreshProjectsButtonEl.disabled = false;
        refreshProjectsButtonEl.textContent = 'Atualizar agora';
      }
    }
  })();

  state.loadingProjectsRequest = request;
  return request;
}

function startPolling() {
  window.clearInterval(state.pollTimer);
  state.pollTimer = window.setInterval(async () => {
    if (!state.user || isPageHidden()) return;
    if (state.meta) {
      applyTrackingCacheUpdateLabel(lastSyncEl, state.meta, { prefix: 'Última atualização do cache' });
      const clientSyncEl = document.getElementById('client-dashboard-sync');
      if (clientSyncEl) applyTrackingCacheUpdateLabel(clientSyncEl, state.meta, { prefix: 'Última atualização do cache' });
      maybeTriggerTrackingCacheWatchdog('polling').catch(() => {});
    }

    const now = Date.now();
    if (now - state.lastProjectsFetchAt >= PROJECTS_REFRESH_MS) {
      await loadProjects({ background: true });
    }

    if (now - state.lastManualAlertsFetchAt >= ALERTS_REFRESH_MS) {
      await loadManualAlerts({ background: true });
    }

    if (now - state.lastAlertResponsesFetchAt >= ALERTS_REFRESH_MS && !adminModalEl?.classList.contains('hidden')) {
      await loadAlertResponses({ background: true });
    }

    if (isStageUpdatesWorkspaceOpen() && now - state.lastStageUpdatesFetchAt >= ALERTS_REFRESH_MS) {
      await loadStageUpdates({ background: true });
    }
  }, 15000);
}


let clientChartTooltipEl = null;

function getClientChartTooltipEl() {
  if (clientChartTooltipEl) return clientChartTooltipEl;
  clientChartTooltipEl = document.createElement('div');
  clientChartTooltipEl.className = 'client-chart-tooltip';
  clientChartTooltipEl.setAttribute('aria-hidden', 'true');
  document.body.appendChild(clientChartTooltipEl);
  return clientChartTooltipEl;
}

function showClientChartTooltip(event, text) {
  if (!text) return;
  const tooltip = getClientChartTooltipEl();
  tooltip.innerHTML = escapeHtml(text).replace(/\n/g, '<br>');
  tooltip.style.left = `${Math.min(window.innerWidth - 220, Math.max(12, event.clientX + 14))}px`;
  tooltip.style.top = `${Math.max(12, event.clientY - 18)}px`;
  tooltip.classList.add('visible');
  tooltip.setAttribute('aria-hidden', 'false');
}

function hideClientChartTooltip() {
  if (!clientChartTooltipEl) return;
  clientChartTooltipEl.classList.remove('visible');
  clientChartTooltipEl.setAttribute('aria-hidden', 'true');
}

function bindClientChartTooltips() {
  document.addEventListener('mousemove', (event) => {
    const target = event.target?.closest?.('[data-client-chart-tooltip]');
    if (!target) return;
    showClientChartTooltip(event, target.getAttribute('data-client-chart-tooltip') || '');
  });
  document.addEventListener('mouseover', (event) => {
    const target = event.target?.closest?.('[data-client-chart-tooltip]');
    if (!target) return;
    showClientChartTooltip(event, target.getAttribute('data-client-chart-tooltip') || '');
  });
  document.addEventListener('mouseout', (event) => {
    const target = event.target?.closest?.('[data-client-chart-tooltip]');
    if (!target) return;
    hideClientChartTooltip();
  });
}

function bindEvents() {
  bindClientChartTooltips();
  if (attentionPopupCloseEl) {
    attentionPopupCloseEl.addEventListener('click', () => closeAttentionPopup());
  }
  if (attentionPopupActionEl) {
    attentionPopupActionEl.addEventListener('click', () => closeAttentionPopup({ openTarget: true }));
  }
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      showNextAttentionPopup();
      sendPresenceHeartbeat({ force: true });
      if (state.user) {
        loadProjects({ background: true }).catch(() => {});
        loadManualAlerts({ background: true }).catch(() => {});
        if (isStageUpdatesWorkspaceOpen()) loadStageUpdates({ background: true }).catch(() => {});
      }
    }
  });
  if (refreshProjectsButtonEl) {
    refreshProjectsButtonEl.addEventListener('click', async () => {
      let manualRefreshCompleted = false;
      try {
        clearProjectsCache();
        const previousCacheTimeMs = getCurrentTrackingCacheTimestampMs();
        if (refreshProjectsButtonEl) {
          refreshProjectsButtonEl.disabled = true;
          refreshProjectsButtonEl.textContent = 'Atualizando...';
        }
        const syncResult = await triggerTrackingCacheSync({ force: true, manual: true, auto: false });
        const confirmedCacheUpdatedAt = syncResult?.background || syncResult?.accepted
          ? await waitForManualTrackingCacheUpdate(previousCacheTimeMs)
          : assertManualTrackingSyncCompleted(syncResult, previousCacheTimeMs);
        if (syncResult?.cacheUpdatedAt) {
          state.meta = {
            ...(state.meta || {}),
            persistentCacheUpdatedAt: confirmedCacheUpdatedAt,
            persistentCacheAgeMs: syncResult.cacheAgeMs ?? 0,
            cacheUpdatedAt: confirmedCacheUpdatedAt,
            cacheAgeMs: syncResult.cacheAgeMs ?? 0,
          };
          applyTrackingCacheUpdateLabel(lastSyncEl, state.meta, { prefix: 'Última atualização do cache' });
        }
        if (syncResult?.staleCacheKept && lastSyncEl) {
          const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
          lastSyncEl.textContent = `${label} • ${describeTrackingSyncWarning(syncResult)}`;
        }
        // v37.24: não exibir popup se a sincronização manual falhar mas existe cache operacional.
        // O painel deve continuar usando a última base válida e mostrar a condição inline.
        if (syncResult && syncResult.ok === false) {
          const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
          if (lastSyncEl && hasVisibleOperationalCache()) {
            lastSyncEl.textContent = `${label} • ${describeTrackingSyncWarning(syncResult)}`;
          } else {
            throw new Error(syncResult.error || 'Falha ao sincronizar cache.');
          }
        }
        await loadProjects({
          force: false,
          skipLocalCache: true,
          suppressLoadingState: false,
          preferServerCache: true,
          requireFreshCacheAfter: confirmedCacheUpdatedAt,
        });
        manualRefreshCompleted = true;
      } catch (error) {
        if (lastSyncEl && hasVisibleOperationalCache()) {
          const label = formatTrackingCacheUpdateLabel(state.meta, { prefix: 'Última atualização do cache' });
          lastSyncEl.textContent = `${label} • cache mantido; tentativa manual não concluída`;
          console.warn('[tracking-sync] Atualização manual não concluída:', error?.message || error);
        } else {
          window.alert(error?.message || 'Falha ao atualizar agora.');
        }
      } finally {
        if (refreshProjectsButtonEl) {
          refreshProjectsButtonEl.disabled = false;
          refreshProjectsButtonEl.textContent = manualRefreshCompleted ? 'Atualizado' : 'Tentar novamente';
          if (manualRefreshCompleted) {
            window.setTimeout(() => {
              if (refreshProjectsButtonEl) refreshProjectsButtonEl.textContent = 'Atualizar agora';
            }, 2500);
          }
        }
      }
    });
  }
  if (exportFilteredProjectsEl) {
    exportFilteredProjectsEl.addEventListener('click', downloadFilteredProjectsExcel);
  }

  if (sectorAlertsContentEl) {
    sectorAlertsContentEl.addEventListener('click', async (event) => {
      const button = event.target.closest('[data-enable-push]');
      if (!button) return;
      button.disabled = true;
      try {
        const ok = await syncPushSubscription(true);
        if (!ok) window.alert('Permita as notificações do navegador e instale o app para receber push no telefone.');
        renderManualAlerts();
      } catch (error) {
        window.alert(error?.message || 'Falha ao ativar push.');
      } finally {
        button.disabled = false;
      }
    });
  }

  searchInputEl.addEventListener("input", (event) => {
    state.searchQuery = event.target.value;
    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
    tableShellEl.scrollTop = 0;
  });

  clearSearchEl.addEventListener("click", () => {
    state.searchQuery = "";
    state.demandFilter = "";
    state.projectTypeFilter = "";
    state.weekFilter = "";
    state.statusFilters = [];
    searchInputEl.value = "";
    if (demandFilterEl) demandFilterEl.value = "";
    if (projectTypeFilterEl) projectTypeFilterEl.value = "";
    if (weekFilterEl) weekFilterEl.value = "";
    renderStatusFilterMenu();
    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
    tableShellEl.scrollTop = 0;
    searchInputEl.focus();
  });

  if (demandFilterEl) {
    demandFilterEl.addEventListener("change", (event) => {
      state.demandFilter = event.target.value;
      applyFilter();
      renderStats();
      renderTable();
      renderSelectedProjectCard();
      tableShellEl.scrollTop = 0;
    });
  }

  if (projectTypeFilterEl) {
    projectTypeFilterEl.addEventListener("change", (event) => {
      state.projectTypeFilter = event.target.value;
      applyFilter();
      renderStats();
      renderTable();
      renderSelectedProjectCard();
      tableShellEl.scrollTop = 0;
    });
  }

  if (weekFilterEl) {
    weekFilterEl.addEventListener("change", (event) => {
      state.weekFilter = event.target.value;
      applyFilter();
      renderStats();
      renderTable();
      renderSelectedProjectCard();
      tableShellEl.scrollTop = 0;
    });
  }

  if (statusFilterToggleEl) {
    renderStatusFilterMenu();
    statusFilterToggleEl.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleStatusFilterMenu();
    });
  }

  if (statusFilterMenuEl) {
    statusFilterMenuEl.addEventListener("click", (event) => {
      event.stopPropagation();
      const allTarget = event.target.closest('[data-status-filter-all]');
      if (allTarget) {
        state.statusFilters = [];
        renderStatusFilterMenu();
        applyFilter();
        renderStats();
        renderTable();
        renderSelectedProjectCard();
        tableShellEl.scrollTop = 0;
        return;
      }
      const optionTarget = event.target.closest('[data-status-filter]');
      if (!optionTarget) return;
      const value = String(optionTarget.getAttribute('data-status-filter') || '').trim();
      if (!value) return;
      const current = new Set(getSelectedStatusFilters());
      if (current.has(value)) current.delete(value);
      else current.add(value);
      const next = Array.from(current);
      state.statusFilters = next.length === PROJECT_STATUS_FILTER_OPTIONS.length ? [] : next;
      renderStatusFilterMenu();
      applyFilter();
      renderStats();
      renderTable();
      renderSelectedProjectCard();
      tableShellEl.scrollTop = 0;
    });
  }

  const attachProjectDrillCard = (cardEl, mode) => {
    if (!cardEl) return;
    const openDrill = () => openProjectDrillPanel(mode);
    cardEl.addEventListener("click", openDrill);
    cardEl.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openDrill();
      }
    });
  };

  attachProjectDrillCard(totalProjectsCardEl, 'total');
  attachProjectDrillCard(startedProjectsCardEl, 'started');
  attachProjectDrillCard(notStartedCardEl, 'not-started');
  attachProjectDrillCard(onHoldCardEl, 'hold');
  attachProjectDrillCard(productionCardEl, 'production');
  attachProjectDrillCard(inspectionCardEl, 'inspection');
  attachProjectDrillCard(paintingCardEl, 'painting');
  attachProjectDrillCard(awaitingShipmentCardEl, 'awaiting');
  attachProjectDrillCard(totalWeightCardEl, 'total-weight');
  attachProjectDrillCard(weldedWeightCardEl, 'welded');
  attachProjectDrillCard(backlogWeldingCardEl, 'backlog');
  attachProjectDrillCard(paintingM2CardEl, 'painting-m2');

  const attachPcpDeadlineCard = (cardId, mode) => {
    const cardEl = document.getElementById(cardId);
    if (!cardEl) return;
    const openModal = () => openPcpDeadlinesModal(mode);
    cardEl.addEventListener('click', openModal);
    cardEl.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openModal();
      }
    });
  };
  attachPcpDeadlineCard('pcp-overdue-card', 'overdue');
  attachPcpDeadlineCard('pcp-due-soon-card', 'due-soon');

  const pcpDeadlinesModalEl = document.getElementById('pcp-deadlines-modal');
  const pcpDeadlinesCloseEl = document.getElementById('pcp-deadlines-close');
  const pcpDeadlinesContentEl = document.getElementById('pcp-deadlines-content');
  if (pcpDeadlinesCloseEl) pcpDeadlinesCloseEl.addEventListener('click', closePcpDeadlinesModal);
  if (pcpDeadlinesModalEl) {
    pcpDeadlinesModalEl.addEventListener('click', (event) => {
      if (event.target.closest('[data-close-pcp-deadlines="true"]')) closePcpDeadlinesModal();
    });
  }
  if (pcpDeadlinesContentEl) {
    pcpDeadlinesContentEl.addEventListener('click', (event) => {
      const toggle = event.target.closest('[data-pcp-deadline-toggle]');
      if (toggle) {
        const key = toggle.dataset.pcpDeadlineToggle || '';
        const block = Array.from(pcpDeadlinesContentEl.querySelectorAll('[data-pcp-deadline-client-block]')).find((item) => item.dataset.pcpDeadlineClientBlock === key);
        const body = Array.from(pcpDeadlinesContentEl.querySelectorAll('[data-pcp-deadline-client-body]')).find((item) => item.dataset.pcpDeadlineClientBody === key);
        if (!block || !body) return;
        const expanded = !block.classList.contains('is-expanded');
        block.classList.toggle('is-expanded', expanded);
        body.classList.toggle('hidden', !expanded);
        toggle.setAttribute('aria-expanded', expanded ? 'true' : 'false');
        return;
      }
      const row = event.target.closest('tr[data-pcp-deadline-project-id]');
      if (!row) return;
      pcpDeadlinesContentEl.querySelectorAll('tr[data-pcp-deadline-project-id]').forEach((item) => item.classList.remove('project-drill-row-selected'));
      row.classList.add('project-drill-row-selected');
      const projectId = Number(row.dataset.pcpDeadlineProjectId);
      const project = state.projects.find((item) => Number(item.rowId) === projectId);
      if (!project) return;
      state.selectedProjectId = projectId;
      renderTable();
      renderSelectedProjectCard();
    });
    pcpDeadlinesContentEl.addEventListener('dblclick', (event) => {
      if (event.target.closest('[data-pcp-deadline-toggle]')) return;
      const row = event.target.closest('tr[data-pcp-deadline-project-id]');
      if (!row) return;
      const projectId = Number(row.dataset.pcpDeadlineProjectId);
      const project = state.projects.find((item) => Number(item.rowId) === projectId);
      if (!project) return;
      state.selectedProjectId = projectId;
      renderTable();
      openProjectModal(project);
    });
  }

  if (projectDrillCloseEl) {
    projectDrillCloseEl.addEventListener("click", closeProjectDrillPanel);
  }

  if (projectDrillBackEl) {
    projectDrillBackEl.addEventListener("click", () => {
      if (state.projectDrill.selectedVesselKey) {
        state.projectDrill.selectedVesselKey = '';
      } else if (state.projectDrill.selectedClientKey) {
        state.projectDrill.selectedClientKey = '';
      }
      renderProjectDrillPanel();
    });
  }

  if (projectDrillContentEl) {
    projectDrillContentEl.addEventListener("click", (event) => {
      const clientButton = event.target.closest("[data-drill-client]");
      if (clientButton) {
        const clientKey = clientButton.dataset.drillClient || '';
        setProjectDrillLevel({ clientKey, vesselKey: '' });
        return;
      }

      const vesselButton = event.target.closest("[data-drill-vessel]");
      if (vesselButton) {
        const vesselKey = vesselButton.dataset.drillVessel || '';
        setProjectDrillLevel({ clientKey: state.projectDrill.selectedClientKey, vesselKey });
        return;
      }

      const projectRow = event.target.closest("tr[data-drill-project-id]");
      if (projectRow) {
        const projectId = Number(projectRow.dataset.drillProjectId);
        const project = state.projects.find((item) => item.rowId === projectId);
        if (!project) return;
        state.selectedProjectId = projectId;
        renderTable();
        projectRow.classList.add("project-drill-row-selected");
      }
    });

    projectDrillContentEl.addEventListener("dblclick", (event) => {
      const projectRow = event.target.closest("tr[data-drill-project-id]");
      if (!projectRow) return;
      const projectId = Number(projectRow.dataset.drillProjectId);
      const project = state.projects.find((item) => item.rowId === projectId);
      if (!project) return;
      state.selectedProjectId = projectId;
      renderTable();
      openProjectModal(project);
    });
  }

  if (projectDrillBreadcrumbEl) {
    projectDrillBreadcrumbEl.addEventListener("click", (event) => {
      const clientButton = event.target.closest("[data-drill-client]");
      if (!clientButton) return;
      const clientKey = clientButton.dataset.drillClient || '';
      setProjectDrillLevel({ clientKey, vesselKey: '' });
    });
  }

  document.addEventListener('click', (event) => {
    if (!statusFilterBoxEl || statusFilterMenuEl?.classList.contains('hidden')) return;
    if (!statusFilterBoxEl.contains(event.target)) closeStatusFilterMenu();
  });

  if (closeDetailDrawerEl) {
    closeDetailDrawerEl.addEventListener("click", () => {
      state.selectedProjectDrawerOpen = false;
      renderSelectedProjectCard();
    });
  }

  bodyEl.addEventListener("click", (event) => {
    const row = event.target.closest("tr[data-project-id]");
    if (!row) return;
    const projectId = Number(row.dataset.projectId);
    const project = state.projects.find((item) => item.rowId === projectId);
    if (!project) return;

    window.clearTimeout(state.rowClickTimer);
    state.rowClickTimer = window.setTimeout(() => {
      state.selectedProjectId = projectId;
      state.selectedProjectDrawerOpen = true;
      renderTable();
      renderSelectedProjectCard();
      state.rowClickTimer = null;
    }, 220);
  });

  bodyEl.addEventListener("dblclick", (event) => {
    const row = event.target.closest("tr[data-project-id]");
    if (!row) return;
    const projectId = Number(row.dataset.projectId);
    const project = state.projects.find((item) => item.rowId === projectId);
    if (!project) return;

    window.clearTimeout(state.rowClickTimer);
    state.rowClickTimer = null;
    state.selectedProjectId = projectId;
    renderTable();
    renderSelectedProjectCard();
    openProjectModal(project);
  });

  modalEl.addEventListener("click", (event) => {
    if (event.target.matches("[data-close-modal='true']")) {
      closeProjectModal();
      return;
    }
    const signalButton = event.target.closest('[data-open-project-signal]');
    if (signalButton) {
      const project = state.projects.find((item) => String(item.rowId) === String(signalButton.dataset.openProjectSignal));
      if (project) openProjectSignalModal(project);
      return;
    }
    const resolveButton = event.target.closest('[data-resolve-signal]');
    if (resolveButton) {
      resolveSignal(resolveButton.dataset.resolveSignal);
    }
  });

  modalCloseEl.addEventListener("click", closeProjectModal);

  if (alertModalCloseEl) {
    alertModalCloseEl.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      closeAlertModal();
    });
  }

  if (alertModalEl) {
    alertModalEl.addEventListener("click", (event) => {
      if (event.target.matches("[data-close-alert='true']")) {
        closeAlertModal();
        return;
      }

      const filterButton = event.target.closest("[data-alert-filter]");
      if (filterButton) {
        state.alertFilter = filterButton.dataset.alertFilter || "all";
        renderAlertModal();
        return;
      }

      const sectorButton = event.target.closest("[data-alert-sector]");
      if (sectorButton) {
        state.alertSectorFilter = sectorButton.dataset.alertSector || "all";
        renderAlertModal();
        return;
      }

      const clientSearchInput = event.target.closest("[data-alert-client-search]");
      if (clientSearchInput) {
        return;
      }

      const downloadPdfButton = event.target.closest("[data-alert-download-pdf]");
      if (downloadPdfButton) {
        downloadAlertsPdf();
        return;
      }

      const alertItem = event.target.closest("[data-alert-project-id], [data-alert-project-number]");
      if (alertItem) {
        const project = findProjectFromAlertElement(alertItem);
        if (!project) return;
        closeAlertModal();
        state.selectedProjectId = project.rowId;
        applyFilter();
        renderTable();
        renderSelectedProjectCard();
        openProjectModal(project);
      }
    });

    alertModalEl.addEventListener("input", (event) => {
      const clientInput = event.target.closest("[data-alert-client-search]");
      if (!clientInput) return;
      const caret = clientInput.selectionStart ?? clientInput.value.length;
      state.alertClientQuery = clientInput.value || "";
      renderAlertModal();
      const nextInput = alertModalEl.querySelector("[data-alert-client-search]");
      if (nextInput) {
        nextInput.focus();
        nextInput.setSelectionRange(caret, caret);
      }
    });
  }

  if (openAlertsButtonEl) {
    openAlertsButtonEl.addEventListener("click", () => {
      renderAlertModal();
      openAlertModal(true, { manual: true });
    });
  }

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;

    if (alertModalEl && !alertModalEl.classList.contains("hidden")) {
      closeAlertModal();
      return;
    }

    if (modalEl && !modalEl.classList.contains("hidden")) {
      closeProjectModal();
      return;
    }

    if (loginModalEl && !loginModalEl.classList.contains("hidden")) {
      closeLoginModal();
      return;
    }
  });

if (loginFormEl) {
  loginFormEl.addEventListener("submit", handleLoginSubmit);
}

if (openLoginButtonEl) {
  openLoginButtonEl.addEventListener("click", () => {
    openLoginModal();
  });
}

if (openClientApiButtonEl) {
  openClientApiButtonEl.addEventListener("click", () => {
    openClientApiModal();
  });
}

if (loginCloseEl) {
  loginCloseEl.addEventListener("click", closeLoginModal);
}

const adminUserSectorEl = document.getElementById("admin-user-sector");
const adminUserRoleEl = document.getElementById("admin-user-role");
const adminUserProjectPmsFieldEl = document.getElementById("admin-user-project-pms-field");
const adminUserProjectPmsSearchEl = document.getElementById("admin-user-project-pms-search");
const adminUserProjectPmsOptionsEl = document.getElementById("admin-user-project-pms-options");
const adminUserQualityCompetenciesFieldEl = document.getElementById("admin-user-quality-competencies-field");
if (adminUserSectorEl) {
  adminUserSectorEl.addEventListener("change", (event) => {
    const next = normalizeSectorValue(event.target.value);
    const selected = new Set(getSelectedAdminAlertSectors());
    if (next) {
      selected.add(next);
      setSelectedAdminAlertSectors([...selected]);
    }
    updateAdminProjectPmAliasesVisibility();
  });
}

if (adminUserRoleEl) {
  adminUserRoleEl.addEventListener("change", (event) => {
    const role = event.target.value;
    if ((role === "admin" || role === "client") && adminUserSectorEl) adminUserSectorEl.value = "all";
    const disabled = role === "admin" || role === "client";
    document.querySelectorAll('[data-admin-alert-sector-option]').forEach((input) => {
      input.disabled = disabled;
    });
    updateAdminClientFieldsVisibility();
    updateAdminProjectPmAliasesVisibility();
  });
}

if (adminUserClientLogoImportEl) {
  adminUserClientLogoImportEl.addEventListener('click', importAdminClientLogoWithEditor);
}

if (adminUserClientLogoUrlEl) {
  adminUserClientLogoUrlEl.addEventListener('input', () => {
    resetAdminLogoEditor(adminUserClientLogoUrlEl.value || '');
  });
}

if (adminClientLogoEditorEl) {
  adminClientLogoEditorEl.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-logo-adjust]');
    if (!button) return;
    const action = button.dataset.logoAdjust;
    if (action === 'zoomIn') adminLogoEditorState.zoom = Math.min(3, adminLogoEditorState.zoom + 0.08);
    if (action === 'zoomOut') adminLogoEditorState.zoom = Math.max(0.35, adminLogoEditorState.zoom - 0.08);
    if (action === 'left') adminLogoEditorState.x -= 8;
    if (action === 'right') adminLogoEditorState.x += 8;
    if (action === 'up') adminLogoEditorState.y -= 8;
    if (action === 'down') adminLogoEditorState.y += 8;
    if (action === 'reset') resetAdminLogoEditor();
    if (action === 'apply') {
      await applyAdminClientLogoAdjustment();
      return;
    }
    updateAdminLogoEditorPreview();
  });
}


function setClientPlatformImageLine(platformName, imageUrl) {
  const key = String(platformName || '').trim();
  const src = String(imageUrl || '').trim();
  if (!key || !src || !adminUserClientPlatformImagesEl) return;
  const map = parseClientPlatformImages(adminUserClientPlatformImagesEl.value || '');
  map[key] = src;
  adminUserClientPlatformImagesEl.value = formatClientPlatformImages(map);
  adminUserClientPlatformImagesEl.dispatchEvent(new Event('input', { bubbles: true }));
}

async function importAdminClientPlatformImage() {
  const platformName = adminUserClientPlatformNameEl?.value || '';
  if (!String(platformName || '').trim()) {
    window.alert('Informe o nome da plataforma/vessel antes de importar a imagem. Ex.: FORTE, FRADE, BRAVO.');
    return;
  }
  const file = adminUserClientPlatformImageFileEl?.files?.[0];
  if (!file) {
    window.alert('Selecione a imagem da plataforma primeiro.');
    return;
  }
  try {
    const dataUrl = await readImageFileAsOptimizedDataUrl(file, { maxWidth: 520, maxHeight: 340, quality: 0.68 });
    setClientPlatformImageLine(platformName, dataUrl);
    if (adminUserClientPlatformImageFileEl) adminUserClientPlatformImageFileEl.value = '';
    if (adminUserFeedbackEl) adminUserFeedbackEl.textContent = `Foto da plataforma ${platformName} importada. Salve o usuário para gravar.`;
  } catch (error) {
    window.alert(error.message || 'Falha ao importar foto da plataforma.');
  }
}

if (adminUserClientPlatformImageImportEl) {
  adminUserClientPlatformImageImportEl.addEventListener('click', importAdminClientPlatformImage);
}

document.querySelectorAll('[data-admin-alert-sector-option]').forEach((input) => {
  input.addEventListener('change', updateAdminProjectPmAliasesVisibility);
});

if (adminUserProjectPmsSearchEl) {
  adminUserProjectPmsSearchEl.addEventListener('input', (event) => {
    setAdminProjectPmSearchQuery(event.target.value);
  });
}

if (adminUserProjectPmsOptionsEl) {
  adminUserProjectPmsOptionsEl.addEventListener('change', (event) => {
    const input = event.target?.closest?.('input[data-admin-project-pm-option]');
    if (!input) return;
    toggleAdminProjectPmAlias(input.value, input.checked);
  });
}

document.querySelectorAll('[data-admin-quality-competency-option]').forEach((input) => {
  input.addEventListener('change', (event) => {
    toggleAdminQualityCompetency(event.target.value, event.target.checked);
  });
});

updateAdminClientFieldsVisibility();
updateAdminProjectPmAliasesVisibility();
updateAdminQualityCompetenciesVisibility();

if (loginModalEl) {
  loginModalEl.addEventListener("click", (event) => {
    if (event.target === loginModalEl || event.target.matches(".modal-backdrop")) {
      closeLoginModal();
    }
  });
}

if (logoutButtonEl) {
  logoutButtonEl.addEventListener("click", handleLogout);
}

if (openChangePasswordButtonEl) {
  openChangePasswordButtonEl.addEventListener("click", openChangePasswordModal);
}

if (changePasswordCloseEl) {
  changePasswordCloseEl.addEventListener("click", closeChangePasswordModal);
}

const changePasswordCancelEl = document.getElementById("change-password-cancel");
if (changePasswordCancelEl) {
  changePasswordCancelEl.addEventListener("click", closeChangePasswordModal);
}

if (changePasswordModalEl) {
  changePasswordModalEl.addEventListener("click", (event) => {
    if (event.target?.dataset?.closeChangePassword === "true") {
      closeChangePasswordModal();
    }
  });
}

if (changePasswordFormEl) {
  changePasswordFormEl.addEventListener("submit", handleChangePasswordSubmit);
}

if (projectViewTabsEl) {
  projectViewTabsEl.addEventListener('click', (event) => {
    const button = event.target.closest('[data-project-view]');
    if (!button) return;
    const nextView = button.dataset.projectView === 'mine' ? 'mine' : 'all';
    if (nextView === state.projectView) return;
    state.projectView = nextView;
    updatePrimaryUserActionUi();
    renderProjectViewTabs();
    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
  });
}

if (openSectorAlertsEl) {
  openSectorAlertsEl.addEventListener("click", () => {
    if (!state.user) {
      openLoginModal();
      return;
    }
    state.sectorAlertsMode = 'default';
    if (!shouldUseSectorScopedToggle() && userHasProjectsScope()) {
      state.projectView = state.projectView === 'mine' ? 'all' : 'mine';
      updatePrimaryUserActionUi();
      renderProjectViewTabs();
      applyFilter();
      renderStats();
      renderTable();
      renderSelectedProjectCard();
      renderAlertBadge();
      if (alertModalEl && !alertModalEl.classList.contains('hidden')) {
        renderAlertModal();
      }
      if (tableShellEl) tableShellEl.scrollTop = 0;
      return;
    }
    state.sectorScopedView = !state.sectorScopedView;
    saveSectorScopedViewPreference(state.sectorScopedView);
    state.alertSectorFilter = state.sectorScopedView ? normalizeAlertSectorFilterValue(getPrimaryUserSector()) || 'all' : 'all';
    updatePrimaryUserActionUi();
    applyFilter();
    renderStats();
    renderTable();
    renderSelectedProjectCard();
    renderAlertBadge();
    if (alertModalEl && !alertModalEl.classList.contains('hidden')) {
      renderAlertModal();
    }
    if (tableShellEl) tableShellEl.scrollTop = 0;
  });
}

if (openMyProjectSignalsEl) {
  openMyProjectSignalsEl.addEventListener('click', () => {
    if (!state.user) {
      openLoginModal();
      return;
    }
    state.sectorAlertsMode = 'my-project-signals';
    const titleEl = document.getElementById('sector-alerts-title');
    if (titleEl) titleEl.textContent = 'Minhas sinalizações ao PCP';
    openSectorAlertsModal();
  });
}

if (openProjectSignalsEl) {
  openProjectSignalsEl.addEventListener('click', () => {
    if (!state.user) {
      openLoginModal();
      return;
    }
    state.sectorAlertsMode = 'project-signals';
    const titleEl = document.getElementById('sector-alerts-title');
    if (titleEl) titleEl.textContent = 'Alertas enviados por Projetos';
    openSectorAlertsModal();
  });
}

if (openStageUpdatesEl) {
  openStageUpdatesEl.addEventListener('click', () => {
    if (!state.user) {
      openLoginModal();
      return;
    }
    state.stageUpdatesSearchQuery = '';
    state.stagePcpPointingMode = false;
    if (isPcpStageUser()) ensurePcpStageSectorDefault();
    syncStageDraftsForCurrentSector();

    if (canValidateStageWorkspace()) {
      openStageValidationWorkspaceInline();
      return;
    }

    // Abre a tela imediatamente. O carregamento/filtragem acontece depois para não parecer travado.
    openStageUpdatesModal({ loading: true });

    loadStageUpdates()
      .then(() => {
        if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
          renderStageUpdatesModal();
        }
      })
      .catch(() => {
        // loadStageUpdates já renderiza fallback seguro.
      });
  });
}

if (stageUpdatesCloseEl) {
  stageUpdatesCloseEl.addEventListener('click', closeStageUpdatesModal);
}

if (stageUpdatesModalEl) {
  stageUpdatesModalEl.addEventListener('click', (event) => {
    if (event.target.matches('[data-close-stage-updates="true"]')) {
      closeStageUpdatesModal();
      return;
    }
    const openPcpPointingButton = event.target.closest('[data-stage-open-pcp-pointing="true"]');
    if (openPcpPointingButton) {
      const selectEl = stageUpdatesModalEl.querySelector('[data-pcp-stage-sector-select="true"]');
      const selectedSector = normalizeSectorValue(selectEl?.value || state.pcpStageSelectedSector || '');
      if (!STAGE_WORKSPACE_SECTORS.includes(selectedSector)) {
        window.alert('Selecione o setor que o PCP irá apontar.');
        return;
      }
      state.pcpStageSelectedSector = selectedSector;
      state.stagePcpPointingMode = true;
      state.stageUpdatesSearchQuery = '';
      syncStageDraftsForCurrentSector();
      renderStageUpdatesModal();
      return;
    }
    if (event.target.closest('[data-stage-back-validation="true"]')) {
      state.stagePcpPointingMode = false;
      state.stageUpdatesSearchQuery = '';
      syncStageDraftsForCurrentSector();
      renderStageUpdatesModal();
      return;
    }
    const masterCheck = event.target.closest('[data-stage-master-check="true"]');
    if (masterCheck) {
      const pending = getFilteredStageUpdatesForValidation().filter((item) => isPendingStageStatus(item.status));
      const ids = masterCheck.checked ? pending.filter(isStageUpdateSelectableForTracking).map((item) => item.id) : [];
      setStageSelection(ids);
      renderStageUpdatesModal();
      return;
    }
    const itemCheck = event.target.closest('[data-stage-item-check]');
    if (itemCheck) {
      const id = String(itemCheck.dataset.stageItemCheck || '').trim();
      const current = new Set(state.stageSelectedIds || []);
      if (itemCheck.checked) current.add(id);
      else current.delete(id);
      setStageSelection(Array.from(current));
      renderStageUpdatesModal();
      return;
    }
    const dateMasterCheck = event.target.closest('[data-stage-date-master-check="true"]');
    if (dateMasterCheck) {
      const ids = dateMasterCheck.checked ? (state.stageDatePendencies || []).map((item) => item.id) : [];
      setStageDateSelection(ids);
      renderStageUpdatesModal();
      return;
    }
    const dateItemCheck = event.target.closest('[data-stage-date-item-check]');
    if (dateItemCheck) {
      const id = String(dateItemCheck.dataset.stageDateItemCheck || '').trim();
      const current = new Set(state.stageDateSelectedIds || []);
      if (dateItemCheck.checked) current.add(id);
      else current.delete(id);
      setStageDateSelection(Array.from(current));
      renderStageUpdatesModal();
      return;
    }
    const trackingUpdateButton = event.target.closest('[data-stage-tracking-update]');
    if (trackingUpdateButton) {
      sendStageTrackingUpdate([trackingUpdateButton.dataset.stageTrackingUpdate], { forceRewrite: false });
      return;
    }
    const trackingRewriteButton = event.target.closest('[data-stage-tracking-rewrite]');
    if (trackingRewriteButton) {
      sendStageTrackingUpdate([trackingRewriteButton.dataset.stageTrackingRewrite], { forceRewrite: true });
      return;
    }
    if (event.target.closest('[data-stage-tracking-bulk="true"]')) {
      sendStageTrackingUpdate(state.stageSelectedIds || [], { forceRewrite: true });
      return;
    }
    if (event.target.closest('[data-stage-conclude-bulk-ok="true"]')) {
      concludeStageUpdatesBulkOk();
      return;
    }
    if (event.target.closest('[data-stage-load-date-pending="true"]')) {
      loadStageHistoryDatePendencies();
      return;
    }
    const dateFixButton = event.target.closest('[data-stage-date-fix]');
    if (dateFixButton) {
      sendStageTrackingUpdate([dateFixButton.dataset.stageDateFix], { dateOnly: true, forceRewrite: true });
      return;
    }
    if (event.target.closest('[data-stage-date-bulk="true"]')) {
      sendStageTrackingUpdate(state.stageDateSelectedIds || [], { dateOnly: true, forceRewrite: true });
      return;
    }
    if (event.target.closest('[data-stage-date-fix-all="true"]')) {
      sendStageTrackingUpdate((state.stageDatePendencies || []).map((item) => item.id), { dateOnly: true, forceRewrite: true });
      return;
    }
    const deleteButton = event.target.closest('[data-stage-delete]');
    if (deleteButton) {
      deleteStageUpdatePending(deleteButton.dataset.stageDelete);
      return;
    }
    const concludeButton = event.target.closest('[data-stage-conclude]');
    if (concludeButton) {
      concludeStageUpdate(concludeButton.dataset.stageConclude);
      return;
    }
    if (event.target.closest('[data-stage-bulk-send="true"]')) {
      handleStageWorkspaceBulkSubmit();
      return;
    }
    if (event.target.closest('[data-stage-clear-drafts="true"]')) {
      clearAllStageDrafts();
      renderStageUpdatesModal();
      return;
    }
    if (event.target.closest('[data-stage-toggle-batch="true"]')) {
      state.stageBatchValidationMode = !state.stageBatchValidationMode;
      renderStageUpdatesModal();
      return;
    }
    if (event.target.closest('[data-stage-conclude-bulk="true"]')) {
      const ids = (Array.isArray(state.stageUpdates) ? state.stageUpdates : []).filter((item) => isPendingStageStatus(item.status)).map((item) => item.id);
      concludeStageUpdatesBulk(ids);
      return;
    }
    const actionButton = event.target.closest('[data-stage-send="true"], [data-stage-review="true"]');
    if (!actionButton) return;
    const rowEl = actionButton.closest('tr');
    const formEl = rowEl?.querySelector('[data-stage-update-form="true"]');
    if (!formEl) return;
    const actionType = actionButton.matches('[data-stage-review="true"]') ? 'review' : 'advance';
    handleStageWorkspaceSubmit(formEl, actionType);
  });
  stageUpdatesModalEl.addEventListener('change', (event) => {
    const pcpSectorEl = event.target.closest('[data-pcp-stage-sector-switch="true"]');
    if (pcpSectorEl) {
      const selectedSector = normalizeSectorValue(pcpSectorEl.value || '');
      if (STAGE_WORKSPACE_SECTORS.includes(selectedSector)) {
        state.pcpStageSelectedSector = selectedSector;
        state.stagePcpPointingMode = true;
        state.stageUpdatesSearchQuery = '';
        syncStageDraftsForCurrentSector();
        renderStageUpdatesModal();
      }
      return;
    }
    const pcpSelectEl = event.target.closest('[data-pcp-stage-sector-select="true"]');
    if (pcpSelectEl) {
      const selectedSector = normalizeSectorValue(pcpSelectEl.value || '');
      if (STAGE_WORKSPACE_SECTORS.includes(selectedSector)) {
        state.pcpStageSelectedSector = selectedSector;
      }
    }
  });
  stageUpdatesModalEl.addEventListener('input', (event) => {
    const searchEl = event.target.closest('[data-stage-search="true"]');
    if (searchEl) {
      const caretPosition = searchEl.selectionStart ?? String(searchEl.value || '').length;
      state.stageUpdatesSearchQuery = searchEl.value || '';
      renderStageUpdatesModal();
      refocusStageSearchInput(caretPosition);
      return;
    }
    const progressEl = event.target.closest('[data-stage-progress="true"]');
    if (progressEl) {
      const rowEl = progressEl.closest('tr');
      const dateEl = rowEl?.querySelector('[name="completionDate"]');
      if (dateEl && Number(progressEl.value) === 100 && !dateEl.value) {
        dateEl.value = new Date().toISOString().slice(0, 10);
      }
      persistStageDraftFromRow(rowEl);
      renderStageUpdatesModal();
      return;
    }
    const draftField = event.target.closest('[name="completionDate"], [name="note"]');
    if (draftField) {
      const rowEl = draftField.closest('tr');
      persistStageDraftFromRow(rowEl);
    }
  });
}


if (sectorAlertsCloseEl) {
  sectorAlertsCloseEl.addEventListener("click", closeSectorAlertsModal);
}

if (sectorAlertsModalEl) {
  sectorAlertsModalEl.addEventListener("click", (event) => {
    if (event.target.matches("[data-close-sector-alerts='true']")) {
      closeSectorAlertsModal();
      return;
    }
    const button = event.target.closest("[data-ack-alert]");
    if (button) {
      acknowledgeManualAlert(button.dataset.ackAlert);
      return;
    }
    const replyButton = event.target.closest("[data-reply-alert]");
    if (replyButton) {
      openAlertResponseModal(replyButton.dataset.replyAlert);
      return;
    }
    const resolveButton = event.target.closest('[data-resolve-signal]');
    if (resolveButton) {
      resolveSignal(resolveButton.dataset.resolveSignal);
    }
  });
}

if (alertResponseCloseEl) {
  alertResponseCloseEl.addEventListener("click", closeAlertResponseModal);
}

if (alertResponseCancelEl) {
  alertResponseCancelEl.addEventListener("click", closeAlertResponseModal);
}

if (alertResponseModalEl) {
  alertResponseModalEl.addEventListener("click", (event) => {
    if (event.target.matches("[data-close-alert-response='true']")) {
      closeAlertResponseModal();
    }
  });
}

if (alertResponseFormEl) {
  alertResponseFormEl.addEventListener("submit", handleAlertResponseSubmit);
}

if (projectSignalCloseEl) {
  projectSignalCloseEl.addEventListener('click', closeProjectSignalModal);
}

if (projectSignalCancelEl) {
  projectSignalCancelEl.addEventListener('click', closeProjectSignalModal);
}

if (projectSignalModalEl) {
  projectSignalModalEl.addEventListener('click', (event) => {
    if (event.target.matches('[data-close-project-signal="true"]')) {
      closeProjectSignalModal();
    }
  });
}

if (projectSignalFormEl) {
  projectSignalFormEl.addEventListener('submit', handleProjectSignalSubmit);
}

if (openAdminPanelEl) {
  openAdminPanelEl.addEventListener("click", () => {
    if (state.user?.role !== "admin") return;
    openAdminModal();
  });
}

if (adminCloseEl) {
  adminCloseEl.addEventListener("click", closeAdminModal);
}

if (adminModalEl) {
  adminModalEl.addEventListener("click", (event) => {
    if (event.target.matches("[data-close-admin='true']")) {
      closeAdminModal();
    }
  });
}

adminTabTriggerEls.forEach((button) => {
  button.addEventListener('click', () => setAdminActiveTab(button.dataset.adminTabTrigger));
});

if (adminUserFormEl) {
  adminUserFormEl.addEventListener("submit", handleAdminUserSubmit);
}

if (adminUserCancelEditEl) {
  adminUserCancelEditEl.addEventListener("click", () => {
    resetAdminUserForm();
    adminUserFeedbackEl.textContent = "";
  });
}

if (adminSyncButtonEl) {
  adminSyncButtonEl.addEventListener("click", syncAdminDataToGithub);
}

if (adminAlertFormEl) {
  adminAlertFormEl.addEventListener("submit", handleAdminAlertSubmit);
}

if (adminAlertSearchEl) {
  adminAlertSearchEl.addEventListener("input", (event) => {
    state.adminAlertSearchQuery = String(event.target.value || "");
    renderAdminAlertsList();
  });
}

if (adminUsersListEl) {
  adminUsersListEl.addEventListener("click", (event) => {
    const roleButton = event.target.closest("[data-user-role][data-user-id]");
    if (roleButton) {
      updateUserRole(roleButton.dataset.userId, roleButton.dataset.userRole);
      return;
    }
    const editButton = event.target.closest("[data-user-edit]");
    if (editButton) {
      startEditUser(editButton.dataset.userEdit);
    }
  });
}


  modalContentEl.addEventListener("click", (event) => {
    const isoSortButton = event.target.closest("[data-modal-iso-sort]");
    if (isoSortButton) {
      state.modalIsoSortMode = isoSortButton.dataset.modalIsoSort || 'urgency';
      const project = getSelectedProject();
      if (project) renderModal(project);
      return;
    }

    const clientPanelButton = event.target.closest('[data-open-client-panel]');
    if (clientPanelButton) {
      const project = state.projects.find((item) => String(item.rowId) === String(clientPanelButton.dataset.openClientPanel));
      if (project && canManageClientBspPanel(project)) openClientBspExecutiveForPmEdit(project);
      else if (project && canOpenClientBspPanel(project)) openClientBspExecutive(project);
      return;
    }

    const backlogCard = event.target.closest("#modal-open-backlog");
    if (backlogCard) {
      const project = getSelectedProject();
      if (project) {
        state.modalPendingOnly = true;
        renderModal(project);
      }
      return;
    }

    const row = event.target.closest("tr[data-modal-row='true']");
    if (!row) return;
    modalContentEl.querySelectorAll("tr[data-modal-row='true'].modal-row-selected").forEach((item) => {
      if (item !== row) item.classList.remove("modal-row-selected");
    });
    row.classList.toggle("modal-row-selected");
  });


  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    if (attentionPopupEl && !attentionPopupEl.classList.contains('hidden')) {
      return;
    }
    if (loginModalEl && !loginModalEl.classList.contains("hidden")) {
      closeLoginModal();
      return;
    }
    if (adminModalEl && !adminModalEl.classList.contains("hidden")) {
      closeAdminModal();
      return;
    }
    if (alertResponseModalEl && !alertResponseModalEl.classList.contains("hidden")) {
      closeAlertResponseModal();
      return;
    }
    if (sectorAlertsModalEl && !sectorAlertsModalEl.classList.contains("hidden")) {
      closeSectorAlertsModal();
      return;
    }
    if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      closeStageUpdatesModal();
      return;
    }
    if (alertModalEl && !alertModalEl.classList.contains("hidden")) {
      closeAlertModal();
      return;
    }
    closeProjectModal();
  });
}

function closeLoginModal() {
  if (!loginModalEl) return;
  if (!state.user) return;
  document.body.classList.remove('auth-locked');
  loginModalEl.classList.add("hidden");
  loginModalEl.setAttribute("aria-hidden", "true");
  if (
    modalEl.classList.contains("hidden") &&
    alertModalEl.classList.contains("hidden") &&
    sectorAlertsModalEl.classList.contains("hidden") &&
    adminModalEl.classList.contains("hidden")
  ) {
    document.body.classList.remove("modal-open");
  }
}

function openLoginModal(message = "") {
  if (!loginModalEl) return;
  if (!state.user) document.body.classList.add('auth-locked');
  if (loginFeedbackEl) loginFeedbackEl.textContent = message || "Faça login para acessar o painel operacional.";
  loginModalEl.classList.remove("hidden");
  loginModalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  window.setTimeout(() => loginUsernameEl?.focus(), 40);
}


const LOGIN_PROGRESS_STEPS = [
  { at: 8, title: 'Validando acesso...', message: 'Estamos conferindo suas credenciais e preparando sua sessão.', detail: 'Autenticação em andamento.' },
  { at: 22, title: 'Conectando ao Portal STEP...', message: 'Estamos conectando seu usuário ao ambiente correto.', detail: 'Sessão validada.' },
  { at: 38, title: 'Carregando BSPs...', message: 'Estamos carregando as BSPs e organizando os projetos por cliente e unidade.', detail: 'BSPs em processamento.' },
  { at: 55, title: 'Carregando POs...', message: 'Estamos carregando as POs, demandas e referências de fabricação.', detail: 'POs e demandas em processamento.' },
  { at: 72, title: 'Atualizando indicadores...', message: 'Estamos calculando pesos, status, pendências e alertas operacionais.', detail: 'Indicadores sendo preparados.' },
  { at: 88, title: 'Definindo dashboards...', message: 'Estamos definindo os dashboards e montando a visualização final.', detail: 'Dashboard quase pronto.' },
  { at: 100, title: 'Tudo pronto.', message: 'Dados carregados com sucesso. Abrindo o painel operacional.', detail: 'Concluído.' },
];

function getLoginProgressStep(percent) {
  return LOGIN_PROGRESS_STEPS.reduce((selected, step) => (percent >= step.at ? step : selected), LOGIN_PROGRESS_STEPS[0]);
}

function setLoginProgress(percent, options = {}) {
  const nextPercent = Math.max(0, Math.min(100, Math.round(percent || 0)));
  state.loginProgressValue = nextPercent;
  const step = options.step || getLoginProgressStep(nextPercent);
  if (loginProgressTitleEl) loginProgressTitleEl.textContent = options.title || step.title;
  if (loginProgressMessageEl) loginProgressMessageEl.textContent = options.message || step.message;
  if (loginProgressDetailEl) loginProgressDetailEl.textContent = options.detail || step.detail;
  if (loginProgressPercentEl) loginProgressPercentEl.textContent = `${nextPercent}%`;
  if (loginProgressFillEl) loginProgressFillEl.style.width = `${nextPercent}%`;
}

function stopLoginProgressTimer() {
  if (state.loginProgressTimer) {
    window.clearInterval(state.loginProgressTimer);
    state.loginProgressTimer = null;
  }
}

function startLoginProgress(options = {}) {
  stopLoginProgressTimer();
  state.loginProgressActive = true;
  state.loginProgressValue = 0;
  if (loginProgressOverlayEl) loginProgressOverlayEl.classList.remove('hidden');
  document.body.classList.add('modal-open');
  setLoginProgress(options.initialPercent || 6, {
    title: options.title || 'Validando acesso...',
    message: options.message || 'Estamos conferindo suas credenciais e preparando sua sessão.',
    detail: options.detail || 'Autenticação em andamento.',
  });

  const targetBeforeDataReady = Number(options.targetBeforeDataReady || 92);
  state.loginProgressTimer = window.setInterval(() => {
    if (!state.loginProgressActive) return;
    const current = Number(state.loginProgressValue || 0);
    if (current >= targetBeforeDataReady) return;
    const increment = current < 35 ? 4 : current < 70 ? 3 : 1;
    setLoginProgress(Math.min(targetBeforeDataReady, current + increment));
  }, 420);
}

function hideLoginProgressOverlay() {
  stopLoginProgressTimer();
  state.loginProgressActive = false;
  if (loginProgressOverlayEl) loginProgressOverlayEl.classList.add('hidden');
  if (
    (!loginModalEl || loginModalEl.classList.contains('hidden')) &&
    (!modalEl || modalEl.classList.contains('hidden')) &&
    (!alertModalEl || alertModalEl.classList.contains('hidden')) &&
    (!sectorAlertsModalEl || sectorAlertsModalEl.classList.contains('hidden')) &&
    (!adminModalEl || adminModalEl.classList.contains('hidden'))
  ) {
    document.body.classList.remove('modal-open');
  }
}

function failLoginProgress(message) {
  stopLoginProgressTimer();
  state.loginProgressActive = false;
  setLoginProgress(Math.max(1, state.loginProgressValue || 1), {
    title: 'Não foi possível carregar o painel',
    message: message || 'Ocorreu uma falha durante o carregamento. Tente novamente.',
    detail: 'Falha no carregamento.',
  });
  window.setTimeout(hideLoginProgressOverlay, 1200);
}

function scheduleClientPoBackgroundRefresh() {
  if (!CLIENT_PORTAL_RELEASE_WITH_PO_PENDING || !isClientUser()) return;
  if (state.clientPoBackgroundRefreshScheduled) return;
  state.clientPoBackgroundRefreshScheduled = true;
  window.setTimeout(() => {
    revalidateProjectsInBackground(true)
      .catch(() => {})
      .finally(() => {
        state.clientPoBackgroundRefreshScheduled = false;
      });
  }, CLIENT_PORTAL_PO_BACKGROUND_REFRESH_DELAY_MS);
}

function hasDashboardDataReady() {
  const projects = Array.isArray(state.projects) ? state.projects : [];
  if (!state.user) return false;

  if (isClientUser()) {
    // v37.17: cliente com escopo explícito pode legitimamente ter carteira vazia
    // após filtro exato, como BW ENERGY separado de BW LNG. Isso não pode prender o login.
    if (projects.length === 0 && state.meta?.emptyClientScopeResolved) return true;
    const bspsText = String(document.getElementById('client-stat-bsps')?.textContent || '').trim();
    const vesselGrid = document.getElementById('client-vessel-grid');
    const hasVisibleClientCards = Boolean(vesselGrid && vesselGrid.querySelector('[data-client-vessel]'));
    const baseReady = bspsText !== '' && bspsText !== '--' && bspsText !== '0' && hasVisibleClientCards;
    if (!baseReady) return false;

    const poReady = isClientProjectsPayloadPoReady({ projects, meta: state.meta || {} });
    if (!poReady) {
      scheduleClientPoBackgroundRefresh();
      if (lastSyncEl) {
        lastSyncEl.textContent = state.meta?.lastSync
          ? `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")} • POs atualizando em segundo plano`
          : 'BSPs carregadas • POs atualizando em segundo plano';
      }
    }
    return poReady || CLIENT_PORTAL_RELEASE_WITH_PO_PENDING;
  }

  const countText = String(searchCountEl?.textContent || '').trim();
  const hasTableRows = Boolean(bodyEl && bodyEl.querySelector('tr[data-project-id]'));
  return hasTableRows || (Array.isArray(state.filteredProjects) && state.filteredProjects.length > 0 && countText !== '0 resultado(s)');
}

function waitForNextRenderFrame() {
  return new Promise((resolve) => window.requestAnimationFrame(() => window.requestAnimationFrame(resolve)));
}

async function ensureDashboardDataReadyBeforeRelease(options = {}) {
  const maxAttempts = Number(options.maxAttempts || 4);
  const retryDelayMs = Number(options.retryDelayMs || 900);
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    await waitForNextRenderFrame();
    if (hasDashboardDataReady()) return true;

    const clientHasNoProjects = isClientUser() && (!Array.isArray(state.projects) || state.projects.length === 0);
    const poStillPending = isClientUser() && !isClientProjectsPayloadPoReady({ projects: state.projects || [], meta: state.meta || {} });
    const clientPoPending = poStillPending && !CLIENT_PORTAL_RELEASE_WITH_PO_PENDING && !clientHasNoProjects;
    if (poStillPending && CLIENT_PORTAL_RELEASE_WITH_PO_PENDING && !clientHasNoProjects) scheduleClientPoBackgroundRefresh();
    setLoginProgress(Math.min(96, 88 + attempt * 2), {
      title: clientHasNoProjects ? 'Carregando BSPs do cliente...' : (clientPoPending ? 'Carregando POs...' : 'Conferindo dados na tela...'),
      message: clientHasNoProjects
        ? 'Estamos forçando uma leitura limpa da carteira do cliente para evitar abrir o Portal vazio.'
        : (clientPoPending
          ? 'As BSPs já foram localizadas. Agora estamos buscando a base completa de POs sem atualizar a página.'
          : 'Estamos garantindo que as BSPs e dashboards já apareceram no painel. As POs podem continuar atualizando em segundo plano.'),
      detail: `Validação dos dados ${attempt}/${maxAttempts}.`,
    });

    try {
      await loadProjects({
        // v37.17: validação pós-login não chama Smartsheet.
        // Ela relê apenas o cache Supabase; atualização pesada é manual/agendada.
        force: false,
        skipLocalCache: true,
        suppressLoadingState: true,
        preferServerCache: true,
        requireData: false,
        requireClientPo: false,
      });
    } catch (error) {
      lastError = error;
      console.warn('[Login] Revalidação de dados ainda pendente:', error?.message || error);
    }

    await waitForNextRenderFrame();
    if (hasDashboardDataReady()) return true;

    await new Promise((resolve) => window.setTimeout(resolve, retryDelayMs));
  }

  throw new Error(lastError?.message || 'Os dados ainda não apareceram na tela. O painel não será liberado vazio; tente novamente em alguns segundos.');
}

async function completeLoginProgress() {
  stopLoginProgressTimer();
  await ensureDashboardDataReadyBeforeRelease({ maxAttempts: 1, retryDelayMs: 300 });
  setLoginProgress(100, LOGIN_PROGRESS_STEPS[LOGIN_PROGRESS_STEPS.length - 1]);
  await new Promise((resolve) => window.setTimeout(resolve, 520));
  hideLoginProgressOverlay();
}

function setupLoginPasswordToggle() {
  if (!toggleLoginPasswordEl || !loginPasswordEl) return;
  const sync = () => {
    const visible = loginPasswordEl.type === "text";
    toggleLoginPasswordEl.textContent = visible ? "Ocultar" : "Mostrar";
    toggleLoginPasswordEl.setAttribute("aria-label", visible ? "Ocultar senha" : "Mostrar senha");
  };
  toggleLoginPasswordEl.addEventListener("click", () => {
    loginPasswordEl.type = loginPasswordEl.type === "password" ? "text" : "password";
    sync();
  });
  sync();
}

function setupAdminPasswordToggle() {
  const passwordEl = document.getElementById("admin-user-password");
  if (!adminUserTogglePasswordEl || !passwordEl) return;
  const sync = () => {
    const visible = passwordEl.type === "text";
    adminUserTogglePasswordEl.textContent = visible ? "Ocultar" : "Mostrar";
    adminUserTogglePasswordEl.setAttribute("aria-label", visible ? "Ocultar senha do usuário" : "Mostrar senha do usuário");
  };
  adminUserTogglePasswordEl.addEventListener("click", () => {
    passwordEl.type = passwordEl.type === "password" ? "text" : "password";
    sync();
  });
  sync();
}


function closeChangePasswordModal() {
  if (!changePasswordModalEl) return;
  changePasswordModalEl.classList.add("hidden");
  changePasswordModalEl.setAttribute("aria-hidden", "true");
  if (changePasswordFormEl) changePasswordFormEl.reset();
  if (changePasswordFeedbackEl) changePasswordFeedbackEl.textContent = "";
  if (
    modalEl.classList.contains("hidden") &&
    alertModalEl.classList.contains("hidden") &&
    sectorAlertsModalEl.classList.contains("hidden") &&
    stageUpdatesModalEl.classList.contains('hidden') &&
    loginModalEl.classList.contains("hidden") &&
    adminModalEl.classList.contains("hidden")
  ) {
    document.body.classList.remove("modal-open");
  }
}

function openChangePasswordModal() {
  if (!state.user || !changePasswordModalEl) return;
  changePasswordModalEl.classList.remove("hidden");
  changePasswordModalEl.setAttribute("aria-hidden", "false");
  if (changePasswordFormEl) changePasswordFormEl.reset();
  if (changePasswordFeedbackEl) changePasswordFeedbackEl.textContent = "";
  document.body.classList.add("modal-open");
  window.setTimeout(() => {
    if (changePasswordCurrentEl) changePasswordCurrentEl.focus();
  }, 50);
}

async function handleChangePasswordSubmit(event) {
  event.preventDefault();
  if (!state.user || !changePasswordFeedbackEl) return;
  const currentPassword = String(changePasswordCurrentEl?.value || '').trim();
  const newPassword = String(changePasswordNewEl?.value || '').trim();
  const confirmPassword = String(changePasswordConfirmEl?.value || '').trim();

  if (!currentPassword || !newPassword || !confirmPassword) {
    changePasswordFeedbackEl.textContent = 'Preencha todos os campos.';
    return;
  }
  if (newPassword.length < 6) {
    changePasswordFeedbackEl.textContent = 'A nova senha deve ter pelo menos 6 caracteres.';
    return;
  }
  if (newPassword !== confirmPassword) {
    changePasswordFeedbackEl.textContent = 'A confirmação da nova senha não confere.';
    return;
  }
  if (currentPassword === newPassword) {
    changePasswordFeedbackEl.textContent = 'A nova senha precisa ser diferente da atual.';
    return;
  }

  try {
    changePasswordFeedbackEl.textContent = 'Alterando senha...';
    const response = await fetch('/api/change-password', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ currentPassword, newPassword }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || 'Não foi possível alterar a senha.');
    }
    changePasswordFeedbackEl.textContent = 'Senha alterada com sucesso.';
    window.setTimeout(() => {
      closeChangePasswordModal();
    }, 700);
  } catch (error) {
    changePasswordFeedbackEl.textContent = error.message || 'Não foi possível alterar a senha.';
  }
}

function updateSessionUi() {
  const user = state.user;
  if (!user) {
    document.body.classList.add('auth-locked');
    state.projectView = 'all';
    state.sectorScopedView = false;
    state.alertSectorFilter = 'all';
    sessionUserNameEl.textContent = "Acesso bloqueado";
    sessionUserMetaEl.textContent = "Faça login para visualizar os projetos, indicadores e detalhes do painel.";
    updatePrimaryUserActionUi();
    renderProjectViewTabs();
    sessionStatusEl.textContent = "bloqueado";
    logoutButtonEl.classList.add("hidden");
    if (openChangePasswordButtonEl) openChangePasswordButtonEl.classList.add("hidden");
    if (openClientApiButtonEl) openClientApiButtonEl.classList.add("hidden");
    if (openIsoQrButtonEl) openIsoQrButtonEl.classList.add("hidden");
    openAdminPanelEl.classList.add("hidden");
    if (openLoginButtonEl) openLoginButtonEl.classList.remove("hidden");
    setClientDashboardMode();
    return;
  }

  document.body.classList.remove('auth-locked');

  if (shouldUseSectorScopedToggle(user)) {
    state.projectView = 'all';
    state.sectorScopedView = loadSectorScopedViewPreference(user);
    state.alertSectorFilter = state.sectorScopedView ? normalizeAlertSectorFilterValue(getPrimaryUserSector(user)) || 'all' : 'all';
  }

  sessionUserNameEl.textContent = user.name || user.username;
  const linkedSectors = getUserAlertSectors(user);
  sessionUserMetaEl.textContent = isClientUser(user)
    ? `Cliente • ${getClientPortalName(user)}`
    : `${user.role === "admin" ? "Administrador" : "Setor"} • ${sectorLabel(user.sector)}${user.role !== "admin" && linkedSectors.length > 1 ? ` • Alertas: ${formatSectorList(linkedSectors)}` : ""}`;
  setClientDashboardMode();
  updatePrimaryUserActionUi();
  sessionStatusEl.textContent = "online";
  logoutButtonEl.classList.remove("hidden");
  if (openChangePasswordButtonEl) openChangePasswordButtonEl.classList.remove("hidden");
  if (openLoginButtonEl) openLoginButtonEl.classList.add("hidden");
  if (openClientApiButtonEl) {
    openClientApiButtonEl.classList.toggle("hidden", !isClientUser(user));
  }
  if (openIsoQrButtonEl) {
    openIsoQrButtonEl.classList.toggle("hidden", isClientUser(user));
  }
  if (user.role === "admin") {
    openAdminPanelEl.classList.remove("hidden");
  } else {
    openAdminPanelEl.classList.add("hidden");
  }

  if (githubSyncBadgeEl) {
    githubSyncBadgeEl.textContent = `GitHub sync: ${state.githubSyncEnabled ? "online" : "local"}`;
  }
}

function resetDashboardForLoggedOutState() {
  state.projects = [];
  state.filteredProjects = [];
  state.stats = null;
  state.meta = null;
  state.alerts = [];
  state.selectedProjectId = null;
  resetClientBspOverridesState();
  if (bodyEl) bodyEl.innerHTML = `<tr><td colspan="21" class="loading-cell">Faça login para visualizar os projetos.</td></tr>`;
  if (detailCardEl) detailCardEl.innerHTML = `<div class="detail-placeholder">Painel protegido. Entre com seu usuário e senha para visualizar as informações.</div>`;
  if (searchCountEl) searchCountEl.textContent = '0 resultado(s)';
  updateExportFilteredProjectsButton();
  if (sheetNameEl) sheetNameEl.textContent = 'Acesso restrito';
  if (lastSyncEl) lastSyncEl.textContent = 'Faça login para carregar os dados.';
  if (alertBadgeCountEl) alertBadgeCountEl.textContent = '0';
  renderProjectViewTabs();
  renderStats();
}

async function bootstrapSession() {
  try {
    const response = await fetch("/api/auth-me", { credentials: "same-origin", cache: "no-store" });
    const data = await response.json().catch(() => null);
    if (!data?.authenticated) {
      state.user = null;
      updateSessionUi();
      resetDashboardForLoggedOutState();
      openLoginModal("Faça login para visualizar o painel.");
      return false;
    }
    state.user = data.user;
    state.githubSyncEnabled = Boolean(data.githubSyncEnabled);
    updateSessionUi();
    closeLoginModal();
    startPresenceHeartbeat();
    syncPushSubscription(false).catch(() => {});
    return true;
  } catch {
    state.user = null;
    updateSessionUi();
    resetDashboardForLoggedOutState();
    openLoginModal("Faça login para visualizar o painel.");
    return false;
  }
}

function getUserAutomaticAlerts() {
  if (!state.user) return [];
  if (state.user.role === "admin") {
    return Array.isArray(state.alerts) ? [...state.alerts] : [];
  }

  const allowedSectors = new Set(getUserAlertSectors(state.user));
  return (Array.isArray(state.alerts) ? state.alerts : [])
    .filter((alert) => allowedSectors.has(normalizeSectorValue(alert?.sector)))
    .filter((alert) => {
      if (!userHasProjectsScope(state.user)) return true;
      const relatedProject = state.projects.find((project) => {
        const alertNumber = normalizeText(alert?.projectNumber || alert?.projectDisplay || '');
        const projectNumber = normalizeText(project?.projectNumber || project?.projectDisplay || '');
        return alertNumber && projectNumber && alertNumber === projectNumber;
      });
      return relatedProject ? projectBelongsToUser(relatedProject, state.user) : false;
    })
    .sort((a, b) => {
      if ((a?.daysRemaining ?? 0) !== (b?.daysRemaining ?? 0)) {
        return (a?.daysRemaining ?? 0) - (b?.daysRemaining ?? 0);
      }
      return String(a?.projectDisplay || "").localeCompare(String(b?.projectDisplay || ""), "pt-BR");
    });
}

function renderManualAlerts(targetAlerts = state.manualAlerts, targetEl = sectorAlertsContentEl) {
  if (!targetEl) return;
  if (!state.user) {
    targetEl.innerHTML = '<div class="detail-placeholder">Faça login para visualizar alertas direcionados ao seu setor.</div>';
    return;
  }

  const manualAlerts = Array.isArray(targetAlerts) ? targetAlerts : [];
  const automaticAlerts = getUserAutomaticAlerts();

  if (!manualAlerts.length && !automaticAlerts.length) {
    targetEl.innerHTML = '<div class="detail-placeholder">Nenhum alerta específico para este login no momento.</div>';
    return;
  }

  const manualHtml = manualAlerts.length
    ? `
      <section class="manual-alert-section">
        <div class="admin-list-item-meta">
          <span class="manual-alert-tag">Alerta Operacional</span>
          <span>${manualAlerts.length} alerta(s)</span>
        </div>
        <div class="manual-alert-section-list">
          ${manualAlerts.map((alert) => {
            const resolved = getSignalResolutionInfo(alert.id);
            return `
            <article class="manual-alert-item manual-alert-item--operational ${resolved ? 'manual-alert-item--resolved' : ''}">
              <div class="admin-list-item-meta">
                ${getSignalStatusBadge(alert)}
                <span class="manual-alert-tag">${escapeHtml(sectorLabel(alert.sector))}</span>
                <span>${escapeHtml(new Date(alert.createdAt).toLocaleString("pt-BR"))}</span>
                <span>Aberta por: ${escapeHtml(alert.createdBy || 'Usuário')}</span>
              </div>
              <strong>${escapeHtml(alert.title || "Sinalização")}</strong>
              <p>${escapeHtml(alert.message || "").replace(/\n/g, '<br>')}</p>
              <div class="manual-alert-actions">
                ${resolved
                  ? `<span class="manual-alert-tag manual-alert-tag--resolved-by">Resolvida por: ${escapeHtml(resolved.username)}</span>${resolved.date ? `<span class="manual-alert-tag">${escapeHtml(new Date(resolved.date).toLocaleString('pt-BR'))}</span>` : ''}`
                  : `${canResolveSignal() ? `<button class="primary-button" type="button" data-resolve-signal="${escapeHtml(alert.id)}">Marcar como resolvida</button>` : ''}`}
              </div>
              ${resolved && resolved.note ? `<div class="response-thread"><div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(resolved.note)}</p></div></div>` : ''}
            </article>
          `;}).join("")}
        </div>
      </section>
    `
    : `
      <section class="manual-alert-section">
        <div class="admin-list-item-meta">
          <span class="manual-alert-tag">Alerta Operacional</span>
          <span>Nenhum alerta operacional para o seu setor.</span>
        </div>
      </section>
    `;

  const automaticHtml = automaticAlerts.length
    ? `
      <section class="manual-alert-section">
        <div class="admin-list-item-meta">
          <span class="manual-alert-tag manual-alert-tag--high">Automáticos</span>
          <span>${automaticAlerts.length} alerta(s) de prazo${userHasProjectsScope(state.user) ? ' dos seus projetos' : ` para ${escapeHtml(formatSectorList(getUserAlertSectors(state.user)))}`}</span>
        </div>
        <div class="manual-alert-section-list">
          ${automaticAlerts.map((alert) => {
            const severity = getAlertSeverity(alert);
            const severityLabel = severity === "urgent" ? "Urgente" : "Médio";
            const dateLabel = alert.daysRemaining < 0
              ? `${Math.abs(alert.daysRemaining)} dia(s) em atraso`
              : `${alert.daysRemaining} dia(s) para o prazo`;
            return `
              <article class="manual-alert-item manual-alert-item--automatic">
                <div class="admin-list-item-meta">
                  <span class="manual-alert-tag manual-alert-tag--${severity === "urgent" ? "urgent" : "high"}">${severityLabel}</span>
                  <span class="manual-alert-tag">${escapeHtml(sectorLabel(alert.sector))}</span>
                  <span>${escapeHtml(alert.plannedFinishDate || "Sem data")}</span>
                  <span>${escapeHtml(dateLabel)}</span>
                </div>
                <strong>${escapeHtml(alert.title || "Alerta automático")}</strong>
                <p>${escapeHtml(alert.message || "")}</p>
                <div class="manual-alert-actions">
                  <span class="manual-alert-tag">${escapeHtml(alert.projectDisplay || alert.projectNumber || "Projeto")}</span>
                  <span class="manual-alert-tag">Cliente: ${escapeHtml(alert.client || "—")}</span>
                </div>
              </article>
            `;
          }).join("")}
        </div>
      </section>
    `
    : `
      <section class="manual-alert-section">
        <div class="admin-list-item-meta">
          <span class="manual-alert-tag manual-alert-tag--high">Automáticos</span>
          <span>Nenhum alerta automático de prazo para o seu setor.</span>
        </div>
      </section>
    `;

  targetEl.innerHTML = `
    <div class="manual-alert-summary">
      <span class="manual-alert-tag">Setor principal: ${escapeHtml(sectorLabel(state.user.sector))}</span>
      <span class="manual-alert-tag">Recebe alertas de: ${escapeHtml(formatSectorList(getUserAlertSectors(state.user)))}</span>
      <span class="manual-alert-tag">Total: ${manualAlerts.length + automaticAlerts.length} alerta(s)</span>
    </div>
    ${manualHtml}
    ${automaticHtml}
  `;
}

function renderMyProjectSignals(targetEl = sectorAlertsContentEl) {
  if (!targetEl) return;
  if (!state.user) {
    targetEl.innerHTML = '<div class="detail-placeholder">Faça login para visualizar as sinalizações que você enviou ao PCP.</div>';
    return;
  }
  const signals = getMyProjectSignals();
  if (!signals.length) {
    targetEl.innerHTML = '<div class="detail-placeholder">Você ainda não enviou nenhuma sinalização ao PCP.</div>';
    return;
  }
  const pendingCount = signals.filter((alert) => !getSignalResolutionInfo(alert.id)).length;
  const resolvedCount = signals.length - pendingCount;
  targetEl.innerHTML = `
    <div class="manual-alert-summary">
      <span class="manual-alert-tag">Minhas sinalizações</span>
      <span class="manual-alert-tag">Enviadas ao PCP</span>
      <span class="manual-alert-tag">Pendentes: ${pendingCount}</span>
      <span class="manual-alert-tag">Resolvidas: ${resolvedCount}</span>
    </div>
    <section class="manual-alert-section">
      <div class="admin-list-item-meta">
        <span class="manual-alert-tag">Acompanhamento do usuário</span>
        <span>${signals.length} registro(s)</span>
      </div>
      <div class="manual-alert-section-list">
        ${signals.map((alert) => {
          const resolved = getSignalResolutionInfo(alert.id);
          return `
            <article class="manual-alert-item manual-alert-item--operational ${resolved ? 'manual-alert-item--resolved' : ''}">
              <div class="admin-list-item-meta">
                ${getSignalStatusBadge(alert)}
                <span class="manual-alert-tag">PCP</span>
                <span>${escapeHtml(new Date(alert.createdAt).toLocaleString('pt-BR'))}</span>
              </div>
              <strong>${escapeHtml(alert.title || 'Sinalização')}</strong>
              <p>${escapeHtml(alert.message || '').replace(/\n/g, '<br>')}</p>
              <div class="manual-alert-actions">
                <span class="manual-alert-tag">Aberta por: ${escapeHtml(alert.createdBy || 'Usuário')}</span>
                ${resolved
                  ? `<span class="manual-alert-tag manual-alert-tag--resolved-by">Resolvida por: ${escapeHtml(resolved.username)}</span>${resolved.date ? `<span class="manual-alert-tag">${escapeHtml(new Date(resolved.date).toLocaleString('pt-BR'))}</span>` : ''}`
                  : `<span class="manual-alert-tag manual-alert-tag--pending">Aguardando PCP</span>`}
              </div>
              ${resolved && resolved.note ? `<div class="response-thread"><div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(resolved.note)}</p></div></div>` : ''}
            </article>
          `;
        }).join('')}
      </div>
    </section>
  `;
}

function renderProjectUserSignals(targetEl = sectorAlertsContentEl) {
  if (!targetEl) return;
  if (!state.user) {
    targetEl.innerHTML = '<div class="detail-placeholder">Faça login para visualizar as sinalizações enviadas por usuários de Projetos.</div>';
    return;
  }
  const signals = getProjectUserSignals();
  if (!signals.length) {
    targetEl.innerHTML = '<div class="detail-placeholder">Nenhuma sinalização enviada por usuários de Projetos foi encontrada.</div>';
    return;
  }
  targetEl.innerHTML = `
    <div class="manual-alert-summary">
      <span class="manual-alert-tag">Fila do PCP</span>
      <span class="manual-alert-tag">Origem: Projetos</span>
      <span class="manual-alert-tag">Total: ${signals.length} sinalização(ões)</span>
    </div>
    <section class="manual-alert-section">
      <div class="admin-list-item-meta">
        <span class="manual-alert-tag">Alertas enviados por Projetos</span>
        <span>${signals.length} registro(s)</span>
      </div>
      <div class="manual-alert-section-list">
        ${signals.map((alert) => {
          const resolved = getSignalResolutionInfo(alert.id);
          return `
            <article class="manual-alert-item manual-alert-item--operational ${resolved ? 'manual-alert-item--resolved' : ''}">
              <div class="admin-list-item-meta">
                ${getSignalStatusBadge(alert)}
                <span class="manual-alert-tag">PCP</span>
                <span>${escapeHtml(new Date(alert.createdAt).toLocaleString('pt-BR'))}</span>
                <span>Aberta por: ${escapeHtml(alert.createdBy || 'Usuário')}</span>
              </div>
              <strong>${escapeHtml(alert.title || 'Sinalização')}</strong>
              <div class="manual-alert-actions">
                ${resolved
                  ? `<span class="manual-alert-tag manual-alert-tag--resolved-by">Resolvida por: ${escapeHtml(resolved.username)}</span>${resolved.date ? `<span class="manual-alert-tag">${escapeHtml(new Date(resolved.date).toLocaleString('pt-BR'))}</span>` : ''}`
                  : `${canResolveSignal() ? `<button class="primary-button" type="button" data-resolve-signal="${escapeHtml(alert.id)}">Marcar como resolvida</button>` : ''}`}
              </div>
              ${resolved && resolved.note ? `<div class="response-thread"><div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(resolved.note)}</p></div></div>` : ''}
            </article>
          `;
        }).join('')}
      </div>
    </section>
  `;
}

async function loadManualAlerts(options = {}) {
  if (!state.user) return;
  if (options.background && shouldSkipBackgroundRequest(options)) return;
  const now = Date.now();
  if (!options.force && options.background && now - state.lastManualAlertsFetchAt < ALERTS_REFRESH_MS) return;
  try {
    const response = await fetch(`/api/sector-alerts?t=${Date.now()}`, { credentials: "same-origin", cache: "no-store" });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao carregar alertas operacionais.");
    }
    state.lastManualAlertsFetchAt = Date.now();
    state.githubSyncEnabled = Boolean(data.githubSyncEnabled ?? state.githubSyncEnabled);
    state.manualAlerts = data.alerts || [];
    state.projectSignals = data.projectSignals || [];
    updateSessionUi();
    renderManualAlerts();
    detectNewUserAlerts();
    if (state.user?.role === "admin") {
      renderAdminAlertsList();
      renderAdminAlertResponses();
    }
  } catch (error) {
    state.manualAlerts = [];
    state.projectSignals = [];
    if (sectorAlertsContentEl) {
      sectorAlertsContentEl.innerHTML = `<div class="detail-placeholder">${escapeHtml(error?.message || "Falha ao carregar alertas operacionais.")}</div>`;
    } else {
      renderManualAlerts([], sectorAlertsContentEl);
    }
    console.warn(error);
  }
}

function openSectorAlertsModal() {
  if (!sectorAlertsModalEl) return;
  if (state.sectorAlertsMode === 'project-signals') {
    renderProjectUserSignals();
  } else if (state.sectorAlertsMode === 'my-project-signals') {
    renderMyProjectSignals();
  } else {
    renderManualAlerts();
  }
  sectorAlertsModalEl.classList.remove("hidden");
  sectorAlertsModalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function closeSectorAlertsModal() {
  if (!sectorAlertsModalEl) return;
  state.sectorAlertsMode = 'default';
  sectorAlertsModalEl.classList.add("hidden");
  sectorAlertsModalEl.setAttribute("aria-hidden", "true");
  if (
    modalEl.classList.contains("hidden") &&
    alertModalEl.classList.contains("hidden") &&
    adminModalEl.classList.contains("hidden") &&
    loginModalEl.classList.contains("hidden")
  ) {
    document.body.classList.remove("modal-open");
  }
}



function getAlertResponsesForAlert(alertId) {
  return (Array.isArray(state.alertResponses) ? state.alertResponses : []).filter((item) => String(item.alertId) === String(alertId));
}

function getAdminReplyStatusLabel(status) {
  const value = String(status || '').toLowerCase();
  if (value === 'respondido') return 'Respondido pelo admin';
  if (value === 'lido') return 'Lido';
  if (value === 'resolvida') return 'Resolvida';
  return 'Aguardando retorno';
}

function renderAdminResponsesThread(alertId) {
  const responses = getAlertResponsesForAlert(alertId);
  if (!responses.length) {
    return `
      <div class="admin-alert-ack-box">
        <strong>Respostas do setor</strong>
        <div class="admin-list-item-meta">
          <span>Nenhuma resposta recebida ainda.</span>
        </div>
      </div>
    `;
  }
  return `
    <div class="admin-alert-ack-box">
      <strong>Respostas do setor</strong>
      <div class="admin-list-item-meta">
        <span>${responses.length} resposta(s)</span>
        <span>Última: ${escapeHtml(responses[0]?.createdAt ? new Date(responses[0].createdAt).toLocaleString('pt-BR') : 'Sem data')}</span>
      </div>
      <div class="admin-alert-ack-list">
        ${responses.map((item) => `
          <div class="admin-alert-ack-item admin-alert-response-item">
            <span><strong>${escapeHtml(item.username || item.userEmail || 'Usuário')}</strong></span>
            <span>${escapeHtml(item.createdAt ? new Date(item.createdAt).toLocaleString('pt-BR') : 'Sem data')}</span>
            <span>Status: ${escapeHtml(getAdminReplyStatusLabel(item.status))}</span>
            <div class="response-bubble"><p>${escapeHtml(item.responseText || '')}</p></div>
            ${item.adminReply ? `<div class="response-bubble response-bubble--admin"><strong>Admin</strong><p>${escapeHtml(item.adminReply)}</p></div>` : ''}
          </div>
        `).join('')}
      </div>
    </div>
  `;
}

function openAlertResponseModal(alertId) {
  const alert = (Array.isArray(state.manualAlerts) ? state.manualAlerts : []).find((item) => String(item.id) === String(alertId));
  if (!alert || !alertResponseModalEl) return;
  state.selectedAlertForResponse = alert;
  if (alertResponseAlertIdEl) alertResponseAlertIdEl.value = alert.id || '';
  if (alertResponseTitleEl) alertResponseTitleEl.textContent = `Responder: ${alert.title || 'Alerta operacional'}`;
  if (alertResponseSubtitleEl) alertResponseSubtitleEl.textContent = `Sua resposta será enviada ao admin para o alerta do setor ${sectorLabel(alert.sector)}.`;
  if (alertResponseTextEl) alertResponseTextEl.value = '';
  if (alertResponseFeedbackEl) alertResponseFeedbackEl.textContent = '';
  alertResponseModalEl.classList.remove('hidden');
  alertResponseModalEl.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  window.setTimeout(() => alertResponseTextEl?.focus(), 40);
}

function closeAlertResponseModal() {
  if (!alertResponseModalEl) return;
  alertResponseModalEl.classList.add('hidden');
  alertResponseModalEl.setAttribute('aria-hidden', 'true');
  state.selectedAlertForResponse = null;
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

function openProjectSignalModal(project) {
  if (!projectSignalModalEl || !project) return;
  if (!canCreateProjectSignal(project)) {
    window.alert('Você só pode enviar sinalização para BSPs que estejam vinculadas ao seu nome.');
    return;
  }
  state.selectedProjectForSignal = project;
  if (projectSignalProjectIdEl) projectSignalProjectIdEl.value = String(project.rowId || '');
  if (projectSignalHeadingEl) projectSignalHeadingEl.textContent = `Nova sinalização • ${project.projectDisplay || project.projectNumber || 'Projeto'}`;
  if (projectSignalSubtitleEl) projectSignalSubtitleEl.textContent = 'A informação será enviada ao PCP para análise e fechamento.';
  if (projectSignalTitleEl) projectSignalTitleEl.value = '';
  if (projectSignalDescriptionEl) projectSignalDescriptionEl.value = '';
  if (projectSignalFeedbackEl) projectSignalFeedbackEl.textContent = '';
  projectSignalModalEl.classList.remove('hidden');
  projectSignalModalEl.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  window.setTimeout(() => projectSignalTitleEl?.focus(), 40);
}

function closeProjectSignalModal() {
  if (!projectSignalModalEl) return;
  projectSignalModalEl.classList.add('hidden');
  projectSignalModalEl.setAttribute('aria-hidden', 'true');
  state.selectedProjectForSignal = null;
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

async function handleProjectSignalSubmit(event) {
  event.preventDefault();
  if (!projectSignalFeedbackEl) return;
  const projectId = String(projectSignalProjectIdEl?.value || '').trim();
  const project = state.projects.find((item) => String(item.rowId) === projectId);
  const title = String(projectSignalTitleEl?.value || '').trim();
  const description = String(projectSignalDescriptionEl?.value || '').trim();
  if (!project || !title || !description) {
    projectSignalFeedbackEl.textContent = 'Preencha título e descrição da sinalização.';
    return;
  }
  if (!canCreateProjectSignal(project)) {
    projectSignalFeedbackEl.textContent = 'Você só pode enviar sinalização para BSPs que estejam vinculadas ao seu nome.';
    return;
  }
  projectSignalFeedbackEl.textContent = 'Enviando sinalização ao PCP...';
  const projectRef = project.projectNumber || project.projectDisplay || `Projeto ${project.rowId}`;
  const payload = {
    sector: 'pcp',
    projectRowId: project.rowId,
    title: `${projectRef} • ${title}`,
    message: `Projeto: ${projectDisplayWithClient(project)}
Informado por: ${state.user?.name || state.user?.username || 'Usuário'}

${description}`,
    priority: 'normal',
    requiresAck: false,
  };
  try {
    const response = await fetch('/api/sector-alerts', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao criar sinalização.');
    projectSignalFeedbackEl.textContent = 'Sinalização enviada ao PCP.';
    await loadManualAlerts();
    await loadAlertResponses();
    if (state.selectedProjectId && String(state.selectedProjectId) === projectId) {
      renderModal(project);
    }
    window.setTimeout(closeProjectSignalModal, 500);
  } catch (error) {
    projectSignalFeedbackEl.textContent = error.message || 'Falha ao criar sinalização.';
  }
}

async function resolveSignal(alertId) {
  if (!alertId) return;
  const note = window.prompt('Adicionar observação de fechamento? (opcional)', '');
  try {
    const response = await fetch('/api/alert-responses', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ alertId, responseText: String(note || '').trim(), status: 'resolvida' }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao marcar sinalização como resolvida.');
    await loadAlertResponses();
    await loadManualAlerts();
    const currentProject = state.projects.find((item) => item.rowId === state.selectedProjectId);
    if (currentProject && !modalEl.classList.contains('hidden')) renderModal(currentProject);
  } catch (error) {
    window.alert(error.message || 'Falha ao marcar sinalização como resolvida.');
  }
}

async function handleAlertResponseSubmit(event) {
  event.preventDefault();
  if (!alertResponseFeedbackEl) return;
  const alertId = String(alertResponseAlertIdEl?.value || '').trim();
  const responseText = String(alertResponseTextEl?.value || '').trim();
  if (!alertId || !responseText) {
    alertResponseFeedbackEl.textContent = 'Digite a resposta antes de enviar.';
    return;
  }
  alertResponseFeedbackEl.textContent = 'Enviando resposta...';
  try {
    const response = await fetch('/api/alert-responses', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ alertId, responseText }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao enviar resposta.');
    alertResponseFeedbackEl.textContent = 'Resposta enviada ao admin.';
    await loadAlertResponses();
    await loadManualAlerts();
    window.setTimeout(closeAlertResponseModal, 500);
  } catch (error) {
    alertResponseFeedbackEl.textContent = error.message || 'Falha ao enviar resposta.';
  }
}

async function loadAlertResponses(options = {}) {
  if (!state.user) {
    state.alertResponses = [];
    return;
  }
  if (options.background && shouldSkipBackgroundRequest(options)) return;
  const now = Date.now();
  if (!options.force && options.background && now - state.lastAlertResponsesFetchAt < ALERTS_REFRESH_MS) return;
  try {
    const response = await fetch('/api/alert-responses', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar respostas das sinalizações.');
    state.lastAlertResponsesFetchAt = Date.now();
    state.alertResponses = Array.isArray(data.responses) ? data.responses : [];
    if (state.user?.role === 'admin') {
      renderAdminAlertResponses();
      renderAdminAlertsList();
    }
  } catch (error) {
    state.alertResponses = [];
    if (state.user?.role === 'admin' && adminAlertResponsesListEl) {
      adminAlertResponsesListEl.innerHTML = `<div class="empty-state">${escapeHtml(error.message || 'Falha ao carregar respostas das sinalizações.')}</div>`;
    }
  }
}

function renderAdminAlertResponses() {
  if (!adminAlertResponsesListEl) return;
  const responses = Array.isArray(state.alertResponses) ? state.alertResponses : [];
  if (!responses.length) {
    adminAlertResponsesListEl.innerHTML = '<div class="empty-state">Nenhuma resposta recebida ainda.</div>';
    return;
  }
  adminAlertResponsesListEl.innerHTML = responses.map((item) => `
    <article class="admin-list-item">
      <strong>${escapeHtml(item.username || item.userEmail || 'Usuário')}</strong>
      <div class="admin-list-item-meta">
        <span>Setor: ${escapeHtml(sectorLabel(item.sector))}</span>
        <span>Status: ${escapeHtml(getAdminReplyStatusLabel(item.status || 'enviado'))}</span>
        <span>${escapeHtml(item.createdAt ? new Date(item.createdAt).toLocaleString('pt-BR') : 'Sem data')}</span>
      </div>
      <p>${escapeHtml(item.responseText || '')}</p>
      <div class="admin-list-item-meta">
        <span>Alerta: ${escapeHtml(item.alertTitle || ((state.manualAlerts || []).find((alert) => String(alert.id) === String(item.alertId))?.title) || item.alertId || 'Alerta')}</span>
      </div>
      ${item.adminReply ? `<div class="response-bubble response-bubble--admin"><strong>Resposta do admin</strong><p>${escapeHtml(item.adminReply)}</p></div>` : ''}
    </article>
  `).join('');
}


function formatPresenceDate(value) {
  if (!value) return 'Nunca registrado';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return 'Data inválida';
  return date.toLocaleString('pt-BR');
}

function formatPresenceElapsed(value) {
  if (!value) return 'sem registro';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return 'sem registro';
  const diffMs = Math.max(0, Date.now() - date.getTime());
  const diffSeconds = Math.floor(diffMs / 1000);
  if (diffSeconds < 45) return 'agora';
  const diffMinutes = Math.floor(diffSeconds / 60);
  if (diffMinutes < 60) return `há ${diffMinutes} min`;
  const diffHours = Math.floor(diffMinutes / 60);
  if (diffHours < 24) return `há ${diffHours} h`;
  const diffDays = Math.floor(diffHours / 24);
  return `há ${diffDays} dia${diffDays > 1 ? 's' : ''}`;
}

function getPresenceViewName() {
  if (adminModalEl && !adminModalEl.classList.contains('hidden')) return 'Painel admin';
  if (stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) return canValidateStageWorkspace() ? 'Validação PCP' : 'Apontamentos';
  if (sectorAlertsModalEl && !sectorAlertsModalEl.classList.contains('hidden')) return 'Meus alertas';
  if (alertModalEl && !alertModalEl.classList.contains('hidden')) return 'Alertas de prazo';
  if (modalEl && !modalEl.classList.contains('hidden')) return 'Detalhamento de projeto';
  if (state.projectView === 'mine') return 'Meus projetos';
  return 'Painel operacional';
}

async function sendPresenceHeartbeat({ force = false } = {}) {
  if (!state.user) return;
  if (!force && document.visibilityState === 'hidden') return;
  try {
    await fetch('/api/presence', {
      method: 'POST',
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        viewName: getPresenceViewName(),
        viewUrl: `${window.location.pathname}${window.location.search}${window.location.hash}`,
        viewTitle: document.title || 'STEP - Painel Operacional',
      }),
    });
  } catch (error) {
    console.warn('Falha ao atualizar presença do usuário:', error);
  }
}

function startPresenceHeartbeat() {
  window.clearInterval(state.presenceHeartbeatTimer);
  if (!state.user) return;
  sendPresenceHeartbeat({ force: true });
  state.presenceHeartbeatTimer = window.setInterval(() => sendPresenceHeartbeat(), PRESENCE_HEARTBEAT_MS);
}

function stopPresenceHeartbeat() {
  window.clearInterval(state.presenceHeartbeatTimer);
  state.presenceHeartbeatTimer = null;
}

function renderAdminPresence(users = []) {
  if (!adminPresenceSummaryEl || !adminPresenceListEl) return;
  const list = Array.isArray(users) ? users : [];
  const onlineUsers = list
    .filter((user) => Boolean(user.online || user.presence?.online))
    .sort((a, b) => new Date(b.lastSeenAt || b.presence?.lastSeenAt || 0) - new Date(a.lastSeenAt || a.presence?.lastSeenAt || 0));

  adminPresenceSummaryEl.textContent = `${onlineUsers.length} online • ${list.length} usuário(s)`;

  if (!onlineUsers.length) {
    adminPresenceListEl.innerHTML = '<div class="empty-state">Nenhum usuário online agora.</div>';
    return;
  }

  adminPresenceListEl.innerHTML = onlineUsers.map((user) => {
    const presence = user.presence || {};
    const lastSeenAt = user.lastSeenAt || presence.lastSeenAt;
    const lastViewAt = user.lastViewAt || presence.lastViewAt || lastSeenAt;
    const lastViewName = user.lastViewName || presence.lastViewName || 'Painel operacional';
    const lastViewTitle = user.lastViewTitle || presence.lastViewTitle || '';
    return `
      <article class="presence-item presence-item--online">
        <div class="presence-item-head">
          <span class="presence-dot presence-dot--online"></span>
          <strong>${escapeHtml(user.name || user.username || 'Usuário')}</strong>
          <span class="presence-badge presence-badge--online">Online</span>
        </div>
        <div class="admin-list-item-meta">
          <span>Login: ${escapeHtml(user.username || '')}</span>
          <span>Setor: ${escapeHtml(sectorLabel(user.sector))}</span>
          <span>Último sinal: ${escapeHtml(formatPresenceElapsed(lastSeenAt))}</span>
          <span>Última visualização: ${escapeHtml(lastViewName)}${lastViewAt ? ` • ${escapeHtml(formatPresenceDate(lastViewAt))}` : ''}</span>
          ${lastViewTitle && lastViewTitle !== lastViewName ? `<span>Tela: ${escapeHtml(lastViewTitle)}</span>` : ''}
        </div>
      </article>
    `;
  }).join('');
}


function readImageFileAsOptimizedDataUrl(file, options = {}) {
  const maxWidth = Number(options.maxWidth || 900);
  const maxHeight = Number(options.maxHeight || 520);
  const outputWidth = Number(options.outputWidth || 0);
  const outputHeight = Number(options.outputHeight || 0);
  const padding = Math.max(0, Number(options.padding || 0));
  const quality = Number(options.quality || 0.78);
  const background = String(options.background || '#041a2d');
  const mimeType = String(options.mimeType || 'image/jpeg');
  const allowUpscale = Boolean(options.allowUpscale);

  return new Promise((resolve, reject) => {
    if (!file || !String(file.type || '').startsWith('image/')) {
      reject(new Error('Selecione um arquivo de imagem válido.'));
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => reject(new Error('Não foi possível ler a imagem selecionada.'));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error('Não foi possível processar a imagem selecionada.'));
      img.onload = () => {
        const naturalWidth = img.naturalWidth || img.width || maxWidth;
        const naturalHeight = img.naturalHeight || img.height || maxHeight;
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');

        if (outputWidth > 0 && outputHeight > 0) {
          canvas.width = Math.max(1, Math.round(outputWidth));
          canvas.height = Math.max(1, Math.round(outputHeight));
          ctx.fillStyle = background;
          ctx.fillRect(0, 0, canvas.width, canvas.height);

          const innerWidth = Math.max(1, canvas.width - (padding * 2));
          const innerHeight = Math.max(1, canvas.height - (padding * 2));
          const containRatio = Math.min(innerWidth / naturalWidth, innerHeight / naturalHeight);
          const ratio = allowUpscale ? containRatio : Math.min(containRatio, 1);
          const drawWidth = Math.max(1, Math.round(naturalWidth * ratio));
          const drawHeight = Math.max(1, Math.round(naturalHeight * ratio));
          const drawX = Math.round((canvas.width - drawWidth) / 2);
          const drawY = Math.round((canvas.height - drawHeight) / 2);

          ctx.imageSmoothingEnabled = true;
          ctx.imageSmoothingQuality = 'high';
          ctx.drawImage(img, drawX, drawY, drawWidth, drawHeight);
          resolve(canvas.toDataURL(mimeType, quality));
          return;
        }

        let width = naturalWidth;
        let height = naturalHeight;
        const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
        width = Math.max(1, Math.round(width * ratio));
        height = Math.max(1, Math.round(height * ratio));
        canvas.width = width;
        canvas.height = height;
        ctx.fillStyle = background;
        ctx.fillRect(0, 0, width, height);
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL(mimeType, quality));
      };
      img.src = String(reader.result || '');
    };
    reader.readAsDataURL(file);
  });
}


function updateAdminLogoEditorPreview() {
  if (!adminClientLogoEditorEl || !adminClientLogoPreviewImgEl) return;
  const src = String(adminLogoEditorState.source || adminUserClientLogoUrlEl?.value || '').trim();
  if (!src) {
    adminClientLogoEditorEl.classList.add('hidden');
    adminClientLogoPreviewImgEl.removeAttribute('src');
    return;
  }
  adminClientLogoEditorEl.classList.remove('hidden');
  adminClientLogoPreviewImgEl.src = src;
  adminClientLogoPreviewImgEl.style.transform = `translate(${adminLogoEditorState.x}px, ${adminLogoEditorState.y}px) scale(${adminLogoEditorState.zoom})`;
}

function resetAdminLogoEditor(source = '') {
  adminLogoEditorState = {
    source: String(source || adminUserClientLogoUrlEl?.value || '').trim(),
    zoom: 1,
    x: 0,
    y: 0,
  };
  updateAdminLogoEditorPreview();
}

function drawAdjustedClientLogoToDataUrl() {
  return new Promise((resolve, reject) => {
    const src = String(adminLogoEditorState.source || adminUserClientLogoUrlEl?.value || '').trim();
    if (!src) {
      reject(new Error('Importe ou informe uma logo antes de aplicar o ajuste.'));
      return;
    }

    const img = new Image();
    img.onload = () => {
      const outputWidth = 720;
      const outputHeight = 420;
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      canvas.width = outputWidth;
      canvas.height = outputHeight;
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, outputWidth, outputHeight);

      const naturalWidth = img.naturalWidth || img.width || outputWidth;
      const naturalHeight = img.naturalHeight || img.height || outputHeight;
      const baseRatio = Math.max(outputWidth / naturalWidth, outputHeight / naturalHeight);
      const ratio = baseRatio * Math.max(0.2, Number(adminLogoEditorState.zoom || 1));
      const drawWidth = naturalWidth * ratio;
      const drawHeight = naturalHeight * ratio;
      const drawX = ((outputWidth - drawWidth) / 2) + Number(adminLogoEditorState.x || 0);
      const drawY = ((outputHeight - drawHeight) / 2) + Number(adminLogoEditorState.y || 0);

      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';
      ctx.drawImage(img, drawX, drawY, drawWidth, drawHeight);
      resolve(canvas.toDataURL('image/png', 0.92));
    };
    img.onerror = () => reject(new Error('Não foi possível renderizar a logo ajustada.'));
    img.src = src;
  });
}

async function importAdminClientLogoWithEditor() {
  const file = adminUserClientLogoFileEl?.files?.[0];
  if (!file) {
    window.alert('Selecione a imagem da logo do cliente primeiro.');
    return;
  }
  try {
    const dataUrl = await readImageFileAsOptimizedDataUrl(file, {
      maxWidth: 1400,
      maxHeight: 900,
      quality: 0.92,
      background: '#ffffff',
      mimeType: 'image/png'
    });
    if (adminUserClientLogoUrlEl) {
      adminUserClientLogoUrlEl.value = dataUrl;
      adminUserClientLogoUrlEl.dispatchEvent(new Event('input', { bubbles: true }));
    }
    resetAdminLogoEditor(dataUrl);
    if (adminUserFeedbackEl) adminUserFeedbackEl.textContent = 'Logo importada. Ajuste a prévia e clique em Aplicar ajuste antes de salvar.';
  } catch (error) {
    window.alert(error.message || 'Falha ao importar logo do cliente.');
  }
}

async function applyAdminClientLogoAdjustment() {
  try {
    const adjusted = await drawAdjustedClientLogoToDataUrl();
    if (adminUserClientLogoUrlEl) {
      adminUserClientLogoUrlEl.value = adjusted;
      adminUserClientLogoUrlEl.dispatchEvent(new Event('input', { bubbles: true }));
    }
    resetAdminLogoEditor(adjusted);
    if (adminUserFeedbackEl) adminUserFeedbackEl.textContent = 'Ajuste da logo aplicado. Salve o usuário para gravar.';
  } catch (error) {
    window.alert(error.message || 'Falha ao aplicar ajuste da logo.');
  }
}


async function importAdminClientImage(fileInput, targetInput, label) {
  const file = fileInput?.files?.[0];
  if (!file) {
    window.alert(`Selecione a imagem de ${label} primeiro.`);
    return;
  }
  try {
    const isLogo = String(label || '').toLowerCase().includes('logo');
    const dataUrl = await readImageFileAsOptimizedDataUrl(file, isLogo
      ? { outputWidth: 720, outputHeight: 420, padding: 56, quality: 0.92, background: '#ffffff', allowUpscale: true, mimeType: 'image/png' }
      : { maxWidth: 520, maxHeight: 340, quality: 0.68 });
    if (targetInput) {
      targetInput.value = dataUrl;
      targetInput.dispatchEvent(new Event('input', { bubbles: true }));
    }
    if (adminUserFeedbackEl) adminUserFeedbackEl.textContent = `Imagem de ${label} importada. Salve o usuário para gravar.`;
  } catch (error) {
    window.alert(error.message || `Falha ao importar imagem de ${label}.`);
  }
}

function updateAdminClientFieldsVisibility() {
  const role = document.getElementById('admin-user-role')?.value || 'sector';
  if (adminUserClientFieldsEl) adminUserClientFieldsEl.classList.toggle('hidden', role !== 'client');
  if (adminUserClientKeyEl) adminUserClientKeyEl.disabled = role !== 'client';
  if (adminUserClientNameEl) adminUserClientNameEl.disabled = role !== 'client';
  if (adminUserClientLogoUrlEl) adminUserClientLogoUrlEl.disabled = role !== 'client';
  if (adminUserClientLogoFileEl) adminUserClientLogoFileEl.disabled = role !== 'client';
  if (adminUserClientLogoImportEl) adminUserClientLogoImportEl.disabled = role !== 'client';
  if (adminUserClientPlatformImageUrlEl) adminUserClientPlatformImageUrlEl.disabled = role !== 'client';
  if (adminUserClientPlatformNameEl) adminUserClientPlatformNameEl.disabled = role !== 'client';
  if (adminUserClientPlatformImagesEl) adminUserClientPlatformImagesEl.disabled = role !== 'client';
  if (adminUserClientPlatformImageFileEl) adminUserClientPlatformImageFileEl.disabled = role !== 'client';
  if (adminUserClientPlatformImageImportEl) adminUserClientPlatformImageImportEl.disabled = role !== 'client';
}

function resetAdminUserForm() {
  if (adminUserFormEl) adminUserFormEl.reset();
  if (adminUserIdEl) adminUserIdEl.value = "";
  if (adminUserCancelEditEl) adminUserCancelEditEl.classList.add("hidden");
  if (adminUserSubmitLabelEl) adminUserSubmitLabelEl.textContent = "Criar usuário";
  syncOperationRegionButtons('BR');
  setSelectedAdminAlertSectors([document.getElementById("admin-user-sector")?.value || "pintura"]);
  state.adminProjectPmSearchQuery = "";
  const projectPmSearchEl = document.getElementById("admin-user-project-pms-search");
  if (projectPmSearchEl) projectPmSearchEl.value = "";
  setAdminProjectPmAliases([]);
  setAdminQualityCompetencies([]);
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = '';
  if (adminUserClientNameEl) adminUserClientNameEl.value = '';
  if (adminUserClientLogoUrlEl) adminUserClientLogoUrlEl.value = '';
  resetAdminLogoEditor('');
  if (adminUserClientLogoFileEl) adminUserClientLogoFileEl.value = '';
  if (adminUserClientPlatformImageUrlEl) adminUserClientPlatformImageUrlEl.value = '';
  if (adminUserClientPlatformNameEl) adminUserClientPlatformNameEl.value = '';
  if (adminUserClientPlatformImagesEl) adminUserClientPlatformImagesEl.value = '';
  if (adminUserClientPlatformImageFileEl) adminUserClientPlatformImageFileEl.value = '';
  if (adminUserCanViewClientPanelEl) adminUserCanViewClientPanelEl.checked = false;
  updateAdminClientFieldsVisibility();
  updateAdminProjectPmAliasesVisibility();
  updateAdminQualityCompetenciesVisibility();
}

function startEditUser(userId) {
  const list = adminUsersListEl?._cachedUsers || [];
  const user = list.find((item) => String(item.id) === String(userId));
  if (!user) return;
  document.getElementById("admin-user-name").value = user.name || "";
  document.getElementById("admin-user-username").value = user.username || "";
  document.getElementById("admin-user-password").value = "";
  document.getElementById("admin-user-role").value = user.role === "admin" ? "admin" : (user.role === "client" ? "client" : "sector");
  syncOperationRegionButtons(user?.operationRegion || user?.siteKey || user?.portalSite || 'BR');
  document.getElementById("admin-user-sector").value = user.role === "client" ? "all" : (user.sector || "all");
  setSelectedAdminAlertSectors(Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector]);
  state.adminProjectPmSearchQuery = "";
  const projectPmSearchEl = document.getElementById("admin-user-project-pms-search");
  if (projectPmSearchEl) projectPmSearchEl.value = "";
  setAdminProjectPmAliases(user.projectPmAliases || []);
  setAdminQualityCompetencies(user.qualityCompetencies || []);
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = user.clientKey || buildClientKey(user.clientName || '', getOperationRegion(user));
  if (adminUserClientNameEl) adminUserClientNameEl.value = user.clientName || '';
  if (adminUserClientLogoUrlEl) adminUserClientLogoUrlEl.value = user.clientLogoUrl || '';
  resetAdminLogoEditor(user.clientLogoUrl || '');
  if (adminUserClientLogoFileEl) adminUserClientLogoFileEl.value = '';
  if (adminUserClientPlatformImageUrlEl) {
    const platformUrl = user.clientPlatformImageUrl || '';
    adminUserClientPlatformImageUrlEl.value = ''; // imagem padrão desativada: use apenas fotos por plataforma
  }
  if (adminUserClientPlatformNameEl) adminUserClientPlatformNameEl.value = '';
  if (adminUserClientPlatformImagesEl) adminUserClientPlatformImagesEl.value = formatClientPlatformImages(user.clientPlatformImages || '');
  if (adminUserClientPlatformImageFileEl) adminUserClientPlatformImageFileEl.value = '';
  if (adminUserCanViewClientPanelEl) adminUserCanViewClientPanelEl.checked = user.canViewClientPanel === true;
  updateAdminClientFieldsVisibility();
  updateAdminProjectPmAliasesVisibility();
  updateAdminQualityCompetenciesVisibility();
  if (adminUserIdEl) adminUserIdEl.value = user.id || "";
  if (adminUserCancelEditEl) adminUserCancelEditEl.classList.remove("hidden");
  if (adminUserSubmitLabelEl) adminUserSubmitLabelEl.textContent = "Salvar usuário";
  adminUserFeedbackEl.textContent = `Editando ${user.name || user.username}.`;
}

async function syncAdminDataToGithub() {
  if (!adminUserFeedbackEl) return;
  adminUserFeedbackEl.textContent = "Sincronizando com o GitHub...";
  try {
    const response = await fetch("/api/admin-github-config", {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "sync" }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao sincronizar com o GitHub.");
    }
    state.githubSyncEnabled = true;
    adminUserFeedbackEl.textContent = `${data.message || "Sincronizado com sucesso com o GitHub."}`;
    updateSessionUi();
    await loadAdminData();
  } catch (error) {
    adminUserFeedbackEl.textContent = error.message || "Falha ao sincronizar com o GitHub. Verifique GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify.";
    state.githubSyncEnabled = false;
    updateSessionUi();
  }
}

function renderAdminUsersList(users = []) {
  if (!adminUsersListEl) return;
  adminUsersListEl._cachedUsers = users;
  if (!users.length) {
    adminUsersListEl.innerHTML = '<div class="empty-state">Nenhum usuário cadastrado.</div>';
    return;
  }
  adminUsersListEl.innerHTML = users.map((user) => {
    const isSelf = user.id === state.user?.id;
    const presence = user.presence || {};
    const online = Boolean(user.online || presence.online);
    const lastSeenAt = user.lastSeenAt || presence.lastSeenAt;
    const lastLoginAt = user.lastLoginAt || presence.lastLoginAt;
    const lastViewAt = user.lastViewAt || presence.lastViewAt;
    const lastViewName = user.lastViewName || presence.lastViewName || '';
    return `
      <article class="admin-list-item ${online ? 'admin-list-item--online' : ''}">
        <div class="admin-user-title-row">
          <strong>${escapeHtml(user.name)}</strong>
          <span class="presence-badge ${online ? 'presence-badge--online' : 'presence-badge--offline'}">
            <span class="presence-dot ${online ? 'presence-dot--online' : 'presence-dot--offline'}"></span>
            ${online ? 'Online agora' : 'Offline'}
          </span>
        </div>
        <div class="admin-list-item-meta">
          <span>Login: ${escapeHtml(user.username)}</span>
          <span>Perfil: ${escapeHtml(user.role === "admin" ? "Admin notificações" : (user.role === "client" ? "Cliente" : "Setor"))}</span>
          <span>Setor principal: ${escapeHtml(sectorLabel(user.sector))}</span>
          <span>Recebe alertas de: ${escapeHtml(formatSectorList(Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector]))}</span>
          ${user.role === 'client' ? `<span>Cliente: ${escapeHtml(user.clientName || user.clientKey || '—')}</span>` : ''}
          ${user.canViewClientPanel ? '<span>Painel do cliente: <strong>Liberado</strong></span>' : ''}
          <span>Ambiente: <strong>${escapeHtml(user.operationRegion || user.siteKey || user.portalSite || 'BR')}</strong></span>
          ${(userHasProjectsScope(user) && Array.isArray(user.projectPmAliases) && user.projectPmAliases.length) ? `<span>PMs adicionais: ${escapeHtml(user.projectPmAliases.join(', '))}</span>` : ''}
          ${userHasQualityScope(user) ? `<span>Competências da Qualidade: ${escapeHtml(formatQualityCompetencies(user.qualityCompetencies || []))}</span>` : ''}
          <span>${user.active ? "Ativo" : "Inativo"}</span>
          <span>Última atividade: ${escapeHtml(formatPresenceDate(lastSeenAt))}${lastSeenAt ? ` (${escapeHtml(formatPresenceElapsed(lastSeenAt))})` : ''}</span>
          <span>Último login: ${escapeHtml(formatPresenceDate(lastLoginAt))}</span>
          <span>Última visualização: ${escapeHtml(lastViewName || 'Sem registro')}${lastViewAt ? ` • ${escapeHtml(formatPresenceDate(lastViewAt))}` : ''}</span>
        </div>
        <div class="manual-alert-actions">
          <button class="ghost-button ghost-button--compact" type="button" data-user-edit="${escapeHtml(user.id)}">Editar</button>
          ${user.role === "admin"
            ? `<button class="ghost-button ghost-button--compact" type="button" data-user-role="sector" data-user-id="${escapeHtml(user.id)}" ${isSelf ? 'disabled' : ''}>Remover permissão admin</button>`
            : `<button class="primary-button" type="button" data-user-role="admin" data-user-id="${escapeHtml(user.id)}">Permitir como admin</button>`}
        </div>
      </article>
    `;
  }).join("");
}

function getFilteredAdminAlerts() {
  const baseAlerts = Array.isArray(state.manualAlerts) ? state.manualAlerts : [];
  const query = normalizeText(state.adminAlertSearchQuery);
  if (!query) return baseAlerts;
  return baseAlerts.filter((alert) => {
    const acknowledgements = Array.isArray(alert?.acknowledgements) ? alert.acknowledgements : [];
    const haystack = [
      alert?.title,
      alert?.message,
      sectorLabel(alert?.sector),
      priorityLabel(alert?.priority),
      alert?.createdBy,
      alert?.createdAt ? new Date(alert.createdAt).toLocaleString("pt-BR") : "",
      ...acknowledgements.flatMap((ack) => [ack?.username, ack?.userId, sectorLabel(ack?.sector), ack?.acknowledgedAt ? new Date(ack.acknowledgedAt).toLocaleString("pt-BR") : ""]),
    ].join(" ");
    return normalizeText(haystack).includes(query);
  });
}

function renderAdminAlertsList() {
  if (!adminAlertsListEl) return;
  const filteredAlerts = getFilteredAdminAlerts();
  if (!filteredAlerts.length) {
    adminAlertsListEl.innerHTML = `<div class="empty-state">${state.adminAlertSearchQuery ? "Nenhum alerta encontrado para a pesquisa informada." : "Nenhum alerta operacional registrado."}</div>`;
    return;
  }
  adminAlertsListEl.innerHTML = filteredAlerts.map((alert) => {
    const acknowledgements = Array.isArray(alert.acknowledgements) ? alert.acknowledgements : [];
    const ackHtml = alert.requiresAck
      ? (acknowledgements.length
        ? `
          <div class="admin-alert-ack-box">
            <strong>Registro de confirmações</strong>
            <div class="admin-list-item-meta">
              <span>${acknowledgements.length} confirmação(ões)</span>
              <span>Última: ${escapeHtml(new Date(acknowledgements[0].acknowledgedAt).toLocaleString("pt-BR"))}</span>
            </div>
            <div class="admin-alert-ack-list">
              ${acknowledgements.map((ack) => `
                <div class="admin-alert-ack-item">
                  <span><strong>${escapeHtml(ack.username || ack.userId || "Usuário")}</strong></span>
                  <span>Setor: ${escapeHtml(sectorLabel(ack.sector))}</span>
                  <span>${escapeHtml(new Date(ack.acknowledgedAt).toLocaleString("pt-BR"))}</span>
                </div>
              `).join("")}
            </div>
          </div>
        `
        : `
          <div class="admin-alert-ack-box">
            <strong>Registro de confirmações</strong>
            <div class="admin-list-item-meta">
              <span>Aguardando confirmação do setor.</span>
            </div>
          </div>
        `)
      : `
        <div class="admin-alert-ack-box">
          <strong>Registro de confirmações</strong>
          <div class="admin-list-item-meta">
            <span>Alerta informativo sem exigência de leitura.</span>
          </div>
        </div>
      `;

    return `
      <article class="admin-list-item">
        <strong>${escapeHtml(alert.title || "Alerta Operacional")}</strong>
        <div class="admin-list-item-meta">
          <span>Setor: ${escapeHtml(sectorLabel(alert.sector))}</span>
          <span>Prioridade: ${escapeHtml(priorityLabel(alert.priority))}</span>
          <span>${escapeHtml(new Date(alert.createdAt).toLocaleString("pt-BR"))}</span>
          <span>${alert.requiresAck ? "Exige leitura" : "Informativo"}</span>
          <span>${alert.lastAckAt ? `Última confirmação: ${escapeHtml(new Date(alert.lastAckAt).toLocaleString("pt-BR"))}` : "Sem confirmação ainda"}</span>
        </div>
        <p>${escapeHtml(alert.message || "")}</p>
        <div class="admin-list-item-meta">
          <span>${alert.lastAckAt ? "Permaneceu 24h no setor após a leitura" : "Ainda visível no setor até a primeira leitura"}</span>
          <span>Registro permanente no admin</span>
        </div>
        ${ackHtml}
        ${renderAdminResponsesThread(alert.id)}
      </article>
    `;
  }).join("");
}

async function loadAdminData(options = {}) {
  if (state.user?.role !== "admin") return;
  if (options.background && shouldSkipBackgroundRequest(options)) return;
  const now = Date.now();
  if (!options.force && options.background && now - state.lastAdminDataFetchAt < ADMIN_REFRESH_MS) return;
  try {
    const response = await fetch("/api/admin-users", { credentials: "same-origin", cache: "no-store" });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao carregar usuários.");
    }
    state.lastAdminDataFetchAt = Date.now();
    state.githubSyncEnabled = Boolean(data.githubSyncEnabled ?? state.githubSyncEnabled);
    updateSessionUi();
    const remoteUsers = Array.isArray(data.users) ? data.users : [];
    state.userPresence = Array.isArray(data.presence) ? data.presence : [];
    if (state.githubSyncEnabled) {
      renderAdminPresence(remoteUsers);
      renderAdminUsersList(remoteUsers);
      return;
    }
    const localUsers = readLocalUsers().map((user) => ({
      id: user.id,
      name: user.name,
      username: user.username,
      role: user.role,
      sector: user.sector,
      alertSectors: Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector],
      projectPmAliases: Array.isArray(user.projectPmAliases) ? user.projectPmAliases : [],
      qualityCompetencies: Array.isArray(user.qualityCompetencies) ? user.qualityCompetencies : [],
      canViewClientPanel: user.canViewClientPanel === true,
      active: user.active !== false,
      createdAt: user.createdAt || null,
    }));
    const merged = [];
    const seen = new Set();
    for (const user of [...remoteUsers, ...localUsers]) {
      const key = normalizeLoginValue(user.username);
      if (seen.has(key)) continue;
      seen.add(key);
      merged.push(user);
    }
    renderAdminPresence(merged);
    renderAdminUsersList(merged);
  } catch (error) {
    const localUsers = readLocalUsers().map((user) => ({
      id: user.id,
      name: user.name,
      username: user.username,
      role: user.role,
      sector: user.sector,
      alertSectors: Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector],
      projectPmAliases: Array.isArray(user.projectPmAliases) ? user.projectPmAliases : [],
      qualityCompetencies: Array.isArray(user.qualityCompetencies) ? user.qualityCompetencies : [],
      canViewClientPanel: user.canViewClientPanel === true,
      active: user.active !== false,
      createdAt: user.createdAt || null,
    }));
    if (localUsers.length) {
      renderAdminPresence(localUsers);
      renderAdminUsersList(localUsers);
    } else {
      renderAdminPresence([]);
      adminUsersListEl.innerHTML = `<div class="empty-state">${escapeHtml(error.message || "Falha ao carregar usuários.")}</div>`;
    }
  }
  renderAdminAlertsList();
  await loadAlertResponses();
}

function openAdminModal() {
  if (!adminModalEl) return;
  if (adminAlertSearchEl) adminAlertSearchEl.value = state.adminAlertSearchQuery || "";
  setAdminActiveTab(state.adminActiveTab || 'usuario');
  adminModalEl.classList.remove("hidden");
  adminModalEl.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
  loadAdminData({ force: true });
  window.clearInterval(adminResponsesPollTimer);
  adminResponsesPollTimer = window.setInterval(() => {
    if (!adminModalEl.classList.contains('hidden') && state.user?.role === 'admin' && !isPageHidden()) {
      loadAdminData({ background: true });
    }
  }, ADMIN_REFRESH_MS);
}

function closeAdminModal() {
  if (!adminModalEl) return;
  setAdminActiveTab('usuario');
  window.clearInterval(adminResponsesPollTimer);
  adminResponsesPollTimer = null;
  adminModalEl.classList.add("hidden");
  adminModalEl.setAttribute("aria-hidden", "true");
  if (
    modalEl.classList.contains("hidden") &&
    alertModalEl.classList.contains("hidden") &&
    sectorAlertsModalEl.classList.contains("hidden") &&
    stageUpdatesModalEl.classList.contains('hidden') &&
    loginModalEl.classList.contains("hidden")
  ) {
    document.body.classList.remove("modal-open");
  }
}
