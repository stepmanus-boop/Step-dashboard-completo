/* STEP Dashboard v37.11 - Núcleo, estado, helpers, notificações e filtros iniciais. Arquivo gerado a partir da divisão segura do app.js. */
const PROJECTS_REFRESH_MS = 600000; // v11 performance: reduz polling pesado para 10 min
const PROJECTS_CACHE_TTL_MS = 20 * 60 * 1000; // v11 performance: cache local válido por 20 min
const ALERTS_REFRESH_MS = 60000;
const PRESENCE_HEARTBEAT_MS = 90000;
const AUTH_REFRESH_MS = 300000;
const ADMIN_REFRESH_MS = 60000;
const ALERT_NOTIFICATION_COOLDOWN_MS = 4 * 60 * 60 * 1000;
const PROJECTS_CACHE_KEY = 'step_dashboard_projects_cache_v15_first_sync_safe';
// v36.80: acelera o login do Portal do Cliente.
// O painel abre assim que as BSPs aparecem; POs continuam atualizando em segundo plano.
const CLIENT_PORTAL_RELEASE_WITH_PO_PENDING = true;
const CLIENT_PORTAL_PO_BACKGROUND_REFRESH_DELAY_MS = 1200;
const STAGE_TRACKING_VALIDATION_COOLDOWN_MS = 30000;

// Replanejamento WIP mantido no código para uso futuro, mas oculto do painel do cliente nesta versão.
const CLIENT_REPLANNING_UI_ENABLED = false;

let adminResponsesPollTimer = null;
let projectsWarmupPromise = null;
let projectsWarmupResetTimer = null;

const state = {
  projects: [],
  filteredProjects: [],
  projectView: 'all',
  projectDrill: { open: false, mode: 'total', selectedClientKey: '', selectedVesselKey: '' },
  clientPortal: { selectedVesselKey: '', selectedProjectId: null, rowClickTimer: null, vesselClickTimer: null },
  clientPoBackgroundRefreshScheduled: false,
  clientApi: { keys: [], loading: false, newToken: '', newTokenKeyId: '', feedback: '' },
  clientBspOverrides: { items: [], byProjectRowId: {}, byProjectNumber: {}, loading: false, loaded: false, feedback: '', editingProjectId: null, activeExecutiveProjectId: null },
  sectorScopedView: false,
  stats: null,
  meta: null,
  alerts: [],
  searchQuery: "",
  demandFilter: "",
  projectTypeFilter: "",
  weekFilter: "",
  statusFilters: [],
  alertFilter: "all",
  alertSectorFilter: "all",
  alertClientQuery: "",
  selectedProjectId: null,
  selectedProjectDrawerOpen: false,
  modalPendingOnly: false,
  modalIsoSortMode: 'urgency',
  rowClickTimer: null,
  pollTimer: null,
  presenceHeartbeatTimer: null,
  loadingProjectsRequest: null,
  lastProjectsFetchAt: 0,
  lastManualAlertsFetchAt: 0,
  lastAlertResponsesFetchAt: 0,
  lastStageUpdatesFetchAt: 0,
  lastAdminDataFetchAt: 0,
  lastAuthRefreshAt: 0,
  projectsLoadedFromCache: false,
  loginProgressTimer: null,
  loginProgressValue: 0,
  loginProgressActive: false,
  economicMode: true,
  user: null,
  githubSyncEnabled: false,
  manualAlerts: [],
  projectSignals: [],
  adminAlertSearchQuery: "",
  adminActiveTab: "usuario",
  adminProjectPmAliasesDraft: [],
  adminProjectPmSearchQuery: "",
  adminQualityCompetenciesDraft: [],
  userPresence: [],
  alertResponses: [],
  selectedAlertForResponse: null,
  manualAlertSignature: "",
  automaticAlertSignature: "",
  pushSupported: false,
  pushSubscribed: false,
  selectedProjectForSignal: null,
  sectorAlertsMode: 'default',
  stageUpdates: [],
  stageUpdatesSearchQuery: '',
  stageSubmittingKeys: {},
  stageDrafts: {},
  stageBulkSubmitting: false,
  stagePcpPointingMode: false,
  pcpStageSelectedSector: '',
  stageBatchValidationMode: false,
  stageSelectedIds: [],
  stageDatePendencies: [],
  stageDatePendingLoaded: false,
  stageDatePendingLoading: false,
  stageTrackingSubmitting: false,
  stageDateSelectedIds: [],
  attentionPopupQueue: [],
  attentionPopupCurrent: null,
  incomingAlertState: {
    manual: { initialized: false, ids: [] },
    projectSignals: { initialized: false, ids: [] },
    automatic: { initialized: false, ids: [] },
    stageUpdates: { initialized: false, ids: [] },
  },
};

const bodyEl = document.getElementById("projects-body");
const detailCardEl = document.getElementById("detail-card");
const detailDrawerEl = document.getElementById("detail-drawer");
const closeDetailDrawerEl = document.getElementById("close-detail-drawer");
const sheetNameEl = document.getElementById("sheet-name");
const lastSyncEl = document.getElementById("last-sync");
const refreshProjectsButtonEl = document.getElementById("refresh-projects-button");
const footerVersionEl = document.getElementById("footer-version");
const searchInputEl = document.getElementById("project-search");
const clearSearchEl = document.getElementById("clear-search");
const exportFilteredProjectsEl = document.getElementById("export-filtered-projects");
const demandFilterEl = document.getElementById("demand-filter");
const projectTypeFilterEl = document.getElementById("project-type-filter");
const weekFilterEl = document.getElementById("week-filter");
const statusFilterToggleEl = document.getElementById("status-filter-toggle");
const statusFilterMenuEl = document.getElementById("status-filter-menu");
const statusFilterBoxEl = document.getElementById("status-filter-box");
const searchCountEl = document.getElementById("search-count");
const tableShellEl = document.getElementById("table-shell");
const projectViewTabsEl = document.getElementById("project-view-tabs");
const totalProjectsCardEl = document.getElementById("total-projects-card");
const startedProjectsCardEl = document.getElementById("started-projects-card");
const notStartedCardEl = document.getElementById("not-started-card");
const onHoldCardEl = document.getElementById("on-hold-card");
const productionCardEl = document.getElementById("production-card");
const inspectionCardEl = document.getElementById("inspection-card");
const paintingCardEl = document.getElementById("painting-card");
const awaitingShipmentCardEl = document.getElementById("awaiting-shipment-card");
const totalWeightCardEl = document.getElementById("total-weight-card");
const weldedWeightCardEl = document.getElementById("welded-weight-card");
const backlogWeldingCardEl = document.getElementById("backlog-welding-card");
const paintingM2CardEl = document.getElementById("painting-m2-card");
const projectDrillPanelEl = document.getElementById("project-drill-panel");
const projectDrillTitleEl = document.getElementById("project-drill-title");
const projectDrillSubtitleEl = document.getElementById("project-drill-subtitle");
const projectDrillBreadcrumbEl = document.getElementById("project-drill-breadcrumb");
const projectDrillContentEl = document.getElementById("project-drill-content");
const projectDrillBackEl = document.getElementById("project-drill-back");
const projectDrillCloseEl = document.getElementById("project-drill-close");

function getProjectDrillCards() {
  return [
    totalProjectsCardEl,
    startedProjectsCardEl,
    notStartedCardEl,
    onHoldCardEl,
    productionCardEl,
    inspectionCardEl,
    paintingCardEl,
    awaitingShipmentCardEl,
    totalWeightCardEl,
    weldedWeightCardEl,
    backlogWeldingCardEl,
    paintingM2CardEl,
  ].filter(Boolean);
}

const modalEl = document.getElementById("project-modal");
const modalContentEl = document.getElementById("modal-content");
const modalTitleEl = document.getElementById("modal-title");
const modalSubtitleEl = document.getElementById("modal-subtitle");
const modalCloseEl = document.getElementById("modal-close");
const alertModalEl = document.getElementById("alert-modal");
const alertModalContentEl = document.getElementById("alert-modal-content");
const alertModalCloseEl = document.getElementById("alert-modal-close");
const alertBadgeCountEl = document.getElementById("alert-badge-count");
const openAlertsButtonEl = document.getElementById("open-alerts-button");

const loginModalEl = document.getElementById("login-modal");
const loginFormEl = document.getElementById("login-form");
const loginUsernameEl = document.getElementById("login-username");
const loginPasswordEl = document.getElementById("login-password");
const loginFeedbackEl = document.getElementById("login-feedback");
const loginProgressOverlayEl = document.getElementById("login-progress-overlay");
const loginProgressTitleEl = document.getElementById("login-progress-title");
const loginProgressMessageEl = document.getElementById("login-progress-message");
const loginProgressFillEl = document.getElementById("login-progress-fill");
const loginProgressPercentEl = document.getElementById("login-progress-percent");
const loginProgressDetailEl = document.getElementById("login-progress-detail");
const toggleLoginPasswordEl = document.getElementById("toggle-login-password");
const loginCloseEl = document.getElementById("login-close");
const sessionUserNameEl = document.getElementById("session-user-name");
const sessionUserMetaEl = document.getElementById("session-user-meta");
const sessionStatusEl = document.getElementById("session-status");
const logoutButtonEl = document.getElementById("logout-button");
const openChangePasswordButtonEl = document.getElementById("open-change-password-button");
const openClientApiButtonEl = document.getElementById("open-client-api-button");
const openLoginButtonEl = document.getElementById("open-login-button");
const changePasswordModalEl = document.getElementById("change-password-modal");
const changePasswordFormEl = document.getElementById("change-password-form");
const changePasswordCurrentEl = document.getElementById("change-password-current");
const changePasswordNewEl = document.getElementById("change-password-new");
const changePasswordConfirmEl = document.getElementById("change-password-confirm");
const changePasswordFeedbackEl = document.getElementById("change-password-feedback");
const changePasswordCloseEl = document.getElementById("change-password-close");
const openSectorAlertsEl = document.getElementById("open-sector-alerts");
const openMyProjectSignalsEl = document.getElementById("open-my-project-signals");
const openProjectSignalsEl = document.getElementById("open-project-signals");
const openStageUpdatesEl = document.getElementById("open-stage-updates");
const openIsoQrButtonEl = document.getElementById("open-iso-qr-button");

const SECTOR_SCOPED_VIEW_STORAGE_PREFIX = 'step_sector_scoped_view:';

function getSectorScopedViewStorageKey(user = state.user) {
  if (!user) return '';
  const username = String(user.username || user.name || '').trim().toLowerCase();
  return username ? `${SECTOR_SCOPED_VIEW_STORAGE_PREFIX}${username}` : '';
}

function loadSectorScopedViewPreference(user = state.user) {
  const key = getSectorScopedViewStorageKey(user);
  if (!key) return false;
  try {
    return window.localStorage.getItem(key) === '1';
  } catch {
    return false;
  }
}

function saveSectorScopedViewPreference(value, user = state.user) {
  const key = getSectorScopedViewStorageKey(user);
  if (!key) return;
  try {
    window.localStorage.setItem(key, value ? '1' : '0');
  } catch {}
}

const STAGE_DRAFTS_STORAGE_PREFIX = 'step_stage_drafts:';

function getStageDraftStorageKey(user = state.user, sector = getStageWorkspaceSector()) {
  if (!user) return '';
  const username = String(user.username || user.name || '').trim().toLowerCase();
  const normalizedSector = String(sector || 'all').trim().toLowerCase();
  return username ? `${STAGE_DRAFTS_STORAGE_PREFIX}${username}:${normalizedSector}` : '';
}

function loadStageDrafts(user = state.user, sector = getStageWorkspaceSector()) {
  const key = getStageDraftStorageKey(user, sector);
  if (!key) return {};
  try {
    const raw = window.localStorage.getItem(key);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function saveStageDrafts(drafts = state.stageDrafts, user = state.user, sector = getStageWorkspaceSector()) {
  const key = getStageDraftStorageKey(user, sector);
  if (!key) return;
  try {
    window.localStorage.setItem(key, JSON.stringify(drafts || {}));
  } catch {}
}

function getStageDraftKey(projectRowId, spoolIso, sector = getStageWorkspaceSector()) {
  return `${String(projectRowId || '').trim()}::${String(spoolIso || '').trim().toLowerCase()}::${String(sector || '').trim().toLowerCase()}`;
}

function getStageDraft(projectRowId, spoolIso, sector = getStageWorkspaceSector()) {
  const key = getStageDraftKey(projectRowId, spoolIso, sector);
  return state.stageDrafts?.[key] || null;
}

function upsertStageDraft(projectRowId, spoolIso, sector, patch = {}) {
  const key = getStageDraftKey(projectRowId, spoolIso, sector);
  const nextDraft = {
    ...(state.stageDrafts?.[key] || {}),
    projectRowId: String(projectRowId || '').trim(),
    spoolIso: String(spoolIso || '').trim(),
    sector: String(sector || '').trim(),
    progress: '',
    completionDate: '',
    note: '',
    ...patch,
  };
  state.stageDrafts = { ...(state.stageDrafts || {}), [key]: nextDraft };
  saveStageDrafts();
  return nextDraft;
}

function removeStageDraft(projectRowId, spoolIso, sector = getStageWorkspaceSector()) {
  const key = getStageDraftKey(projectRowId, spoolIso, sector);
  if (!state.stageDrafts?.[key]) return;
  const next = { ...(state.stageDrafts || {}) };
  delete next[key];
  state.stageDrafts = next;
  saveStageDrafts();
}

function clearAllStageDrafts() {
  state.stageDrafts = {};
  saveStageDrafts();
}

function getStageDraftEntries(sector = getStageWorkspaceSector()) {
  return Object.values(state.stageDrafts || {}).filter((item) => String(item?.sector || '').trim().toLowerCase() === String(sector || '').trim().toLowerCase());
}

function getReadyStageDraftEntries(sector = getStageWorkspaceSector()) {
  return getStageDraftEntries(sector).filter((item) => item && String(item.projectRowId || '').trim() && String(item.spoolIso || '').trim() && Number(item.progress || 0) > 0);
}

function syncStageDraftsForCurrentSector() {
  state.stageDrafts = loadStageDrafts();
}

const sectorAlertsModalEl = document.getElementById("sector-alerts-modal");
const sectorAlertsCloseEl = document.getElementById("sector-alerts-close");
const sectorAlertsContentEl = document.getElementById("sector-alerts-content");
const alertResponseModalEl = document.getElementById("alert-response-modal");
const alertResponseCloseEl = document.getElementById("alert-response-close");
const alertResponseCancelEl = document.getElementById("alert-response-cancel");
const alertResponseFormEl = document.getElementById("alert-response-form");
const alertResponseAlertIdEl = document.getElementById("alert-response-alert-id");
const alertResponseTitleEl = document.getElementById("alert-response-title");
const alertResponseSubtitleEl = document.getElementById("alert-response-subtitle");
const alertResponseTextEl = document.getElementById("alert-response-text");
const alertResponseFeedbackEl = document.getElementById("alert-response-feedback");
const adminAlertResponsesListEl = document.getElementById("admin-alert-responses-list");
const openAdminPanelEl = document.getElementById("open-admin-panel");
const adminModalEl = document.getElementById("admin-modal");
const adminCloseEl = document.getElementById("admin-close");
const adminUserFormEl = document.getElementById("admin-user-form");
const adminUserFeedbackEl = document.getElementById("admin-user-feedback");
const adminAlertFormEl = document.getElementById("admin-alert-form");
const adminAlertFeedbackEl = document.getElementById("admin-alert-feedback");
const adminUsersListEl = document.getElementById("admin-users-list");
const adminPresenceSummaryEl = document.getElementById("admin-presence-summary");
const adminPresenceListEl = document.getElementById("admin-presence-list");
const adminAlertsListEl = document.getElementById("admin-alerts-list");
const adminAlertSearchEl = document.getElementById("admin-alert-search");
const githubSyncBadgeEl = document.getElementById("github-sync-badge");
const adminSyncButtonEl = document.getElementById("admin-sync-button");
const adminUserCancelEditEl = document.getElementById("admin-user-cancel-edit");
const adminUserTogglePasswordEl = document.getElementById("admin-user-toggle-password");
const adminUserIdEl = document.getElementById("admin-user-id");
const adminUserSubmitLabelEl = document.getElementById("admin-user-submit-label");
const adminUserOperationRegionEl = document.getElementById("admin-user-operation-region");
const adminUserClientFieldsEl = document.getElementById("admin-user-client-fields");
const adminUserClientKeyEl = document.getElementById("admin-user-client-key");
const adminUserClientNameEl = document.getElementById("admin-user-client-name");
const adminUserClientLogoUrlEl = document.getElementById("admin-user-client-logo-url");
const adminUserClientLogoFileEl = document.getElementById("admin-user-client-logo-file");
const adminUserClientLogoImportEl = document.getElementById("admin-user-client-logo-import");
const adminClientLogoEditorEl = document.getElementById("admin-client-logo-editor");
const adminClientLogoPreviewImgEl = document.getElementById("admin-client-logo-preview-img");
const adminUserClientPlatformImageUrlEl = document.getElementById("admin-user-client-platform-image-url");
const adminUserClientPlatformImageFileEl = document.getElementById("admin-user-client-platform-image-file");
const adminUserClientPlatformImageImportEl = document.getElementById("admin-user-client-platform-image-import");
const adminUserClientPlatformNameEl = document.getElementById("admin-user-client-platform-name");
const adminUserClientPlatformImagesEl = document.getElementById("admin-user-client-platform-images");
const adminUserCanViewClientPanelEl = document.getElementById("admin-user-can-view-client-panel");
const adminTabTriggerEls = Array.from(document.querySelectorAll('[data-admin-tab-trigger]'));
const adminTabPanelEls = Array.from(document.querySelectorAll('[data-admin-tab-panel]'));
const projectSignalModalEl = document.getElementById('project-signal-modal');
const projectSignalCloseEl = document.getElementById('project-signal-close');
const projectSignalCancelEl = document.getElementById('project-signal-cancel');
const projectSignalFormEl = document.getElementById('project-signal-form');
const projectSignalProjectIdEl = document.getElementById('project-signal-project-id');
const projectSignalTitleEl = document.getElementById('project-signal-title');
const projectSignalDescriptionEl = document.getElementById('project-signal-description');
const projectSignalFeedbackEl = document.getElementById('project-signal-feedback');
const projectSignalHeadingEl = document.getElementById('project-signal-heading');
const projectSignalSubtitleEl = document.getElementById('project-signal-subtitle');
const stageUpdatesModalEl = document.getElementById('stage-updates-modal');
const stageUpdatesCloseEl = document.getElementById('stage-updates-close');
const stageUpdatesContentEl = document.getElementById('stage-updates-content');
const attentionPopupEl = document.getElementById('attention-popup-modal');
const attentionPopupTitleEl = document.getElementById('attention-popup-title');
const attentionPopupMetaEl = document.getElementById('attention-popup-meta');
const attentionPopupBodyEl = document.getElementById('attention-popup-body');
const attentionPopupActionEl = document.getElementById('attention-popup-action');
const attentionPopupCloseEl = document.getElementById('attention-popup-close');

const installAppButtonEl = document.getElementById("install-app-button");
const connectionStatusEl = document.getElementById("connection-status");
let deferredInstallPrompt = null;
let adminLogoEditorState = { source: '', zoom: 1, x: 0, y: 0 };



function setAdminActiveTab(tab) {
  const validTabs = new Set(['usuario', 'historico', 'alerta']);
  const nextTab = validTabs.has(tab) ? tab : 'usuario';
  state.adminActiveTab = nextTab;
  adminTabTriggerEls.forEach((button) => {
    const active = button.dataset.adminTabTrigger === nextTab;
    button.classList.toggle('is-active', active);
    button.setAttribute('aria-selected', active ? 'true' : 'false');
  });
  adminTabPanelEls.forEach((panel) => {
    const active = panel.dataset.adminTabPanel === nextTab;
    panel.classList.toggle('is-active', active);
    panel.hidden = !active;
  });
}
function isIosDevice() {
  return /iphone|ipad|ipod/i.test(window.navigator.userAgent || "");
}

function isStandaloneMode() {
  return window.matchMedia?.('(display-mode: standalone)')?.matches || window.navigator.standalone === true;
}

function updateConnectionStatus() {
  if (!connectionStatusEl) return;
  const offline = window.navigator.onLine === false;
  connectionStatusEl.textContent = offline ? 'Offline' : 'Online';
  connectionStatusEl.classList.toggle('connection-status--offline', offline);
}

function setupInstallExperience() {
  if (!installAppButtonEl) return;

  const refreshInstallButton = () => {
    const canShow = !isStandaloneMode() && (!!deferredInstallPrompt || isIosDevice());
    installAppButtonEl.classList.toggle('hidden', !canShow);
  };

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredInstallPrompt = event;
    refreshInstallButton();
  });

  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    refreshInstallButton();
  });

  installAppButtonEl.addEventListener('click', async () => {
    if (deferredInstallPrompt) {
      deferredInstallPrompt.prompt();
      try { await deferredInstallPrompt.userChoice; } catch (_) {}
      deferredInstallPrompt = null;
      refreshInstallButton();
      return;
    }

    if (isIosDevice()) {
      window.alert('No iPhone/iPad, abra no Safari, toque em Compartilhar e depois em “Adicionar à Tela de Início”.');
    }
  });

  refreshInstallButton();
}

async function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) return;
  try {
    let refreshing = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (refreshing) return;
      refreshing = true;
      window.location.reload();
    });

    const registration = await navigator.serviceWorker.register('./sw.js?v=41', { updateViaCache: 'none' });
    if (typeof registration.update === 'function') {
      registration.update().catch(() => {});
    }
  } catch (error) {
    console.warn('Falha ao registrar service worker.', error);
  }
}


function getAlertStorageKey(kind = 'manual', userId = '') {
  return `step-last-alerts:${kind}:${userId || 'guest'}`;
}

function buildAlertSignature(list, mapper) {
  return JSON.stringify((Array.isArray(list) ? list : []).map(mapper).sort());
}

function urlBase64ToUint8Array(base64String) {
  const padding = '='.repeat((4 - (base64String.length % 4)) % 4);
  const base64 = (base64String + padding).replace(/-/g, '+').replace(/_/g, '/');
  const rawData = atob(base64);
  return Uint8Array.from([...rawData].map((char) => char.charCodeAt(0)));
}

async function showBrowserNotification(title, body, tag) {
  if (!('Notification' in window) || Notification.permission !== 'granted') return;
  try {
    const registration = await navigator.serviceWorker.getRegistration();
    if (registration?.showNotification) {
      await registration.showNotification(title, { body, tag, icon: '/assets/icon-192.png', badge: '/assets/icon-192.png', data: { url: '/' }, requireInteraction: true, renotify: true, vibrate: [220, 120, 220] });
    } else {
      new Notification(title, { body, tag, requireInteraction: true, renotify: true });
    }
  } catch (error) {
    console.warn('Falha ao exibir notificação.', error);
  }
}

async function syncPushSubscription(forcePrompt = false) {
  if (!state.user || !('serviceWorker' in navigator) || !('PushManager' in window)) return false;
  state.pushSupported = true;
  const registration = await navigator.serviceWorker.ready;
  let permission = Notification.permission;
  if (forcePrompt && permission === 'default') {
    permission = await Notification.requestPermission();
  }
  if (permission !== 'granted') return false;
  const statusRes = await fetch('/api/push-subscriptions', { credentials: 'same-origin', cache: 'no-store' }).catch(() => null);
  const status = statusRes ? await statusRes.json().catch(() => null) : null;
  const vapidPublicKey = status?.vapidPublicKey || '';
  if (!vapidPublicKey) return false;
  let subscription = await registration.pushManager.getSubscription();
  if (!subscription) {
    subscription = await registration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: urlBase64ToUint8Array(vapidPublicKey),
    });
  }
  await fetch('/api/push-subscriptions', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ subscription }),
  });
  state.pushSubscribed = true;
  return true;
}

function readAlertNotificationState(key) {
  try {
    const raw = window.localStorage.getItem(key);
    const parsed = raw ? JSON.parse(raw) : null;
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

function writeAlertNotificationState(key, signature, notifiedAt = null, notifiedWindow = '') {
  window.localStorage.setItem(key, JSON.stringify({ signature, notifiedAt, notifiedWindow }));
}


function buildIncomingAlertId(kind, item = {}) {
  if (kind === 'manual') return String(item.id || '').trim();
  if (kind === 'projectSignals') return String(item.id || '').trim();
  if (kind === 'stageUpdates') return String(item.id || '').trim();
  if (kind === 'automatic') {
    return [item.projectNumber || item.projectDisplay || '', item.sector || '', item.daysRemaining ?? ''].join('::');
  }
  return String(item.id || item.key || '').trim();
}

function getIncomingAlertState(kind) {
  if (!state.incomingAlertState?.[kind]) {
    state.incomingAlertState = { ...(state.incomingAlertState || {}), [kind]: { initialized: false, ids: [] } };
  }
  return state.incomingAlertState[kind];
}

function playAttentionTone() {
  try {
    const AudioCtx = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtx) return;
    const ctx = new AudioCtx();
    const oscillator = ctx.createOscillator();
    const gain = ctx.createGain();
    oscillator.type = 'sine';
    oscillator.frequency.value = 880;
    gain.gain.setValueAtTime(0.001, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.08, ctx.currentTime + 0.02);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.45);
    oscillator.connect(gain);
    gain.connect(ctx.destination);
    oscillator.start();
    oscillator.stop(ctx.currentTime + 0.46);
    oscillator.onended = () => ctx.close().catch(() => {});
  } catch {}
}

function openAttentionPopupTarget(item) {
  if (!item) return;
  const kind = String(item.kind || '').trim().toLowerCase();
  if (kind === 'stage-updates') {
    state.stageBatchValidationMode = Boolean(state.user && normalizeSectorValue(state.user?.sector) === 'pcp');
    openStageUpdatesModal();
    return;
  }
  if (kind === 'pcp-deadline-7') {
    if (typeof openPcpDeadlinesModal === 'function') openPcpDeadlinesModal('due-7');
    return;
  }
  if (kind === 'automatic') {
    openAlertModal(true, { manual: true });
    return;
  }
  if (kind === 'projectsignals') {
    state.sectorAlertsMode = 'project-signals';
    openSectorAlertsModal();
    return;
  }
  state.sectorAlertsMode = 'default';
  openSectorAlertsModal();
}

function renderAttentionPopup(item) {
  if (!attentionPopupEl || !item) return;
  if (attentionPopupTitleEl) attentionPopupTitleEl.textContent = item.title || 'Novo alerta';
  if (attentionPopupMetaEl) attentionPopupMetaEl.textContent = item.meta || '';
  if (attentionPopupBodyEl) attentionPopupBodyEl.textContent = item.message || 'Você recebeu uma nova notificação.';
  if (attentionPopupActionEl) {
    attentionPopupActionEl.textContent = item.actionLabel || 'Abrir alerta';
    attentionPopupActionEl.dataset.attentionAction = item.kind || 'manual';
  }
}

function showNextAttentionPopup() {
  if (!attentionPopupEl || state.attentionPopupCurrent || !state.attentionPopupQueue.length) return;
  state.attentionPopupCurrent = state.attentionPopupQueue.shift();
  renderAttentionPopup(state.attentionPopupCurrent);
  attentionPopupEl.classList.remove('hidden');
  attentionPopupEl.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  playAttentionTone();
}

function queueAttentionPopup(item) {
  if (!item?.dedupeKey) return;
  const currentKey = state.attentionPopupCurrent?.dedupeKey;
  const queuedKeys = new Set((state.attentionPopupQueue || []).map((entry) => entry?.dedupeKey).filter(Boolean));
  if (item.dedupeKey === currentKey || queuedKeys.has(item.dedupeKey)) return;
  state.attentionPopupQueue = [...(state.attentionPopupQueue || []), item];
  if (document.visibilityState === 'visible') {
    showNextAttentionPopup();
  }
}

function closeAttentionPopup(options = {}) {
  if (!attentionPopupEl || !state.attentionPopupCurrent) return;
  const current = state.attentionPopupCurrent;
  attentionPopupEl.classList.add('hidden');
  attentionPopupEl.setAttribute('aria-hidden', 'true');
  state.attentionPopupCurrent = null;
  if (options.openTarget) {
    openAttentionPopupTarget(current);
  }
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
  if (document.visibilityState === 'visible') {
    window.setTimeout(showNextAttentionPopup, 30);
  }
}

function buildAttentionPopupItem(kind, item = {}) {
  const normalizedKind = String(kind || '').trim();
  const baseId = buildIncomingAlertId(kind, item);
  if (!baseId) return null;
  if (normalizedKind === 'manual') {
    return {
      kind: 'manual',
      dedupeKey: `manual:${baseId}`,
      title: item.title || 'Novo alerta operacional',
      meta: `Setor: ${sectorLabel(item.sector)}${item.priority ? ` • Prioridade: ${String(item.priority).toUpperCase()}` : ''}`,
      message: item.message || 'Você recebeu um novo alerta operacional.',
      actionLabel: 'Abrir alerta',
    };
  }
  if (normalizedKind === 'projectSignals') {
    return {
      kind: 'projectsignals',
      dedupeKey: `projectSignals:${baseId}`,
      title: item.title || 'Nova sinalização para o PCP',
      meta: `Projetos • ${item.createdBy || 'Usuário'}`,
      message: item.message || 'Uma nova sinalização foi enviada para validação.',
      actionLabel: 'Abrir sinalização',
    };
  }
  if (normalizedKind === 'automatic') {
    return {
      kind: 'automatic',
      dedupeKey: `automatic:${baseId}`,
      title: 'Prazo em alerta',
      meta: `${item.projectDisplay || item.projectNumber || 'Projeto'} • ${sectorLabel(item.sector)}`,
      message: `${item.projectDisplay || item.projectNumber || 'Projeto'} requer atenção do seu setor.`,
      actionLabel: 'Abrir alertas',
    };
  }
  if (normalizedKind === 'stageUpdates') {
    return {
      kind: 'stage-updates',
      dedupeKey: `stageUpdates:${baseId}`,
      title: item.status && String(item.status).toLowerCase().includes('review') ? 'Nova revisão para o PCP' : 'Novo apontamento para validação',
      meta: `${item.projectDisplay || item.projectNumber || 'Projeto'} • ${item.spoolIso || 'Spool'} • ${sectorLabel(item.sector)}`,
      message: item.note || 'Um novo apontamento foi enviado e aguarda validação do PCP.',
      actionLabel: 'Abrir apontamentos',
    };
  }
  return null;
}

function syncIncomingAlerts(kind, items = []) {
  const bucket = getIncomingAlertState(kind);
  const currentIds = (Array.isArray(items) ? items : []).map((item) => buildIncomingAlertId(kind, item)).filter(Boolean);
  if (!bucket.initialized) {
    bucket.initialized = true;
    bucket.ids = currentIds;
    return;
  }
  const previousIds = new Set(bucket.ids || []);
  (Array.isArray(items) ? items : []).forEach((item) => {
    const itemId = buildIncomingAlertId(kind, item);
    if (!itemId || previousIds.has(itemId)) return;
    const popupItem = buildAttentionPopupItem(kind, item);
    if (popupItem) queueAttentionPopup(popupItem);
  });
  bucket.ids = currentIds;
}

function getProjectAlertWindow(date = new Date()) {
  const hours = Number(date.getHours());
  if (hours === 9) return `${date.toISOString().slice(0, 10)}:09`;
  if (hours === 14) return `${date.toISOString().slice(0, 10)}:14`;
  return '';
}

function shouldNotifyAlert(stateEntry, signature, options = {}) {
  if (!signature) return false;
  if (!stateEntry?.signature) return false;
  if (stateEntry.signature === signature) return false;
  const scheduleOnly = Boolean(options.scheduleOnly);
  if (scheduleOnly) {
    const activeWindow = getProjectAlertWindow();
    if (!activeWindow) return false;
    return stateEntry?.notifiedWindow !== activeWindow;
  }
  const lastNotifiedAt = Number(stateEntry.notifiedAt || 0);
  return !lastNotifiedAt || (Date.now() - lastNotifiedAt) >= ALERT_NOTIFICATION_COOLDOWN_MS;
}

function detectNewUserAlerts() {
  if (!state.user || state.user.role === 'admin') return;
  const manualAlerts = Array.isArray(state.manualAlerts) ? state.manualAlerts : [];
  const projectSignals = Array.isArray(state.projectSignals) ? state.projectSignals : [];
  const automaticAlerts = getUserAutomaticAlerts();
  syncIncomingAlerts('manual', manualAlerts);
  syncIncomingAlerts('projectSignals', projectSignals);
  syncIncomingAlerts('automatic', automaticAlerts);
  const manualSignature = buildAlertSignature(manualAlerts, (item) => `${item.id}:${item.updatedAt || item.createdAt || ''}`);
  const automaticSignature = buildAlertSignature(automaticAlerts, (item) => `${item.projectNumber || item.projectDisplay}:${item.sector}:${item.daysRemaining}`);
  const manualKey = getAlertStorageKey('manual', state.user.sub || state.user.username);
  const autoKey = getAlertStorageKey('automatic', state.user.sub || state.user.username);
  const prevManual = readAlertNotificationState(manualKey);
  const prevAuto = readAlertNotificationState(autoKey);

  let manualNotifiedAt = prevManual?.notifiedAt || null;
  let autoNotifiedAt = prevAuto?.notifiedAt || null;
  let manualNotifiedWindow = prevManual?.notifiedWindow || '';
  let autoNotifiedWindow = prevAuto?.notifiedWindow || '';
  const scheduledProjectAlerts = userHasProjectsScope(state.user) && state.projectView === 'mine';
  const activeWindow = scheduledProjectAlerts ? getProjectAlertWindow() : '';

  if (shouldNotifyAlert(prevManual, manualSignature, { scheduleOnly: scheduledProjectAlerts })) {
    const latest = manualAlerts[0];
    if (latest) {
      showBrowserNotification('Novo alerta operacional', latest.title || latest.message || 'Você recebeu um novo alerta.', `manual-${latest.id}`);
      manualNotifiedAt = Date.now();
      manualNotifiedWindow = activeWindow || '';
    }
  }
  if (shouldNotifyAlert(prevAuto, automaticSignature, { scheduleOnly: scheduledProjectAlerts })) {
    const latestAuto = automaticAlerts[0];
    if (latestAuto) {
      showBrowserNotification('Prazo em alerta', `${latestAuto.projectDisplay || latestAuto.projectNumber || 'Projeto'} requer atenção do seu setor.`, `auto-${latestAuto.projectNumber || latestAuto.projectDisplay}`);
      autoNotifiedAt = Date.now();
      autoNotifiedWindow = activeWindow || '';
    }
  }
  writeAlertNotificationState(manualKey, manualSignature, manualNotifiedAt, manualNotifiedWindow);
  writeAlertNotificationState(autoKey, automaticSignature, autoNotifiedAt, autoNotifiedWindow);
  state.manualAlertSignature = manualSignature;
  state.automaticAlertSignature = automaticSignature;
}

function formatNumber(value, fractionDigits = 0) {
  if (value == null || Number.isNaN(value)) return "—";
  return Number(value).toLocaleString("pt-BR", {
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  });
}

function formatPercent(value) {
  if (value == null || Number.isNaN(value)) return "—";
  return `${Number(value).toLocaleString("pt-BR", {
    minimumFractionDigits: value % 1 === 0 ? 0 : 1,
    maximumFractionDigits: 1,
  })}%`;
}


function getTrackingCacheTimestamp(meta = state.meta) {
  if (!meta || typeof meta !== 'object') return null;
  return meta.persistentCacheUpdatedAt || meta.cacheUpdatedAt || meta.lastSync || meta.persistentCacheServedAt || null;
}

function getTrackingCacheAgeMs(meta = state.meta) {
  const timestamp = getTrackingCacheTimestamp(meta);
  if (!timestamp) return null;
  const ms = new Date(timestamp).getTime();
  if (!Number.isFinite(ms) || ms <= 0) return null;
  return Math.max(0, Date.now() - ms);
}

function formatDurationPtBr(ms) {
  const totalSeconds = Math.max(0, Math.floor(Number(ms || 0) / 1000));
  if (totalSeconds < 60) return 'agora mesmo';
  const totalMinutes = Math.floor(totalSeconds / 60);
  if (totalMinutes < 60) return `há ${totalMinutes} min`;
  const totalHours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (totalHours < 24) {
    return minutes ? `há ${totalHours}h ${minutes}min` : `há ${totalHours}h`;
  }
  const days = Math.floor(totalHours / 24);
  const hours = totalHours % 24;
  if (days < 7) {
    return hours ? `há ${days}d ${hours}h` : `há ${days}d`;
  }
  return `há ${days}d`;
}

function formatTrackingCacheUpdateLabel(meta = state.meta, options = {}) {
  const prefix = options.prefix || 'Última atualização do cache';
  const timestamp = getTrackingCacheTimestamp(meta);
  if (!timestamp) return `${prefix}: --`;

  const date = new Date(timestamp);
  if (!Number.isFinite(date.getTime())) return `${prefix}: --`;

  const ageMs = getTrackingCacheAgeMs(meta);
  const ageLabel = ageMs == null ? '' : ` • ${formatDurationPtBr(ageMs)}`;
  const refreshAfterMs = Number(meta?.persistentCacheAutoRefreshAfterMs || meta?.cacheAutoRefreshAfterMs || 15 * 60 * 1000);
  const stale = ageMs != null && refreshAfterMs > 0 && ageMs >= refreshAfterMs;
  const staleLabel = stale ? ' • aguardando rotina de atualização' : '';

  // v37.23: Supabase grava timestamptz em UTC (+00).
  // A exibição do painel sempre converte para horário de Brasília para evitar
  // parecer que a atualização não aconteceu por diferença de fuso.
  const brDateTime = date.toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  return `${prefix}: ${brDateTime} (BR)${ageLabel}${staleLabel}`;
}

function applyTrackingCacheUpdateLabel(target, meta = state.meta, options = {}) {
  const node = typeof target === 'string' ? document.getElementById(target) : target;
  if (!node) return;
  node.textContent = formatTrackingCacheUpdateLabel(meta, options);
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9_ -]+/g, "");
}



function normalizeCompactText(value) {
  return normalizeText(value).replace(/[_ -]+/g, "");
}


const CLIENT_TRATATIVA_OBSERVATION_RULES = [
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

function textMatchesAliasRule(text, aliases = []) {
  const normalized = normalizeText(text || "");
  const compact = normalizeCompactText(text || "");
  if (!normalized && !compact) return false;
  return aliases.some((alias) => {
    const normalizedAlias = normalizeText(alias || "");
    const compactAlias = normalizeCompactText(alias || "");
    return Boolean(
      (normalizedAlias && normalized.includes(normalizedAlias))
      || (compactAlias && compact.includes(compactAlias))
    );
  });
}

function getProjectObservationContexts(project) {
  if (!project) return [];
  const contexts = [];
  const add = (source, text) => {
    const value = String(text || "").trim();
    if (value) contexts.push({ source, text: value });
  };

  add("BSP", project.observations);
  add("BSP", project.OBSERVATIONS);
  add("BSP", project.observation);
  add("BSP", project.note);
  add("BSP", project.notes);
  add("BSP", project.comments);

  if (Array.isArray(project.spools)) {
    project.spools.forEach((spool, index) => {
      const tag = spool?.iso || spool?.drawing || spool?.description || `Tag ${index + 1}`;
      add(`Tag ${tag}`, spool?.observations);
      add(`Tag ${tag}`, spool?.OBSERVATIONS);
      add(`Tag ${tag}`, spool?.observation);
      add(`Tag ${tag}`, spool?.note);
      add(`Tag ${tag}`, spool?.notes);
      add(`Tag ${tag}`, spool?.comments);
    });
  }

  return contexts;
}

function getProjectTratativaObservationMatches(project) {
  const matches = [];
  const seen = new Set();
  const seenLabelText = new Set();
  const addMatch = (match) => {
    const label = String(match?.label || "").trim();
    const source = String(match?.source || "BSP").trim() || "BSP";
    const text = String(match?.text || "").trim();
    if (!label || !text) return;
    const key = `${label}|${source}|${text}`;
    const labelTextKey = `${label}|${normalizeText(text)}`;
    if (seen.has(key) || seenLabelText.has(labelTextKey)) return;
    seen.add(key);
    seenLabelText.add(labelTextKey);
    matches.push({ label, source, text });
  };

  if (Array.isArray(project?.tratativaObservationMatches)) {
    project.tratativaObservationMatches.forEach(addMatch);
  }

  if (Array.isArray(project?.spools)) {
    project.spools.forEach((spool) => {
      if (Array.isArray(spool?.tratativaObservationMatches)) {
        spool.tratativaObservationMatches.forEach(addMatch);
      }
    });
  }

  for (const context of getProjectObservationContexts(project)) {
    for (const rule of CLIENT_TRATATIVA_OBSERVATION_RULES) {
      if (!textMatchesAliasRule(context.text, rule.aliases)) continue;
      addMatch({ label: rule.label, source: context.source, text: context.text });
    }
  }
  return matches;
}

function projectHasTratativaObservation(project) {
  return getProjectTratativaObservationMatches(project).length > 0;
}

function getProjectTratativaReason(project) {
  const matches = getProjectTratativaObservationMatches(project);
  if (!matches.length) return "";
  return matches.map((match) => `${match.label} • ${match.source}: ${match.text}`).join(" | ");
}

function renderClientTratativaNotice(project) {
  const matches = getProjectTratativaObservationMatches(project);
  if (!matches.length) return "";
  return `
    <section class="client-observation-notice client-observation-notice--tratativa">
      <div>
        <span>Em tratativa</span>
        <strong>Observação crítica ativa</strong>
      </div>
      <ul>
        ${matches.map((match) => `<li><strong>${escapeHtml(match.label)}</strong> — ${escapeHtml(match.source)}: ${escapeHtml(match.text)}</li>`).join("")}
      </ul>
    </section>
  `;
}

function renderClientOnHoldNotice(project) {
  if (!isProjectOnHold(project)) return "";
  const reason = getProjectHoldReason(project);
  return `
    <section class="client-observation-notice client-observation-notice--hold">
      <div>
        <span>ON HOLD</span>
        <strong>BSP EM ON HOLD — em espera operacional</strong>
      </div>
      <p>${escapeHtml(reason || "On Hold identificado")}</p>
    </section>
  `;
}

function buildSearchIndex(parts) {
  const values = (parts || []).filter(Boolean).map((item) => String(item));
  const expanded = [];

  values.forEach((value) => {
    expanded.push(value);
    const compact = normalizeCompactText(value);
    if (compact) expanded.push(compact);
    const digitsOnly = String(value).replace(/\D+/g, "");
    if (digitsOnly) expanded.push(digitsOnly);
  });

  return normalizeText(expanded.join(" | "));
}

function matchesFlexibleSearch(values, query) {
  const rawQuery = String(query || '').trim();
  const normalizedQuery = normalizeText(rawQuery).trim();
  const compactQuery = normalizeCompactText(rawQuery).trim();
  const digitsQuery = rawQuery.replace(/\D+/g, "");

  if (!normalizedQuery && !compactQuery && !digitsQuery) return true;

  const index = buildSearchIndex(values || []);
  return Boolean(
    (normalizedQuery && index.includes(normalizedQuery))
    || (compactQuery && index.includes(compactQuery))
    || (digitsQuery && index.includes(digitsQuery))
  );
}

function refocusStageSearchInput(caretPosition = null) {
  window.requestAnimationFrame(() => {
    const input = stageUpdatesModalEl?.querySelector('[data-stage-search="true"]');
    if (!input) return;
    input.focus();
    const position = Number.isFinite(Number(caretPosition))
      ? Number(caretPosition)
      : String(input.value || '').length;
    try {
      input.setSelectionRange(position, position);
    } catch {}
  });
}

function normalizeLoginValue(value) {
  return normalizeText(value || "");
}

function getLocalUsersStorageKey() {
  return "step-admin-local-users";
}

function readLocalUsers() {
  try {
    const raw = window.localStorage.getItem(getLocalUsersStorageKey());
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function writeLocalUsers(users) {
  try {
    window.localStorage.setItem(getLocalUsersStorageKey(), JSON.stringify(Array.isArray(users) ? users : []));
  } catch {}
}

function upsertLocalUser(user) {
  const users = readLocalUsers();
  const key = normalizeLoginValue(user.username);
  const next = users.filter((item) => normalizeLoginValue(item.username) !== key);
  next.push(user);
  writeLocalUsers(next);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}


const AVAILABLE_SECTORS = [
  { value: "engenharia", label: "Engenharia" },
  { value: "suprimento", label: "Suprimento" },
  { value: "pintura", label: "Pintura" },
  { value: "inspecao", label: "Qualidade" },
  { value: "pendente_envio", label: "Logística" },
  { value: "producao", label: "Produção" },
  { value: "calderaria", label: "Calderaria" },
  { value: "solda", label: "Solda" },
  { value: "pcp", label: "PCP" },
  { value: "projetos", label: "Projetos" },
];

const QUALITY_COMPETENCY_OPTIONS = [
  { value: "dimensional_inicial", label: "Inspeção Dimensional Inicial" },
  { value: "dimensional_final", label: "Inspeção Dimensional Final" },
  { value: "nde", label: "END / NDE" },
  { value: "th", label: "TH" },
  { value: "final_inspection_qc", label: "Final Inspection QC" },
];

function normalizeSectorValue(value) {
  const normalized = normalizeText(value)
    .replace(/[\s-]+/g, '_')
    .replace(/__+/g, '_');

  if (!normalized) return "";
  if (["enviado", "sent", "shipped", "delivered_final"].includes(normalized)) return "enviado";
  if (["envio", "pendenteenvio", "pendente_envio", "pendente_de_envio", "pending_shipment", "awaiting_shipment", "logistica", "logistica_", "logistics", "expedicao", "shipping"].includes(normalized)) return "pendente_envio";
  if (["inspecao", "inspection", "qualidade", "quality"].includes(normalized)) return "inspecao";
  if (["engenharia", "engineering"].includes(normalized)) return "engenharia";
  if (["suprimento", "suprimentos", "supply", "supply_chain", "procurement"].includes(normalized)) return "suprimento";
  if (["pintura", "painting", "coating"].includes(normalized)) return "pintura";
  if (["producao", "production"].includes(normalized)) return "producao";
  if (["calderaria", "boilermaker", "fabrication"].includes(normalized)) return "calderaria";
  if (["solda", "welding"].includes(normalized)) return "solda";
  if (["pcp", "planejamento", "planejamento_controle_producao", "planning", "planning_control"].includes(normalized)) return "pcp";
  if (["projetos", "projeto", "project", "projects", "pm"].includes(normalized)) return "projetos";
  if (["all", "todos", "todo", "geral", "tudo"].includes(normalized)) return "all";
  return normalized;
}

function getUserAlertSectors(user = state.user) {
  if (!user) return [];
  const values = [];
  if (user.sector && user.sector !== "all") values.push(user.sector);
  if (Array.isArray(user.alertSectors)) values.push(...user.alertSectors);
  const seen = new Set();
  return values.map((item) => normalizeSectorValue(item)).filter((item) => item && item !== "all" && !seen.has(item) && seen.add(item));
}

function formatSectorList(values = []) {
  const labels = getUniqueSectorLabels(values);
  return labels.length ? labels.join(", ") : "—";
}

function getUniqueSectorLabels(values = []) {
  const seen = new Set();
  const labels = [];
  for (const value of values) {
    const normalized = normalizeSectorValue(value);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    labels.push(sectorLabel(normalized));
  }
  return labels;
}

function getSelectedAdminAlertSectors() {
  return Array.from(document.querySelectorAll('[data-admin-alert-sector-option]:checked')).map((input) => normalizeSectorValue(input.value));
}

function setSelectedAdminAlertSectors(values = []) {
  const allowed = new Set((Array.isArray(values) ? values : []).map((item) => normalizeSectorValue(item)));
  document.querySelectorAll('[data-admin-alert-sector-option]').forEach((input) => {
    input.checked = allowed.has(normalizeSectorValue(input.value));
  });
}

function normalizeProjectPmAliases(input = []) {
  const values = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const seen = new Set();
  const aliases = [];
  for (const value of values) {
    const item = String(value || '').trim();
    if (!item) continue;
    const key = normalizeText(item);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    aliases.push(item);
  }
  return aliases;
}

function splitProjectPmNames(value = '') {
  return String(value || '')
    .split(/[\n;,|/]+/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function getAvailableProjectPmAliases(extraValues = []) {
  const values = [];
  if (Array.isArray(extraValues)) values.push(...extraValues);
  for (const project of Array.isArray(state.projects) ? state.projects : []) {
    values.push(...splitProjectPmNames(project?.pm || ''));
  }
  return normalizeProjectPmAliases(values).sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));
}

function getAdminProjectPmAliases() {
  return normalizeProjectPmAliases(state.adminProjectPmAliasesDraft || []);
}

function updateAdminProjectPmAliasCount() {
  const countEl = document.getElementById('admin-user-project-pms-count');
  if (!countEl) return;
  const selected = getAdminProjectPmAliases();
  countEl.textContent = selected.length
    ? `${selected.length} PM${selected.length > 1 ? 's' : ''} adicional${selected.length > 1 ? 'is' : ''} selecionado${selected.length > 1 ? 's' : ''}: ${selected.join(', ')}`
    : 'Nenhum PM adicional selecionado.';
}

function renderAdminProjectPmAliasOptions() {
  const optionsEl = document.getElementById('admin-user-project-pms-options');
  if (!optionsEl) return;

  const selectedValues = getAdminProjectPmAliases();
  const selectedKeys = new Set(selectedValues.map((item) => normalizeText(item)));
  const search = normalizeText(state.adminProjectPmSearchQuery || document.getElementById('admin-user-project-pms-search')?.value || '');
  const allOptions = getAvailableProjectPmAliases(selectedValues);
  const filteredOptions = allOptions.filter((name) => !search || normalizeText(name).includes(search));

  if (!allOptions.length) {
    optionsEl.innerHTML = '<div class="pm-select-empty">Nenhum nome de PM encontrado nos projetos carregados.</div>';
    updateAdminProjectPmAliasCount();
    return;
  }

  if (!filteredOptions.length) {
    optionsEl.innerHTML = '<div class="pm-select-empty">Nenhum PM encontrado para essa busca.</div>';
    updateAdminProjectPmAliasCount();
    return;
  }

  optionsEl.innerHTML = filteredOptions.map((name) => {
    const checked = selectedKeys.has(normalizeText(name)) ? 'checked' : '';
    const disabled = adminUserFormHasProjectsScope() ? '' : 'disabled';
    return `
      <label class="check-row pm-select-row">
        <input type="checkbox" data-admin-project-pm-option value="${escapeHtml(name)}" ${checked} ${disabled} />
        ${escapeHtml(name)}
      </label>
    `;
  }).join('');
  updateAdminProjectPmAliasCount();
}

function setAdminProjectPmAliases(values = []) {
  state.adminProjectPmAliasesDraft = normalizeProjectPmAliases(values);
  renderAdminProjectPmAliasOptions();
}

function setAdminProjectPmSearchQuery(value = '') {
  state.adminProjectPmSearchQuery = String(value || '');
  renderAdminProjectPmAliasOptions();
}

function toggleAdminProjectPmAlias(value, checked) {
  const current = getAdminProjectPmAliases();
  const key = normalizeText(value);
  const next = checked
    ? normalizeProjectPmAliases([...current, value])
    : current.filter((item) => normalizeText(item) !== key);
  state.adminProjectPmAliasesDraft = next;
  updateAdminProjectPmAliasCount();
}

function adminUserFormHasProjectsScope() {
  const role = document.getElementById('admin-user-role')?.value || 'sector';
  if (role === 'admin') return false;
  const sector = normalizeSectorValue(document.getElementById('admin-user-sector')?.value || '');
  return sector === 'projetos' || getSelectedAdminAlertSectors().includes('projetos');
}

function updateAdminProjectPmAliasesVisibility() {
  const field = document.getElementById('admin-user-project-pms-field');
  const searchInput = document.getElementById('admin-user-project-pms-search');
  const optionsEl = document.getElementById('admin-user-project-pms-options');
  if (!field) return;
  const show = adminUserFormHasProjectsScope();
  field.classList.toggle('hidden', !show);
  if (searchInput) searchInput.disabled = !show;
  if (optionsEl) {
    optionsEl.querySelectorAll('input[data-admin-project-pm-option]').forEach((input) => {
      input.disabled = !show;
    });
  }
  if (!show) {
    state.adminProjectPmAliasesDraft = [];
    state.adminProjectPmSearchQuery = '';
    if (searchInput) searchInput.value = '';
  }
  renderAdminProjectPmAliasOptions();
  updateAdminQualityCompetenciesVisibility();
}


function normalizeQualityCompetencies(input = []) {
  const rawValues = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const allowed = new Set(QUALITY_COMPETENCY_OPTIONS.map((item) => item.value));
  const aliases = {
    inicial: 'dimensional_inicial',
    dimensional_inicial: 'dimensional_inicial',
    initial_dimensional: 'dimensional_inicial',
    initial_dimensional_inspection: 'dimensional_inicial',
    dimensional_final: 'dimensional_final',
    final_dimensional: 'dimensional_final',
    final_dimensional_inspection: 'dimensional_final',
    nde: 'nde',
    end: 'nde',
    non_destructive: 'nde',
    non_destructive_examination: 'nde',
    th: 'th',
    hydro: 'th',
    hydro_test: 'th',
    teste_hidrostatico: 'th',
    final_inspection: 'final_inspection_qc',
    final_inspection_qc: 'final_inspection_qc',
    inspection_finish_qc: 'final_inspection_qc',
  };
  const seen = new Set();
  const values = [];
  for (const value of rawValues) {
    const key = normalizeText(value).replace(/[\s-]+/g, '_').replace(/__+/g, '_');
    const normalized = aliases[key] || key;
    if (!allowed.has(normalized) || seen.has(normalized)) continue;
    seen.add(normalized);
    values.push(normalized);
  }
  return values;
}

function qualityCompetencyLabel(value) {
  const normalized = normalizeQualityCompetencies([value])[0] || String(value || '');
  return QUALITY_COMPETENCY_OPTIONS.find((item) => item.value === normalized)?.label || value || '—';
}

function userHasQualityScope(user = state.user) {
  if (!user || user.role === 'admin') return false;
  return normalizeSectorValue(user.sector) === 'inspecao' || getUserAlertSectors(user).includes('inspecao');
}

function adminUserFormHasQualityScope() {
  const role = document.getElementById('admin-user-role')?.value || 'sector';
  if (role === 'admin') return false;
  const sector = normalizeSectorValue(document.getElementById('admin-user-sector')?.value || '');
  return sector === 'inspecao' || getSelectedAdminAlertSectors().includes('inspecao');
}

function getAdminQualityCompetencies() {
  return normalizeQualityCompetencies(state.adminQualityCompetenciesDraft || []);
}

function setAdminQualityCompetencies(values = []) {
  state.adminQualityCompetenciesDraft = normalizeQualityCompetencies(values);
  const selected = new Set(state.adminQualityCompetenciesDraft);
  document.querySelectorAll('[data-admin-quality-competency-option]').forEach((input) => {
    input.checked = selected.has(input.value);
  });
  updateAdminQualityCompetenciesVisibility();
}

function updateAdminQualityCompetenciesVisibility() {
  const field = document.getElementById('admin-user-quality-competencies-field');
  if (!field) return;
  const show = adminUserFormHasQualityScope();
  field.classList.toggle('hidden', !show);
  document.querySelectorAll('[data-admin-quality-competency-option]').forEach((input) => {
    input.disabled = !show;
  });
  if (!show) state.adminQualityCompetenciesDraft = [];
}

function toggleAdminQualityCompetency(value, checked) {
  const current = getAdminQualityCompetencies();
  const normalized = normalizeQualityCompetencies([value])[0];
  if (!normalized) return;
  state.adminQualityCompetenciesDraft = checked
    ? normalizeQualityCompetencies([...current, normalized])
    : current.filter((item) => item !== normalized);
}

function sectorLabel(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "pintura") return "Pintura";
  if (normalized === "inspecao") return "Qualidade";
  if (normalized === "engenharia") return "Engenharia";
  if (normalized === "suprimento") return "Suprimento";
  if (normalized === "pendente_envio") return "Logística";
  if (normalized === "enviado") return "Enviado";
  if (normalized === "producao") return "Produção";
  if (normalized === "calderaria") return "Calderaria";
  if (normalized === "solda") return "Solda";
  if (normalized === "pcp") return "PCP";
  if (normalized === "projetos") return "Projetos";
  if (normalized === "all") return "Todos";
  return value || "—";
}

function priorityLabel(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "urgent") return "Urgente";
  if (normalized === "high") return "Alta";
  if (normalized === "low") return "Baixa";
  return "Normal";
}

function parseDateObject(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate()));
  }

  const raw = String(value).trim();
  if (!raw) return null;

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


function parseClientSafeDateObject(value, options = {}) {
  if (value == null || value === '') return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const parsedDate = parseDateObject(value);
    const year = parsedDate?.getUTCFullYear?.();
    const minYear = Number(options.minYear) || 2020;
    const maxYear = Number(options.maxYear) || 2055;
    return year >= minYear && year <= maxYear ? parsedDate : null;
  }

  const raw = String(value).trim();
  if (!raw) return null;

  // Evita que valores de percentual, booleanos, IDs ou números soltos virem datas
  // por acidente. Ex.: new Date('0') = 01/01/2000, o que achatava a Curva S.
  if (/^-?\d+(?:[.,]\d+)?$/.test(raw)) return null;
  if (/%/.test(raw)) return null;
  if (/^(sim|não|nao|yes|no|true|false|concluido|concluído|completed|n\/a|na)$/i.test(raw)) return null;

  const hasDateShape = /\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/.test(raw)
    || /\b\d{4}-\d{2}-\d{2}\b/.test(raw);
  if (!hasDateShape) return null;

  const parsedDate = parseDateObject(raw);
  if (!parsedDate) return null;
  const year = parsedDate.getUTCFullYear();
  const minYear = Number(options.minYear) || 2020;
  const maxYear = Number(options.maxYear) || 2055;
  if (year < minYear || year > maxYear) return null;
  return parsedDate;
}

function compareProjectsByPlannedFinishDate(a, b) {
  const left = parseDateObject(a?.plannedFinishDate);
  const right = parseDateObject(b?.plannedFinishDate);

  if (left && right) {
    const diff = left.getTime() - right.getTime();
    if (diff !== 0) return diff;
  } else if (left && !right) {
    return -1;
  } else if (!left && right) {
    return 1;
  }

  return String(a?.projectDisplay || "").localeCompare(String(b?.projectDisplay || ""), "pt-BR");
}

function getWeekAnchor(year) {
  const jan1 = new Date(Date.UTC(year, 0, 1));
  const anchor = new Date(jan1);
  anchor.setUTCDate(jan1.getUTCDate() - jan1.getUTCDay());
  return anchor;
}

function getCurrentBrazilDate() {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Sao_Paulo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const parts = formatter.formatToParts(new Date());
  const year = Number(parts.find((item) => item.type === "year")?.value);
  const month = Number(parts.find((item) => item.type === "month")?.value);
  const day = Number(parts.find((item) => item.type === "day")?.value);
  return new Date(Date.UTC(year, month - 1, day));
}

function getCurrentBrazilYear() {
  return getCurrentBrazilDate().getUTCFullYear();
}

function formatProductionWeekLabel(weekNumber, weekYear) {
  return `Semana ${weekNumber} - ${weekYear}`;
}

function getProductionWeekLabelFromDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
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

function getCurrentProductionWeekLabel() {
  return state.meta?.currentWeek || getProductionWeekLabelFromDate(getCurrentBrazilDate());
}

function parseWeekLabel(label) {
  const text = String(label || "").trim();
  const weekMatch = text.match(/Semana\s+(\d+)/i);
  const yearMatch = text.match(/-\s*(\d{4})$/);
  return {
    week: weekMatch ? Number(weekMatch[1]) : Number.MAX_SAFE_INTEGER,
    year: yearMatch ? Number(yearMatch[1]) : getCurrentBrazilYear(),
  };
}

function compareWeekLabels(a, b) {
  const left = parseWeekLabel(a);
  const right = parseWeekLabel(b);
  if (left.year !== right.year) return left.year - right.year;
  if (left.week !== right.week) return left.week - right.week;
  return String(a || "").localeCompare(String(b || ""), "pt-BR");
}

function getAlertSeverity(alert) {
  const type = String(alert?.type || "").toLowerCase();
  if (type.includes("conference")) return "medium";
  if (type.includes("overdue") || type.includes("urgent") || type.includes("deadline")) return "urgent";
  return "medium";
}

function normalizeAlertSectorFilterValue(value) {
  const normalized = normalizeCompactText(value);
  if (!normalized) return "";
  if (["envio", "pendenteenvio", "pendentedeenvio", "awaitingshipment", "pendingshipment", "shipping", "logistica", "logistics", "expedicao"].includes(normalized)) {
    return "envio";
  }
  if (["inspecao", "inspection"].includes(normalized)) return "inspecao";
  return normalized;
}

function alertBelongsToUser(alert) {
  if (!userHasProjectsScope() || state.projectView !== 'mine') return true;
  const project = (() => {
    const projectId = Number(alert?.projectRowId || 0);
    if (projectId) {
      const direct = state.projects.find((item) => item.rowId === projectId);
      if (direct) return direct;
    }
    const projectNumber = normalizeText(alert?.projectNumber || alert?.projectDisplay || '');
    if (!projectNumber) return null;
    return state.projects.find((item) => normalizeText(item.projectNumber) === projectNumber || normalizeText(item.projectDisplay) === projectNumber) || null;
  })();
  if (!project) return false;
  return projectBelongsToUser(project);
}

function getVisibleAlertsSource() {
  const alerts = Array.isArray(state.alerts) ? state.alerts : [];
  if (userHasProjectsScope() && state.projectView === 'mine') {
    return alerts.filter((alert) => alertBelongsToUser(alert));
  }
  if (isSectorScopedViewActive()) {
    return alerts.filter((alert) => alertMatchesScopedSector(alert));
  }
  return alerts;
}

function getFilteredAlerts() {
  let alerts = [...getVisibleAlertsSource()];
  const clientQuery = normalizeText(state.alertClientQuery).trim();

  if (state.alertFilter === "medium") {
    alerts = alerts.filter((alert) => getAlertSeverity(alert) === "medium");
  } else if (state.alertFilter === "urgent") {
    alerts = alerts.filter((alert) => getAlertSeverity(alert) === "urgent");
  }

  if (state.alertSectorFilter && state.alertSectorFilter !== "all") {
    alerts = alerts.filter((alert) => normalizeAlertSectorFilterValue(alert.sector) === state.alertSectorFilter);
  }

  if (clientQuery) {
    alerts = alerts.filter((alert) => {
      const haystack = normalizeText([alert.client, alert.projectDisplay, alert.projectNumber].filter(Boolean).join(" | "));
      return haystack.includes(clientQuery);
    });
  }

  return alerts;
}

function getAlertFilterSummary() {
  const severityMap = { all: 'Tudo', medium: 'Médio', urgent: 'Urgente' };
  const sectorMap = {
    all: 'Todos os setores',
    solda: 'Solda',
    calderaria: 'Calderaria',
    inspecao: 'Qualidade',
    pintura: 'Pintura',
    envio: 'Logística',
  };

  return {
    severity: severityMap[state.alertFilter] || 'Tudo',
    sector: sectorMap[state.alertSectorFilter] || 'Todos os setores',
    client: String(state.alertClientQuery || '').trim() || 'Todos os clientes',
  };
}

function sanitizeFileName(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9-_]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .toLowerCase();
}

function buildAlertPdfFileName() {
  const summary = getAlertFilterSummary();
  const parts = ['alertas'];
  if (summary.sector !== 'Todos os setores') parts.push(summary.sector);
  if (summary.client !== 'Todos os clientes') parts.push(summary.client);
  const stamp = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
  parts.push(stamp);
  return `${sanitizeFileName(parts.join('-')) || 'alertas-relatorio'}.pdf`;
}

async function loadImageAsDataUrl(src) {
  if (!src) return null;
  try {
    const response = await fetch(src);
    if (!response.ok) return null;
    const blob = await response.blob();
    return await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    console.warn('Não foi possível carregar a logo para o PDF.', error);
    return null;
  }
}

async function downloadAlertsPdf() {
  const filteredAlerts = getFilteredAlerts();
  if (!filteredAlerts.length) {
    window.alert('Nenhum alerta encontrado para exportar em PDF.');
    return;
  }

  const jsPdfApi = window.jspdf?.jsPDF;
  if (!jsPdfApi) {
    window.alert('A biblioteca de PDF não foi carregada. Atualize a página e tente novamente.');
    return;
  }

  const doc = new jsPdfApi({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const summary = getAlertFilterSummary();
  const generatedAt = new Date().toLocaleString('pt-BR');
  const logoDataUrl = await loadImageAsDataUrl('./assets/step-logo.png');

  if (logoDataUrl) {
    try {
      doc.addImage(logoDataUrl, 'PNG', 14, 10, 34, 11);
    } catch (error) {
      console.warn('Não foi possível renderizar a logo no PDF.', error);
    }
  }

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.text('Relatório de alertas para impressão', 52, 16);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  const subtitle = `Filtro: ${summary.severity} | Setor: ${summary.sector} | Cliente: ${summary.client} | Total: ${filteredAlerts.length}`;
  doc.text(subtitle, 14, 28);
  doc.text(`Gerado em: ${generatedAt}`, 14, 34);

  const rows = filteredAlerts.map((alert) => {
    const severity = getAlertSeverity(alert) === 'urgent' ? 'Urgente' : 'Médio';
    const daysLabel = alert.daysRemaining < 0
      ? `${Math.abs(alert.daysRemaining)} dia(s) em atraso`
      : `${alert.daysRemaining} dia(s) para o término`;

    return [
      String(alert.projectDisplay || alert.projectNumber || '—'),
      String(alert.client || '—'),
      String(alert.sector || '—'),
      String(alert.title || '—'),
      String(alert.plannedFinishDate || '—'),
      daysLabel,
      String(alert.currentStageGroup || alert.currentStage || '—'),
      String(formatPercent(alert.coatingPercent)),
      severity,
      String(alert.message || '—'),
    ];
  });

  doc.autoTable({
    startY: 40,
    head: [[
      'Projeto', 'Cliente', 'Setor', 'Alerta', 'Término planejado',
      'Prazo', 'Etapa atual', 'Pintura', 'Prioridade', 'Detalhe'
    ]],
    body: rows,
    tableWidth: 'auto',
    styles: {
      font: 'helvetica',
      fontSize: 7,
      cellPadding: 1.4,
      overflow: 'linebreak',
      valign: 'middle',
      lineColor: [220, 228, 236],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: [22, 83, 126],
      textColor: [255, 255, 255],
      fontStyle: 'bold',
      fontSize: 7,
      cellPadding: 1.5,
      halign: 'left',
    },
    columnStyles: {
      0: { cellWidth: 24 },
      1: { cellWidth: 22 },
      2: { cellWidth: 14 },
      3: { cellWidth: 20 },
      4: { cellWidth: 19 },
      5: { cellWidth: 22 },
      6: { cellWidth: 24 },
      7: { cellWidth: 11 },
      8: { cellWidth: 13 },
      9: { cellWidth: 58 },
    },
    margin: { top: 40, right: 8, bottom: 14, left: 8 },
    didDrawPage(data) {
      const footer = `STEP • Página ${data.pageNumber}`;
      doc.setFontSize(9);
      doc.text(footer, pageWidth - 14, pageHeight - 6, { align: 'right' });
    },
  });

  doc.save(buildAlertPdfFileName());
}


function getClientProjectDisplayCode(project) {
  const base = String(project?.projectDisplay || project?.projectNumber || 'BSP').trim() || 'BSP';
  const poDisplay = String(project?.customerPoDisplay || '').trim();
  if (project?.clientDisplayCode) return String(project.clientDisplayCode).trim();
  return `${base} - ${poDisplay || 'Aguardando PO'}`;
}

function projectDisplayWithClient(project) {
  const projectName = String(project?.projectDisplay || '').trim();
  const clientName = String(project?.client || '').trim();
  return clientName ? `${projectName} - ${clientName}` : (projectName || '—');
}

function getProjectClientLabel(project) {
  return String(project?.client || '').trim() || '—';
}

function getProjectVesselLabel(project) {
  return String(project?.vessel || project?.unit || project?.unidade || '').trim() || '—';
}

function getProjectSignalMatchKey(project) {
  return normalizeText(project?.projectNumber || project?.projectDisplay || '').trim();
}

function getProjectSignals(project) {
  const projectKey = getProjectSignalMatchKey(project);
  if (!projectKey) return [];
  const source = Array.isArray(state.projectSignals) && state.projectSignals.length ? state.projectSignals : (Array.isArray(state.manualAlerts) ? state.manualAlerts : []);
  return source
    .filter((alert) => {
      const titleKey = normalizeText(alert?.title || '').trim();
      const messageKey = normalizeText(alert?.message || '').trim();
      return titleKey.includes(projectKey) || messageKey.includes(projectKey);
    })
    .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());
}

function getSignalResolutionInfo(alertId) {
  const responses = getAlertResponsesForAlert(alertId);
  const resolved = [...responses]
    .filter((item) => String(item?.status || '').toLowerCase() === 'resolvida')
    .sort((a, b) => new Date(b.updatedAt || b.createdAt || 0).getTime() - new Date(a.updatedAt || a.createdAt || 0).getTime())[0];
  if (!resolved) return null;
  return {
    username: resolved.username || resolved.userEmail || 'Usuário',
    date: resolved.updatedAt || resolved.createdAt || null,
    note: resolved.responseText || '',
  };
}

function getSignalStatusBadge(alert) {
  const resolved = getSignalResolutionInfo(alert?.id);
  return resolved
    ? '<span class="manual-alert-tag manual-alert-tag--resolved">Resolvida</span>'
    : '<span class="manual-alert-tag manual-alert-tag--pending">Pendente</span>';
}

function canCreateProjectSignal(project = null, user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  if (!(normalizeSectorValue(user.sector) === 'projetos' || userHasProjectsScope(user))) return false;
  if (!project) return true;
  return projectBelongsToUser(project, user);
}

function canResolveSignal(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return normalizeSectorValue(user.sector) === 'pcp' || getUserAlertSectors(user).includes('pcp');
}

function isProjectUserSignal(alert) {
  if (!alert) return false;
  const sector = normalizeSectorValue(alert?.sector);
  const message = normalizeText(alert?.message || '').trim();
  if (sector !== 'pcp') return false;
  return message.includes('projeto') && message.includes('informado por');
}

function getProjectUserSignals(source = null) {
  const list = Array.isArray(source)
    ? source
    : (Array.isArray(state.projectSignals) && state.projectSignals.length ? state.projectSignals : (Array.isArray(state.manualAlerts) ? state.manualAlerts : []));
  return list
    .filter((alert) => isProjectUserSignal(alert))
    .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());
}

function canViewProjectSignals(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return normalizeSectorValue(user.sector) === 'pcp' || getUserAlertSectors(user).includes('pcp');
}

function isMyCreatedSignal(alert, user = state.user) {
  if (!alert || !user) return false;
  return String(alert.createdBy || '').trim().toLowerCase() === String(user.username || '').trim().toLowerCase();
}

function getMyProjectSignals(user = state.user, source = null) {
  const list = Array.isArray(source)
    ? source
    : (Array.isArray(state.projectSignals) && state.projectSignals.length ? state.projectSignals : (Array.isArray(state.manualAlerts) ? state.manualAlerts : []));
  return list
    .filter((alert) => isProjectUserSignal(alert) && isMyCreatedSignal(alert, user))
    .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());
}

function canViewMyProjectSignals(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return normalizeSectorValue(user.sector) === 'projetos' || userHasProjectsScope(user);
}


const STAGE_WORKSPACE_SECTORS = ['engenharia', 'suprimento', 'pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];

function isPcpStageUser(user = state.user) {
  return Boolean(user && normalizeSectorValue(user.sector) === 'pcp');
}

function getStageSectorOptionsHtml(selected = '') {
  const current = normalizeSectorValue(selected);
  return STAGE_WORKSPACE_SECTORS.map((sector) => `<option value="${escapeHtml(sector)}" ${current === sector ? 'selected' : ''}>${escapeHtml(sectorLabel(sector))}</option>`).join('');
}

function ensurePcpStageSectorDefault() {
  if (!isPcpStageUser()) return '';
  const current = normalizeSectorValue(state.pcpStageSelectedSector);
  if (STAGE_WORKSPACE_SECTORS.includes(current)) return current;
  state.pcpStageSelectedSector = 'solda';
  return state.pcpStageSelectedSector;
}

const STAGE_PROGRESS_OPTIONS = [25, 50, 75, 100];

function normalizeStageWorkspaceText(value) {
  return normalizeText(value || '')
    .replace(/[–—−]/g, '-')
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



function parseStageWorkspacePercent(value) {
  if (value == null || value === '' || value === 'N/A') return null;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return null;
    return value >= 0 && value <= 1 ? value * 100 : value;
  }
  let raw = String(value || '').trim();
  if (!raw || raw === 'N/A') return null;
  raw = raw.replace('%', '').replace(/\s/g, '').replace(',', '.');
  const parsed = Number(raw.replace(/[^\d.-]/g, ''));
  if (!Number.isFinite(parsed)) return null;
  return parsed >= 0 && parsed <= 1 ? parsed * 100 : parsed;
}

function hasStageWorkspaceValue(stageValues, key) {
  const value = stageValues?.[key];
  if (value == null) return false;
  const text = String(value).trim();
  return Boolean(text && text !== 'N/A' && text.toLowerCase() !== 'não' && text.toLowerCase() !== 'nao');
}

function getStageWorkspacePercent(stageValues, key) {
  return parseStageWorkspacePercent(stageValues?.[key]) ?? 0;
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

  const th = getStageWorkspacePercent(stageValues, 'Hydro Test Pressure (QC)');
  const nde = parseStageWorkspacePercent(stageValues?.['Non Destructive Examination (QC)']);
  const finalDimensional = getStageWorkspacePercent(stageValues, 'Final Dimensional Inpection/3D (QC)');
  const initialDimensional = getStageWorkspacePercent(stageValues, 'Initial Dimensional Inspection/3D');
  const finalInspection = getStageWorkspacePercent(stageValues, 'Final Inspection');
  const includeHydro = shouldClientShowHydro({ ...(project || {}), projectType: spool?.projectType || project?.projectType, spools: spool ? [spool] : project?.spools });

  if (includeHydro && (text.includes('hydro') || text.includes('teste hidrostatico') || /\bth\b/.test(text) || text.includes('aguardando em th') || (th > 0 && th < 100))) return 'th';
  if (text.includes('nde') || /\bend\b/.test(text) || text.includes('non destructive') || (nde != null && nde > 0 && nde < 100)) return 'nde';
  if (text.includes('final dimensional') || text.includes('final dimensional inspection') || text.includes('final dimensional inpection') || finalDimensional > 0) return 'dimensional_final';
  if (text.includes('initial dimensional') || text.includes('inspecao dimensional inicial') || initialDimensional > 0) return 'dimensional_inicial';
  if (text.includes('final inspection') || text.includes('inspection finish') || finalInspection > 0) return 'final_inspection_qc';

  return 'dimensional_final';
}

function isQualityCompetenceAllowedForUser(project, spool, user = state.user) {
  if (!user || user.role === 'admin') return true;
  if (normalizeSectorValue(user.sector) === 'pcp') return true;
  const competencies = normalizeQualityCompetencies(user.qualityCompetencies || []);
  if (!competencies.length) return true;
  const competence = getSpoolQualityCompetence(project, spool);
  return !competence || competencies.includes(competence);
}

function getSpoolCompetenceSector(project, spool) {
  const stageValues = spool?.stageValues || project?.stageValues || {};
  const spoolOwnText = normalizeStageWorkspaceText([
    spool?.currentStatus,
    spool?.stage,
    spool?.flow?.status,
    spool?.currentSector,
    spool?.operationalSector,
    spool?.flow?.sector,
  ].filter(Boolean).join(' '));
  const spoolSaysPintura = spoolOwnText.includes('pintura')
    || spoolOwnText.includes('paint')
    || spoolOwnText.includes('coating')
    || spoolOwnText.includes('surface preparation')
    || spoolOwnText.includes('intermediaria');
  // v35.9: finalizado nunca entra em apontamento, mesmo que tenha texto antigo de setor.
  const finished = Boolean(spool?.finished || spool?.projectFinishedFlag)
    || spoolOwnText.includes('finalizado')
    || spoolOwnText.includes('concluido')
    || spoolOwnText.includes('concluído');
  if (finished) return '';

  // v36.2: a demanda do setor vem primeiro da etapa/status do próprio ISO.
  const ownTextSector = inferStageSectorFromOwnText(spoolOwnText);
  if (ownTextSector) return ownTextSector;

  const coating = Math.max(
    getStageWorkspacePercent(stageValues, 'Surface preparation and/or coating'),
    getStageWorkspacePercent(stageValues, 'HDG / FBE.  (PAINT)'),
    parseStageWorkspacePercent(spool?.coatingPercent) ?? 0
  );
  const finalInspection = getStageWorkspacePercent(stageValues, 'Final Inspection');
  const packageDelivered = getStageWorkspacePercent(stageValues, 'Package and Delivered');
  const th = getStageWorkspacePercent(stageValues, 'Hydro Test Pressure (QC)');
  const nde = parseStageWorkspacePercent(stageValues?.['Non Destructive Examination (QC)']);
  const finalDimensional = getStageWorkspacePercent(stageValues, 'Final Dimensional Inpection/3D (QC)');
  const fullWelding = getStageWorkspacePercent(stageValues, 'Full welding execution');
  const initialDimensional = getStageWorkspacePercent(stageValues, 'Initial Dimensional Inspection/3D');
  const spoolAssemble = getStageWorkspacePercent(stageValues, 'Spool Assemble and tack weld');
  const weldingPreparation = getStageWorkspacePercent(stageValues, 'Welding Preparation');
  const withdrewMaterial = getStageWorkspacePercent(stageValues, 'Withdrew Material');
  const materialSeparation = getStageWorkspacePercent(stageValues, 'Material Separation');
  const procurement = Math.max(
    getStageWorkspacePercent(stageValues, 'Procuremnt Status %'),
    getStageWorkspacePercent(stageValues, 'Material Release to Fabrication')
  );
  const drawing = getStageWorkspacePercent(stageValues, 'Drawing Execution Advance%');
  const fabricationStarted = Boolean(spool?.fabricationStartDate || hasStageWorkspaceValue(stageValues, 'Fabrication Start Date'));
  const boilermakerDone = hasStageWorkspaceValue(stageValues, 'Boilermaker Finish Date');
  const projectFinishDate = hasStageWorkspaceValue(stageValues, 'Project Finish Date');
  const includeHydro = shouldClientShowHydro({ ...(project || {}), projectType: spool?.projectType || project?.projectType, spools: spool ? [spool] : project?.spools });

  if (projectFinishDate || packageDelivered >= 100) return '';
  if (coating >= 100) return 'pendente_envio';
  if (coating > 0 || (includeHydro && th >= 100) || (!includeHydro && finalDimensional >= 100)) return 'pintura';
  if (fullWelding > 0 && fullWelding < 100) return 'solda';
  if ((includeHydro && th > 0) || (nde != null && nde > 0) || finalDimensional > 0 || (includeHydro && finalDimensional >= 100) || fullWelding >= 100 || initialDimensional > 0 || boilermakerDone || spoolAssemble >= 100) {
    if (initialDimensional >= 100 && fullWelding <= 0) return 'solda';
    return 'inspecao';
  }
  if (fullWelding > 0 || initialDimensional >= 100) return 'solda';
  if (spoolAssemble > 0 || weldingPreparation > 0 || weldingPreparation >= 100 || withdrewMaterial > 0) return 'calderaria';
  if (fabricationStarted || materialSeparation >= 100) return 'producao';
  if (materialSeparation > 0 || procurement > 0 || procurement >= 100 || drawing >= 100) return 'suprimento';
  if (drawing > 0 || drawing >= 0) {
    const textForDrawing = normalizeStageWorkspaceText([spool?.currentStatus, spool?.stage, spool?.flow?.status, project?.currentStage].filter(Boolean).join(' '));
    if (textForDrawing.includes('detalhamento') || textForDrawing.includes('drawing') || !textForDrawing) return 'engenharia';
  }

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

  if (text.includes('finalizado')) return '';
  if (text.includes('package and delivered') || text.includes('final inspection') || text.includes('unitizacao') || text.includes('preparado para envio') || text.includes('logistica')) return 'pendente_envio';
  if (text.includes('pintura') || text.includes('paint') || text.includes('coating') || text.includes('surface preparation') || text.includes('acabamento') || text.includes('intermediaria') || text === 'j f') return 'pintura';
  if (text.includes('hydro') || /\bth\b/.test(text) || text.includes('dimensional') || text.includes('inspection') || text.includes('inspecao') || text.includes('qualidade') || text.includes('nde') || text.includes('end')) return 'inspecao';
  if (text.includes('full welding') || text.includes('solda') || text === 'solda') return 'solda';
  if (text.includes('pre montagem') || text.includes('spool assemble') || text.includes('tack weld') || text.includes('welding preparation') || text.includes('boilermaker') || text.includes('calderaria')) return 'calderaria';
  if (text.includes('corte') || text.includes('limpeza') || text.includes('fabrication start') || text.includes('producao')) return 'producao';
  if (text.includes('separacao de material') || text.includes('material separation') || text.includes('estoque') || text.includes('procure') || text.includes('suprimento')) return 'suprimento';
  if (text.includes('detalhamento') || text.includes('drawing') || text.includes('engenharia')) return 'engenharia';

  return normalizeSectorValue(spool?.currentSector || spool?.operationalSector || spool?.flow?.sector || project?.currentSector || project?.operationalSector || project?.sectorSummary);
}

function isSpoolReleasedForStageSector(project, spool, sector = getStageWorkspaceSector()) {
  const currentSector = normalizeSectorValue(sector);
  if (!currentSector) return false;

  const directText = normalizeStageWorkspaceText([
    spool?.currentStatus,
    spool?.stage,
    spool?.currentSector,
    spool?.operationalSector,
    spool?.flow?.status,
    spool?.flow?.sector,
    spool?.etapaAtual,
  ].filter(Boolean).join(' '));
  const directSector = inferStageSectorFromOwnText(directText);

  // v36.2: se o próprio ISO informa etapa, ela manda.
  // Ex.: login Pintura só vê "Pintura/Aguardando início de pintura/Intermediária".
  // Solda, Corte, Inspeção etc. ficam fora da Pintura.
  if (directSector) {
    if (currentSector === 'inspecao' && directSector === 'inspecao') return isQualityCompetenceAllowedForUser(project, spool);
    return currentSector === directSector;
  }

  const competenceSector = getSpoolCompetenceSector(project, spool);
  if (!competenceSector || currentSector !== competenceSector) return false;
  if (currentSector === 'inspecao') return isQualityCompetenceAllowedForUser(project, spool);
  return true;
}

function filterProjectForStageSector(project, sector = getStageWorkspaceSector()) {
  const originalSpools = Array.isArray(project?.spools) ? project.spools : [];
  const releasedSpools = originalSpools.filter((spool) => isSpoolReleasedForStageSector(project, spool, sector));
  if (!releasedSpools.length) return null;
  return {
    ...project,
    spools: releasedSpools,
    stageWorkspaceTotalSpools: originalSpools.length,
    stageWorkspaceReleasedSpools: releasedSpools.length,
  };
}

function getStageWorkspaceRawProjectMatches() {
  const query = String(state.stageUpdatesSearchQuery || '').trim();
  const source = Array.isArray(state.projects) ? state.projects : [];
  const matches = !query ? source : source.filter((project) => {
    const projectValues = [
      project.projectNumber,
      project.projectDisplay,
      project.projectPrefix,
      project.client,
      project.currentStage,
      project.projectStatus,
      project.jobProcessStatus,
      project.projectType,
      getProjectTypeLabel(project),
      ...(project.spools || []).flatMap((spool) => [spool.iso, spool.description, spool.drawing, spool.currentStatus, spool.stage, spool.currentSector]),
    ];
    return matchesFlexibleSearch(projectValues, query);
  });
  return matches;
}

function getStageWorkspaceBlockedInfo() {
  const sector = getStageWorkspaceSector();
  const rawMatches = getStageWorkspaceRawProjectMatches();
  const blocked = rawMatches
    .map((project) => {
      const spools = Array.isArray(project?.spools) ? project.spools : [];
      const released = spools.filter((spool) => isSpoolReleasedForStageSector(project, spool, sector));
      if (released.length) return null;
      const first = spools[0] || null;
      return {
        project,
        stage: first ? getSpoolStageLabel(project, first) : (getProjectCurrentStageDisplay(project) || 'Etapa não identificada'),
        sector: first ? getSpoolCompetenceSector(project, first) : normalizeSectorValue(project?.currentSector || project?.operationalSector || project?.sectorSummary),
      };
    })
    .filter(Boolean);
  return { count: blocked.length, first: blocked[0] || null };
}

function getStageWorkspaceSector(user = state.user) {
  const ownSector = normalizeSectorValue(user?.sector);
  if (ownSector === 'pcp' && state.stagePcpPointingMode) {
    const selected = normalizeSectorValue(state.pcpStageSelectedSector);
    return STAGE_WORKSPACE_SECTORS.includes(selected) ? selected : '';
  }
  return ownSector;
}

function getStageWorkspaceLabel(sector = getStageWorkspaceSector()) {
  return sectorLabel(sector) || 'Etapa';
}

function canOpenStageWorkspace(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  const sector = getStageWorkspaceSector(user);
  return sector === 'pcp' || STAGE_WORKSPACE_SECTORS.includes(sector);
}

function canValidateStageWorkspace(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return getStageWorkspaceSector(user) === 'pcp';
}

function stageWorkspaceSearchProjects() {
  const sector = getStageWorkspaceSector();
  return getStageWorkspaceRawProjectMatches()
    .map((project) => filterProjectForStageSector(project, sector))
    .filter(Boolean)
    .slice(0, 8);
}

function getStageUpdatesForCurrentSector(source = null, sector = getStageWorkspaceSector()) {
  const list = Array.isArray(source) ? source : (Array.isArray(state.stageUpdates) ? state.stageUpdates : []);
  return list.filter((item) => normalizeSectorValue(item?.sector) === sector);
}

function isPendingStageStatus(status) {
  const normalized = String(status || 'pending').trim().toLowerCase();
  return ['pending', 'pending_advance', 'pending_review'].includes(normalized);
}

function isResolvedStageStatus(status) {
  const normalized = String(status || '').trim().toLowerCase();
  return ['resolved', 'resolved_advance', 'resolved_review'].includes(normalized);
}

function isReviewStageStatus(status) {
  const normalized = String(status || '').trim().toLowerCase();
  return ['pending_review', 'resolved_review'].includes(normalized);
}

function stageUpdateActionLabel(status) {
  return isReviewStageStatus(status) ? 'Revisão' : 'Enviado';
}

function stageUpdateResolveLabel(status) {
  return isReviewStageStatus(status) ? 'Revisão concluída' : 'Concluído';
}

function stageUpdatePendingLabel(status) {
  return isReviewStageStatus(status) ? 'Revisão PCP' : 'Pendente';
}

function getStageTrackingInfo(item) {
  const progress = Number(item?.progress || 0);
  const current = Number(item?.trackingProgress);
  const hasCurrent = Number.isFinite(current);
  const matched = Boolean(item?.trackingMatched) || (hasCurrent && current >= progress);
  const status = String(item?.trackingStatus || '').trim().toLowerCase();
  const checking = !hasCurrent && ['checking', 'timeout', 'pending_check'].includes(status);

  return {
    current: hasCurrent ? current : null,
    matched,
    label: hasCurrent
      ? (matched ? `Tracking OK ${formatPercent(current)}` : `Aguardando tracking ${formatPercent(current)}/${formatPercent(progress)}`)
      : (checking ? 'Validando Tracking...' : 'Conferir Tracking'),
    className: hasCurrent
      ? (matched ? 'stage-badge--tracking-ok' : 'stage-badge--tracking-waiting')
      : (checking ? 'stage-badge--tracking-waiting' : 'stage-badge--tracking-missing'),
  };
}

function stageTrackingBadgeHtml(item) {
  const info = getStageTrackingInfo(item);
  return `<span class="stage-badge ${info.className}">${escapeHtml(info.label)}</span>`;
}

function getPendingStageUpdate(projectRowId, spoolIso, sector = getStageWorkspaceSector()) {
  return getStageUpdatesForCurrentSector().find((item) =>
    isPendingStageStatus(item.status)
    && Number(item.projectRowId || 0) === Number(projectRowId || 0)
    && String(item.spoolIso || '').trim().toLowerCase() === String(spoolIso || '').trim().toLowerCase()
  ) || null;
}

function getLatestResolvedStageUpdate(projectRowId, spoolIso, sector = getStageWorkspaceSector()) {
  return getStageUpdatesForCurrentSector().filter((item) =>
    isResolvedStageStatus(item.status)
    && Number(item.projectRowId || 0) === Number(projectRowId || 0)
    && String(item.spoolIso || '').trim().toLowerCase() === String(spoolIso || '').trim().toLowerCase()
  ).sort((a,b)=> new Date(b.resolvedAt || b.createdAt || 0) - new Date(a.resolvedAt || a.createdAt || 0))[0] || null;
}

function getMyStageUpdates() {
  const username = String(state.user?.username || '').trim().toLowerCase();
  return (Array.isArray(state.stageUpdates) ? state.stageUpdates : [])
    .filter((item) => String(item.createdBy || '').trim().toLowerCase() === username)
    .sort((a,b)=> new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
}

function formatStageDate(value) {
  if (!value) return '—';
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toLocaleString('pt-BR');
  return escapeHtml(String(value));
}

function renderProjectSignals(project) {
  const signals = getProjectSignals(project);
  const actionButton = canCreateProjectSignal(project)
    ? `<button class="primary-button" type="button" data-open-project-signal="${escapeHtml(project.rowId)}">Nova sinalização ao PCP</button>`
    : '';
  const itemsHtml = signals.length
    ? signals.map((alert) => {
        const resolved = getSignalResolutionInfo(alert.id);
        return `
          <article class="project-signal-item ${resolved ? 'project-signal-item--resolved' : ''}">
            <div class="admin-list-item-meta">
              ${getSignalStatusBadge(alert)}
              <span>${escapeHtml(new Date(alert.createdAt).toLocaleString('pt-BR'))}</span>
              <span>Aberta por: ${escapeHtml(alert.createdBy || 'Usuário')}</span>
            </div>
            <strong>${escapeHtml(alert.title || 'Sinalização')}</strong>
            <p>${escapeHtml(alert.message || '').replace(/\n/g, '<br>')}</p>
            <div class="manual-alert-actions">
              ${resolved
                ? `<span class="manual-alert-tag manual-alert-tag--resolved-by">Resolvida por: ${escapeHtml(resolved.username)}</span>${resolved.date ? `<span class="manual-alert-tag">${escapeHtml(new Date(resolved.date).toLocaleString('pt-BR'))}</span>` : ''}`
                : `${canResolveSignal() ? `<button class="ghost-button" type="button" data-resolve-signal="${escapeHtml(alert.id)}">Marcar como resolvida</button>` : ''}`}
            </div>
            ${resolved && resolved.note ? `<div class="response-thread"><div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(resolved.note)}</p></div></div>` : ''}
          </article>
        `;
      }).join('')
    : '<div class="empty-inline">Nenhuma sinalização registrada para esta BSP.</div>';
  return `
    <section class="project-signals-section">
      <div class="project-signals-head">
        <div>
          <span class="manual-alert-tag">Sinalizações</span>
          <strong>Sinalizações do projeto</strong>
        </div>
        ${actionButton}
      </div>
      <div class="project-signals-list">${itemsHtml}</div>
    </section>
  `;
}

function uiStateLabel(stateValue) {
  if (stateValue === "completed") return "Finalizado";
  if (stateValue === "awaiting_shipment") return "Aguardando envio";
  if (stateValue === "preparing_shipment") return "Preparando para envio";
  if (stateValue === "in_progress") return "Em produção";
  return "Não iniciado";
}

function translateProjectStatus(projectStatus, uiState) {
  if (uiState === "completed") return "Finalizado";
  if (uiState === "awaiting_shipment") return "Aguardando envio";
  if (uiState === "preparing_shipment") return "Preparando para envio";
  if (uiState === "not_started") return "Não iniciado";

  const normalized = String(projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (["ONGOING", "ON GOING", "IN PROGRESS", "EM PRODUCAO", "EM PRODUÇÃO"].includes(normalized)) {
    return "Em produção";
  }
  if (["PREPARING SHIPMENT", "PREPARING FOR SHIPMENT", "PREPARANDO PARA ENVIO"].includes(normalized)) {
    return "Preparando para envio";
  }
  if (["ON HOLD", "HOLD", "PAUSED", "EM ESPERA"].includes(normalized)) {
    return uiState === "not_started" ? "Em espera" : "Em produção";
  }
  if (["COMPLETED", "DONE", "FINISHED", "CONCLUIDO", "CONCLUÍDO", "FINALIZADO"].includes(normalized)) {
    return "Finalizado";
  }
  return projectStatus || uiStateLabel(uiState);
}

function hasPreparingShipmentWindow(project) {
  const projectCoating = Number(project?.stageValues?.['Surface preparation and/or coating'] ?? NaN);
  const projectFinalInspection = Number(project?.stageValues?.['Final Inspection'] ?? NaN);
  const projectPackageDelivered = Number(project?.stageValues?.['Package and Delivered'] ?? project?.stageValues?.['Unitização e envio'] ?? NaN);
  const projectMatches = Number.isFinite(projectCoating)
    && projectCoating >= 100
    && Number.isFinite(projectFinalInspection)
    && projectFinalInspection >= 25
    && projectFinalInspection < 100
    && (!Number.isFinite(projectPackageDelivered) || projectPackageDelivered < 100);

  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const spoolMatches = spools.some((spool) => {
    const coating = Number(spool?.stageValues?.['Surface preparation and/or coating'] ?? NaN);
    const finalInspection = Number(spool?.stageValues?.['Final Inspection'] ?? NaN);
    const packageDelivered = Number(spool?.stageValues?.['Package and Delivered'] ?? spool?.stageValues?.['Unitização e envio'] ?? NaN);
    return Number.isFinite(coating)
      && coating >= 100
      && Number.isFinite(finalInspection)
      && finalInspection >= 25
      && finalInspection < 100
      && (!Number.isFinite(packageDelivered) || packageDelivered < 100);
  });

  return spoolMatches || projectMatches;
}

function getLogisticsProgressSnapshot(source) {
  const stageValues = source?.stageValues || {};
  const coating = Number(stageValues['Surface preparation and/or coating'] ?? NaN);
  const finalInspection = Number(stageValues['Final Inspection'] ?? NaN);
  const packageDelivered = Number(stageValues['Package and Delivered'] ?? stageValues['Unitização e envio'] ?? NaN);
  return {
    coating,
    finalInspection,
    packageDelivered,
    hasCoating: Number.isFinite(coating),
    hasFinalInspection: Number.isFinite(finalInspection),
    hasPackageDelivered: Number.isFinite(packageDelivered),
  };
}

function isUnitizationInTratativaSnapshot(snapshot) {
  return Boolean(snapshot?.hasCoating)
    && snapshot.coating >= 100
    && Boolean(snapshot?.hasFinalInspection)
    && snapshot.finalInspection >= 25
    && snapshot.finalInspection < 100;
}

function isAwaitingShipmentSnapshot(snapshot) {
  return Boolean(snapshot?.hasCoating)
    && snapshot.coating >= 100
    && Boolean(snapshot?.hasFinalInspection)
    && snapshot.finalInspection >= 100
    && Boolean(snapshot?.hasPackageDelivered)
    && snapshot.packageDelivered >= 25
    && snapshot.packageDelivered < 100;
}

function projectHasLogisticsWindow(project, predicate) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const openSpools = spools.filter((spool) => spool?.flow?.state !== 'completed' && spool?.flow?.status !== 'Finalizado');
  const sourceSpools = openSpools.length ? openSpools : spools;
  if (sourceSpools.length) {
    return sourceSpools.some((spool) => predicate(getLogisticsProgressSnapshot(spool)));
  }
  return predicate(getLogisticsProgressSnapshot(project));
}

function projectHasUnitizationInTratativa(project) {
  return projectHasLogisticsWindow(project, isUnitizationInTratativaSnapshot);
}

function projectHasAwaitingShipmentPackage(project) {
  return projectHasLogisticsWindow(project, isAwaitingShipmentSnapshot);
}

function isPreparedShipmentSpool(spool) {
  const snapshot = getLogisticsProgressSnapshot(spool);
  return snapshot.hasCoating
    && snapshot.coating >= 100
    && snapshot.hasFinalInspection
    && snapshot.finalInspection >= 100
    && (!snapshot.hasPackageDelivered || snapshot.packageDelivered < 25);
}

function isProjectPreparedForShipment(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const openSpools = spools.filter((spool) => spool?.flow?.state !== 'completed' && spool?.flow?.status !== 'Finalizado');
  if (openSpools.length) {
    return openSpools.every((spool) => isPreparedShipmentSpool(spool));
  }
  const projectCoating = Number(project?.stageValues?.['Surface preparation and/or coating'] ?? NaN);
  const projectFinalInspection = Number(project?.stageValues?.['Final Inspection'] ?? NaN);
  const projectPackageDelivered = Number(project?.stageValues?.['Package and Delivered'] ?? project?.stageValues?.['Unitização e envio'] ?? NaN);
  return Number.isFinite(projectCoating)
    && projectCoating >= 100
    && Number.isFinite(projectFinalInspection)
    && projectFinalInspection >= 100
    && (!Number.isFinite(projectPackageDelivered) || projectPackageDelivered < 25);
}

function getPreparedShipmentTags(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const openSpools = spools.filter((spool) => spool?.flow?.state !== 'completed' && spool?.flow?.status !== 'Finalizado');
  if (openSpools.length) {
    return openSpools.filter((spool) => isPreparedShipmentSpool(spool)).length;
  }
  return isProjectPreparedForShipment(project)
    ? Number(project?.quantitySpools || 1)
    : 0;
}

function isAwaitingShipmentFlowStatus(value) {
  return normalizeText(value || '').replace(/[^a-z0-9]+/g, '') === 'aguardandoenvio';
}

function getAwaitingShipmentTags(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const openSpools = spools.filter((spool) => spool?.flow?.state !== 'completed' && spool?.flow?.status !== 'Finalizado');
  const sourceSpools = openSpools.length ? openSpools : spools;

  if (sourceSpools.length) {
    return sourceSpools.filter((spool) => {
      const snapshot = getLogisticsProgressSnapshot(spool);
      return isAwaitingShipmentSnapshot(snapshot) || isAwaitingShipmentFlowStatus(spool?.flow?.status || spool?.currentStatus || spool?.stage);
    }).length;
  }

  return projectHasAwaitingShipmentPackage(project) || isAwaitingShipmentFlowStatus(project?.flow?.status || project?.currentStatus || project?.currentStage || project?.statusSummary)
    ? Number(project?.quantitySpools || 1)
    : 0;
}

function getProjectStatusPresentation(project) {
  // v37.71: no painel principal, a coluna Status deve mostrar o status real do Tracking/Smartsheet
  // (ex.: TH, Solda, Pintura, Aguardando END), não o estado genérico "Em produção".
  const trackingStatusText = project?.statusSummary || project?.currentStatus || project?.flow?.status || project?.currentStage || project?.stage || '';
  if (trackingStatusText) {
    const normalizedStatus = normalizeText(trackingStatusText);
    const state = project?.finished || project?.uiState === 'completed' || normalizedStatus.includes('finalizado')
      ? 'completed'
      : (project?.uiState === 'awaiting_shipment' ? 'preparing_shipment' : (project?.uiState || project?.operationalState || 'in_progress'));
    return { text: trackingStatusText, state };
  }

  if (projectHasTratativaObservation(project)) {
    return { text: 'Em tratativa', state: 'in_progress', reason: getProjectTratativaReason(project) };
  }

  if (projectHasAwaitingShipmentPackage(project) && project?.uiState !== 'completed') {
    return { text: 'Aguardando envio', state: 'awaiting_shipment' };
  }

  const preparedForShipment = isProjectPreparedForShipment(project);
  if (preparedForShipment && project?.uiState !== 'completed') {
    return { text: 'Preparado para envio', state: 'preparing_shipment' };
  }

  if (hasPreparingShipmentWindow(project) && !['completed', 'awaiting_shipment'].includes(project?.uiState)) {
    return { text: 'Preparado para envio', state: 'preparing_shipment' };
  }

  const state = ['awaiting_shipment', 'completed'].includes(project?.uiState) ? 'completed' : (project?.uiState || 'not_started');
  return {
    text: translateProjectStatus(project?.projectStatus, project?.uiState),
    state,
  };
}

function isProjectFinalizedForDisplay(project) {
  if (hasClientIncompleteProductionEvidence(project)) return false;

  const statusText = normalizeText([
    project?.statusSummary,
    project?.currentStatus,
    project?.currentStage,
    project?.projectStatus,
    project?.flow?.status,
  ].filter(Boolean).join(' '));
  return Boolean(
    project?.finished
    || project?.uiState === 'completed'
    || project?.operationalState === 'completed'
    || project?.flow?.state === 'completed'
    || statusText.includes('finalizado')
  );
}

function getProjectSectorSummary(project) {
  // v37.75: quando a BSP estiver sinalizada como ON HOLD no Smartsheet,
  // a Etapa Atual deve aparecer como On Hold, sem misturar com Produção/Qualidade/Pintura.
  if (isProjectOnHold(project)) return 'On Hold';
  if (isProjectFinalizedForDisplay(project)) return 'Enviado';
  return project?.sectorSummary || project?.currentStageGroup || project?.currentSector || project?.operationalSector || '';
}

function getProjectCurrentStageDisplay(project) {
  if (isProjectOnHold(project)) return 'On Hold';
  if (isProjectFinalizedForDisplay(project)) return 'Enviado';
  return project?.currentStageGroup || getProjectSectorSummary(project) || simplifyCurrentStage(project);
}

function getFlowSectorKey(flow = {}) {
  return normalizeSectorValue(flow.sector || flow.currentSector || flow.operationalSector || '');
}

function getProjectOpenFlowItems(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const source = spools.length
    ? spools.map((spool) => ({ flow: spool.flow || { status: spool.stage || spool.currentStatus, sector: spool.currentSector || spool.operationalSector, state: spool.operationalState || spool.uiState }, spool }))
    : [{ flow: project?.flow || { status: project?.currentStage || project?.statusSummary, sector: getProjectSectorSummary(project), state: project?.operationalState || project?.uiState }, spool: null }];
  return source.filter((item) => item.flow?.state !== 'completed' && item.flow?.status !== 'Finalizado');
}

function getProjectSectorKeys(project) {
  const items = getProjectOpenFlowItems(project);
  const keys = new Set();
  for (const item of items) {
    const key = getFlowSectorKey(item.flow);
    if (key) keys.add(key);
  }
  if (!keys.size && project?.finished) keys.add('pendente_envio');
  return keys;
}


function classifyStageSector(value) {
  const normalized = normalizeText(value || "");
  if (!normalized) return '';

  if (normalized.includes('final inspection') || normalized.includes('unitizacao') || normalized.includes('unitizacao e envio') || normalized.includes('package and delivered') || normalized.includes('envio')) {
    return 'pendente_envio';
  }
  if (normalized.includes('inspection') || normalized.includes('inspecao') || normalized.includes('dimensional') || normalized.includes('hydro test') || normalized === 'th') {
    return 'inspecao';
  }
  if (normalized.includes('paint') || normalized.includes('coating') || normalized.includes('surface preparation') || normalized.includes('hdg') || normalized.includes('fbe')) {
    return 'pintura';
  }
  if (normalized.includes('solda') || normalized.includes('weld')) {
    return 'solda';
  }
  if (normalized.includes('calderaria') || normalized.includes('fabrication') || normalized.includes('fit-up') || normalized.includes('montagem')) {
    return 'calderaria';
  }
  if (normalized.includes('production') || normalized.includes('producao') || normalized.includes('produção')) {
    return 'producao';
  }
  return '';
}

function simplifyCurrentStage(project) {
  const directSummary = getProjectSectorSummary(project);
  if (directSummary) return directSummary;
  const uiState = String(project?.uiState || "").trim().toLowerCase();
  const sector = normalizeText(project?.operationalSector || "");
  const stage = normalizeText(project?.currentStage || "");

  if (uiState === "completed" || stage.includes("finalizado")) {
    return "Enviado";
  }

  if (
    uiState === "awaiting_shipment" ||
    stage.includes("final inspection") ||
    stage.includes("unitizacao") ||
    stage.includes("unitizacao e envio") ||
    stage.includes("package and delivered") ||
    stage.includes("envio") ||
    sector.includes("logistica") || sector.includes("pendente de envio") ||
    sector.includes("envio")
  ) {
    return "Logística";
  }

  if (
    sector.includes("inspecao") ||
    stage.includes("inspection") ||
    stage.includes("inspecao") ||
    stage.includes("dimensional") ||
    stage.includes("hydro test") ||
    stage.includes("th")
  ) {
    return "Qualidade";
  }

  if (
    sector.includes("pintura") ||
    stage.includes("paint") ||
    stage.includes("coating") ||
    stage.includes("surface preparation") ||
    stage.includes("hdg") ||
    stage.includes("fbe")
  ) {
    return "Pintura";
  }

  return "Produção";
}

function stageStatusClass(status) {
  if (status === "completed") return "completed";
  if (status === "in_progress") return "in_progress";
  if (status === "waiting") return "waiting";
  return "ignored";
}

function setClock(targetTimeId, targetDateId, locale, timeZone) {
  const now = new Date();
  const timeText = new Intl.DateTimeFormat(locale, {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
    timeZone,
  }).format(now);

  const dateText = new Intl.DateTimeFormat(locale, {
    weekday: "short",
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone,
  }).format(now);

  document.getElementById(targetTimeId).textContent = timeText;
  document.getElementById(targetDateId).textContent = dateText;
}


function percentStateClass(value) {
  if (value == null || Number.isNaN(value)) return "";
  if (Number(value) >= 100) return "value-complete";
  if (Number(value) > 0) return "value-progress";
  return "";
}

function tableCellClass(value, type = "percent") {
  if (type !== "percent") return "";
  return percentStateClass(value);
}

function getCurrentDashboardLanguage() {
  try {
    const stored = localStorage.getItem('step-dashboard-language');
    if (stored === 'en-US' || stored === 'es-ES' || stored === 'pt-BR') return stored;
  } catch (_) {}
  return document.documentElement?.lang || 'pt-BR';
}

function startClocks() {
  const tick = () => {
    const lang = getCurrentDashboardLanguage();
    const brLocale = lang === 'pt-BR' ? 'pt-BR' : lang;
    const ptLocale = lang === 'pt-BR' ? 'pt-PT' : lang;
    setClock("clock-br-time", "clock-br-date", brLocale, "America/Sao_Paulo");
    setClock("clock-pt-time", "clock-pt-date", ptLocale, "Europe/Lisbon");
  };
  tick();
  window.addEventListener('step:language-change', tick);
  window.setInterval(tick, 1000);
}

function formatProjectTypeLabel(value) {
  const raw = String(value || "").trim();
  if (!raw) return "—";
  const normalized = normalizeText(raw);

  if (normalized.includes("support") || normalized === "sup" || normalized.includes("suporte")) {
    return "SUP";
  }
  if (normalized.includes("frame") || normalized.includes("structure") || normalized.includes("estrutura")) {
    return "Estrutura";
  }
  if (normalized.includes("spool")) {
    return "Spool";
  }
  return raw;
}

function getProjectTypeLabel(project) {
  return formatProjectTypeLabel(project?.projectType || project?.type || project?.project_type);
}

function compareProjectTypeLabels(a, b) {
  const order = new Map([
    ["spool", 1],
    ["estrutura", 2],
    ["sup", 3],
  ]);
  const na = normalizeText(a);
  const nb = normalizeText(b);
  const oa = order.get(na) || 99;
  const ob = order.get(nb) || 99;
  if (oa !== ob) return oa - ob;
  return String(a).localeCompare(String(b), "pt-BR", { numeric: true, sensitivity: "base" });
}

function enrichProjects(projects) {
  return (projects || []).map((project) => {
    const searchParts = [
      project.projectDisplay,
      project.projectNumber,
      project.projectPrefix,
      project.currentStage,
      project.projectStatus,
      project.projectType,
      getProjectTypeLabel(project),
      project.client,
      project.vessel,
      project.observations,
      project.customerPo,
      project.clientFocalPoint,
      project.clientFocalPointDisplay,
      ...(project.customerPoList || []),
      ...(project.clientFocalPointList || []),
      ...(project.spools || []).flatMap((spool) => [spool.iso, spool.description, spool.drawing, spool.observations]),
    ];

    return {
      ...project,
      currentStageGroup: simplifyCurrentStage(project),
      _searchText: buildSearchIndex(searchParts),
    };
  });
}

const PROJECT_STATUS_FILTER_OPTIONS = ["Aguardando envio", "Em produção", "Em tratativa", "Finalizado", "Não iniciado"];

function normalizeStatusFilterValue(value) {
  return normalizeText(String(value || '').trim());
}

function getSelectedStatusFilters() {
  const valid = new Set(PROJECT_STATUS_FILTER_OPTIONS.map((item) => normalizeStatusFilterValue(item)));
  return Array.from(new Set((Array.isArray(state.statusFilters) ? state.statusFilters : [])
    .map((item) => String(item || '').trim())
    .filter((item) => valid.has(normalizeStatusFilterValue(item)))));
}

function areAllStatusFiltersSelected() {
  const selected = getSelectedStatusFilters();
  return !selected.length || selected.length === PROJECT_STATUS_FILTER_OPTIONS.length;
}

function isStatusFilterSelected(option) {
  if (areAllStatusFiltersSelected()) return true;
  const normalizedOption = normalizeStatusFilterValue(option);
  return getSelectedStatusFilters().some((item) => normalizeStatusFilterValue(item) === normalizedOption);
}

function getStatusFilterButtonLabel() {
  const selected = getSelectedStatusFilters();
  if (!selected.length || selected.length === PROJECT_STATUS_FILTER_OPTIONS.length) return 'Todos os status';
  if (selected.length === 1) return selected[0];
  return `${selected.length} status selecionados`;
}

function syncStatusFilterButtonLabel() {
  if (!statusFilterToggleEl) return;
  statusFilterToggleEl.textContent = getStatusFilterButtonLabel();
}

function renderStatusFilterMenu() {
  if (!statusFilterMenuEl) return;
  const allChecked = areAllStatusFiltersSelected();
  statusFilterMenuEl.innerHTML = [
    `<label class="status-filter-option" data-status-filter-all="1"><input type="checkbox" data-status-filter-all="1" ${allChecked ? 'checked' : ''}><span>Todos os status</span></label>`,
    ...PROJECT_STATUS_FILTER_OPTIONS.map((option) => `<label class="status-filter-option" data-status-filter="${option}"><input type="checkbox" data-status-filter="${option}" ${isStatusFilterSelected(option) ? 'checked' : ''}><span>${option}</span></label>`),
  ].join('');
  syncStatusFilterButtonLabel();
}

function closeStatusFilterMenu() {
  if (!statusFilterMenuEl || !statusFilterToggleEl) return;
  statusFilterMenuEl.classList.add('hidden');
  statusFilterToggleEl.classList.remove('is-open');
  statusFilterToggleEl.setAttribute('aria-expanded', 'false');
}

function openStatusFilterMenu() {
  if (!statusFilterMenuEl || !statusFilterToggleEl) return;
  renderStatusFilterMenu();
  statusFilterMenuEl.classList.remove('hidden');
  statusFilterToggleEl.classList.add('is-open');
  statusFilterToggleEl.setAttribute('aria-expanded', 'true');
}

function toggleStatusFilterMenu() {
  if (!statusFilterMenuEl) return;
  if (statusFilterMenuEl.classList.contains('hidden')) openStatusFilterMenu();
  else closeStatusFilterMenu();
}

function getProjectStatusFilterLabel(project) {
  if (projectHasTratativaObservation(project)) return 'Em tratativa';
  const presentationText = normalizeText(getProjectStatusPresentation(project)?.text || '');
  const projectStatusText = normalizeText(project?.projectStatus || '');
  const currentStageText = normalizeText(project?.currentStage || '');
  const uiState = String(project?.uiState || '').trim();

  if (uiState === 'completed' || presentationText.includes('finalizado')) {
    return 'Finalizado';
  }
  if (projectHasAwaitingShipmentPackage(project) || uiState === 'awaiting_shipment' || presentationText.includes('aguardando envio')) {
    return 'Aguardando envio';
  }
  if (projectHasUnitizationInTratativa(project) || presentationText.includes('tratativa') || projectStatusText.includes('tratativa') || currentStageText.includes('tratativa')) {
    return 'Em tratativa';
  }
  if (presentationText.includes('preparado para envio') || presentationText.includes('preparando para envio')) {
    return 'Em tratativa';
  }
  if (uiState === 'not_started' || presentationText.includes('nao iniciado') || presentationText.includes('não iniciado') || presentationText.includes('em espera')) {
    return 'Não iniciado';
  }
  return 'Em produção';
}

function projectMatchesStatusFilter(project) {
  const selected = getSelectedStatusFilters();
  if (!selected.length || selected.length === PROJECT_STATUS_FILTER_OPTIONS.length) return true;
  const label = getProjectStatusFilterLabel(project);
  const normalizedLabel = normalizeStatusFilterValue(label);
  return selected.some((item) => normalizeStatusFilterValue(item) === normalizedLabel);
}

function buildDemandOptions() {
  if (!demandFilterEl) return;
  const selected = state.demandFilter || "";
  const hiddenDemandOptions = new Set([
    normalizeText("Project Finished?"),
    normalizeText("Drawing Execution"),
    normalizeText("Emissão de detalhamento"),
  ]);

  const options = Array.from(
    new Set(
      state.projects
        .map((project) => getProjectCurrentStageDisplay(project))
        .filter(Boolean)
        .filter((option) => !hiddenDemandOptions.has(normalizeText(option)))
    )
  ).sort((a, b) => a.localeCompare(b, "pt-BR"));

  demandFilterEl.innerHTML = [
    '<option value="">Todas as demandas</option>',
    ...options.map((option) => `<option value="${option}">${option}</option>`),
  ].join("");

  demandFilterEl.value = options.includes(selected) ? selected : "";
  if (!options.includes(selected)) state.demandFilter = "";
}

function buildProjectTypeOptions() {
  if (!projectTypeFilterEl) return;
  const selected = state.projectTypeFilter || "";
  const options = Array.from(
    new Set(
      state.projects
        .map((project) => getProjectTypeLabel(project))
        .filter((option) => option && option !== "—")
    )
  ).sort(compareProjectTypeLabels);

  projectTypeFilterEl.innerHTML = [
    '<option value="">Todos os tipos</option>',
    ...options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`),
  ].join("");

  projectTypeFilterEl.value = options.includes(selected) ? selected : "";
  if (!options.includes(selected)) state.projectTypeFilter = "";
}

function buildWeekOptions() {
  if (!weekFilterEl) return;
  const selected = state.weekFilter || "";
  const currentWeek = getCurrentProductionWeekLabel();
  const weekLabels = Array.from(
    new Set([
      currentWeek,
      ...state.projects.flatMap((project) => {
        const spoolWeeks = (project.spools || []).map((spool) => spool.weldingWeek).filter(Boolean);
        if (spoolWeeks.length) return spoolWeeks;
        return project.weldingWeek ? [project.weldingWeek] : [];
      }),
    ])
  ).sort(compareWeekLabels);

  const options = ['<option value="">Todas as semanas</option>'];
  for (const label of weekLabels) {
    options.push(`<option value="${label}">${label}</option>`);
  }

  weekFilterEl.innerHTML = options.join("");
  weekFilterEl.value = weekLabels.includes(selected) ? selected : "";
  if (!weekLabels.includes(selected)) state.weekFilter = "";
}

function getActiveWeekLabel() {
  return state.weekFilter || "Todas as semanas";
}

function projectMatchesWeekFilter(project, weekLabel = state.weekFilter) {
  if (!weekLabel) return true;
  const normalizedWeek = String(weekLabel || '').trim();
  if (!normalizedWeek) return true;
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  if (spools.length) {
    return spools.some((spool) => String(spool?.weldingWeek || '').trim() === normalizedWeek);
  }
  return String(project?.weldingWeek || '').trim() === normalizedWeek;
}

function hasActiveProjectTableFilters() {
  const hasQuery = Boolean(normalizeText(state.searchQuery).trim());
  const hasDemand = Boolean(normalizeText(state.demandFilter).trim());
  const hasType = Boolean(normalizeText(state.projectTypeFilter).trim());
  const hasWeek = Boolean(String(state.weekFilter || '').trim());
  const hasStatus = !areAllStatusFiltersSelected();
  return hasQuery || hasDemand || hasType || hasWeek || hasStatus;
}

function getStatsProjectsSource() {
  // v37.68: os cards devem usar a mesma base real exibida na tabela/Excel.
  // A lista state.filteredProjects já vem da base do Smartsheet após aplicar busca/filtros.
  // Antes, quando não havia filtro ativo, os cards voltavam para outra fonte e o Total ficava
  // menor que o botão Baixar Excel, mesmo com 136 projetos raiz no Tracking.
  const filtered = Array.isArray(state.filteredProjects) ? state.filteredProjects : [];
  const source = filtered.length ? filtered : getVisibleProjectsSource();
  return source.filter((project) => projectMatchesWeekFilter(project));
}

function isProjectStatusOnHold(projectStatus) {
  const normalized = normalizeText(projectStatus || "");
  const compact = normalized.replace(/[^a-z0-9]+/g, "");
  return compact === "onhold"
    || compact === "hold"
    || compact === "pausado"
    || compact === "paused"
    || compact === "emespera"
    || normalized.includes("on hold")
    || normalized.includes("em hold")
    || normalized.includes("projeto em hold")
    || normalized.includes("hold conforme")
    || normalized.includes("hold")
    || normalized.includes("em espera")
    || normalized.includes("pausado")
    || normalized.includes("paused")
    || normalized.includes("paralisado")
    || normalized.includes("suspenso");
}

function isProjectStatusPending(projectStatus) {
  const normalized = normalizeText(projectStatus || "");
  const compact = normalized.replace(/[^a-z0-9]+/g, "");
  return compact === "pending" || normalized === "pending";
}

function getProjectStatusTexts(project) {
  if (!project) return [];
  const texts = [
    project.projectStatus,
    project["PROJECT STATUS"],
    project.status,
    project.currentStatus,
    project.currentStage,
    project.statusSummary,
    project.sectorSummary,
    project.operationalState,
    project.uiState,
    project.flow?.status,
    project.flow?.state,
  ];

  if (project.stageValues && typeof project.stageValues === 'object') {
    texts.push(project.stageValues['Project Finished?'], project.stageValues['Project Finish Date']);
  }

  if (Array.isArray(project.spools)) {
    project.spools.forEach((spool) => {
      texts.push(
        spool?.projectStatus,
        spool?.["PROJECT STATUS"],
        spool?.status,
        spool?.currentStatus,
        spool?.stage,
        spool?.uiState,
        spool?.operationalState,
        spool?.flow?.status,
        spool?.flow?.state
      );
      if (spool?.stageValues && typeof spool.stageValues === 'object') {
        texts.push(spool.stageValues['Project Finished?'], spool.stageValues['Project Finish Date']);
      }
    });
  }

  return texts.filter((value) => value != null && String(value).trim() !== '');
}

function isProjectPending(project) {
  return getProjectStatusTexts(project).some((value) => isProjectStatusPending(value));
}

function isMeaningfulFinishValue(value) {
  if (value == null) return false;
  const raw = String(value).trim();
  if (!raw) return false;
  const compact = normalizeText(raw).replace(/[^a-z0-9]+/g, '');
  return !['na', 'n/a', 'none', 'null', 'false', 'no', 'nao', 'não', '0'].includes(compact);
}

function isProjectStatusFinished(projectStatus) {
  const normalized = normalizeText(projectStatus || "");
  const compact = normalized.replace(/[^a-z0-9]+/g, "");
  return ['finalizado', 'concluido', 'finished', 'completed', 'delivered', 'entregue', 'enviado'].includes(compact)
    || normalized === 'finalizado'
    || normalized === 'concluido'
    || normalized.includes('project finished')
    || normalized.includes('projeto finalizado');
}

function hasProjectFinishDateMarker(project) {
  if (!project) return false;
  const values = [
    project.projectFinishDate,
    project.finishDate,
    project.finishedDate,
    project.shipmentDate,
    project.stageValues?.['Project Finish Date'],
    project.stageValues?.['PROJECT FINISH DATE'],
  ];
  return values.some(isMeaningfulFinishValue);
}

function hasProjectFinishedBooleanMarker(project) {
  if (!project) return false;
  const values = [project.finished, project.projectFinishedFlag, project.stageValues?.['Project Finished?'], project.stageValues?.['PROJECT FINISHED?']];
  return values.some((value) => {
    if (typeof value === 'boolean') return value;
    const compact = normalizeText(value || '').replace(/[^a-z0-9]+/g, '');
    return ['true', 'yes', 'sim', 'y', '1', 'finalizado', 'concluido', 'completed', 'finished'].includes(compact);
  });
}

function areAllProjectSpoolsFinished(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  return spools.length > 0 && spools.every((spool) => !hasClientIncompleteProductionEvidence(spool) && Boolean(
    spool?.finished
    || spool?.projectFinishedFlag
    || spool?.uiState === 'completed'
    || spool?.operationalState === 'completed'
    || spool?.flow?.state === 'completed'
    || spool?.flow?.status === 'Finalizado'
    || isProjectStatusFinished(spool?.projectStatus)
    || isProjectStatusFinished(spool?.status)
    || isProjectStatusFinished(spool?.currentStatus)
    || isMeaningfulFinishValue(spool?.stageValues?.['Project Finish Date'])
  ));
}

function isProjectFinishedForTotal(project) {
  if (!project) return false;
  if (hasClientIncompleteProductionEvidence(project)) return false;
  return Boolean(
    hasProjectFinishedBooleanMarker(project)
    || hasProjectFinishDateMarker(project)
    || project.uiState === 'completed'
    || project.operationalState === 'completed'
    || project.flow?.state === 'completed'
    || isProjectStatusFinished(project.projectStatus)
    || isProjectStatusFinished(project.status)
    || isProjectStatusFinished(project.currentStage)
    || isProjectStatusFinished(project.currentStatus)
    || isProjectStatusFinished(project.statusSummary)
    || isProjectStatusFinished(project.flow?.status)
    || areAllProjectSpoolsFinished(project)
  );
}

function isProjectExcludedFromTotal(project) {
  return isProjectOnHold(project) || isProjectPending(project) || isProjectFinishedForTotal(project);
}

function getProjectRootJobProcessStatus(project) {
  const candidates = [
    project?.jobProcessStatus,
    project?.stageValues?.['Job Process Status'],
    project?.stageValues?.['JOB PROCESS STATUS'],
    project?.currentStage,
  ];
  const value = candidates.find((item) => item != null && String(item).trim() !== '');
  return value == null ? '' : String(value).trim();
}

function isProjectRootJobStatusDelivered(project) {
  const normalized = normalizeText(getProjectRootJobProcessStatus(project));
  return normalized.includes('package and delivered') || normalized.includes('project finished');
}

function isProjectRootJobStatusNotStarted(project) {
  const normalized = normalizeText(getProjectRootJobProcessStatus(project));
  return normalized.includes('fabrication not started');
}

function normalizeFlowSortText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function getFlowSortWeight(flow = {}) {
  const status = normalizeFlowSortText(flow?.status || '');
  const sector = normalizeFlowSortText(flow?.sector || '');

  if (status === 'finalizado' || status.includes('project finished')) return 999;
  if (status.includes('package and delivered') || status === 'delivered' || status === 'enviado') return 130;
  if (status.includes('ag emissao de detalhamento') || status === 'emissao de detalhamento') return 10;
  if (status.includes('verificando estoque') || status.includes('aguardando material')) return 20;
  if (status.includes('separacao de material') || status.includes('material separation')) return 30;
  if (status.includes('corte e limpeza') || status.includes('fabrication start')) return 40;
  if (status.includes('pre montagem') || status.includes('welding preparation')) return 50;
  if (status.includes('inspecao dimensional de ajuste') || status.includes('dma 3d')) return 60;
  if (status === 'solda' || status.includes('full welding')) return 70;
  if (status.includes('inspecao dimensional final') || status.includes('inspection qc')) return 80;
  if (status.includes('aguardando end')) return 90;
  if (status === 'th') return 100;
  if (status === 'pintura') return 110;
  if (status.includes('aguardando inicio de pintura')) return 111;
  if (status === 'j f' || status === 'jf') return 112;
  if (status.includes('intermediaria')) return 113;
  if (status.includes('acabamento')) return 114;
  if (status === 'concluido') return 115;
  if (status.includes('unitizacao e inspecao') || status.includes('unitizacao')) return 120;
  if (status.includes('preparado para envio') || status.includes('preparando para envio') || status.includes('aguardando envio')) return 130;

  if (sector === 'engenharia') return 10;
  if (sector === 'suprimento') return 20;
  if (sector === 'producao') return 40;
  if (sector === 'qualidade') return 80;
  if (sector === 'pintura') return 110;
  if (sector === 'logistica') return 120;
  return 500;
}

function getProjectCardBucketFromFlow(flow = {}) {
  const sectorKey = getFlowSectorKey(flow);
  const weight = getFlowSortWeight(flow);

  if (sectorKey === 'pendente_envio' || sectorKey === 'enviado' || (weight >= 120 && weight < 500)) return 'awaiting';
  if (sectorKey === 'pintura' || (weight >= 110 && weight < 500)) return 'painting';
  if (sectorKey === 'inspecao' || (weight >= 80 && weight < 500)) return 'inspection';
  if (['producao', 'solda', 'calderaria'].includes(sectorKey) || (weight >= 40 && weight < 500)) return 'production';
  return 'not_started';
}

function getProjectDelayedStageStats(project) {
  const openItems = getProjectOpenFlowItems(project);
  const fallbackFlow = project?.flow || {
    status: getProjectRootJobProcessStatus(project),
    sector: getProjectSectorSummary(project) || project?.operationalSector || project?.currentStageGroup || project?.currentSector,
    state: project?.operationalState || project?.uiState,
  };
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
  // A Etapa Atual segue mostrando todas as etapas; os cards contam a menor etapa operacional.
  // Ex.: Pintura + Solda na mesma BSP => conta em Produção/Solda.
  if (!project) return 'unknown';
  if (isProjectOnHold(project)) return 'hold';
  if (isProjectPending(project)) return 'pending';
  if (isProjectFinishedForTotal(project)) return 'finished';
  const delayed = getProjectDelayedStageStats(project);
  if (delayed.bucket === 'not_started') return 'not_started';
  return 'started';
}

function isProjectStartedForStats(project) {
  if (!project) return { started: false, tags: 0 };
  const bucket = getProjectExclusiveCardBucket(project);
  if (bucket !== 'started') return { started: false, tags: 0 };
  return { started: true, tags: getProjectItemCountForSummary(project) };
}


function getProjectItemCountForSummary(project) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  return Number(project?.quantitySpools || spools.length || 1);
}

function getProjectPreStartTagsForSummary(project) {
  const openItems = getProjectOpenFlowItems(project);
  return openItems.filter((item) => {
    const sectorKey = getFlowSectorKey(item.flow);
    return sectorKey === 'engenharia' || sectorKey === 'suprimento';
  }).length;
}

function isProjectNotStartedForSummary(project) {
  return getProjectExclusiveCardBucket(project) === 'not_started';
}

function getProjectNotStartedTagsForSummary(project) {
  if (!isProjectNotStartedForSummary(project)) return 0;
  return getProjectItemCountForSummary(project);
}

function getProjectStartedTagsForSummary(project) {
  return getProjectExclusiveCardBucket(project) === 'started' ? getProjectItemCountForSummary(project) : 0;
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

  if (project.stageValues && typeof project.stageValues === 'object') {
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
      if (spool?.stageValues && typeof spool.stageValues === 'object') {
        texts.push(...Object.values(spool.stageValues));
      }
    });
  }

  return texts.filter((value) => value != null && String(value).trim() !== '');
}

function isProjectOnHold(project) {
  return getProjectHoldContextTexts(project).some((value) => isProjectStatusOnHold(value));
}

function buildClientStats(projects) {
  const visibleProjects = Array.isArray(projects) ? projects : [];
  const activeProjects = visibleProjects.filter((project) => !isProjectExcludedFromTotal(project));
  const stats = {
    // Total de Projetos deve bater com a mesma base exibida na tabela e no botão Baixar Excel.
    // Os cards abaixo são visões operacionais sobre essa base; produção/qualidade/pintura/envio
    // são subgrupos de status e podem se sobrepor ao card de iniciados.
    totalProjects: visibleProjects.length,
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
  for (const project of visibleProjects) {
    const spools = Array.isArray(project.spools) ? project.spools : [];
    const tags = Number(project.quantitySpools || spools.length || 0);
    const isFinishedProject = isProjectFinishedForTotal(project);
    stats.totalSpools += tags;
    stats.totalWeightKg += Number(project.kilos || 0);
    stats.totalWeldedWeightKg += Number(project.weldedWeightKg || 0);
    const openPaintingM2 = spools.length
      ? spools.filter((spool) => spool.flow?.state !== 'completed' && spool.flow?.status !== 'Finalizado').reduce((total, spool) => total + Number(spool.m2Painting || 0), 0)
      : 0;
    stats.totalPaintingM2 += isFinishedProject ? 0 : (openPaintingM2 > 0 ? openPaintingM2 : Number(project.m2Painting || 0));
    const cardBucket = getProjectExclusiveCardBucket(project);

    if (cardBucket === 'hold') {
      stats.notStartedHold += 1;
      stats.notStartedHoldTags += tags;
      continue;
    }

    if (cardBucket === 'pending') {
      continue;
    }

    if (cardBucket === 'finished') {
      stats.completed += 1;
      stats.completedTags += tags;
      continue;
    }

    if (cardBucket === 'not_started') {
      stats.notStarted += 1;
      stats.notStartedTags += tags;
      continue;
    }

    stats.startedProjects += 1;
    stats.startedTags += tags;

    progressAccumulator += Number(project.overallProgress || 0);

    const delayedStage = getProjectDelayedStageStats(project);
    const delayedTags = delayedStage.tagCount || tags || 1;
    if (delayedStage.bucket === 'production') {
      stats.inProgress += 1;
      stats.inProgressTags += delayedTags;
    }
    if (delayedStage.bucket === 'inspection') {
      stats.inspectionProjects += 1;
      stats.inspectionTags += delayedTags;
    }
    if (delayedStage.bucket === 'painting') {
      stats.paintingProjects += 1;
      stats.paintingTags += delayedTags;
    }
    if (delayedStage.bucket === 'awaiting') {
      stats.awaitingShipment += 1;
      stats.awaitingShipmentTags += delayedTags;
    }
  }

  stats.averageOverallProgress = activeProjects.length ? progressAccumulator / activeProjects.length : 0;
  return stats;
}

function getTotalWeldedWeightAllProjects() {
  return getStatsProjectsSource().reduce((total, project) => {
    const spools = project.spools || [];
    if (spools.length) {
      return total + spools.reduce((spoolTotal, spool) => spoolTotal + (spool.weldedWeightKg || 0), 0);
    }

    return total + (project.weldedWeightKg || 0);
  }, 0);
}

function getTotalFinishedWeightAllProjects() {
  return getStatsProjectsSource().reduce((total, project) => {
    if (!project?.finished) return total;
    return total + Number(project?.kilos || 0);
  }, 0);
}

function getWeldedWeightForWeek(weekLabel) {
  if (!weekLabel || weekLabel === "Todas as semanas") return getTotalWeldedWeightAllProjects();
  return getStatsProjectsSource().reduce((total, project) => {
    const spools = project.spools || [];
    if (spools.length) {
      return total + spools.reduce((spoolTotal, spool) => {
        if (spool.weldingWeek !== weekLabel) return spoolTotal;
        return spoolTotal + (spool.weldedWeightKg || 0);
      }, 0);
    }

    if (project.weldingWeek !== weekLabel) return total;
    return total + (project.weldedWeightKg || 0);
  }, 0);
}

function userHasProjectsScope(user = state.user) {
  if (!user || user.role === "admin") return false;
  return getUserAlertSectors(user).includes("projetos") || normalizeSectorValue(user.sector) === "projetos";
}

function formatQualityCompetencies(values = []) {
  const normalized = normalizeQualityCompetencies(values);
  return normalized.length ? normalized.map(qualityCompetencyLabel).join(', ') : 'Todas as competências da Qualidade';
}

function shouldUseSectorScopedToggle(user = state.user) {
  if (!user || user.role === 'admin') return false;
  const primarySector = getPrimaryUserSector(user);
  return Boolean(primarySector) && primarySector !== 'projetos';
}

function getPrimaryUserSector(user = state.user) {
  if (!user) return '';
  const linkedSectors = getUserAlertSectors(user);
  if (linkedSectors.length) {
    const firstOperational = linkedSectors.find((sector) => sector !== 'pcp' && sector !== 'projetos');
    if (firstOperational) return firstOperational;
    return linkedSectors[0];
  }
  return normalizeSectorValue(user.sector);
}

function isSectorScopedViewActive(user = state.user) {
  return Boolean(user) && shouldUseSectorScopedToggle(user) && state.sectorScopedView;
}

function getScopedDemandLabelsForUser(user = state.user) {
  const sector = getPrimaryUserSector(user);
  if (sector === 'pendente_envio') return ['Logística', 'Pendente de envio'];
  if (sector === 'inspecao') return ['Qualidade', 'Inspeção'];
  if (sector === 'pintura') return ['Pintura'];
  if (sector === 'solda') return ['Solda'];
  if (sector === 'calderaria') return ['Calderaria'];
  if (sector === 'engenharia') return ['Engenharia'];
  if (sector === 'suprimento') return ['Suprimento'];
  if (sector === 'producao') return ['Produção'];
  return [];
}

function getProjectSectorForScopedView(project) {
  const operationalSector = normalizeSectorValue(project?.operationalSector || '');
  const currentStageSector = normalizeSectorValue(classifyStageSector(project?.currentStage || ''));
  const jobProcessSector = normalizeSectorValue(classifyStageSector(project?.jobProcessStatus || ''));
  const currentGroup = normalizeSectorValue(project?.currentStageGroup || simplifyCurrentStage(project));
  const operationalState = normalizeSectorValue(project?.operationalState || project?.uiState || '');
  const weldingProgress = Number(project?.stageValues?.['Full welding execution'] ?? project?.stageValues?.['SOLDA'] ?? NaN);
  const hasWeldingProgress = Number.isFinite(weldingProgress);
  const projectInsideWeldingWindow = hasWeldingProgress && weldingProgress >= 25 && weldingProgress < 100;
  const isWeldingCompleted = hasWeldingProgress && weldingProgress >= 100;
  const coatingProgress = Number(project?.stageValues?.['Surface preparation and/or coating'] ?? NaN);
  const finalInspectionProgress = Number(project?.stageValues?.['Final Inspection'] ?? NaN);
  const packageDeliveredProgress = Number(project?.stageValues?.['Package and Delivered'] ?? project?.stageValues?.['Unitização e envio'] ?? NaN);
  const hasCoatingProgress = Number.isFinite(coatingProgress);
  const hasFinalInspectionProgress = Number.isFinite(finalInspectionProgress);
  const hasPackageDeliveredProgress = Number.isFinite(packageDeliveredProgress);
  const projectInsideLogisticsWindow = hasCoatingProgress && coatingProgress >= 100 && hasFinalInspectionProgress && finalInspectionProgress >= 100 && hasPackageDeliveredProgress && packageDeliveredProgress >= 25 && packageDeliveredProgress < 100;
  const isLogisticsCompleted = hasCoatingProgress && coatingProgress >= 100 && hasPackageDeliveredProgress && packageDeliveredProgress >= 100;
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const spoolWeldingProgressValues = spools
    .map((spool) => Number(spool?.stageValues?.['Full welding execution'] ?? spool?.stageValues?.['SOLDA'] ?? NaN))
    .filter((value) => Number.isFinite(value));
  const spoolLogisticsPairs = spools
    .map((spool) => {
      const coating = Number(spool?.stageValues?.['Surface preparation and/or coating'] ?? NaN);
      const finalInspection = Number(spool?.stageValues?.['Final Inspection'] ?? NaN);
      const packageDelivered = Number(spool?.stageValues?.['Package and Delivered'] ?? spool?.stageValues?.['Unitização e envio'] ?? NaN);
      return {
        coating,
        finalInspection,
        packageDelivered,
        hasCoating: Number.isFinite(coating),
        hasFinalInspection: Number.isFinite(finalInspection),
        hasPackageDelivered: Number.isFinite(packageDelivered),
      };
    })
    .filter((pair) => pair.hasCoating || pair.hasFinalInspection || pair.hasPackageDelivered);
  const hasSpoolWeldingProgress = spoolWeldingProgressValues.length > 0;
  const hasSpoolInsideWeldingWindow = spoolWeldingProgressValues.some((value) => value >= 25 && value < 100);
  const areAllSpoolsOutsideWeldingWindow = hasSpoolWeldingProgress && !hasSpoolInsideWeldingWindow;
  const isInsideWeldingWindow = hasSpoolInsideWeldingWindow || (!hasSpoolWeldingProgress && projectInsideWeldingWindow);
  const hasSpoolLogisticsProgress = spoolLogisticsPairs.length > 0;
  const hasSpoolInsideLogisticsWindow = spoolLogisticsPairs.some((pair) => pair.hasCoating && pair.coating >= 100 && pair.hasFinalInspection && pair.finalInspection >= 100 && pair.hasPackageDelivered && pair.packageDelivered >= 25 && pair.packageDelivered < 100);
  const areAllSpoolsOutsideLogisticsWindow = hasSpoolLogisticsProgress && !hasSpoolInsideLogisticsWindow;
  const isInsideLogisticsWindow = hasSpoolInsideLogisticsWindow || (!hasSpoolLogisticsProgress && projectInsideLogisticsWindow);

  if (currentGroup === 'pendente_envio' || operationalState === 'pendente_envio') {
    return 'pendente_envio';
  }
  if (currentGroup === 'inspecao') {
    return 'inspecao';
  }
  if (currentGroup === 'pintura') {
    return 'pintura';
  }

  if (isInsideLogisticsWindow) {
    return 'pendente_envio';
  }

  if (isInsideWeldingWindow) {
    return 'solda';
  }

  if (jobProcessSector === 'solda') {
    if (!hasWeldingProgress && !hasSpoolWeldingProgress) {
      return 'solda';
    }
  } else if (jobProcessSector === 'calderaria') {
    return 'calderaria';
  } else if (jobProcessSector === 'inspecao') {
    return 'inspecao';
  } else if (jobProcessSector === 'pintura') {
    return 'pintura';
  } else if (jobProcessSector === 'pendente_envio') {
    return 'pendente_envio';
  }

  if (currentStageSector === 'solda') {
    if (!hasWeldingProgress && !hasSpoolWeldingProgress) {
      return 'solda';
    }
  } else if (currentStageSector === 'calderaria') {
    return 'calderaria';
  } else if (currentStageSector === 'inspecao') {
    return 'inspecao';
  } else if (currentStageSector === 'pintura') {
    return 'pintura';
  } else if (currentStageSector === 'pendente_envio') {
    return 'pendente_envio';
  }

  if (operationalSector === 'pendente_envio') {
    if (!hasFinalInspectionProgress && !hasPackageDeliveredProgress && !hasSpoolLogisticsProgress) {
      return 'pendente_envio';
    }
    if (isInsideLogisticsWindow) {
      return 'pendente_envio';
    }
  }
  if (operationalSector === 'inspecao') {
    return 'inspecao';
  }
  if (operationalSector === 'pintura') {
    return 'pintura';
  }
  if (operationalSector === 'solda') {
    if (!hasWeldingProgress && !hasSpoolWeldingProgress) {
      return 'solda';
    }
    if (isInsideWeldingWindow) {
      return 'solda';
    }
  }
  if (operationalSector === 'calderaria') {
    return 'calderaria';
  }
  if (operationalSector === 'producao') {
    return isWeldingCompleted ? 'producao' : 'producao';
  }

  if (hasPackageDeliveredProgress && hasFinalInspectionProgress) {
    if (finalInspectionProgress < 100 || packageDeliveredProgress < 25) {
      if (currentGroup === 'pendente_envio') {
        return 'pintura';
      }
    }
    if (isLogisticsCompleted) {
      return currentGroup && currentGroup !== 'pendente_envio' ? currentGroup : (operationalSector && operationalSector !== 'pendente_envio' ? operationalSector : (jobProcessSector && jobProcessSector !== 'pendente_envio' ? jobProcessSector : (currentStageSector && currentStageSector !== 'pendente_envio' ? currentStageSector : 'pintura')));
    }
  }

  if (hasSpoolLogisticsProgress && areAllSpoolsOutsideLogisticsWindow) {
    if (spoolLogisticsPairs.every((pair) => (pair.hasPackageDelivered && pair.packageDelivered >= 100) || (!pair.hasPackageDelivered && pair.hasFinalInspection && pair.finalInspection >= 100))) {
      return currentGroup && currentGroup !== 'pendente_envio' ? currentGroup : (operationalSector && operationalSector !== 'pendente_envio' ? operationalSector : (jobProcessSector && jobProcessSector !== 'pendente_envio' ? jobProcessSector : (currentStageSector && currentStageSector !== 'pendente_envio' ? currentStageSector : 'pintura')));
    }
    if (spoolLogisticsPairs.every((pair) => !pair.hasFinalInspection || pair.finalInspection < 100 || !pair.hasPackageDelivered || pair.packageDelivered < 25)) {
      if (currentGroup === 'pendente_envio') {
        return 'pintura';
      }
    }
  }

  if (hasWeldingProgress && weldingProgress < 25) {
    return currentGroup === 'solda' ? 'producao' : (currentGroup || jobProcessSector || currentStageSector || operationalSector || 'producao');
  }

  if (areAllSpoolsOutsideWeldingWindow) {
    if (spoolWeldingProgressValues.every((value) => value >= 100)) {
      return currentGroup && currentGroup !== 'solda' ? currentGroup : (operationalSector && operationalSector !== 'solda' ? operationalSector : (jobProcessSector && jobProcessSector !== 'solda' ? jobProcessSector : (currentStageSector && currentStageSector !== 'solda' ? currentStageSector : 'producao')));
    }
    if (spoolWeldingProgressValues.every((value) => value < 25)) {
      return currentGroup === 'solda' ? 'producao' : (currentGroup || jobProcessSector || currentStageSector || operationalSector || 'producao');
    }
  }

  if (isWeldingCompleted) {
    return currentGroup && currentGroup !== 'solda' ? currentGroup : (operationalSector && operationalSector !== 'solda' ? operationalSector : (jobProcessSector && jobProcessSector !== 'solda' ? jobProcessSector : (currentStageSector && currentStageSector !== 'solda' ? currentStageSector : 'producao')));
  }

  return currentGroup || jobProcessSector || currentStageSector || operationalSector || 'all';
}

function projectMatchesScopedSector(project, user = state.user) {
  const sector = getPrimaryUserSector(user);
  if (!sector) return true;

  const sectorKeys = getProjectSectorKeys(project);
  const hasAny = (...keys) => keys.some((key) => sectorKeys.has(key));

  if (sector === 'pendente_envio') return hasAny('pendente_envio');
  if (sector === 'inspecao') return hasAny('inspecao');
  if (sector === 'pintura') return hasAny('pintura');
  if (sector === 'solda') return hasAny('solda', 'producao');
  if (sector === 'calderaria') return hasAny('calderaria', 'producao');
  if (sector === 'producao') return hasAny('producao', 'solda', 'calderaria');
  if (sector === 'engenharia') return hasAny('engenharia');
  if (sector === 'suprimento') return hasAny('suprimento');

  const labels = getScopedDemandLabelsForUser(user).map((item) => normalizeText(item).trim()).filter(Boolean);
  if (!labels.length) return true;
  const currentGroup = normalizeText(project?.currentStageGroup || simplifyCurrentStage(project)).trim();
  return labels.some((label) => currentGroup.includes(label));
}

function alertMatchesScopedSector(alert, user = state.user) {
  const sector = getPrimaryUserSector(user);
  if (!sector) return true;
  return normalizeSectorValue(alert?.sector) === sector;
}

function updatePrimaryUserActionUi() {
  if (isClientUser()) {
    if (openSectorAlertsEl) openSectorAlertsEl.classList.add('hidden');
    if (openMyProjectSignalsEl) openMyProjectSignalsEl.classList.add('hidden');
    if (openProjectSignalsEl) openProjectSignalsEl.classList.add('hidden');
    if (openStageUpdatesEl) openStageUpdatesEl.classList.add('hidden');
    return;
  }
  if (!openSectorAlertsEl) return;
  const sectorScopedToggle = shouldUseSectorScopedToggle();
  const projectsScope = !sectorScopedToggle && userHasProjectsScope();
  const viewingMine = projectsScope && state.projectView === "mine";
  const sectorScopedView = isSectorScopedViewActive();
  openSectorAlertsEl.textContent = projectsScope
    ? (viewingMine ? "Todos os projetos" : "Meus projetos")
    : (sectorScopedView ? "Todos os alertas" : "Meus alertas");
  openSectorAlertsEl.title = projectsScope
    ? (viewingMine
        ? "Voltar para a visualização com todos os projetos"
        : "Visualizar apenas os projetos vinculados ao seu nome na coluna PM")
    : (sectorScopedView
        ? "Voltar para a visualização com todos os projetos e alertas"
        : "Visualizar apenas os projetos e alertas do seu setor monitorado");
  const titleEl = document.getElementById("sector-alerts-title");
  if (titleEl && state.sectorAlertsMode !== 'project-signals') {
    titleEl.textContent = projectsScope ? "Meus projetos" : "Meus alertas por setor";
  }
  if (openMyProjectSignalsEl) {
    const canViewMine = canViewMyProjectSignals();
    const pendingCount = getMyProjectSignals().filter((alert) => !getSignalResolutionInfo(alert.id)).length;
    openMyProjectSignalsEl.classList.toggle('hidden', !canViewMine);
    openMyProjectSignalsEl.title = 'Acompanhar as sinalizações que você enviou ao PCP';
    openMyProjectSignalsEl.textContent = canViewMine && pendingCount > 0
      ? `Minhas sinalizações (${pendingCount})`
      : 'Minhas sinalizações';
  }
  if (openProjectSignalsEl) {
    const canView = canViewProjectSignals();
    openProjectSignalsEl.classList.toggle('hidden', !canView);
    openProjectSignalsEl.title = 'Visualizar apenas os alertas enviados pelos usuários de Projetos';
  }
  if (openStageUpdatesEl) {
    const canOpen = canOpenStageWorkspace();
    openStageUpdatesEl.classList.toggle('hidden', !canOpen);
    openStageUpdatesEl.textContent = canValidateStageWorkspace() ? 'Validação PCP / Apontamentos' : 'Apontamentos';
    openStageUpdatesEl.title = canValidateStageWorkspace()
      ? 'Validar apontamentos enviados pelos setores e consultar o histórico'
      : 'Informar o avanço da sua etapa por spool';
  }
}

function tokenizeNormalizedNames(values = []) {
  const set = new Set();
  const source = Array.isArray(values) ? values : [values];
  for (const value of source) {
    const normalized = normalizeText(value).trim();
    if (!normalized) continue;
    set.add(normalized);
    for (const part of normalized.split(/[^a-z0-9]+/)) {
      if (part) set.add(part);
    }
  }
  return set;
}

function projectBelongsToUser(project, user = state.user) {
  if (!project || !userHasProjectsScope(user)) return false;
  const pmValue = String(project.pm || '').trim();
  if (!pmValue) return false;
  const candidates = tokenizeNormalizedNames([
    user.name,
    user.username,
    String(user.username || '').split('@')[0],
    ...(Array.isArray(user.projectPmAliases) ? user.projectPmAliases : []),
  ]);
  if (!candidates.size) return false;
  const normalizedPm = normalizeText(pmValue).trim();
  const pmTokens = tokenizeNormalizedNames(pmValue.split(/[;,|/]+/));
  for (const candidate of candidates) {
    if (normalizedPm === candidate || normalizedPm.includes(candidate)) return true;
    if (pmTokens.has(candidate)) return true;
  }
  return false;
}




function resetClientBspOverridesState() {
  state.clientBspOverrides = { items: [], byProjectRowId: {}, byProjectNumber: {}, loading: false, loaded: false, feedback: '', editingProjectId: null, activeExecutiveProjectId: null };
}

function userCanViewClientPanel(user = state.user) {
  if (!user) return false;
  if (user.role === 'admin' || user.role === 'client') return true;
  return user.canViewClientPanel === true;
}

function canManageClientBspPanel(project = null, user = state.user) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  if (!userHasProjectsScope(user)) return false;
  return project ? projectBelongsToUser(project, user) : true;
}

function canOpenClientBspPanel(project = null, user = state.user) {
  return isClientUser(user) || canManageClientBspPanel(project, user) || userCanViewClientPanel(user);
}

function openClientBspExecutiveForPmEdit(project) {
  if (!project || !canManageClientBspPanel(project)) return;
  state.clientBspOverrides.editingProjectId = project.rowId;
  openClientBspExecutive(project, { keepEditing: true, scrollToEditor: true });
}

function getClientBspOverrideProjectRowId(project) {
  return String(project?.rowId ?? project?.rowNumber ?? '').trim();
}

function getClientBspOverrideProjectNumber(project) {
  return String(project?.projectNumber || project?.projectDisplay || '').trim();
}

function normalizeClientBspOverride(row = {}) {
  if (!row || typeof row !== 'object') return null;
  const projectRowId = String(row.projectRowId ?? row.project_row_id ?? '').trim();
  const projectNumber = String(row.projectNumber ?? row.project_number ?? '').trim();
  const customFields = row.customFields && typeof row.customFields === 'object'
    ? row.customFields
    : (row.custom_fields && typeof row.custom_fields === 'object' ? row.custom_fields : {});
  const systemDate = (camelKey, snakeKey = '') => row[camelKey] ?? (snakeKey ? row[snakeKey] : '') ?? customFields[`__${camelKey}`] ?? customFields[camelKey] ?? '';
  return {
    id: row.id || '',
    region: row.region || '',
    projectRowId,
    projectNumber,
    projectDisplay: row.projectDisplay ?? row.project_display ?? '',
    clientKey: row.clientKey ?? row.client_key ?? '',
    clientName: row.clientName ?? row.client_name ?? '',
    vessel: row.vessel || '',
    pm: row.pm || '',
    drawingsStartOverride: systemDate('drawingsStartOverride'),
    drawingsFinishOverride: systemDate('drawingsFinishOverride'),
    procurementStartOverride: systemDate('procurementStartOverride'),
    procurementFinishOverride: systemDate('procurementFinishOverride'),
    fabricationStartOverride: row.fabricationStartOverride ?? row.fabrication_start_override ?? '',
    boilermakerFinishOverride: row.boilermakerFinishOverride ?? row.boilermaker_finish_override ?? '',
    weldingFinishOverride: row.weldingFinishOverride ?? row.welding_finish_override ?? '',
    inspectionFinishOverride: row.inspectionFinishOverride ?? row.inspection_finish_override ?? '',
    thFinishOverride: row.thFinishOverride ?? row.th_finish_override ?? '',
    coatingFinishOverride: row.coatingFinishOverride ?? row.coating_finish_override ?? '',
    projectFinishOverride: row.projectFinishOverride ?? row.project_finish_override ?? '',
    executiveStatus: row.executiveStatus ?? row.executive_status ?? '',
    executiveNote: row.executiveNote ?? row.executive_note ?? '',
    delayReason: row.delayReason ?? row.delay_reason ?? '',
    customFields,
    visibleToClient: row.visibleToClient ?? row.visible_to_client ?? true,
    createdBy: row.createdBy ?? row.created_by ?? '',
    createdByName: row.createdByName ?? row.created_by_name ?? '',
    createdAt: row.createdAt ?? row.created_at ?? null,
    updatedBy: row.updatedBy ?? row.updated_by ?? '',
    updatedByName: row.updatedByName ?? row.updated_by_name ?? '',
    updatedAt: row.updatedAt ?? row.updated_at ?? null,
  };
}

function setClientBspOverrides(items = []) {
  const normalized = (Array.isArray(items) ? items : []).map(normalizeClientBspOverride).filter(Boolean);
  const byProjectRowId = {};
  const byProjectNumber = {};
  for (const item of normalized) {
    if (item.projectRowId) byProjectRowId[item.projectRowId] = item;
    if (item.projectNumber) byProjectNumber[normalizeText(item.projectNumber)] = item;
  }
  state.clientBspOverrides.items = normalized;
  state.clientBspOverrides.byProjectRowId = byProjectRowId;
  state.clientBspOverrides.byProjectNumber = byProjectNumber;
  state.clientBspOverrides.loaded = true;
}

function getClientBspOverride(project) {
  const rowId = getClientBspOverrideProjectRowId(project);
  if (rowId && state.clientBspOverrides.byProjectRowId[rowId]) return state.clientBspOverrides.byProjectRowId[rowId];
  const numberKey = normalizeText(getClientBspOverrideProjectNumber(project));
  if (numberKey && state.clientBspOverrides.byProjectNumber[numberKey]) return state.clientBspOverrides.byProjectNumber[numberKey];
  return null;
}

function hasClientBspOverrideContent(override) {
  if (!override) return false;
  const values = [
    override.drawingsStartOverride,
    override.drawingsFinishOverride,
    override.procurementStartOverride,
    override.procurementFinishOverride,
    override.fabricationStartOverride,
    override.boilermakerFinishOverride,
    override.weldingFinishOverride,
    override.inspectionFinishOverride,
    override.thFinishOverride,
    override.coatingFinishOverride,
    override.projectFinishOverride,
    override.executiveStatus,
    override.executiveNote,
    override.delayReason,
  ];
  const custom = override.customFields && typeof override.customFields === 'object'
    ? Object.values(override.customFields).filter((value) => String(value || '').trim())
    : [];
  return values.some((value) => String(value || '').trim()) || custom.length > 0;
}

function clientDateInputValue(value) {
  const date = parseClientSafeDateObject(value) || parseDateObject(value);
  if (!date) return '';
  const y = date.getUTCFullYear();
  const m = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function clientOverrideDateObject(value) {
  return parseClientSafeDateObject(value) || parseDateObject(value);
}

function clientOverrideCustomFieldsArray(override) {
  const fields = override?.customFields && typeof override.customFields === 'object' ? override.customFields : {};
  return Object.entries(fields)
    .filter(([label]) => !String(label || '').startsWith('__'))
    .map(([label, value]) => ({ label: String(label || '').trim(), value: String(value || '').trim() }))
    .filter((item) => item.label || item.value);
}


function clientSafeJsonForScript(value) {
  return JSON.stringify(value == null ? null : value).replace(/</g, '\u003c').replace(/>/g, '\u003e').replace(/&/g, '\u0026');
}

function normalizeClientIsoScheduleOverride(value = {}) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const startDate = clientDateInputValue(value.startDate || value.start || value.isoStartDate || value.inicio || '');
  const finishDate = clientDateInputValue(value.finishDate || value.finish || value.isoFinishDate || value.termino || '');
  const note = String(value.note || value.observation || value.observacao || '').trim();
  const iso = String(value.iso || value.ISO || '').trim();
  const drawing = String(value.drawing || value.drawingNumber || '').trim();
  const description = String(value.description || '').trim();
  if (!startDate && !finishDate && !note && !iso && !drawing && !description) return null;
  return {
    startDate,
    finishDate,
    note,
    iso,
    drawing,
    description,
    updatedAt: value.updatedAt || value.updated_at || null,
    updatedBy: value.updatedBy || value.updated_by || '',
    updatedByName: value.updatedByName || value.updated_by_name || '',
  };
}

function getClientIsoScheduleOverrideMap(override) {
  const fields = override?.customFields && typeof override.customFields === 'object' ? override.customFields : {};
  const map = fields.__isoDateOverrides || fields.__isoScheduleOverrides || fields.isoDateOverrides;
  return map && typeof map === 'object' && !Array.isArray(map) ? map : {};
}

function getClientSpoolIsoOverrideCandidateKeys(spool, index = 0) {
  const candidates = [
    getClientSpoolPanelKey(spool, index),
    normalizeCompactText(spool?.iso || ''),
    normalizeCompactText(spool?.drawing || ''),
    normalizeCompactText(spool?.projectRef || ''),
    normalizeCompactText([spool?.iso, spool?.drawing].filter(Boolean).join('|')),
  ].map((item) => String(item || '').trim()).filter(Boolean);
  return Array.from(new Set(candidates));
}

function getClientIsoDateOverride(project, spool, index = 0) {
  if (!project || !spool) return null;
  const override = getClientBspOverride(project);
  const map = getClientIsoScheduleOverrideMap(override);
  if (!map || !Object.keys(map).length) return null;
  for (const key of getClientSpoolIsoOverrideCandidateKeys(spool, index)) {
    const normalized = normalizeClientIsoScheduleOverride(map[key]);
    if (normalized) return { ...normalized, key };
  }
  return null;
}

function clientReadSpoolDateCandidates(spool, wantedKeys, candidates) {
  clientReadDateCandidatesFromSource(spool, wantedKeys, candidates);
  clientReadDateCandidatesFromSource(spool?.stageValues, wantedKeys, candidates);
  clientReadDateCandidatesFromMilestones(spool?.milestones, wantedKeys, candidates);
}

function getClientSpoolTrackingDate(spool, wantedKeys, mode = 'last') {
  const candidates = [];
  clientReadSpoolDateCandidates(spool, wantedKeys, candidates);
  const unique = Array.from(new Map(candidates.filter(Boolean).map((date) => [date.getTime(), date])).values()).sort((a, b) => a - b);
  if (!unique.length) return null;
  return mode === 'first' ? unique[0] : unique[unique.length - 1];
}

function getClientSpoolTrackingDates(project, spool) {
  const scopedProject = { ...(project || {}), spools: spool ? [spool] : [], stageValues: spool?.stageValues || project?.stageValues || {} };
  const hydro = shouldClientShowHydro(scopedProject);
  return {
    drawingsStart: getClientSpoolTrackingDate(spool, ['Drawing Start Date', 'Drawings Start Date', 'Engineering Start Date', 'Start Drawing Date'], 'first'),
    drawingsFinish: getClientSpoolTrackingDate(spool, ['Drawing Execution Advance%', 'Drawing Finish Date', 'Drawings Finish Date', 'Engineering Finish Date', 'Drawing'], 'last'),
    procurementStart: getClientSpoolTrackingDate(spool, ['Procurement Start Date', 'Procuremnt Start Date', 'Material Procurement Start Date', 'Materials Start Date'], 'first'),
    procurementFinish: getClientSpoolTrackingDate(spool, ['Procuremnt Status %', 'Procurement Status %', 'Procurement Finish Date', 'Material Separation', 'Material Release to Fabrication', 'Materials Finish Date'], 'last'),
    fabricationStart: getClientSpoolTrackingDate(spool, ['Fabrication Start Date', 'Fab. Início', 'FAB INICIO', 'Fab Inicio'], 'first'),
    boilermakerFinish: getClientSpoolTrackingDate(spool, ['Boilermaker Finish Date', 'Caldeiraria', 'Calderaria Finish Date'], 'last'),
    weldingFinish: getClientSpoolTrackingDate(spool, ['Welding Finish Date', 'Solda', 'Weld Finish Date', 'Full Welding Finish Date'], 'last'),
    inspectionFinish: getClientSpoolTrackingDate(spool, ['Inspection Finish Date (QC)', 'Final Dimensional Inspection Finish Date', 'Inspection Finish'], 'last'),
    thFinish: hydro ? getClientSpoolTrackingDate(spool, ['TH Finish Date', 'Hydro Finish Date', 'Hydro Testing Finish Date'], 'last') : null,
    coatingFinish: getClientSpoolTrackingDate(spool, ['Coating Finish Date', 'Painting Finish Date', 'HDG / FBE DATE RETORNO (PAINT)'], 'last'),
    projectFinish: getClientSpoolTrackingDate(spool, ['Project Finish Date', 'Data de Envio', 'Shipment Date', 'Delivery Date'], 'last'),
    sources: {},
  };
}

function buildClientIsoExecutiveSchedule(project, spool, index = 0) {
  if (!project || !spool) return [];
  const isoOverride = getClientIsoDateOverride(project, spool, index) || {};
  const scopedProject = {
    ...project,
    spools: [spool],
    stageValues: spool.stageValues && typeof spool.stageValues === 'object' && Object.keys(spool.stageValues).length ? spool.stageValues : (project.stageValues || {}),
    plannedStartDate: isoOverride.startDate || spool.plannedStartDate || spool.startDate || project.plannedStartDate,
    plannedFinishDate: isoOverride.finishDate || spool.plannedFinishDate || spool.finishDate || spool.deliveryDate || project.plannedFinishDate,
  };
  const start = parseClientSafeDateObject(isoOverride.startDate) || parseClientSafeDateObject(scopedProject.plannedStartDate) || getClientAnalyticStartDate(project);
  const finish = parseClientSafeDateObject(isoOverride.finishDate) || parseClientSafeDateObject(scopedProject.plannedFinishDate) || getClientAnalyticFinishDate(project);
  return buildClientExecutiveSchedule(scopedProject, {
    start,
    finish,
    realDates: getClientSpoolTrackingDates(project, spool),
    skipReplan: true,
  });
}

async function loadClientBspOverrides(options = {}) {
  if (!state.user || state.clientBspOverrides.loading) return;
  state.clientBspOverrides.loading = true;
  if (!options.silent) state.clientBspOverrides.feedback = 'Carregando ajustes executivos...';
  try {
    const response = await fetch('/api/client-bsp-overrides', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar ajustes executivos.');
    setClientBspOverrides(data.overrides || []);
    state.clientBspOverrides.feedback = '';
    renderClientDashboard();
    if (modalEl && !modalEl.classList.contains('hidden')) {
      const selected = getSelectedProject();
      if (selected) renderModal(selected);
    }
    const activeId = state.clientBspOverrides.activeExecutiveProjectId;
    const execModal = document.getElementById('client-bsp-executive-modal');
    if (activeId && execModal && !execModal.classList.contains('hidden')) {
      const project = state.projects.find((item) => String(item.rowId) === String(activeId));
      if (project) openClientBspExecutive(project, { keepEditing: true });
    }
  } catch (error) {
    const message = error?.message || 'Falha ao carregar ajustes executivos.';
    state.clientBspOverrides.feedback = /client_bsp_overrides|Tabela client_bsp_overrides/i.test(String(message || '')) ? '' : message;
    if (state.clientBspOverrides.feedback) console.warn('Ajustes executivos não carregados:', state.clientBspOverrides.feedback);
  } finally {
    state.clientBspOverrides.loading = false;
  }
}

function scheduleClientBspOverridesLoad() {
  if (!state.user || state.clientBspOverrides.loading || state.clientBspOverrides.loaded) return;
  window.setTimeout(() => loadClientBspOverrides({ silent: true }), 0);
}

function getVisibleProjectsSource() {
  if (state.projectView === 'mine' && userHasProjectsScope()) {
    return state.projects.filter((project) => projectBelongsToUser(project));
  }
  if (isSectorScopedViewActive()) {
    return state.projects.filter((project) => projectMatchesScopedSector(project));
  }
  return state.projects;
}

function renderProjectViewTabs() {
  if (!projectViewTabsEl) return;
  if (shouldUseSectorScopedToggle() || !userHasProjectsScope()) {
    state.projectView = 'all';
    projectViewTabsEl.innerHTML = '';
    projectViewTabsEl.classList.add('hidden');
    return;
  }
  const mineCount = state.projects.filter((project) => projectBelongsToUser(project)).length;
  projectViewTabsEl.classList.add('hidden');
  projectViewTabsEl.innerHTML = `
    <button type="button" class="ghost-button ghost-button--compact ${state.projectView === 'all' ? 'is-active' : ''}" data-project-view="all">Todos os projetos <strong>${state.projects.length}</strong></button>
    <button type="button" class="ghost-button ghost-button--compact ${state.projectView === 'mine' ? 'is-active' : ''}" data-project-view="mine">Meus projetos <strong>${mineCount}</strong></button>
  `;
}

function applyFilter() {
  const query = normalizeText(state.searchQuery).trim();
  const demand = normalizeText(state.demandFilter).trim();
  const selectedProjectType = normalizeText(state.projectTypeFilter).trim();
  const selectedWeek = String(state.weekFilter || '').trim();

  const sourceProjects = getVisibleProjectsSource();

  state.filteredProjects = sourceProjects
    .filter((project) => {
      const matchesQuery = !query || project._searchText.includes(query);
      const matchesDemand = !demand
        || normalizeText(getProjectCurrentStageDisplay(project)).includes(demand)
        || normalizeText(project.currentStage).includes(demand)
        || normalizeText(translateProjectStatus(project.projectStatus, project.uiState)).includes(demand);
      const matchesProjectType = !selectedProjectType || normalizeText(getProjectTypeLabel(project)) === selectedProjectType;
      const matchesWeek = projectMatchesWeekFilter(project, selectedWeek);
      const matchesStatus = projectMatchesStatusFilter(project);
      return matchesQuery && matchesDemand && matchesProjectType && matchesWeek && matchesStatus;
    })
    .sort(compareProjectsByPlannedFinishDate);

  if (!state.filteredProjects.find((project) => project.rowId === state.selectedProjectId)) {
    state.selectedProjectId = state.filteredProjects[0]?.rowId || null;
  }
}

function getSelectedProject() {
  return state.filteredProjects.find((project) => project.rowId === state.selectedProjectId)
    || state.projects.find((project) => project.rowId === state.selectedProjectId)
    || null;
}

function getBacklogKg(project) {
  if (!project) return 0;
  const total = Number(project.kilos || 0);
  const welded = Number(project.weldedWeightKg || 0);
  return Math.max(0, total - welded);
}

function getProjectItemCount(project) {
  const declared = Number(project?.quantitySpools || 0);
  const spoolsCount = Array.isArray(project?.spools) ? project.spools.length : 0;
  return declared > 0 ? declared : spoolsCount;
}


function isClientUser(user = state.user) {
  return Boolean(user && user.role === 'client');
}

function normalizeOperationRegionValue(value = 'BR') {
  // Painel Brasil: sistema separado do painel Portugal.
  // A base Supabase é compartilhada, mas esta build nunca deve trocar para PT.
  return 'BR';
}

function getOperationRegion(user = null) {
  return 'BR';
}

function syncOperationRegionButtons(value = 'BR') {
  const normalized = normalizeOperationRegionValue(value);
  if (adminUserOperationRegionEl) adminUserOperationRegionEl.value = normalized;
  document.querySelectorAll('[data-operation-region-option]').forEach((button) => {
    button.classList.toggle(
      'is-active',
      normalizeOperationRegionValue(button.dataset.operationRegionOption || 'BR') === normalized
    );
  });
  try { window.localStorage.setItem('step_operation_region', 'BR'); } catch {}
}

function buildClientKey(clientName, region = getOperationRegion()) {
  const base = String(clientName || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  if (!base) return '';
  const clean = base.replace(/_(BR|PT)$/i, '');
  return `${clean}_${normalizeOperationRegionValue(region)}`;
}

document.addEventListener('click', (event) => {
  const button = event.target.closest('[data-operation-region-option]');
  if (!button) return;
  const nextRegion = 'BR';
  syncOperationRegionButtons(nextRegion);
  if (adminUserClientNameEl && adminUserClientKeyEl && document.getElementById('admin-user-role')?.value === 'client') {
    const cleanCurrent = String(adminUserClientKeyEl.value || '').replace(/_(BR|PT)$/i, '');
    const source = adminUserClientNameEl.value || cleanCurrent;
    if (source) adminUserClientKeyEl.value = buildClientKey(source, nextRegion);
  }
});

if (adminUserOperationRegionEl) {
  adminUserOperationRegionEl.addEventListener('change', () => {
    const nextRegion = 'BR';
    syncOperationRegionButtons(nextRegion);
    if (adminUserClientNameEl && adminUserClientKeyEl && document.getElementById('admin-user-role')?.value === 'client') {
      const cleanCurrent = String(adminUserClientKeyEl.value || '').replace(/_(BR|PT)$/i, '');
      const source = adminUserClientNameEl.value || cleanCurrent;
      if (source) adminUserClientKeyEl.value = buildClientKey(source, nextRegion);
    }
  });
}

