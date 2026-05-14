const PROJECTS_REFRESH_MS = 180000;
const PROJECTS_CACHE_TTL_MS = 5 * 60 * 1000;
const ALERTS_REFRESH_MS = 60000;
const PRESENCE_HEARTBEAT_MS = 90000;
const AUTH_REFRESH_MS = 300000;
const ADMIN_REFRESH_MS = 60000;
const ALERT_NOTIFICATION_COOLDOWN_MS = 4 * 60 * 60 * 1000;
const PROJECTS_CACHE_KEY = 'step_dashboard_projects_cache_v4_client_scope';

let adminResponsesPollTimer = null;
let projectsWarmupPromise = null;
let projectsWarmupResetTimer = null;

const state = {
  projects: [],
  filteredProjects: [],
  projectView: 'all',
  projectDrill: { open: false, mode: 'total', selectedClientKey: '', selectedVesselKey: '' },
  clientPortal: { selectedVesselKey: '', selectedProjectId: null, rowClickTimer: null, vesselClickTimer: null },
  clientApi: { keys: [], loading: false, newToken: '', newTokenKeyId: '', feedback: '' },
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
const adminAutoLoginListEl = document.getElementById("admin-auto-login-list");
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

    const registration = await navigator.serviceWorker.register('./sw.js?v=32', { updateViaCache: 'none' });
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

  if (text.includes('hydro') || text.includes('teste hidrostatico') || /\bth\b/.test(text) || text.includes('aguardando em th') || (th > 0 && th < 100)) return 'th';
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

  if (projectFinishDate || packageDelivered >= 100) return '';
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

  return {
    current: hasCurrent ? current : null,
    matched,
    label: hasCurrent
      ? (matched ? `Tracking OK ${formatPercent(current)}` : `Aguardando tracking ${formatPercent(current)}/${formatPercent(progress)}`)
      : 'Tracking não localizado',
    className: hasCurrent
      ? (matched ? 'stage-badge--tracking-ok' : 'stage-badge--tracking-waiting')
      : 'stage-badge--tracking-missing',
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
  if (hasClientIncompleteProductionEvidence(project)) {
    const openStage = getClientFirstIncompleteProductionStage(project);
    if (openStage?.key === 'package') return { text: 'Aguardando envio', state: 'awaiting_shipment' };
    if (openStage?.key === 'engineering') return { text: 'Engenharia em andamento', state: 'in_progress' };
    if (openStage?.key === 'procurement') return { text: 'Suprimento em andamento', state: 'in_progress' };
    return { text: 'Em produção', state: 'in_progress' };
  }

  if (projectHasAwaitingShipmentPackage(project) && project?.uiState !== 'completed') {
    return { text: 'Aguardando envio', state: 'awaiting_shipment' };
  }

  const preparedForShipment = isProjectPreparedForShipment(project);
  if (preparedForShipment && project?.uiState !== 'completed') {
    return { text: 'Preparado para envio', state: 'preparing_shipment' };
  }

  const statusText = project?.statusSummary || project?.currentStatus || project?.currentStage || '';
  if (statusText) {
    const state = project?.finished || project?.uiState === 'completed'
      ? 'completed'
      : (project?.uiState === 'awaiting_shipment' ? 'preparing_shipment' : (project?.uiState || project?.operationalState || 'in_progress'));
    return { text: statusText, state };
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
  if (isProjectFinalizedForDisplay(project)) return 'Enviado';
  return project?.sectorSummary || project?.currentStageGroup || project?.currentSector || project?.operationalSector || '';
}

function getProjectCurrentStageDisplay(project) {
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

function startClocks() {
  const tick = () => {
    setClock("clock-br-time", "clock-br-date", "pt-BR", "America/Sao_Paulo");
    setClock("clock-pt-time", "clock-pt-date", "pt-PT", "Europe/Lisbon");
  };
  tick();
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
  if (Array.isArray(state.filteredProjects) && hasActiveProjectTableFilters()) return state.filteredProjects;
  const source = getVisibleProjectsSource();
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

function isProjectStartedForStats(project) {
  if (!project || isProjectExcludedFromTotal(project)) return { started: false, tags: 0 };
  const openItems = getProjectOpenFlowItems(project);
  const startedItems = openItems.filter((item) => {
    const sectorKey = getFlowSectorKey(item.flow);
    return ['producao', 'solda', 'calderaria', 'inspecao', 'pintura', 'pendente_envio'].includes(sectorKey);
  });
  if (startedItems.length) return { started: true, tags: startedItems.length };

  const statusText = normalizeText([project.currentStage, project.currentStatus, project.statusSummary, project.operationalSector, project.currentSector, project.flow?.status, project.flow?.sector].filter(Boolean).join(' '));
  const textualStarted = ['corte', 'fabrication', 'pre montagem', 'solda', 'welding', 'inspecao', 'inspection', 'th', 'pintura', 'painting', 'coating', 'unitizacao', 'envio'].some((term) => statusText.includes(term));
  if (textualStarted) return { started: true, tags: Number(project.quantitySpools || 1) };

  const progress = Number(project.overallProgress || project.currentStagePercent || 0);
  return progress > 0 ? { started: true, tags: Number(project.quantitySpools || 1) } : { started: false, tags: 0 };
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
    const isHoldProject = isProjectOnHold(project);
    const isPendingProject = isProjectPending(project);

    if (isHoldProject) {
      stats.notStartedHold += 1;
      stats.notStartedHoldTags += tags;
      continue;
    }

    if (isPendingProject) {
      continue;
    }

    if (isFinishedProject) {
      stats.completed += 1;
      stats.completedTags += tags;
      continue;
    }

    progressAccumulator += Number(project.overallProgress || 0);

    const openItems = getProjectOpenFlowItems(project);
    const startedSnapshot = isProjectStartedForStats(project);
    if (startedSnapshot.started) {
      stats.startedProjects += 1;
      stats.startedTags += Number(startedSnapshot.tags || tags || 0);
    }
    const countSector = (sectorKey) => openItems.filter((item) => getFlowSectorKey(item.flow) === sectorKey).length;
    const producaoTags = countSector('producao') + countSector('solda') + countSector('calderaria');
    const qualidadeTags = countSector('inspecao');
    const pinturaTags = countSector('pintura');
    const logisticaTags = getAwaitingShipmentTags(project);
    const preStartTags = countSector('engenharia') + countSector('suprimento');

    if (producaoTags) {
      stats.inProgress += 1;
      stats.inProgressTags += producaoTags;
    }
    if (qualidadeTags) {
      stats.inspectionProjects += 1;
      stats.inspectionTags += qualidadeTags;
    }
    if (pinturaTags) {
      stats.paintingProjects += 1;
      stats.paintingTags += pinturaTags;
    }
    if (logisticaTags) {
      stats.awaitingShipment += 1;
      stats.awaitingShipmentTags += logisticaTags;
    }
    if (preStartTags || (!openItems.length && !project.finished)) {
      stats.notStarted += 1;
      stats.notStartedTags += preStartTags || tags;
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


function getOperationRegion(user = null) {
  const fromAdminSelect = adminUserOperationRegionEl?.value;
  const raw = String(
    fromAdminSelect
    || user?.operationRegion
    || user?.region
    || user?.environment
    || window.localStorage.getItem('step_operation_region')
    || 'PT'
  ).trim().toUpperCase();

  return ['BR', 'BRASIL', 'BRAZIL'].includes(raw) ? 'BR' : 'PT';
}


function syncOperationRegionButtons(value = 'PT') {
  const normalized = String(value || 'PT').trim().toUpperCase() === 'BR' ? 'BR' : 'PT';

  if (adminUserOperationRegionEl) {
    adminUserOperationRegionEl.value = normalized;
  }

  document.querySelectorAll('[data-operation-region-option]').forEach((button) => {
    button.classList.toggle(
      'is-active',
      String(button.dataset.operationRegionOption || '').toUpperCase() === normalized
    );
  });

  try {
    window.localStorage.setItem('step_operation_region', normalized);
  } catch {}
}

document.addEventListener('click', (event) => {
  const button = event.target.closest('[data-operation-region-option]');
  if (!button) return;

  const nextRegion = String(button.dataset.operationRegionOption || 'PT').toUpperCase() === 'BR' ? 'BR' : 'PT';
  syncOperationRegionButtons(nextRegion);

  if (adminUserClientNameEl && adminUserClientKeyEl) {
    const cleanCurrent = String(adminUserClientKeyEl.value || '').replace(/_(BR|PT)$/i, '');
    const source = adminUserClientNameEl.value || cleanCurrent;
    adminUserClientKeyEl.value = buildClientKey(source, nextRegion);
  }
});

if (adminUserOperationRegionEl) {
  adminUserOperationRegionEl.addEventListener('change', () => {
    syncOperationRegionButtons(adminUserOperationRegionEl.value || 'PT');
    if (adminUserClientNameEl && adminUserClientKeyEl) {
      const cleanCurrent = String(adminUserClientKeyEl.value || '').replace(/_(BR|PT)$/i, '');
      const source = adminUserClientNameEl.value || cleanCurrent;
      adminUserClientKeyEl.value = buildClientKey(source, getOperationRegion());
    }
  });
}


function buildProjectsApiUrl(params = {}) {
  const query = new URLSearchParams();
  query.set('region', getOperationRegion());
  Object.entries(params || {}).forEach(([key, value]) => {
    if (value === undefined || value === null || value === false || value === '') return;
    query.set(key, String(value));
  });
  return `/api/projects?${query.toString()}`;
}

function buildClientKey(clientName, region = 'PT') {
  const base = String(clientName || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');

  if (!base) return '';
  return `${base}_${region}`;
}

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
        <button class="mini-action-button client-macro-button" type="button" data-client-open-macro-dashboard>Visão executiva da carteira</button>
        <button class="mini-action-button client-api-button" type="button" data-client-open-api>Gerar API</button>
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
  if (syncEl) syncEl.textContent = state.meta?.lastSync ? `Atualização: ${new Date(state.meta.lastSync).toLocaleString('pt-BR')}` : 'Atualização: --';

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

function renderClientSpoolRows(spools, limit = 120) {
  const items = (Array.isArray(spools) ? [...spools] : []).sort(compareClientSpoolsByPriority).slice(0, limit);
  if (!items.length) return '<tr><td colspan="6" class="loading-cell">Nenhuma tag detalhada encontrada para esta BSP.</td></tr>';
  return items.map((spool) => {
    const state = getClientSpoolVisualState(spool);
    const statusText = spool?.currentStatus || spool?.stage || uiStateLabel(spool?.uiState);
    return `<tr class="client-spool-row client-spool-row--${state}"><td><strong>${escapeHtml(spool.iso || '—')}</strong></td><td>${escapeHtml(spool.description || '—')}</td><td><span class="client-spool-chip client-spool-chip--${state}">${escapeHtml(statusText)}</span></td><td>${escapeHtml(spool.currentSector || spool.operationalSector || '—')}</td><td><span class="client-spool-progress client-spool-progress--${state}">${formatPercent(spool.overallProgress)}</span></td><td>${formatNumber(spool.kilos, 2)} kg</td></tr>`;
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
    <div class="client-stage-strip">
      ${['Drawing Execution Advance%', 'Procuremnt Status %', 'Material Separation', 'Full welding execution', 'Non Destructive Examination (QC)', 'Hydro Test Pressure (QC)', 'Surface preparation and/or coating', 'Final Inspection', 'Package and Delivered'].map((key) => {
        const label = (state.meta?.stageOrder || []).find((stage) => stage.key === key)?.label || key;
        const value = stageValues[key];
        return renderClientStageStripCard(label, value);
      }).join('')}
    </div>
    <div class="client-table-wrap client-table-wrap--compact">
      <table class="client-bsp-table"><thead><tr><th>Tag/ISO</th><th>Descrição</th><th>Status</th><th>Etapa</th><th>%</th><th>Peso</th></tr></thead><tbody>
        ${renderClientSpoolRows(spools, 80)}
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
  const fabrication = { hasEvidence: hasClientStageEvidence(project, [
    'Welding Preparation',
    'Spool Assemble and tack weld',
    'Initial Dimensional Inspection/3D',
    'Full welding execution',
    'Non Destructive Examination (QC)',
    'Final Dimensional Inpection/3D (QC)',
    'Hydro Test Pressure (QC)',
    'Surface preparation and/or coating',
    'HDG / FBE.  (PAINT)',
  ]), percent: getClientFabricationProgress(project) };
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

  const stages = [
    { keys: ['Welding Preparation', 'Spool Assemble and tack weld'], weight: 10 },
    { keys: ['Initial Dimensional Inspection/3D'], weight: 8 },
    { keys: ['Full welding execution'], weight: 25 },
    { keys: ['Non Destructive Examination (QC)'], weight: 12 },
    { keys: ['Final Dimensional Inpection/3D (QC)'], weight: 8 },
    { keys: ['Hydro Test Pressure (QC)'], weight: 7 },
    { keys: ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)'], weight: 15 },
  ];
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

function getClientPackageProgress(project) {
  return getClientStageValue(project, ['Package and Delivered', 'Final Inspection']);
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
    const parsed = parseDateObject(value);
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
    const parsed = parseDateObject(value);
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

function buildClientSCurveData(project) {
  const start = getClientAnalyticStartDate(project);
  let finish = getClientAnalyticFinishDate(project);
  if (finish <= start) finish = addUtcDays(start, 30);
  const duration = clientDaysBetween(start, finish);
  const step = Math.max(1, Math.ceil(duration / 14));
  const today = getCurrentBrazilDate();
  const actualNow = getClientOverallProgress(project);
  const plannedToday = getClientPlannedToday(project);
  const points = [];
  for (let day = 0; day <= duration; day += step) {
    const date = addUtcDays(start, day);
    const ratio = day / duration;
    const planned = clientSchedulePlannedPercent(ratio);
    let actual = null;
    if (date <= today) {
      if (plannedToday > 0) {
        actual = clampClientPercent((planned / plannedToday) * actualNow);
      } else {
        actual = actualNow > 0 ? actualNow : 0;
      }
    }
    points.push({ date, planned, actual });
  }
  if (points[points.length - 1]?.date < finish) {
    points.push({ date: finish, planned: 100, actual: finish <= today ? actualNow : null });
  }
  return points;
}

function clientSvgPolyline(points, width, height, getValue) {
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const usable = points.filter((point) => getValue(point) != null);
  if (!usable.length) return '';
  return usable.map((point, index) => {
    const x = pad.left + (points.indexOf(point) / Math.max(1, points.length - 1)) * innerW;
    const y = pad.top + (1 - clampClientPercent(getValue(point)) / 100) * innerH;
    return `${index ? 'L' : 'M'} ${x.toFixed(1)} ${y.toFixed(1)}`;
  }).join(' ');
}

function renderClientSCurveSvg(project) {
  const points = buildClientSCurveData(project);
  const width = 760;
  const height = 260;
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const plannedPath = clientSvgPolyline(points, width, height, (point) => point.planned);
  const actualPath = clientSvgPolyline(points, width, height, (point) => point.actual);
  const grid = [0, 25, 50, 75, 100].map((value) => {
    const y = pad.top + (1 - value / 100) * innerH;
    return `<line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" class="client-chart-grid" /><text x="8" y="${y + 4}" class="client-chart-label">${value}%</text>`;
  }).join('');
  const first = points[0]?.date ? clientFormatDateValue(points[0].date) : '';
  const mid = points[Math.floor(points.length / 2)]?.date ? clientFormatDateValue(points[Math.floor(points.length / 2)].date) : '';
  const last = points[points.length - 1]?.date ? clientFormatDateValue(points[points.length - 1].date) : '';
  const actualCircle = (() => {
    const lastActualIndex = points.map((point, index) => ({ point, index })).filter((item) => item.point.actual != null).pop();
    if (!lastActualIndex) return '';
    const x = pad.left + (lastActualIndex.index / Math.max(1, points.length - 1)) * innerW;
    const y = pad.top + (1 - clampClientPercent(lastActualIndex.point.actual) / 100) * innerH;
    return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="5" class="client-chart-dot" />`;
  })();
  return `
    <svg class="client-scurve-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Curva S planejado versus realizado">
      ${grid}
      <line x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <path d="${plannedPath}" class="client-chart-planned" />
      ${actualPath ? `<path d="${actualPath}" class="client-chart-actual" />${actualCircle}` : ''}
      <text x="${pad.left}" y="${height - 12}" class="client-chart-date">${escapeHtml(first)}</text>
      <text x="${pad.left + innerW / 2 - 38}" y="${height - 12}" class="client-chart-date">${escapeHtml(mid)}</text>
      <text x="${width - pad.right - 72}" y="${height - 12}" class="client-chart-date">${escapeHtml(last)}</text>
    </svg>
  `;
}

function renderClientGauge(percent, label, plannedPercent = null, options = {}) {
  const p = clampClientPercent(percent);
  const plannedBase = plannedPercent == null ? p : clampClientPercent(plannedPercent);
  const planned = Math.max(p, plannedBase);
  const deviation = Math.max(0, planned - p);
  const deliveryDate = options?.deliveryDate ? clientFormatDateValue(options.deliveryDate) : '';
  const noteText = options?.note ? String(options.note) : '';
  const ringStyle = `background: conic-gradient(#0b9b7a 0 ${p}%, #efc14f ${p}% ${planned}%, #d4dce2 ${planned}% 100%)`;
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

function buildClientExecutiveSchedule(project) {
  const start = getClientAnalyticStartDate(project);
  const finish = getClientAnalyticFinishDate(project) || getProjectShipmentDate(project);
  const totalBusinessDays = Math.max(5, countBusinessDaysInclusive(start, finish) || 116);

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
    { key: 'hydro', label: 'Hydro Testing', base: 6, percent: getClientStageValue(project, ['Hydro Test Pressure (QC)']) },
    { key: 'painting', label: 'Painting', base: 14, percent: getClientStageValue(project, ['Surface preparation and/or coating', 'HDG / FBE.  (PAINT)']) },
  ];
  const logisticsTemplate = [
    { key: 'packing', label: 'Packing', base: 1, percent: getClientStageValue(project, ['Package and Delivered']) },
    { key: 'final-inspection', label: 'Final Inspection', base: 1, percent: getClientStageValue(project, ['Final Inspection']) },
    { key: 'delivery', label: 'Delivery', base: 1, percent: getClientPackageProgress(project) },
  ];
  const fabricationGroup = groupDurations.find((item) => item.key === 'fabrication');
  const deliveryGroup = groupDurations.find((item) => item.key === 'delivery');
  const fabricationDurations = scaleDurationVector(fabricationTemplate, fabricationGroup?.duration || 81);
  const logisticsDurations = scaleDurationVector(logisticsTemplate, deliveryGroup?.duration || 3);

  const rows = [];
  let cursor = parseDateObject(start) || parseDateObject(getProjectShipmentDate(project)) || new Date();
  for (const group of groupDurations) {
    const groupStart = new Date(cursor.getTime());
    const groupFinish = addBusinessDaysUtc(groupStart, Math.max(0, (group.duration || 1) - 1)) || groupStart;
    rows.push({ type: 'group', label: group.label, progress: group.percent, duration: group.duration, start: groupStart, finish: groupFinish });
    let children = [];
    if (group.key === 'engineering') {
      children = [{ label: 'Drawings', progress: group.percent, duration: group.duration }];
    } else if (group.key === 'procurement') {
      children = [{ label: 'Materials for Application Acquisition', progress: group.percent, duration: group.duration }];
    } else if (group.key === 'fabrication') {
      children = fabricationDurations.map((item) => ({ label: item.label, progress: item.percent, duration: item.duration }));
    } else if (group.key === 'delivery') {
      children = logisticsDurations.map((item) => ({ label: item.label, progress: item.percent, duration: item.duration }));
    }

    let childCursor = new Date(groupStart.getTime());
    for (const child of children) {
      const childStart = new Date(childCursor.getTime());
      const childFinish = addBusinessDaysUtc(childStart, Math.max(0, (child.duration || 1) - 1)) || childStart;
      rows.push({ type: 'child', label: child.label, progress: child.progress, duration: child.duration, start: childStart, finish: childFinish });
      childCursor = addBusinessDaysUtc(childFinish, 1) || childFinish;
    }

    cursor = addBusinessDaysUtc(groupFinish, 1) || groupFinish;
  }
  return rows;
}

function getClientScheduleVisualState(progress) {
  const value = clampClientPercent(progress);
  if (value >= 99.9) return 'completed';
  if (value <= 0) return 'not-started';
  return 'in-progress';
}

function renderClientExecutiveSchedule(project) {
  const rows = buildClientExecutiveSchedule(project);
  if (!rows.length) return '<div class="client-empty-state">Schedule não disponível para esta BSP.</div>';
  return `
    <div class="client-table-wrap client-table-wrap--compact client-exec-schedule-table">
      <table class="client-bsp-table client-bsp-table--schedule">
        <thead><tr><th>Etapa</th><th>%</th><th>Prazo médio</th><th>Início</th><th>Término</th><th>Status</th></tr></thead>
        <tbody>
          ${rows.map((row) => {
            const state = getClientScheduleVisualState(row.progress);
            const label = row.type === 'group' ? `<strong>${escapeHtml(row.label)}</strong>` : `<span class="client-schedule-child">${escapeHtml(row.label)}</span>`;
            return `<tr class="client-schedule-row client-schedule-row--${state} client-schedule-row--${row.type}"><td>${label}</td><td><span class="client-spool-progress client-spool-progress--${state}">${formatPercent(row.progress)}</span></td><td>${formatNumber(row.duration, 0)}d</td><td>${formatClientDateShort(row.start)}</td><td>${formatClientDateShort(row.finish)}</td><td><span class="client-spool-chip client-spool-chip--${state}">${state === 'completed' ? 'Concluído' : state === 'in-progress' ? 'Em andamento' : 'Não iniciado'}</span></td></tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function getClientStageTimeline(project) {
  const stages = getClientProductionStages(project);
  return stages.map((stage) => ({
    ...stage,
    state: stage.percent >= 100 ? 'done' : stage.percent > 0 ? 'active' : 'future',
  }));
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
    const pdfButton = event.target.closest('[data-client-download-pdf]');
    if (pdfButton) {
      event.preventDefault();
      event.stopPropagation();
      handleClientExecutivePdfDownload(pdfButton);
      return;
    }
    if (event.target.closest('[data-client-exec-close]')) closeClientBspExecutive();
  });
  return modal;
}

function closeClientBspExecutive() {
  const modal = document.getElementById('client-bsp-executive-modal');
  if (modal) modal.classList.add('hidden');
  document.body.classList.remove('client-exec-open');
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
  const width = 760;
  const height = 260;
  const pad = { left: 42, right: 16, top: 18, bottom: 38 };
  const innerW = width - pad.left - pad.right;
  const innerH = height - pad.top - pad.bottom;
  const plannedPath = clientSvgPolyline(points, width, height, (point) => point.planned);
  const actualPath = clientSvgPolyline(points, width, height, (point) => point.actual);
  const grid = [0, 25, 50, 75, 100].map((value) => {
    const y = pad.top + (1 - value / 100) * innerH;
    return `<line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" class="client-chart-grid" /><text x="8" y="${y + 4}" class="client-chart-label">${value}%</text>`;
  }).join('');
  const first = points[0]?.date ? clientFormatDateValue(points[0].date) : '';
  const mid = points[Math.floor(points.length / 2)]?.date ? clientFormatDateValue(points[Math.floor(points.length / 2)].date) : '';
  const last = points[points.length - 1]?.date ? clientFormatDateValue(points[points.length - 1].date) : '';
  const actualCircle = (() => {
    const lastActualIndex = points.map((point, index) => ({ point, index })).filter((item) => item.point.actual != null).pop();
    if (!lastActualIndex) return '';
    const x = pad.left + (lastActualIndex.index / Math.max(1, points.length - 1)) * innerW;
    const y = pad.top + (1 - clampClientPercent(lastActualIndex.point.actual) / 100) * innerH;
    return `<circle cx="${x.toFixed(1)}" cy="${y.toFixed(1)}" r="5" class="client-chart-dot" />`;
  })();
  return `
    <svg class="client-scurve-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Curva S macro planejado versus realizado">
      ${grid}
      <line x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${height - pad.bottom}" class="client-chart-axis" />
      <path d="${plannedPath}" class="client-chart-planned" />
      ${actualPath ? `<path d="${actualPath}" class="client-chart-actual" />${actualCircle}` : ''}
      <text x="${pad.left}" y="${height - 12}" class="client-chart-date">${escapeHtml(first)}</text>
      <text x="${pad.left + innerW / 2 - 38}" y="${height - 12}" class="client-chart-date">${escapeHtml(mid)}</text>
      <text x="${width - pad.right - 72}" y="${height - 12}" class="client-chart-date">${escapeHtml(last)}</text>
    </svg>
  `;
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
    return `<tr class="client-spool-row client-spool-row--${visualState}"><td><strong>${escapeHtml(getClientProjectDisplayCode(project))}</strong></td><td>${escapeHtml(getProjectVesselLabel(project) || '—')}</td><td>${formatNumber(getProjectItemCount(project))}</td><td>${formatNumber(project.kilos, 0)} kg</td><td>${formatNumber(project.weldedWeightKg, 0)} kg</td><td><span class="client-spool-chip client-spool-chip--${visualState}">${escapeHtml(status.text)}</span></td><td><span class="client-spool-progress client-spool-progress--${visualState}">${formatPercent(getClientOverallProgress(project))}</span></td><td>${escapeHtml(project.plannedFinishDate || '—')}</td></tr>`;
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
  const detailStageKeys = ['Drawing Execution Advance%', 'Procuremnt Status %', 'Material Separation', 'Full welding execution', 'Non Destructive Examination (QC)', 'Hydro Test Pressure (QC)', 'Surface preparation and/or coating', 'Final Inspection', 'Package and Delivered'];
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
    columnStyles: { 0: { cellWidth: 44 }, 1: { cellWidth: 24 }, 2: { cellWidth: 13, halign: 'center' }, 3: { cellWidth: 19, halign: 'right' }, 4: { cellWidth: 19, halign: 'right' }, 5: { cellWidth: 30 }, 6: { cellWidth: 16, halign: 'center' }, 7: { cellWidth: 21 } },
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
  const deviationPercent = Math.max(0, plannedToday - overall);
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
        ${renderClientGauge(overall, isUnitScope ? 'unidade' : 'carteira', plannedToday, { note: 'Meta macro até hoje' })}
      </section>
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Fabrication Progress</h3><span>${isUnitScope ? 'Fabricação ponderada da unidade' : 'Fabricação ponderada da carteira'}</span></div>
        ${renderClientGauge(fabrication, 'fabricação', plannedToday, { note: 'Meta macro até hoje' })}
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

    <section class="client-exec-card client-exec-attention">
      <div class="client-exec-card-head"><h3>S-Curve | Attention Points</h3><span>${isUnitScope ? 'Análise automática da unidade' : 'Análise automática da carteira'}</span></div>
      <ul>${attention.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
    </section>

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
}

function openClientBspExecutive(project) {
  if (!project || !isClientUser()) return;
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
  const deviationPercent = Math.max(0, plannedToday - overall);
  const stageValues = project.stageValues || {};
  const spools = Array.isArray(project.spools) ? project.spools : [];
  const detailStageKeys = ['Drawing Execution Advance%', 'Procuremnt Status %', 'Material Separation', 'Full welding execution', 'Non Destructive Examination (QC)', 'Hydro Test Pressure (QC)', 'Surface preparation and/or coating', 'Final Inspection', 'Package and Delivered'];

  content.innerHTML = `
    <header class="client-exec-header">
      <div>
        <p class="client-kicker">Visão Executiva da BSP</p>
        <h2>${escapeHtml(getClientProjectDisplayCode(project))}</h2>
        <p>${escapeHtml(getProjectClientLabel(project))} • ${escapeHtml(getProjectVesselLabel(project))} • <span class="cell-status cell-status--${status.state}">${escapeHtml(status.text)}</span></p>
        <div class="client-exec-header-actions"><button class="client-exec-pdf-button" type="button" data-client-download-pdf data-client-report-type="project" data-client-report-project-id="${escapeHtml(project.rowId)}">Baixar PDF</button></div>
      </div>
      <div class="client-exec-dates">
        <span>Início: <strong>${escapeHtml(startDate || '—')}</strong></span>
        <span>Término planejado: <strong>${escapeHtml(finishDate || '—')}</strong></span>
        <span>Planejado hoje: <strong>${formatPercent(plannedToday)}</strong></span>
        <span>Desvio: <strong>${formatPercent(deviationPercent)}</strong></span>
        <span>Envio efetivo: <strong>${escapeHtml(shipmentDate || '—')}</strong></span>
      </div>
    </header>

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
        ${renderClientGauge(overall, 'concluído', plannedToday, { deliveryDate: shipmentDate, note: 'Meta até hoje' })}
      </section>
      <section class="client-exec-card">
        <div class="client-exec-card-head"><h3>Fabrication Progress</h3><span>Fabricação ponderada + desvio</span></div>
        ${renderClientGauge(fabrication, 'fabricação', plannedToday, { note: 'Meta até hoje' })}
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
        <div class="client-exec-legend"><span><i class="planned"></i> Planejado</span><span><i class="actual"></i> Realizado</span></div>
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
      <div class="client-exec-card-head"><h3>Schedule Executivo da BSP</h3><span>Prazos médios por etapa com base na data inicial e final</span></div>
      ${renderClientExecutiveSchedule(project)}
    </section>

    <section class="client-exec-card client-exec-attention">
      <div class="client-exec-card-head"><h3>S-Curve | Attention Points</h3><span>Análise automática</span></div>
      <ul>${attention.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
    </section>

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
        <table class="client-bsp-table"><thead><tr><th>Tag/ISO</th><th>Descrição</th><th>Status</th><th>Etapa</th><th>%</th><th>Peso</th></tr></thead><tbody>
          ${renderClientSpoolRows(spools, 120)}
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
          <p>Gere uma chave somente leitura para consultar as informações que já aparecem neste painel. Esta API não edita, não altera e não exclui dados.</p>
        </div>
        <button class="client-api-close" type="button" data-client-api-close aria-label="Fechar">×</button>
      </div>
      <div class="client-api-body">
        <section class="client-api-card">
          <div class="client-api-card-head">
            <div>
              <h3>Nova chave de leitura</h3>
              <p>O token completo aparece somente uma vez. Copie e guarde em local seguro. Permissão: somente consulta.</p>
            </div>
          </div>
          <div class="client-api-form-row">
            <label>
              <span>Nome da integração</span>
              <input id="client-api-key-name" type="text" value="API de leitura do Portal do Cliente" maxlength="120" />
            </label>
            <button class="mini-action-button" type="button" data-client-api-create>Criar API somente leitura</button>
          </div>
          <div id="client-api-new-token" class="client-api-token-box hidden"></div>
        </section>
        <section class="client-api-card">
          <div class="client-api-card-head">
            <div>
              <h3>Como consumir</h3>
              <p>Use o endpoint abaixo com o header Authorization. Método permitido: GET.</p>
            </div>
          </div>
          <pre class="client-api-code" id="client-api-example"></pre>
          <div class="client-api-help-grid">
            <span><strong>Somente leitura:</strong> não aceita POST, PUT, PATCH ou DELETE</span>
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
        <span>Copie sua chave somente leitura agora:</span>
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
        <strong>Token somente leitura criado agora</strong>
        <span>Copie este token completo. Ele não será exibido novamente depois que fechar esta janela.</span>
        <code>${escapeHtml(rawToken)}</code>
        <small>Use em Authorization: Bearer ${escapeHtml(rawToken)} • Método permitido: GET</small>
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
          <small>Criada: ${escapeHtml(created)} • Último uso: ${escapeHtml(used)} • ${escapeHtml(status)} • Somente leitura</small>
        </div>
        <div class="client-api-key-actions">
          ${isNewKey ? '<button class="mini-action-button" type="button" data-client-api-copy-token>Copiar token</button>' : ''}
          ${key.active === false
            ? `<button class="mini-action-button mini-action-button--danger" type="button" data-client-api-delete="${escapeHtml(key.id)}">Excluir</button>`
            : `<button class="mini-action-button mini-action-button--danger" type="button" data-client-api-revoke="${escapeHtml(key.id)}">Revogar</button>`}
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

function closeClientApiModal() {
  const modal = document.getElementById('client-api-modal');
  if (modal) modal.classList.add('hidden');
  document.body.classList.remove('client-api-open');
  state.clientApi.newToken = '';
  state.clientApi.newTokenKeyId = '';
}

async function createClientApiKeyFromModal() {
  const input = document.getElementById('client-api-key-name');
  const name = input?.value || 'API de leitura do Portal do Cliente';
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
    state.clientApi.feedback = 'API somente leitura criada com sucesso. Copie o token agora.';
    renderClientApiExample(state.clientApi.newToken || '<SUA_API_KEY>');
    await loadClientApiKeys();
    state.clientApi.feedback = 'API somente leitura criada com sucesso. Copie o token agora.';
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

async function deleteClientApiKeyFromModal(id) {
  if (!id) return;
  const confirmed = window.confirm('Excluir definitivamente esta chave revogada da lista?');
  if (!confirmed) return;
  state.clientApi.feedback = 'Excluindo chave revogada...';
  renderClientApiKeys();
  try {
    const response = await fetch('/api/client-api-keys', {
      method: 'DELETE',
      credentials: 'same-origin',
      cache: 'no-store',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, action: 'delete' }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao excluir API.');
    state.clientApi.feedback = 'Chave excluída da lista.';
    await loadClientApiKeys();
  } catch (error) {
    state.clientApi.feedback = error?.message || 'Falha ao excluir API.';
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
    return;
  }
  const deleteButton = event.target.closest('[data-client-api-delete]');
  if (deleteButton) {
    deleteClientApiKeyFromModal(deleteButton.dataset.clientApiDelete || '');
  }
}

function handleClientDashboardClick(event) {
  const apiButton = event.target.closest('[data-client-open-api]');
  if (apiButton) {
    openClientApiModal();
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

function getProjectSectorTagCount(project, sectorKeys = []) {
  const expected = new Set(sectorKeys);
  if (!expected.size) return 0;
  return getProjectOpenFlowItems(project).filter((item) => expected.has(getFlowSectorKey(item.flow))).length;
}

function getProjectProductionTags(project) {
  return getProjectSectorTagCount(project, ['producao', 'solda', 'calderaria']);
}

function getProjectInspectionTags(project) {
  return getProjectSectorTagCount(project, ['inspecao']);
}

function getProjectPaintingTags(project) {
  return getProjectSectorTagCount(project, ['pintura']);
}

function getProjectStartedTagsForDrill(project) {
  const snapshot = isProjectStartedForStats(project);
  return snapshot.started ? Number(snapshot.tags || getProjectItemCount(project) || 0) : 0;
}

function getProjectNotStartedTagsForDrill(project) {
  if (!project || isProjectExcludedFromTotal(project)) return 0;
  const snapshot = isProjectStartedForStats(project);
  if (snapshot.started) return 0;
  const openItems = getProjectOpenFlowItems(project);
  return openItems.length || Number(project.quantitySpools || getProjectItemCount(project) || 1);
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
    case 'awaiting': return getAwaitingShipmentTags(project);
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

  const active = source.filter((project) => !isProjectExcludedFromTotal(project));

  if (mode === 'started') {
    return active.filter((project) => getProjectStartedTagsForDrill(project) > 0);
  }

  if (mode === 'not-started') {
    return active.filter((project) => getProjectNotStartedTagsForDrill(project) > 0);
  }

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
    return active.filter((project) => getAwaitingShipmentTags(project) > 0);
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

  return active;
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
    group.awaitingTags += getAwaitingShipmentTags(project);
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
    group.awaitingTags += getAwaitingShipmentTags(project);
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
    if (mode === 'awaiting') return getAwaitingShipmentTags(b) - getAwaitingShipmentTags(a) || compareProjectsByPlannedFinishDate(a, b);
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
      return `<td>${formatNumber(getAwaitingShipmentTags(project))}</td><td>${escapeHtml(getProjectShipmentDate(project) || '—')}</td>`;
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

function renderStats() {
  const stats = buildClientStats(getStatsProjectsSource());
  state.visibleStats = stats;
  const totalFinishedWeight = getTotalFinishedWeightAllProjects();
  const totalWeldedWeight = Number(stats.totalWeldedWeightKg || 0);
  const totalBacklogWelding = Math.max(0, Number(stats.totalWeightKg || 0) - totalWeldedWeight);
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
      </div>
    </div>
  `;

  const button = document.getElementById("open-selected-project");
  if (button) {
    button.addEventListener("click", () => openProjectModal(project));
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

    ${renderProjectSignals(project)}

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

function isStageUpdatesWorkspaceOpen() {
  return Boolean(stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden'));
}

function shouldSkipBackgroundRequest(options = {}) {
  return !options.force && isPageHidden();
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
  return loadProjects({
    force,
    background: true,
    skipLocalCache: true,
    suppressLoadingState: true,
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

function writeProjectsCache(payload) {
  try {
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
  return false;
}

function applyProjectsPayload(data, options = {}) {
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
  if (state.user && sectorAlertsModalEl && !sectorAlertsModalEl.classList.contains('hidden')) {
    renderManualAlerts();
  }
}

function updateMeta() {
  if (!state.meta) return;
  sheetNameEl.textContent = state.meta.clientPortal ? 'Base operacional' : (state.meta.sheetName || "Smartsheet");
  lastSyncEl.textContent = `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`;
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
  const shouldUseCache = !force && cached?.payload && !shouldIgnoreCachedProjectsPayload(cached);
  if (shouldUseCache) {
    applyProjectsPayload(cached.payload, { fromCache: true });
    const fresh = isProjectsCacheFresh(cached);
    if (lastSyncEl && state.meta?.lastSync) {
      lastSyncEl.textContent = `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")} • exibindo cache local${fresh ? '' : ' enquanto atualiza'}`;
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
      const preferServerCache = Boolean(options.preferServerCache || (!background && !force));
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
      state.lastProjectsFetchAt = Date.now();
      writeProjectsCache(data);
      applyProjectsPayload(data, { fromCache: false });
      return {
        ok: true,
        fromCache: false,
        revalidating: false,
        projectsCount: projectsFromApi.length,
      };
    } catch (error) {
      const fallbackMessage = error?.message || "Falha ao atualizar dados operacionais.";

      if (options.requireData) {
        throw error;
      }

      if (state.projects.length) {
        const staleSuffix = state.meta?.lastSync
          ? ` | exibindo última atualização válida: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`
          : "";
        lastSyncEl.textContent = `Conexão instável com os dados operacionais${staleSuffix}`;
        console.warn("Falha temporária ao atualizar projetos:", fallbackMessage);
        return;
      }

      bodyEl.innerHTML = `<tr><td colspan="21" class="loading-cell">${escapeHtml(fallbackMessage)}</td></tr>`;
      detailCardEl.innerHTML = `<div class="detail-placeholder">${escapeHtml(fallbackMessage)}</div>`;
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

function bindEvents() {
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
    refreshProjectsButtonEl.addEventListener('click', () => {
      clearProjectsCache();
      loadProjects({ force: true }).catch((error) => window.alert(error?.message || 'Falha ao atualizar agora.'));
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

function hasDashboardDataReady() {
  const projects = Array.isArray(state.projects) ? state.projects : [];
  if (!state.user || projects.length === 0) return false;

  if (isClientUser()) {
    const bspsText = String(document.getElementById('client-stat-bsps')?.textContent || '').trim();
    const vesselGrid = document.getElementById('client-vessel-grid');
    const hasVisibleClientCards = Boolean(vesselGrid && vesselGrid.querySelector('[data-client-vessel]'));
    return bspsText !== '' && bspsText !== '--' && bspsText !== '0' && hasVisibleClientCards;
  }

  const countText = String(searchCountEl?.textContent || '').trim();
  const hasTableRows = Boolean(bodyEl && bodyEl.querySelector('tr[data-project-id]'));
  return hasTableRows || (Array.isArray(state.filteredProjects) && state.filteredProjects.length > 0 && countText !== '0 resultado(s)');
}

function waitForNextRenderFrame() {
  return new Promise((resolve) => window.requestAnimationFrame(() => window.requestAnimationFrame(resolve)));
}

async function ensureDashboardDataReadyBeforeRelease(options = {}) {
  const maxAttempts = Number(options.maxAttempts || 3);
  const retryDelayMs = Number(options.retryDelayMs || 900);

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    await waitForNextRenderFrame();
    if (hasDashboardDataReady()) return true;

    setLoginProgress(Math.min(96, 88 + attempt * 2), {
      title: 'Conferindo dados na tela...',
      message: 'Estamos garantindo que as BSPs, POs e dashboards já apareceram no painel.',
      detail: `Validação visual dos dados ${attempt}/${maxAttempts}.`,
    });

    await loadProjects({
      force: false,
      skipLocalCache: false,
      suppressLoadingState: true,
      preferServerCache: true,
      requireData: false,
    });
    await waitForNextRenderFrame();
    if (hasDashboardDataReady()) return true;

    await new Promise((resolve) => window.setTimeout(resolve, retryDelayMs));
  }

  throw new Error('Os dados ainda não apareceram na tela. O painel não será liberado vazio; tente novamente em alguns segundos.');
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
    openAdminPanelEl.classList.add("hidden");
    if (openLoginButtonEl) openLoginButtonEl.classList.remove("hidden");
    setClientDashboardMode();
    return;
  }

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
              TEMP
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
  syncOperationRegionButtons('PT');
  setSelectedAdminAlertSectors([document.getElementById("admin-user-sector")?.value || "pintura"]);
  state.adminProjectPmSearchQuery = "";
  const projectPmSearchEl = document.getElementById("admin-user-project-pms-search");
  if (projectPmSearchEl) projectPmSearchEl.value = "";
  setAdminProjectPmAliases([]);
  setAdminQualityCompetencies([]);
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = '';
  if (adminUserClientNameEl) adminUserClientNameEl.value = '';
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = '';
  if (adminUserClientLogoUrlEl) adminUserClientLogoUrlEl.value = '';
  resetAdminLogoEditor('');
  if (adminUserClientLogoFileEl) adminUserClientLogoFileEl.value = '';
  if (adminUserClientPlatformImageUrlEl) adminUserClientPlatformImageUrlEl.value = '';
  if (adminUserClientPlatformNameEl) adminUserClientPlatformNameEl.value = '';
  if (adminUserClientPlatformImagesEl) adminUserClientPlatformImagesEl.value = '';
  if (adminUserClientPlatformImageFileEl) adminUserClientPlatformImageFileEl.value = '';
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
  syncOperationRegionButtons(user?.operationRegion || user?.siteKey || user?.portalSite || 'PT');
  document.getElementById("admin-user-sector").value = user.role === "client" ? "all" : (user.sector || "all");
  setSelectedAdminAlertSectors(Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector]);
  state.adminProjectPmSearchQuery = "";
  const projectPmSearchEl = document.getElementById("admin-user-project-pms-search");
  if (projectPmSearchEl) projectPmSearchEl.value = "";
  setAdminProjectPmAliases(user.projectPmAliases || []);
  setAdminQualityCompetencies(user.qualityCompetencies || []);
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = user.clientKey || '';
  if (adminUserClientNameEl) adminUserClientNameEl.value = user.clientName || '';
  if (adminUserClientKeyEl) adminUserClientKeyEl.value = user.clientKey || buildClientKey(user.clientName || '', getOperationRegion(user));
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
  renderAdminAutoLoginList(users);
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
function getAutoLoginRegion() {
  const active = document.querySelector('[data-auto-login-region].is-active');
  const raw = String(active?.dataset?.autoLoginRegion || adminUserOperationRegionEl?.value || getOperationRegion()).toUpperCase();
  return raw === 'BR' ? 'BR' : 'PT';
}

function buildAutoLoginUrl(user, region = getAutoLoginRegion()) {
  const url = new URL(window.location.href);
  url.searchParams.set('login', String(user?.username || '').replace(/__(BR|PT)$/i, ''));
  url.searchParams.set('region', region === 'BR' ? 'BR' : 'PT');
  url.searchParams.set('autoLogin', '1');
  return url.toString();
}

function renderAdminAutoLoginList(users = adminUsersListEl?._cachedUsers || []) {
  if (!adminAutoLoginListEl) return;
  const region = getAutoLoginRegion();
  const filtered = (Array.isArray(users) ? users : [])
    .filter((user) => user && user.active !== false)
    .filter((user) => {
      if (user.role === 'admin') return true;
      const sector = normalizeSectorValue(user.sector);
      if (sector === 'pcp') return true;
      const userRegion = String(user.operationRegion || user.siteKey || user.portalSite || 'PT').toUpperCase();
      return userRegion === region;
    });

  if (!filtered.length) {
    adminAutoLoginListEl.innerHTML = `<div class="empty-state">Nenhum usuário encontrado para ${region}.</div>`;
    return;
  }

  adminAutoLoginListEl.innerHTML = filtered.map((user) => {
    const visibleLogin = String(user.username || '').replace(/__(BR|PT)$/i, '');
    const url = buildAutoLoginUrl(user, region);
    return `
      <article class="admin-list-item">
        <div class="admin-user-title-row">
          <strong>${escapeHtml(user.name || visibleLogin)}</strong>
          <span class="admin-badge">${escapeHtml(region)}</span>
        </div>
        <div class="admin-list-item-meta">
          <span>Login visível: ${escapeHtml(visibleLogin)}</span>
          <span>Perfil: ${escapeHtml(user.role === 'client' ? 'Cliente' : (user.role === 'admin' ? 'Admin' : sectorLabel(user.sector)))}</span>
          ${user.clientKey ? `<span>Client Key: ${escapeHtml(user.clientKey)}</span>` : ''}
        </div>
        <div class="manual-alert-actions">
          <button class="ghost-button ghost-button--compact" type="button" data-copy-auto-login="${escapeHtml(url)}">Copiar link</button>
          <a class="ghost-button ghost-button--compact" href="${escapeHtml(url)}" target="_blank" rel="noopener">Abrir</a>
        </div>
      </article>
    `;
  }).join('');
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
  prewarmProjectsApi();
  try {
    const response = await fetch("/api/auth-login", {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        username: String(loginUsernameEl.value || "").trim(),
        password: String(loginPasswordEl.value || "").trim(),
        operationRegion: getOperationRegion(),
        siteKey: getOperationRegion(),
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

    if (!readProjectsCache()?.payload) {
      await waitForProjectsWarmup(2000);
    }

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
      title: 'Carregando POs...',
      message: 'Estamos carregando as POs, demandas e referências de fabricação.',
      detail: 'POs e demandas em processamento.',
    });
    await projectsPromise;
    setLoginProgress(88, {
      title: 'Definindo dashboards...',
      message: 'Estamos definindo os dashboards e montando a visualização final.',
      detail: 'Dashboard quase pronto.',
    });

    await ensureDashboardDataReadyBeforeRelease({ maxAttempts: 1, retryDelayMs: 300 });

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

async function handleLogout() {
  await fetch("/api/auth-logout", { credentials: "same-origin" });
  state.user = null;
  // Mantém o cache local de projetos entre logout e novo login.
  // A chave já é escopada por papel/usuário/cliente; apagá-la aqui fazia o próximo login
  // perder o carregamento imediato e voltar a depender do Smartsheet/Supabase.
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
  window.clearInterval(state.pollTimer);
  stopPresenceHeartbeat();
  updateSessionUi();
  resetDashboardForLoggedOutState();
  openLoginModal("Sessão encerrada. Faça login novamente para acessar o painel.");
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
    const response = await fetch('/api/stage-updates', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar apontamentos setoriais.');
    state.lastStageUpdatesFetchAt = Date.now();
    state.stageUpdates = Array.isArray(data.updates) ? data.updates : [];
    if (normalizeSectorValue(state.user?.sector) === 'pcp') {
      const pendingForPcp = state.stageUpdates.filter((item) => isPendingStageStatus(item?.status));
      syncIncomingAlerts('stageUpdates', pendingForPcp);
    }
  } catch (error) {
    // v35.7: não bloquear o apontamento do setor quando a consulta do histórico/validação falhar.
    // O setor ainda precisa conseguir abrir a tela, buscar a BSP e enviar o apontamento.
    state.stageUpdates = [];
    state.lastStageUpdatesFetchAt = Date.now();
    if (stageUpdatesContentEl && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      renderStageUpdatesModal();
      const warning = document.createElement('div');
      warning.className = 'stage-inline-warning';
      warning.textContent = error.message || 'Histórico de apontamentos indisponível no momento; você ainda pode lançar novos apontamentos.';
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
                              <div class="stage-row-form" data-stage-update-form="true" data-project-row-id="${escapeHtml(String(projectRowId || ''))}" data-project-number="${escapeHtml(project.projectNumber || '')}" data-spool-iso="${escapeHtml(spool.iso || '')}" data-stage-sector="${escapeHtml(sector || '')}">
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
  const pending = filtered.filter((item) => isPendingStageStatus(item.status));
  const history = filtered.filter((item) => isResolvedStageStatus(item.status)).slice(0, 80);
  const selectable = pending.filter(isStageUpdateSelectableForTracking);
  const selectedIds = getSelectedVisibleStageIds(pending);
  const selectedDateIds = (state.stageDateSelectedIds || []).filter((id) => (state.stageDatePendencies || []).some((item) => String(item.id) === String(id)));
  state.stageSelectedIds = selectedIds;
  state.stageDateSelectedIds = selectedDateIds;
  const submitting = Boolean(state.stageTrackingSubmitting);

  stageUpdatesContentEl.innerHTML = `
    <div class="stage-workspace-shell stage-validation-workspace" id="stage-validation">
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
  if (!progress && !completionDate && !note) {
    removeStageDraft(projectRowId, spoolIso, sector);
    return;
  }
  upsertStageDraft(projectRowId, spoolIso, sector, { progress, completionDate, note });
}

async function handleStageWorkspaceBulkSubmit() {
  const sector = getStageWorkspaceSector();
  const items = getReadyStageDraftEntries(sector).map((item) => ({
    projectRowId: item.projectRowId,
    spoolIso: item.spoolIso,
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
  upsertStageDraft(projectRowId, spoolIso, sector, { progress, completionDate, note, actionType });
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
      body: JSON.stringify({ projectRowId, spoolIso, progress, completionDate, note, sector, actionType }),
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
      operationRegion: String(adminUserOperationRegionEl?.value || 'PT').toUpperCase() === 'BR' ? 'BR' : 'PT',
      siteKey: String(adminUserOperationRegionEl?.value || 'PT').toUpperCase() === 'BR' ? 'BR' : 'PT',
      portalSite: String(adminUserOperationRegionEl?.value || 'PT').toUpperCase() === 'BR' ? 'BR' : 'PT',
      sector: document.getElementById("admin-user-role").value === 'client' ? 'all' : document.getElementById("admin-user-sector").value,
      alertSectors: document.getElementById("admin-user-role").value === 'client' ? [] : getSelectedAdminAlertSectors(),
      projectPmAliases: adminUserFormHasProjectsScope() ? getAdminProjectPmAliases() : [],
      qualityCompetencies: adminUserFormHasQualityScope() ? getAdminQualityCompetencies() : [],
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
      operationRegion: payload.operationRegion || 'BR',
      sector: payload.role === "admin" ? "all" : payload.sector,
      alertSectors: payload.role === "admin" ? [] : payload.alertSectors,
      projectPmAliases: payload.role === "admin" ? [] : payload.projectPmAliases,
      qualityCompetencies: payload.role === "admin" ? [] : payload.qualityCompetencies,
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
  prewarmProjectsApi();
  const authenticated = await bootstrapSession();
  if (authenticated) {
    const autoOpenStageValidation = shouldOpenStageValidationWorkspaceFromUrl() && canValidateStageWorkspace();
    if (autoOpenStageValidation) {
      state.stageUpdatesSearchQuery = '';
      openStageUpdatesModal({ loading: true });
    }
    setProjectsLoadingState('Carregando painel...');
    await loadProjects();
    syncStageDraftsForCurrentSector();
    startPostSessionBackgroundLoads({ autoOpenStageValidation });
    startPolling();
  }
}

init();
