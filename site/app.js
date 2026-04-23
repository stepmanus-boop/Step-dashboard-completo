const DEFAULT_POLL_MS = 10000;
const ALERT_NOTIFICATION_COOLDOWN_MS = 4 * 60 * 60 * 1000;

let adminResponsesPollTimer = null;

const state = {
  projects: [],
  filteredProjects: [],
  projectView: 'all',
  sectorScopedView: false,
  stats: null,
  meta: null,
  alerts: [],
  searchQuery: "",
  demandFilter: "",
  weekFilter: "",
  alertFilter: "all",
  alertSectorFilter: "all",
  alertClientQuery: "",
  selectedProjectId: null,
  modalPendingOnly: false,
  rowClickTimer: null,
  pollTimer: null,
  user: null,
  githubSyncEnabled: false,
  manualAlerts: [],
  projectSignals: [],
  adminAlertSearchQuery: "",
  adminActiveTab: "usuario",
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
};

const bodyEl = document.getElementById("projects-body");
const detailCardEl = document.getElementById("detail-card");
const sheetNameEl = document.getElementById("sheet-name");
const lastSyncEl = document.getElementById("last-sync");
const footerVersionEl = document.getElementById("footer-version");
const searchInputEl = document.getElementById("project-search");
const clearSearchEl = document.getElementById("clear-search");
const demandFilterEl = document.getElementById("demand-filter");
const weekFilterEl = document.getElementById("week-filter");
const searchCountEl = document.getElementById("search-count");
const tableShellEl = document.getElementById("table-shell");
const projectViewTabsEl = document.getElementById("project-view-tabs");
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
const toggleLoginPasswordEl = document.getElementById("toggle-login-password");
const loginCloseEl = document.getElementById("login-close");
const sessionUserNameEl = document.getElementById("session-user-name");
const sessionUserMetaEl = document.getElementById("session-user-meta");
const sessionStatusEl = document.getElementById("session-status");
const logoutButtonEl = document.getElementById("logout-button");
const openLoginButtonEl = document.getElementById("open-login-button");
const openSectorAlertsEl = document.getElementById("open-sector-alerts");
const openMyProjectSignalsEl = document.getElementById("open-my-project-signals");
const openProjectSignalsEl = document.getElementById("open-project-signals");
const openStageUpdatesEl = document.getElementById("open-stage-updates");
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
const adminAlertsListEl = document.getElementById("admin-alerts-list");
const adminAlertSearchEl = document.getElementById("admin-alert-search");
const githubSyncBadgeEl = document.getElementById("github-sync-badge");
const adminSyncButtonEl = document.getElementById("admin-sync-button");
const adminUserCancelEditEl = document.getElementById("admin-user-cancel-edit");
const adminUserTogglePasswordEl = document.getElementById("admin-user-toggle-password");
const adminUserIdEl = document.getElementById("admin-user-id");
const adminUserSubmitLabelEl = document.getElementById("admin-user-submit-label");
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

const installAppButtonEl = document.getElementById("install-app-button");
const connectionStatusEl = document.getElementById("connection-status");
let deferredInstallPrompt = null;


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

    const registration = await navigator.serviceWorker.register('./sw.js?v=5', { updateViaCache: 'none' });
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
      await registration.showNotification(title, { body, tag, icon: '/assets/icon-192.png', badge: '/assets/icon-192.png', data: { url: '/' } });
    } else {
      new Notification(title, { body, tag });
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
  const automaticAlerts = getUserAutomaticAlerts();
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
  { value: "pintura", label: "Pintura" },
  { value: "inspecao", label: "Inspeção" },
  { value: "pendente_envio", label: "Pendente de envio" },
  { value: "producao", label: "Produção" },
  { value: "calderaria", label: "Calderaria" },
  { value: "solda", label: "Solda" },
  { value: "pcp", label: "PCP" },
  { value: "projetos", label: "Projetos" },
];

function normalizeSectorValue(value) {
  const normalized = normalizeText(value)
    .replace(/[\s-]+/g, '_')
    .replace(/__+/g, '_');

  if (!normalized) return "";
  if (["envio", "pendenteenvio", "pendente_envio", "pendente_de_envio", "pending_shipment", "awaiting_shipment", "logistica", "logistics", "expedicao", "shipping"].includes(normalized)) return "pendente_envio";
  if (["inspecao", "inspection"].includes(normalized)) return "inspecao";
  if (["pintura", "painting", "coating"].includes(normalized)) return "pintura";
  if (["producao", "production"].includes(normalized)) return "producao";
  if (["calderaria", "boilermaker", "fabrication"].includes(normalized)) return "calderaria";
  if (["solda", "welding"].includes(normalized)) return "solda";
  if (["pcp", "planejamento", "planejamento_controle_producao", "planning", "planning_control"].includes(normalized)) return "pcp";
  if (["projetos", "projeto", "project", "projects", "pm"].includes(normalized)) return "projetos";
  if (normalized === "all") return "all";
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

function sectorLabel(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized === "pintura") return "Pintura";
  if (normalized === "inspecao") return "Inspeção";
  if (normalized === "pendente_envio") return "Pendente de envio";
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
    inspecao: 'Inspeção',
    pintura: 'Pintura',
    envio: 'Pendente de envio',
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


function projectDisplayWithClient(project) {
  const projectName = String(project?.projectDisplay || '').trim();
  const clientName = String(project?.client || '').trim();
  return clientName ? `${projectName} - ${clientName}` : (projectName || '—');
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


const STAGE_WORKSPACE_SECTORS = ['pintura', 'inspecao', 'pendente_envio', 'producao', 'calderaria', 'solda'];
const STAGE_PROGRESS_OPTIONS = [25, 50, 75, 100];

function getStageWorkspaceSector(user = state.user) {
  return normalizeSectorValue(user?.sector);
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
  const query = normalizeText(state.stageUpdatesSearchQuery || '');
  const source = Array.isArray(state.projects) ? state.projects : [];
  if (!query) return source.slice(0, 8);
  return source.filter((project) => {
    return [project.projectNumber, project.projectDisplay, project.client, project.currentStage]
      .some((value) => normalizeText(value).includes(query));
  }).slice(0, 8);
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
  if (stateValue === "in_progress") return "Em produção";
  return "Não iniciado";
}

function translateProjectStatus(projectStatus, uiState) {
  if (uiState === "completed") return "Finalizado";
  if (uiState === "awaiting_shipment") return "Aguardando envio";
  if (uiState === "not_started") return "Não iniciado";

  const normalized = String(projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (["ONGOING", "ON GOING", "IN PROGRESS", "EM PRODUCAO", "EM PRODUÇÃO"].includes(normalized)) {
    return "Em produção";
  }
  if (["ON HOLD", "HOLD", "PAUSED", "EM ESPERA"].includes(normalized)) {
    return uiState === "not_started" ? "Em espera" : "Em produção";
  }
  if (["COMPLETED", "DONE", "FINISHED", "CONCLUIDO", "CONCLUÍDO", "FINALIZADO"].includes(normalized)) {
    return "Finalizado";
  }
  return projectStatus || uiStateLabel(uiState);
}

function simplifyCurrentStage(project) {
  const uiState = String(project?.uiState || "").trim().toLowerCase();
  const sector = normalizeText(project?.operationalSector || "");
  const stage = normalizeText(project?.currentStage || "");

  if (
    uiState === "awaiting_shipment" ||
    uiState === "completed" ||
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
    return "Inspeção";
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

function enrichProjects(projects) {
  return (projects || []).map((project) => {
    const searchParts = [
      project.projectDisplay,
      project.projectNumber,
      project.projectPrefix,
      project.currentStage,
      project.projectStatus,
      project.client,
      ...(project.spools || []).flatMap((spool) => [spool.iso, spool.description, spool.drawing]),
    ];

    return {
      ...project,
      currentStageGroup: simplifyCurrentStage(project),
      _searchText: buildSearchIndex(searchParts),
    };
  });
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
        .map((project) => project.currentStageGroup || simplifyCurrentStage(project))
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

function getStatsProjectsSource() {
  return getVisibleProjectsSource();
}

function buildClientStats(projects) {
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
    const tags = Number(project.quantitySpools || 0);
    stats.totalSpools += tags;
    stats.totalWeightKg += Number(project.kilos || 0);
    stats.totalWeldedWeightKg += Number(project.weldedWeightKg || 0);
    stats.totalPaintingM2 += Number(project.m2Painting || 0);
    progressAccumulator += Number(project.overallProgress || 0);

    const stateValue = project.operationalState || project.uiState;
    const statusCandidates = [
      project?.projectStatus,
      project?.currentStage,
      project?.operationalState,
      project?.uiState,
    ].filter(Boolean).map((value) => String(value).trim().toLowerCase());
    const excludeFromCompletedCounts = Boolean(project?.projectFinishedFlag) || statusCandidates.some((value) => value.includes("project finished"));

    if (stateValue === "completed") {
      if (!excludeFromCompletedCounts) {
        stats.completed += 1;
        stats.completedTags += tags;
      }
    } else if (stateValue === "awaiting_shipment") {
      stats.awaitingShipment += 1;
      stats.awaitingShipmentTags += tags;
      if (!excludeFromCompletedCounts) {
        stats.completed += 1;
        stats.completedTags += tags;
      }
    } else if (stateValue === "in_inspection") {
      stats.inspectionProjects += 1;
      stats.inspectionTags += tags;
    } else if (stateValue === "in_production") {
      if (project.operationalSector === "Pintura") {
        stats.paintingProjects += 1;
        stats.paintingTags += tags;
      } else {
        stats.inProgress += 1;
        stats.inProgressTags += tags;
      }
    } else {
      const normalizedProjectStatus = String(project?.projectStatus || "").trim().toUpperCase().replace(/\s+/g, " ");
      const isHoldProject = ["ON HOLD", "HOLD", "PAUSED", "EM ESPERA"].includes(normalizedProjectStatus);

      stats.notStarted += 1;
      stats.notStartedTags += tags;

      if (isHoldProject) {
        stats.notStartedHold += 1;
        stats.notStartedHoldTags += tags;
      }
    }
  }

  stats.averageOverallProgress = projects.length ? progressAccumulator / projects.length : 0;
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
    const isFinished = Boolean(project?.finished) || normalizeText(project?.projectStatus).includes("project finished") || normalizeText(project?.jobProcessStatus).includes("project finished");
    if (!isFinished) return total;
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
  return Boolean(user) && !userHasProjectsScope(user) && user.role !== 'admin' && state.sectorScopedView;
}

function getScopedDemandLabelsForUser(user = state.user) {
  const sector = getPrimaryUserSector(user);
  if (sector === 'pendente_envio') return ['Logística', 'Pendente de envio'];
  if (sector === 'inspecao') return ['Inspeção'];
  if (sector === 'pintura') return ['Pintura'];
  if (sector === 'solda' || sector === 'calderaria' || sector === 'producao') return ['Produção'];
  return [];
}

function projectMatchesScopedSector(project, user = state.user) {
  const labels = getScopedDemandLabelsForUser(user).map((item) => normalizeText(item).trim()).filter(Boolean);
  if (!labels.length) return true;
  const currentGroup = normalizeText(project?.currentStageGroup || simplifyCurrentStage(project)).trim();
  return labels.includes(currentGroup);
}

function alertMatchesScopedSector(alert, user = state.user) {
  const sector = getPrimaryUserSector(user);
  if (!sector) return true;
  return normalizeSectorValue(alert?.sector) === sector;
}

function updatePrimaryUserActionUi() {
  if (!openSectorAlertsEl) return;
  const projectsScope = userHasProjectsScope();
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
    titleEl.textContent = projectsScope ? "Meus projetos" : (sectorScopedView ? "Meus alertas por setor" : "Meus alertas por setor");
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
    openStageUpdatesEl.textContent = canValidateStageWorkspace() ? 'Validação PCP' : 'Apontamentos';
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
  const candidates = tokenizeNormalizedNames([user.name, user.username, String(user.username || '').split('@')[0]]);
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
  if (!userHasProjectsScope()) {
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

  const sourceProjects = getVisibleProjectsSource();

  state.filteredProjects = sourceProjects
    .filter((project) => {
      const matchesQuery = !query || project._searchText.includes(query);
      const matchesDemand = !demand
        || normalizeText(project.currentStageGroup || simplifyCurrentStage(project)).includes(demand)
        || normalizeText(project.currentStage).includes(demand)
        || normalizeText(translateProjectStatus(project.projectStatus, project.uiState)).includes(demand);
      return matchesQuery && matchesDemand;
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

function renderStats() {
  const stats = buildClientStats(getStatsProjectsSource());
  state.visibleStats = stats;
  const totalFinishedWeight = getTotalFinishedWeightAllProjects();
  const totalWeldedWeight = Number(stats.totalWeldedWeightKg || 0);
  const totalBacklogWelding = Math.max(0, Number(stats.totalWeightKg || 0) - totalWeldedWeight);
  document.getElementById("stat-projects").textContent = formatNumber(stats.totalProjects);
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
  setTags("stat-awaiting-tags", stats.awaitingShipmentTags);

  const completedEl = document.getElementById("stat-completed");
  if (completedEl) completedEl.textContent = formatNumber(stats.completed);
  setTags("stat-completed-tags", stats.completedTags);
}

function renderTable() {
  if (!state.filteredProjects.length) {
    bodyEl.innerHTML = '<tr><td colspan="17" class="loading-cell">Nenhum projeto encontrado para a busca informada.</td></tr>';
    searchCountEl.textContent = "0 resultado(s)";
    return;
  }

  searchCountEl.textContent = `${state.filteredProjects.length} resultado(s)`;

  bodyEl.innerHTML = state.filteredProjects
    .map((project) => {
      const isActive = project.rowId === state.selectedProjectId;
      const statusText = translateProjectStatus(project.projectStatus, project.uiState);
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
      const statusState = ["awaiting_shipment", "completed"].includes(project.uiState) ? "completed" : project.uiState;

      return `
        <tr class="${rowClass}" data-project-id="${project.rowId}">
          <td>${project.projectDisplay || "—"}</td>
          <td>${project.plannedFinishDate || "—"}</td>
          <td>${formatNumber(project.quantitySpools)}</td>
          <td>${formatNumber(project.weldedWeightKg, 0)}</td>
          <td>${project.weldingWeek || "—"}</td>
          <td>${formatNumber(project.kilos, 2)}</td>
          <td>${formatNumber(project.m2Painting, 3)}</td>
          <td>
            <span class="stage-pill">
              <span class="stage-dot stage-dot--${stageStatusClass(project.currentStageStatus)}"></span>
              <span class="stage-text">${project.currentStageGroup || simplifyCurrentStage(project)}</span>
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
          <td class="cell-finished cell-finished--${project.finished ? "yes" : "no"}">${completedSymbol}</td>
        </tr>
      `;
    })
    .join("");
}

function renderSelectedProjectCard() {
  const project = getSelectedProject();
  if (!project) {
    detailCardEl.innerHTML = '<div class="detail-placeholder">Selecione um projeto na tabela ou pela busca para abrir o popup.</div>';
    return;
  }

  const statusText = translateProjectStatus(project.projectStatus, project.uiState);
  const matchedSpools = project.spools?.length || 0;

  detailCardEl.innerHTML = `
    <div class="detail-hero compact">
      <div class="detail-project-title">
        <div>
          <p class="detail-project-subtitle">Projeto selecionado</p>
          <h3>${projectDisplayWithClient(project)}</h3>
        </div>
        <span class="badge badge--${["awaiting_shipment", "completed"].includes(project.uiState) ? "completed" : project.uiState}">${statusText}</span>
      </div>

      <div class="detail-grid compact-grid">
        <div class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(project.quantitySpools)}</strong></div>
        <div class="metric-chip"><span>Cliente</span><strong>${project.client || "—"}</strong></div>
        <div class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></div>
        <button class="metric-chip metric-chip--button" type="button" id="open-backlog-project">
          <span>Backlog KG</span><strong>${formatNumber(getBacklogKg(project), 0)} kg</strong><small>${formatBacklogItemText(project)}</small>
        </button>
        <div class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></div>
        <div class="metric-chip"><span>Início planejado</span><strong>${project.plannedStartDate || "—"}</strong></div>
        <div class="metric-chip"><span>Término planejado</span><strong>${project.plannedFinishDate || "—"}</strong></div>
        <div class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 0)}kg</strong></div>
        <div class="metric-chip"><span>Painting</span><strong>${formatNumber(project.m2Painting, 3)}</strong></div>
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
          <span class="stage-name">${project.currentStageGroup || simplifyCurrentStage(project)}</span>
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

  const sourceSpools = state.modalPendingOnly ? getPendingSpools(project) : (project.spools || []);
  const sortedSpools = [...sourceSpools].sort((a, b) => {
    const aProgress = Number.isFinite(Number(a?.individualProgress)) ? Number(a.individualProgress) : 999999;
    const bProgress = Number.isFinite(Number(b?.individualProgress)) ? Number(b.individualProgress) : 999999;
    if (aProgress !== bProgress) return aProgress - bProgress;
    return String(a?.iso || '').localeCompare(String(b?.iso || ''), 'pt-BR', { numeric: true, sensitivity: 'base' });
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

      return `
        <tr data-modal-row="true">
          <td>${spool.iso || "—"}</td>
          <td>${spool.description || "—"}</td>
          <td class="modal-observation-cell">${observations}</td>
          <td>${formatNumber(spool.weldedWeightKg, 0)} kg</td>
          <td>${spool.weldingWeek || "—"}</td>
          <td>${formatNumber(spool.kilos, 2)}</td>
          <td>${formatNumber(spool.m2Painting, 3)}</td>
          <td><span class="cell-status cell-status--${["awaiting_shipment", "completed"].includes(spool.uiState) ? "completed" : spool.uiState}">${uiStateLabel(spool.uiState)}</span></td>
          <td class="${percentStateClass(spool.stagePercent)}">${spool.stage || "—"}</td>
          <td class="${percentStateClass(spool.individualProgress)}">${formatPercent(spool.individualProgress)}</td>
          <td class="${percentStateClass(spool.overallProgress)}">${formatPercent(spool.overallProgress)}</td>
          ${stageColumns}
        </tr>
      `;
    })
    .join("");

  const stageHeaders = stageOrder.map((stage) => `<th>${stage.label}</th>`).join("");
  const statusText = translateProjectStatus(project.projectStatus, project.uiState);

  modalTitleEl.textContent = projectDisplayWithClient(project);
  modalSubtitleEl.textContent = `${statusText} • ${state.modalPendingOnly ? getPendingSpools(project).length : (project.spools?.length || 0)} item(ns) interno(s)`;

  modalContentEl.innerHTML = `
    <section class="modal-summary-grid">
      <article class="metric-chip"><span>Qtd. itens</span><strong>${formatNumber(project.quantitySpools)}</strong></article>
      <article class="metric-chip"><span>Cliente</span><strong>${project.client || "—"}</strong></article>
      <article class="metric-chip"><span>Peso total soldado</span><strong>${formatNumber(project.weldedWeightKg, 0)} kg</strong></article>
      <article class="metric-chip metric-chip--button" id="modal-open-backlog" role="button" tabindex="0"><span>Backlog KG</span><strong>${formatNumber(getBacklogKg(project), 0)} kg</strong><small>${formatBacklogItemText(project)}</small></article>
      <article class="metric-chip"><span>Semana finalizado</span><strong>${project.weldingWeek || "—"}</strong></article>
      <article class="metric-chip"><span>Início planejado</span><strong>${project.plannedStartDate || "—"}</strong></article>
      <article class="metric-chip"><span>Término planejado</span><strong>${project.plannedFinishDate || "—"}</strong></article>
      <article class="metric-chip"><span>Peso total</span><strong>${formatNumber(project.kilos, 0)}kg</strong></article>
      <article class="metric-chip"><span>Painting total</span><strong>${formatNumber(project.m2Painting, 3)}</strong></article>
      <article class="metric-chip"><span>% Individual</span><strong>${formatPercent(project.individualProgress)}</strong></article>
      <article class="metric-chip"><span>% Geral</span><strong>${formatPercent(project.overallProgress)}</strong></article>
      <article class="metric-chip"><span>Etapa atual</span><strong>${project.currentStage}</strong></article>
    </section>

    <section class="modal-milestones">
      ${milestoneList || '<div class="empty-inline">Nenhum marco de data disponível.</div>'}
    </section>

    ${renderProjectSignals(project)}

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
            <th>Painting</th>
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
  if (userHasProjectsScope(state.user) && state.projectView === 'mine') {
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
  if (userHasProjectsScope(state.user) && state.projectView === 'mine') {
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
  if (userHasProjectsScope(state.user) && state.projectView === 'mine') {
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
    { key: "inspecao", label: "Inspeção", match: ["Inspeção"] },
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

function openAlertModal(force = false) {
  if (!alertModalEl) return;
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

function updateMeta() {
  if (!state.meta) return;
  sheetNameEl.textContent = state.meta.sheetName || "Smartsheet";
  lastSyncEl.textContent = `Última atualização: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`;
  footerVersionEl.textContent = `Versão da sheet: ${state.meta.version}`;
}

async function loadProjects() {
  try {
    const response = await fetch("/api/projects", { cache: "no-store", credentials: "same-origin" });
    let data = null;

    try {
      data = await response.json();
    } catch (parseError) {
      throw new Error("Falha ao atualizar dados da planilha.");
    }

    if (response.status === 401) {
      state.user = null;
      updateSessionUi();
      resetDashboardForLoggedOutState();
      openLoginModal(data?.error || "Faça login para visualizar o painel.");
      return;
    }

    if (!response.ok || !data.ok) {
      throw new Error(data?.error || "Falha ao carregar projetos.");
    }

    state.projects = enrichProjects(data.projects || []);
    renderProjectViewTabs();
    state.stats = data.stats || null;
    state.meta = data.meta || null;
    state.alerts = data.alerts || [];
    buildDemandOptions();
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
    if (shouldOpenAlertPopup()) {
      openAlertModal(true);
    } else {
      renderAlertModal();
    }
    if (state.user && sectorAlertsModalEl && !sectorAlertsModalEl.classList.contains("hidden")) {
      renderManualAlerts();
    }
  } catch (error) {
    const fallbackMessage = error?.message || "Falha ao atualizar dados da planilha.";

    if (state.projects.length) {
      const staleSuffix = state.meta?.lastSync
        ? ` | exibindo última atualização válida: ${new Date(state.meta.lastSync).toLocaleString("pt-BR")}`
        : "";
      lastSyncEl.textContent = `Conexão instável com a planilha${staleSuffix}`;
      console.warn("Falha temporária ao atualizar projetos:", fallbackMessage);
      return;
    }

    bodyEl.innerHTML = `<tr><td colspan="17" class="loading-cell">${fallbackMessage}</td></tr>`;
    detailCardEl.innerHTML = `<div class="detail-placeholder">${fallbackMessage}</div>`;
  }
}

function startPolling() {
  window.clearInterval(state.pollTimer);
  state.pollTimer = window.setInterval(async () => {
    await loadProjects();
    if (state.user) {
      await loadManualAlerts();
      await loadAlertResponses();
    }
  }, DEFAULT_POLL_MS);
}

function bindEvents() {
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
    state.weekFilter = "";
    searchInputEl.value = "";
    if (demandFilterEl) demandFilterEl.value = "";
    if (weekFilterEl) weekFilterEl.value = "";
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

  if (weekFilterEl) {
    weekFilterEl.addEventListener("change", (event) => {
      state.weekFilter = event.target.value;
      renderStats();
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
      openAlertModal(true);
    });
  }

if (loginFormEl) {
  loginFormEl.addEventListener("submit", handleLoginSubmit);
}

if (openLoginButtonEl) {
  openLoginButtonEl.addEventListener("click", () => {
    openLoginModal();
  });
}

if (loginCloseEl) {
  loginCloseEl.addEventListener("click", closeLoginModal);
}

const adminUserSectorEl = document.getElementById("admin-user-sector");
const adminUserRoleEl = document.getElementById("admin-user-role");
if (adminUserSectorEl) {
  adminUserSectorEl.addEventListener("change", (event) => {
    const next = normalizeSectorValue(event.target.value);
    const selected = new Set(getSelectedAdminAlertSectors());
    if (next) {
      selected.add(next);
      setSelectedAdminAlertSectors([...selected]);
    }
  });
}

if (adminUserRoleEl) {
  adminUserRoleEl.addEventListener("change", (event) => {
    const disabled = event.target.value === "admin";
    document.querySelectorAll('[data-admin-alert-sector-option]').forEach((input) => {
      input.disabled = disabled;
    });
  });
}

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
    if (userHasProjectsScope()) {
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
  openStageUpdatesEl.addEventListener('click', async () => {
    if (!state.user) {
      openLoginModal();
      return;
    }
    state.stageUpdatesSearchQuery = '';
    await loadStageUpdates();
    openStageUpdatesModal();
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
    const concludeButton = event.target.closest('[data-stage-conclude]');
    if (concludeButton) {
      concludeStageUpdate(concludeButton.dataset.stageConclude);
    }
  });
  stageUpdatesModalEl.addEventListener('input', (event) => {
    const searchEl = event.target.closest('[data-stage-search="true"]');
    if (searchEl) {
      state.stageUpdatesSearchQuery = searchEl.value || '';
      renderStageUpdatesModal();
      return;
    }
    const progressEl = event.target.closest('[data-stage-progress="true"]');
    if (progressEl) {
      const rowEl = progressEl.closest('tr');
      const dateEl = rowEl?.querySelector('[name="completionDate"]');
      if (dateEl && Number(progressEl.value) === 100 && !dateEl.value) {
        dateEl.value = new Date().toISOString().slice(0, 10);
      }
    }
  });
  stageUpdatesModalEl.addEventListener('click', (event) => {
    const actionButton = event.target.closest('[data-stage-send="true"], [data-stage-review="true"]');
    if (!actionButton) return;
    const rowEl = actionButton.closest('tr');
    const formEl = rowEl?.querySelector('[data-stage-update-form="true"]');
    if (!formEl) return;
    const actionType = actionButton.matches('[data-stage-review="true"]') ? 'review' : 'advance';
    handleStageWorkspaceSubmit(formEl, actionType);
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
    openAdminPanelEl.classList.add("hidden");
    if (openLoginButtonEl) openLoginButtonEl.classList.remove("hidden");
    return;
  }

  if (!userHasProjectsScope(user)) {
    state.projectView = 'all';
    state.sectorScopedView = false;
    state.alertSectorFilter = 'all';
  }

  sessionUserNameEl.textContent = user.name || user.username;
  const linkedSectors = getUserAlertSectors(user);
  sessionUserMetaEl.textContent = `${user.role === "admin" ? "Administrador" : "Setor"} • ${sectorLabel(user.sector)}${user.role !== "admin" && linkedSectors.length > 1 ? ` • Alertas: ${formatSectorList(linkedSectors)}` : ""}`;
  updatePrimaryUserActionUi();
  sessionStatusEl.textContent = "online";
  logoutButtonEl.classList.remove("hidden");
  if (openLoginButtonEl) openLoginButtonEl.classList.add("hidden");
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
  if (bodyEl) bodyEl.innerHTML = `<tr><td colspan="17" class="loading-cell">Faça login para visualizar os projetos.</td></tr>`;
  if (detailCardEl) detailCardEl.innerHTML = `<div class="detail-placeholder">Painel protegido. Entre com seu usuário e senha para visualizar as informações.</div>`;
  if (searchCountEl) searchCountEl.textContent = '0 resultado(s)';
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

async function loadManualAlerts() {
  if (!state.user) return;
  try {
    const response = await fetch(`/api/sector-alerts?t=${Date.now()}`, { credentials: "same-origin", cache: "no-store" });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao carregar alertas operacionais.");
    }
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

async function loadAlertResponses() {
  if (!state.user) {
    state.alertResponses = [];
    return;
  }
  try {
    const response = await fetch('/api/alert-responses', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar respostas das sinalizações.');
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

function resetAdminUserForm() {
  if (adminUserFormEl) adminUserFormEl.reset();
  if (adminUserIdEl) adminUserIdEl.value = "";
  if (adminUserCancelEditEl) adminUserCancelEditEl.classList.add("hidden");
  if (adminUserSubmitLabelEl) adminUserSubmitLabelEl.textContent = "Criar usuário";
  setSelectedAdminAlertSectors([document.getElementById("admin-user-sector")?.value || "pintura"]);
}

function startEditUser(userId) {
  const list = adminUsersListEl?._cachedUsers || [];
  const user = list.find((item) => String(item.id) === String(userId));
  if (!user) return;
  document.getElementById("admin-user-name").value = user.name || "";
  document.getElementById("admin-user-username").value = user.username || "";
  document.getElementById("admin-user-password").value = "";
  document.getElementById("admin-user-role").value = user.role === "admin" ? "admin" : "sector";
  document.getElementById("admin-user-sector").value = user.sector && user.sector !== "all" ? user.sector : "pintura";
  setSelectedAdminAlertSectors(Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector]);
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
    return `
      <article class="admin-list-item">
        <strong>${escapeHtml(user.name)}</strong>
        <div class="admin-list-item-meta">
          <span>Login: ${escapeHtml(user.username)}</span>
          <span>Perfil: ${escapeHtml(user.role === "admin" ? "Admin notificações" : "Setor")}</span>
          <span>Setor principal: ${escapeHtml(sectorLabel(user.sector))}</span>
          <span>Recebe alertas de: ${escapeHtml(formatSectorList(Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector]))}</span>
          <span>${user.active ? "Ativo" : "Inativo"}</span>
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

async function loadAdminData() {
  if (state.user?.role !== "admin") return;
  try {
    const response = await fetch("/api/admin-users", { credentials: "same-origin", cache: "no-store" });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao carregar usuários.");
    }
    state.githubSyncEnabled = Boolean(data.githubSyncEnabled ?? state.githubSyncEnabled);
    updateSessionUi();
    const remoteUsers = Array.isArray(data.users) ? data.users : [];
    if (state.githubSyncEnabled) {
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
    renderAdminUsersList(merged);
  } catch (error) {
    const localUsers = readLocalUsers().map((user) => ({
      id: user.id,
      name: user.name,
      username: user.username,
      role: user.role,
      sector: user.sector,
      alertSectors: Array.isArray(user.alertSectors) ? user.alertSectors : [user.sector],
      active: user.active !== false,
      createdAt: user.createdAt || null,
    }));
    if (localUsers.length) {
      renderAdminUsersList(localUsers);
    } else {
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
  loadAdminData();
  window.clearInterval(adminResponsesPollTimer);
  adminResponsesPollTimer = window.setInterval(() => {
    if (!adminModalEl.classList.contains('hidden') && state.user?.role === 'admin') {
      loadAlertResponses();
    }
  }, 10000);
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

async function handleLoginSubmit(event) {
  event.preventDefault();
  if (!loginFeedbackEl) return;
  loginFeedbackEl.textContent = "Validando acesso...";
  try {
    const response = await fetch("/api/auth-login", {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        username: String(loginUsernameEl.value || "").trim(),
        password: String(loginPasswordEl.value || "").trim(),
      }),
    });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || "Falha ao entrar.");
    }
    state.user = data.user;
    closeLoginModal();
    await bootstrapSession();
    await loadProjects();
    await loadManualAlerts();
    await loadAlertResponses();
    await loadStageUpdates();
    if (state.user?.role === "admin") {
      await loadAdminData();
    }
    startPolling();
  } catch (error) {
    loginFeedbackEl.textContent = error.message || "Falha ao autenticar.";
  }
}

async function handleLogout() {
  await fetch("/api/auth-logout", { credentials: "same-origin" });
  state.user = null;
  state.manualAlerts = [];
  state.projectSignals = [];
  state.alertResponses = [];
  state.stageUpdates = [];
  window.clearInterval(state.pollTimer);
  updateSessionUi();
  resetDashboardForLoggedOutState();
  openLoginModal("Sessão encerrada. Faça login novamente para acessar o painel.");
}




async function loadStageUpdates() {
  if (!state.user) {
    state.stageUpdates = [];
    return;
  }
  try {
    const response = await fetch('/api/stage-updates', { credentials: 'same-origin', cache: 'no-store' });
    const data = await response.json().catch(() => null);
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao carregar apontamentos setoriais.');
    state.stageUpdates = Array.isArray(data.updates) ? data.updates : [];
  } catch (error) {
    state.stageUpdates = [];
    if (stageUpdatesContentEl && stageUpdatesModalEl && !stageUpdatesModalEl.classList.contains('hidden')) {
      stageUpdatesContentEl.innerHTML = `<div class="empty-state">${escapeHtml(error.message || 'Falha ao carregar apontamentos setoriais.')}</div>`;
    }
  }
}

function renderStageSectorWorkspace() {
  if (!stageUpdatesContentEl) return;
  const sector = getStageWorkspaceSector();
  const stageLabel = getStageWorkspaceLabel(sector);
  const matchedProjects = stageWorkspaceSearchProjects();
  const myUpdates = getMyStageUpdates();
  const pendingMine = myUpdates.filter((item) => isPendingStageStatus(item.status));
  const resolvedMine = myUpdates.filter((item) => isResolvedStageStatus(item.status)).slice(0, 10);
  stageUpdatesContentEl.innerHTML = `
    <div class="stage-workspace-shell">
      <section class="admin-card admin-card--wide">
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Buscar BSP / cliente</span>
            <input type="text" data-stage-search="true" value="${escapeHtml(state.stageUpdatesSearchQuery || '')}" placeholder="Ex.: BSP 25-1246" />
          </label>
          <div class="stage-muted">Etapa atual do seu login: <strong>${escapeHtml(stageLabel)}</strong></div>
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
                      <span class="stage-badge">${spools.length} spool(s)</span>
                    </div>
                  </div>
                </div>
                <div class="table-shell">
                  <table class="stage-inline-table">
                    <thead><tr><th>Spool</th><th>Descrição</th><th>Andamento</th><th>Data conclusão</th><th>Obs.</th><th>Ação</th></tr></thead>
                    <tbody>
                      ${spools.map((spool) => {
                        const projectRowId = project.rowId || project.rowNumber;
                        const pending = getPendingStageUpdate(projectRowId, spool.iso, sector);
                        const lastResolved = getLatestResolvedStageUpdate(projectRowId, spool.iso, sector);
                        const submitKey = `${String(projectRowId || '').trim()}::${String(spool.iso || '').trim().toLowerCase()}::${String(sector || '').trim().toLowerCase()}`;
                        const isSubmitting = Boolean(state.stageSubmittingKeys?.[submitKey]);
                        return `
                          <tr>
                            <td>${escapeHtml(spool.iso || '—')}</td>
                            <td>${escapeHtml(spool.description || '—')}</td>
                            <td>
                              <div class="stage-row-form" data-stage-update-form="true" data-project-row-id="${escapeHtml(String(projectRowId || ''))}" data-project-number="${escapeHtml(project.projectNumber || '')}" data-spool-iso="${escapeHtml(spool.iso || '')}" data-stage-sector="${escapeHtml(sector || '')}">
                                <select name="progress" data-stage-progress="true" ${pending || isSubmitting ? 'disabled' : ''}>
                                  ${STAGE_PROGRESS_OPTIONS.map((value) => `<option value="${value}">${value}%</option>`).join('')}
                                </select>
                              </div>
                            </td>
                            <td><input type="date" name="completionDate" value="" ${pending || isSubmitting ? 'disabled' : ''} /></td>
                            <td><textarea name="note" rows="2" placeholder="Observação opcional" ${pending || isSubmitting ? 'disabled' : ''}></textarea></td>
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
          }).join('')}</div>` : `<div class="empty-state">Pesquise a BSP para visualizar os spools e lançar o avanço da sua etapa.</div>`}
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

function renderStageValidationWorkspace() {
  if (!stageUpdatesContentEl) return;
  const query = normalizeText(state.stageUpdatesSearchQuery || '');
  const all = Array.isArray(state.stageUpdates) ? state.stageUpdates : [];
  const filtered = all.filter((item) => {
    if (!query) return true;
    return [item.projectNumber, item.projectDisplay, item.client, item.spoolIso, item.sector, item.createdByName]
      .some((value) => normalizeText(value).includes(query));
  }).sort((a,b)=> new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
  const pending = filtered.filter((item) => isPendingStageStatus(item.status));
  const history = filtered.filter((item) => isResolvedStageStatus(item.status)).slice(0, 50);
  stageUpdatesContentEl.innerHTML = `
    <div class="stage-workspace-shell">
      <section class="admin-card admin-card--wide">
        <div class="stage-toolbar">
          <label class="stack-field">
            <span>Buscar BSP / spool / setor</span>
            <input type="text" data-stage-search="true" value="${escapeHtml(state.stageUpdatesSearchQuery || '')}" placeholder="Ex.: BSP 25-1246" />
          </label>
          <div class="stage-muted">Pendentes: <strong>${pending.length}</strong> • Histórico: <strong>${history.length}</strong></div>
        </div>
      </section>
      <div class="stage-two-col">
        <section class="admin-card admin-card--wide">
          <div class="admin-card-head"><h4>Pendentes para validação do PCP</h4></div>
          <div class="stage-update-list">
            ${pending.length ? pending.map((item) => `
              <article class="stage-update-card">
                <div class="stage-update-head">
                  <div>
                    <strong>${escapeHtml(item.projectDisplay || item.projectNumber || 'Projeto')} • ${escapeHtml(item.spoolIso || 'Spool')}</strong>
                    <div class="stage-update-meta">
                      <span class="stage-badge stage-badge--sector">${escapeHtml(sectorLabel(item.sector))}</span>
                      <span class="stage-badge ${isReviewStageStatus(item.status) ? 'stage-badge--review' : 'stage-badge--pending'}">${escapeHtml(stageUpdatePendingLabel(item.status))}</span>
                      <span class="stage-badge">${escapeHtml(String(item.progress || 0))}%</span>
                    </div>
                  </div>
                  <button class="primary-button" type="button" data-stage-conclude="${escapeHtml(item.id)}">${escapeHtml(isReviewStageStatus(item.status) ? 'Tratar revisão' : 'Concluir')}</button>
                </div>
                <p><strong>Spool:</strong> ${escapeHtml(item.spoolDescription || '—')}</p>
                <p><strong>${escapeHtml(isReviewStageStatus(item.status) ? 'Motivo da revisão:' : 'Observação do setor:')}</strong> ${escapeHtml(item.note || (isReviewStageStatus(item.status) ? 'Sem motivo informado.' : 'Sem observação do setor.'))}</p>
                <div class="stage-muted">Informado por ${escapeHtml(item.createdByName || item.createdBy || 'Usuário')} em ${escapeHtml(formatStageDate(item.createdAt))}</div>
              </article>`).join('') : `<div class="empty-state">Nenhum apontamento pendente no momento.</div>`}
          </div>
        </section>
        <section class="admin-card admin-card--wide">
          <div class="admin-card-head"><h4>Histórico validado</h4></div>
          <div class="stage-update-list">
            ${history.length ? history.map((item) => `
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
                <div class="stage-muted">Informado por ${escapeHtml(item.createdByName || item.createdBy || 'Usuário')} • ${escapeHtml(isReviewStageStatus(item.status) ? 'revisão tratada' : 'concluído')} por ${escapeHtml(item.resolvedByName || item.resolvedBy || 'PCP')} em ${escapeHtml(formatStageDate(item.resolvedAt))}</div>
                ${item.resolutionNote ? `<div class="response-bubble response-bubble--admin"><strong>Fechamento PCP</strong><p>${escapeHtml(item.resolutionNote)}</p></div>` : ''}
              </article>`).join('') : `<div class="empty-state">Nenhum histórico validado encontrado.</div>`}
          </div>
        </section>
      </div>
    </div>`;
}

function renderStageUpdatesModal() {
  if (!stageUpdatesContentEl) return;
  if (canValidateStageWorkspace()) {
    renderStageValidationWorkspace();
    return;
  }
  renderStageSectorWorkspace();
}

function openStageUpdatesModal() {
  if (!stageUpdatesModalEl) return;
  const titleEl = document.getElementById('stage-updates-title');
  const subtitleEl = document.getElementById('stage-updates-subtitle');
  if (titleEl) titleEl.textContent = canValidateStageWorkspace() ? 'Validação PCP dos apontamentos' : `Apontamentos da etapa • ${getStageWorkspaceLabel()}`;
  if (subtitleEl) subtitleEl.textContent = canValidateStageWorkspace()
    ? 'Conclua os registros validados para que saiam da fila e permaneçam no histórico.'
    : 'Cada setor informa somente a sua própria etapa por spool. O PCP valida e mantém o histórico.';
  renderStageUpdatesModal();
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

async function handleStageWorkspaceSubmit(formEl, actionType = 'advance') {
  const rowEl = formEl?.closest('tr');
  const projectRowId = String(formEl?.dataset?.projectRowId || '').trim();
  const spoolIso = String(formEl?.dataset?.spoolIso || '').trim();
  const sector = String(formEl?.dataset?.stageSector || getStageWorkspaceSector() || '').trim();
  const progress = String(formEl?.querySelector('[name="progress"]')?.value || '').trim();
  const completionDate = String(rowEl?.querySelector('[name="completionDate"]')?.value || '').trim();
  const note = String(rowEl?.querySelector('[name="note"]')?.value || '').trim();
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
    if (!response.ok || !data?.ok) throw new Error(data?.error || 'Falha ao concluir apontamento.');
    await loadStageUpdates();
    renderStageUpdatesModal();
  } catch (error) {
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
      sector: document.getElementById("admin-user-sector").value,
      alertSectors: getSelectedAdminAlertSectors(),
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
  const authenticated = await bootstrapSession();
  if (authenticated) {
    await loadProjects();
    await syncPushSubscription(false).catch(() => {});
    await loadManualAlerts();
    await loadAlertResponses();
    await loadStageUpdates();
    if (state.user?.role === "admin") {
      await loadAdminData();
    }
    startPolling();
  }
}

init();
