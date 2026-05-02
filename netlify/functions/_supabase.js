const { normalizeSectorList, normalizeText, normalizeSectorValue, hashPassword, verifyPassword } = require('./_auth');

const SUPABASE_URL = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || '');
const SUPABASE_ANON_KEY = String(process.env.SUPABASE_ANON_KEY || '');

function isSupabaseConfigured() {
  return Boolean(SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY);
}

function getSupabaseHeaders(prefer = '') {
  const headers = {
    apikey: SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    'Content-Type': 'application/json',
  };
  if (prefer) headers.Prefer = prefer;
  return headers;
}

async function supabaseFetch(path, options = {}) {
  if (!isSupabaseConfigured()) {
    throw new Error('Supabase não configurado. Defina SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY no Netlify.');
  }
  const response = await fetch(`${SUPABASE_URL}${path}`, {
    ...options,
    headers: {
      ...getSupabaseHeaders(),
      ...(options.headers || {}),
    },
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Supabase ${response.status}: ${text}`);
  }
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) return response.json();
  return response.text();
}

function mapUser(row) {
  if (!row) return null;
  return {
    id: row.id,
    name: row.name,
    username: row.username,
    passwordHash: row.password_hash,
    role: row.role === 'admin' ? 'admin' : 'sector',
    sector: normalizeSectorValue(row.sector || (row.role === 'admin' ? 'all' : '')), 
    alertSectors: normalizeSectorList(row.sector || '', Array.isArray(row.alert_sectors) ? row.alert_sectors : []),
    projectPmAliases: Array.isArray(row.project_pm_aliases) ? row.project_pm_aliases.filter(Boolean) : [],
    active: row.active !== false,
    createdAt: row.created_at || null,
    updatedAt: row.updated_at || null,
  };
}


const PRESENCE_ONLINE_WINDOW_MS = Number(process.env.PRESENCE_ONLINE_WINDOW_MS || 2 * 60 * 1000);

function mapPresence(row) {
  if (!row) return null;
  const lastSeenAt = row.last_seen_at || null;
  const lastSeenTime = lastSeenAt ? new Date(lastSeenAt).getTime() : 0;
  const isFresh = Boolean(lastSeenTime && (Date.now() - lastSeenTime) <= PRESENCE_ONLINE_WINDOW_MS);
  const online = row.status === 'online' && isFresh;
  return {
    userId: row.user_id,
    username: row.username || '',
    name: row.name || '',
    role: row.role === 'admin' ? 'admin' : 'sector',
    sector: normalizeSectorValue(row.sector || ''),
    alertSectors: normalizeSectorList(row.sector || '', Array.isArray(row.alert_sectors) ? row.alert_sectors : []),
    status: online ? 'online' : 'offline',
    online,
    lastSeenAt,
    lastLoginAt: row.last_login_at || null,
    lastLogoutAt: row.last_logout_at || null,
    lastViewAt: row.last_view_at || null,
    lastViewName: row.last_view_name || '',
    lastViewUrl: row.last_view_url || '',
    lastViewTitle: row.last_view_title || '',
    userAgent: row.user_agent || '',
    ipAddress: row.ip_address || '',
    updatedAt: row.updated_at || null,
  };
}

function mapAlert(row) {
  if (!row) return null;
  return {
    id: row.id,
    title: row.title,
    message: row.message,
    sector: normalizeSectorValue(row.sector),
    priority: row.priority || 'normal',
    requiresAck: row.require_ack !== false,
    createdBy: row.created_by || '',
    createdAt: row.created_at || null,
    active: row.active !== false,
    expiresAfterReadHours: Number(row.expires_after_read_hours || 24),
    readExpiresAt: row.read_expires_at || null,
    updatedAt: row.updated_at || null,
  };
}

function mapAck(row) {
  if (!row) return null;
  return {
    id: row.id,
    alertId: row.alert_id,
    userId: row.user_id,
    username: row.username,
    sector: normalizeSectorValue(row.sector),
    acknowledgedAt: row.read_at,
  };
}


function mapStageUpdate(row) {
  if (!row) return null;
  return {
    id: row.id,
    projectRowId: Number(row.project_row_id || 0),
    projectNumber: row.project_number || '',
    projectDisplay: row.project_display || '',
    client: row.client || '',
    spoolIso: row.spool_iso || '',
    spoolDescription: row.spool_description || '',
    sector: normalizeSectorValue(row.sector),
    progress: Number(row.progress || 0),
    completionDate: row.completion_date || '',
    note: row.note || '',
    status: row.status || 'pending',
    createdBy: row.created_by || '',
    createdByName: row.created_by_name || '',
    createdAt: row.created_at || null,
    resolvedBy: row.resolved_by || '',
    resolvedByName: row.resolved_by_name || '',
    resolvedAt: row.resolved_at || null,
    resolutionNote: row.resolution_note || '',
    updatedAt: row.updated_at || null,
  };
}

function mapResponse(row) {
  if (!row) return null;
  return {
    id: row.id,
    alertId: row.alert_id,
    userId: row.user_id,
    username: row.username,
    userEmail: row.user_email || '',
    sector: normalizeSectorValue(row.sector),
    responseText: row.response_text || '',
    adminReply: row.admin_reply || '',
    status: row.status || 'enviado',
    createdAt: row.created_at || null,
    updatedAt: row.updated_at || null,
  };
}

async function listUsers() {
  const rows = await supabaseFetch('/rest/v1/users?select=*&order=created_at.desc');
  return (Array.isArray(rows) ? rows : []).map(mapUser);
}

async function getUserByUsername(username) {
  const q = encodeURIComponent(String(username || '').trim());
  const rows = await supabaseFetch(`/rest/v1/users?select=*&username=eq.${q}&limit=1`);
  return mapUser(Array.isArray(rows) ? rows[0] : null);
}

async function getUserById(userId) {
  const q = encodeURIComponent(String(userId || '').trim());
  const rows = await supabaseFetch(`/rest/v1/users?select=*&id=eq.${q}&limit=1`);
  return mapUser(Array.isArray(rows) ? rows[0] : null);
}

async function insertUser(input) {
  const payload = {
    name: input.name,
    username: input.username,
    password_hash: input.passwordHash,
    role: input.role,
    sector: input.sector,
    alert_sectors: input.alertSectors || [],
    project_pm_aliases: Array.isArray(input.projectPmAliases) ? input.projectPmAliases : [],
    active: input.active !== false,
  };
  const rows = await supabaseFetch('/rest/v1/users?select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapUser(Array.isArray(rows) ? rows[0] : null);
}

async function updateUser(userId, updates) {
  const q = encodeURIComponent(String(userId || '').trim());
  const payload = {};
  if ('name' in updates) payload.name = updates.name;
  if ('username' in updates) payload.username = updates.username;
  if ('passwordHash' in updates) payload.password_hash = updates.passwordHash;
  if ('role' in updates) payload.role = updates.role;
  if ('sector' in updates) payload.sector = updates.sector;
  if ('alertSectors' in updates) payload.alert_sectors = updates.alertSectors || [];
  if ('projectPmAliases' in updates) payload.project_pm_aliases = Array.isArray(updates.projectPmAliases) ? updates.projectPmAliases : [];
  if ('active' in updates) payload.active = updates.active !== false;
  const rows = await supabaseFetch(`/rest/v1/users?id=eq.${q}&select=*`, {
    method: 'PATCH',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapUser(Array.isArray(rows) ? rows[0] : null);
}


async function listUserPresence() {
  try {
    const rows = await supabaseFetch('/rest/v1/user_presence?select=*&order=last_seen_at.desc.nullslast');
    return (Array.isArray(rows) ? rows : []).map(mapPresence).filter(Boolean);
  } catch (error) {
    if (String(error.message || '').includes('user_presence')) return [];
    throw error;
  }
}

async function upsertUserPresence(input = {}) {
  const now = input.now || new Date().toISOString();
  const payload = {
    user_id: String(input.userId || '').trim(),
    username: input.username || '',
    name: input.name || '',
    role: input.role === 'admin' ? 'admin' : 'sector',
    sector: normalizeSectorValue(input.sector || ''),
    alert_sectors: normalizeSectorList(input.sector || '', Array.isArray(input.alertSectors) ? input.alertSectors : []),
    status: input.status === 'offline' ? 'offline' : 'online',
    last_seen_at: input.status === 'offline' ? (input.lastSeenAt || now) : now,
    updated_at: now,
    user_agent: String(input.userAgent || '').slice(0, 500),
    ip_address: String(input.ipAddress || '').slice(0, 120),
  };

  if (input.markLogin) payload.last_login_at = now;
  if (input.status === 'offline') payload.last_logout_at = now;

  if ('lastViewName' in input) payload.last_view_name = String(input.lastViewName || '').slice(0, 160);
  if ('lastViewUrl' in input) payload.last_view_url = String(input.lastViewUrl || '').slice(0, 700);
  if ('lastViewTitle' in input) payload.last_view_title = String(input.lastViewTitle || '').slice(0, 200);
  if (input.lastViewName || input.lastViewUrl || input.lastViewTitle) payload.last_view_at = now;

  if (!payload.user_id) return null;

  try {
    const rows = await supabaseFetch('/rest/v1/user_presence?on_conflict=user_id&select=*', {
      method: 'POST',
      headers: getSupabaseHeaders('resolution=merge-duplicates,return=representation'),
      body: JSON.stringify(payload),
    });
    return mapPresence(Array.isArray(rows) ? rows[0] : null);
  } catch (error) {
    if (String(error.message || '').includes('user_presence')) return null;
    throw error;
  }
}

async function markUserPresenceOffline(userId, input = {}) {
  const q = encodeURIComponent(String(userId || '').trim());
  if (!q) return null;
  const now = new Date().toISOString();
  const payload = {
    status: 'offline',
    last_seen_at: now,
    last_logout_at: now,
    updated_at: now,
  };
  if (input.lastViewName) payload.last_view_name = String(input.lastViewName || '').slice(0, 160);
  if (input.lastViewUrl) payload.last_view_url = String(input.lastViewUrl || '').slice(0, 700);
  if (input.lastViewTitle) payload.last_view_title = String(input.lastViewTitle || '').slice(0, 200);
  if (input.lastViewName || input.lastViewUrl || input.lastViewTitle) payload.last_view_at = now;

  try {
    const rows = await supabaseFetch(`/rest/v1/user_presence?user_id=eq.${q}&select=*`, {
      method: 'PATCH',
      headers: getSupabaseHeaders('return=representation'),
      body: JSON.stringify(payload),
    });
    return mapPresence(Array.isArray(rows) ? rows[0] : null);
  } catch (error) {
    if (String(error.message || '').includes('user_presence')) return null;
    throw error;
  }
}

async function listManualAlerts() {
  const rows = await supabaseFetch('/rest/v1/manual_alerts?select=*&order=created_at.desc');
  return (Array.isArray(rows) ? rows : []).map(mapAlert);
}

async function createManualAlert(input) {
  const payload = {
    title: input.title,
    message: input.message,
    sector: input.sector,
    priority: input.priority || 'normal',
    require_ack: input.requiresAck !== false,
    created_by: input.createdBy || '',
    active: input.active !== false,
    expires_after_read_hours: Number(input.expiresAfterReadHours || 24),
  };
  const rows = await supabaseFetch('/rest/v1/manual_alerts?select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapAlert(Array.isArray(rows) ? rows[0] : null);
}

async function listAcknowledgements() {
  const rows = await supabaseFetch('/rest/v1/alert_acknowledgements?select=*&order=read_at.desc');
  return (Array.isArray(rows) ? rows : []).map(mapAck);
}

async function addAcknowledgement(input) {
  const payload = {
    alert_id: input.alertId,
    user_id: input.userId || null,
    username: input.username || '',
    sector: input.sector || '',
  };
  const rows = await supabaseFetch('/rest/v1/alert_acknowledgements?select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapAck(Array.isArray(rows) ? rows[0] : null);
}

async function findAcknowledgement(alertId, userId) {
  const a = encodeURIComponent(String(alertId || '').trim());
  const u = encodeURIComponent(String(userId || '').trim());
  const rows = await supabaseFetch(`/rest/v1/alert_acknowledgements?select=*&alert_id=eq.${a}&user_id=eq.${u}&limit=1`);
  return mapAck(Array.isArray(rows) ? rows[0] : null);
}


async function listAlertResponses(alertId = '') {
  const filter = String(alertId || '').trim()
    ? `&alert_id=eq.${encodeURIComponent(String(alertId || '').trim())}`
    : '';
  try {
    const rows = await supabaseFetch(`/rest/v1/alert_responses?select=*&order=created_at.desc${filter}`);
    return (Array.isArray(rows) ? rows : []).map(mapResponse);
  } catch (error) {
    if (String(error.message || '').includes('alert_responses')) return [];
    throw error;
  }
}

async function createAlertResponse(input) {
  const payload = {
    alert_id: input.alertId,
    user_id: input.userId || null,
    username: input.username || '',
    user_email: input.userEmail || '',
    sector: input.sector || '',
    response_text: input.responseText || '',
    status: input.status || 'enviado',
  };
  const rows = await supabaseFetch('/rest/v1/alert_responses?select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapResponse(Array.isArray(rows) ? rows[0] : null);
}

async function updateAlertResponse(responseId, updates) {
  const q = encodeURIComponent(String(responseId || '').trim());
  const payload = {};
  if ('adminReply' in updates) payload.admin_reply = updates.adminReply || '';
  if ('status' in updates) payload.status = updates.status || 'enviado';
  const rows = await supabaseFetch(`/rest/v1/alert_responses?id=eq.${q}&select=*`, {
    method: 'PATCH',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapResponse(Array.isArray(rows) ? rows[0] : null);
}

function userPasswordMatches(password, stored) {
  if (!stored) return false;
  const raw = String(stored);
  if (raw.startsWith('scrypt$')) return verifyPassword(password, raw);
  return String(password) === raw;
}

async function listPushSubscriptions(userId = '') {
  const filter = String(userId || '').trim()
    ? `&user_id=eq.${encodeURIComponent(String(userId || '').trim())}`
    : '';
  try {
    const rows = await supabaseFetch(`/rest/v1/push_subscriptions?select=*&active=is.true&order=updated_at.desc${filter}`);
    return Array.isArray(rows) ? rows : [];
  } catch (error) {
    if (String(error.message || '').includes('push_subscriptions')) return [];
    throw error;
  }
}

async function upsertPushSubscription(input) {
  const payload = {
    user_id: input.userId,
    username: input.username || '',
    sector: input.sector || '',
    endpoint: input.endpoint,
    subscription_json: input.subscription,
    active: input.active !== false,
  };
  const rows = await supabaseFetch('/rest/v1/push_subscriptions?on_conflict=endpoint&select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('resolution=merge-duplicates,return=representation'),
    body: JSON.stringify(payload),
  });
  return Array.isArray(rows) ? rows[0] : null;
}

async function removePushSubscription(endpoint) {
  const q = encodeURIComponent(String(endpoint || '').trim());
  await supabaseFetch(`/rest/v1/push_subscriptions?endpoint=eq.${q}`, {
    method: 'DELETE',
    headers: getSupabaseHeaders(),
  });
  return true;
}


async function listStageUpdates() {
  try {
    const rows = await supabaseFetch('/rest/v1/stage_updates?select=*&order=created_at.desc');
    return (Array.isArray(rows) ? rows : []).map(mapStageUpdate);
  } catch (error) {
    if (String(error.message || '').includes('stage_updates')) return [];
    throw error;
  }
}

async function createStageUpdate(input) {
  const payload = {
    id: input.id,
    project_row_id: Number(input.projectRowId || 0),
    project_number: input.projectNumber || '',
    project_display: input.projectDisplay || '',
    client: input.client || '',
    spool_iso: input.spoolIso || '',
    spool_description: input.spoolDescription || '',
    sector: input.sector || '',
    progress: Number(input.progress || 0),
    completion_date: input.completionDate || null,
    note: input.note || '',
    status: input.status || 'pending',
    created_by: input.createdBy || '',
    created_by_name: input.createdByName || '',
    resolved_by: input.resolvedBy || null,
    resolved_by_name: input.resolvedByName || null,
    resolved_at: input.resolvedAt || null,
    resolution_note: input.resolutionNote || '',
  };
  const rows = await supabaseFetch('/rest/v1/stage_updates?select=*', {
    method: 'POST',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapStageUpdate(Array.isArray(rows) ? rows[0] : null);
}


async function deleteStageUpdates(updateIds = []) {
  const ids = Array.from(new Set((Array.isArray(updateIds) ? updateIds : [updateIds])
    .map((id) => String(id || '').trim())
    .filter(Boolean)));
  for (const id of ids) {
    const q = encodeURIComponent(id);
    await supabaseFetch(`/rest/v1/stage_updates?id=eq.${q}`, {
      method: 'DELETE',
      headers: getSupabaseHeaders(),
    });
  }
  return ids.length;
}

async function updateStageUpdate(updateId, updates) {
  const q = encodeURIComponent(String(updateId || '').trim());
  const payload = {};
  if ('status' in updates) payload.status = updates.status || 'pending';
  if ('resolvedBy' in updates) payload.resolved_by = updates.resolvedBy || '';
  if ('resolvedByName' in updates) payload.resolved_by_name = updates.resolvedByName || '';
  if ('resolvedAt' in updates) payload.resolved_at = updates.resolvedAt || null;
  if ('resolutionNote' in updates) payload.resolution_note = updates.resolutionNote || '';
  const rows = await supabaseFetch(`/rest/v1/stage_updates?id=eq.${q}&select=*`, {
    method: 'PATCH',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapStageUpdate(Array.isArray(rows) ? rows[0] : null);
}

module.exports = {
  SUPABASE_URL,
  SUPABASE_ANON_KEY,
  isSupabaseConfigured,
  listUsers,
  getUserByUsername,
  getUserById,
  insertUser,
  updateUser,
  listManualAlerts,
  markUserPresenceOffline,
  upsertUserPresence,
  listUserPresence,
  createManualAlert,
  listAcknowledgements,
  addAcknowledgement,
  findAcknowledgement,
  listAlertResponses,
  createAlertResponse,
  updateAlertResponse,
  userPasswordMatches,
  mapUser,
  mapAlert,
  mapAck,
  mapResponse,
  mapStageUpdate,
  mapPresence,
  hashPassword,
  normalizeSectorList,
  normalizeText,
  listPushSubscriptions,
  upsertPushSubscription,
  removePushSubscription,
  listStageUpdates,
  createStageUpdate,
  updateStageUpdate,
  deleteStageUpdates,
};
