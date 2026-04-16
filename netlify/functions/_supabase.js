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
    active: row.active !== false,
    createdAt: row.created_at || null,
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
  if ('active' in updates) payload.active = updates.active !== false;
  const rows = await supabaseFetch(`/rest/v1/users?id=eq.${q}&select=*`, {
    method: 'PATCH',
    headers: getSupabaseHeaders('return=representation'),
    body: JSON.stringify(payload),
  });
  return mapUser(Array.isArray(rows) ? rows[0] : null);
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
  hashPassword,
  normalizeSectorList,
  normalizeText,
};
