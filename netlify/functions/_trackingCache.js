const crypto = require('crypto');

function cleanEnvSecret(value = '') {
  return String(value || '')
    .trim()
    .replace(/^Bearer\s+/i, '')
    .replace(/^[']|[']$/g, '')
    .replace(/^[\"]|[\"]$/g, '')
    .trim();
}

const SUPABASE_URL = String(
  process.env.SUPABASE_URL ||
  process.env.VITE_SUPABASE_URL ||
  ''
).trim().replace(/\/$/, '');

const SUPABASE_SERVICE_ROLE_KEY = cleanEnvSecret(
  process.env.SUPABASE_SERVICE_ROLE_KEY ||
  process.env.SUPABASE_SECRET_KEY ||
  process.env.SB_SECRET_KEY ||
  process.env.SUPABASE_SECRET ||
  ''
);

const TRACKING_CACHE_TABLE = String(process.env.TRACKING_CACHE_TABLE || 'step_tracking_cache').trim() || 'step_tracking_cache';
const TRACKING_CACHE_TIMEOUT_MS = Math.max(1500, Number(process.env.TRACKING_CACHE_TIMEOUT_MS || 3500));
const TRACKING_CACHE_MAX_AGE_MS = Math.max(5 * 60 * 1000, Number(process.env.TRACKING_CACHE_MAX_AGE_MS || 12 * 60 * 60 * 1000));
const TRACKING_CACHE_ENABLED = String(process.env.TRACKING_CACHE_ENABLED || '1') !== '0';

// v37.09 FREE PLAN SAFE:
// - evita crescimento de banco no Supabase Free;
// - evita gravação repetida em cold starts;
// - evita salvar payload gigante por engano.
const TRACKING_CACHE_MIN_WRITE_INTERVAL_MS = Math.max(
  60 * 1000,
  Number(process.env.TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || 15 * 60 * 1000)
);
const TRACKING_CACHE_MAX_PAYLOAD_BYTES = Math.max(
  1 * 1024 * 1024,
  Number(process.env.TRACKING_CACHE_MAX_PAYLOAD_BYTES || 24 * 1024 * 1024)
);


const TRACKING_CACHE_REFRESH_LOCK_TTL_MS = Math.max(
  60 * 1000,
  Number(process.env.TRACKING_CACHE_REFRESH_LOCK_TTL_MS || 5 * 60 * 1000)
);

// v37.22: evita lock preso por timeout da Function/Scheduled Function.
// Não é lock transacional; se a função morrer depois de adquirir o lock,
// outro processo pode limpar o lock antigo após este limite operacional.
const TRACKING_CACHE_MAX_LOCK_HOLD_MS = Math.max(
  45 * 1000,
  Number(process.env.TRACKING_CACHE_MAX_LOCK_HOLD_MS || 90 * 1000)
);

const refreshLockMemory = global.__STEP_TRACKING_CACHE_REFRESH_LOCK_MEMORY__ || {};
global.__STEP_TRACKING_CACHE_REFRESH_LOCK_MEMORY__ = refreshLockMemory;

const lastWriteByKey = global.__STEP_TRACKING_CACHE_LAST_WRITE__ || {};
global.__STEP_TRACKING_CACHE_LAST_WRITE__ = lastWriteByKey;

function isTrackingCacheConfigured() {
  return Boolean(TRACKING_CACHE_ENABLED && SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY);
}

function getHeaders(prefer = '') {
  const headers = {
    apikey: SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    'Content-Type': 'application/json',
  };
  if (prefer) headers.Prefer = prefer;
  return headers;
}

async function supabaseCacheFetch(path, options = {}) {
  if (!isTrackingCacheConfigured()) return null;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), Math.max(1500, Number(options.timeoutMs || TRACKING_CACHE_TIMEOUT_MS)));
  const { timeoutMs: _timeoutMs, ...fetchOptions } = options;
  try {
    const response = await fetch(`${SUPABASE_URL}${path}`, {
      ...fetchOptions,
      signal: controller.signal,
      headers: {
        ...getHeaders(),
        ...(fetchOptions.headers || {}),
      },
    });
    if (!response.ok) {
      const text = await response.text().catch(() => '');
      throw new Error(`Supabase cache ${response.status}: ${text}`);
    }
    const contentType = response.headers.get('content-type') || '';
    if (contentType.includes('application/json')) return response.json();
    return response.text();
  } finally {
    clearTimeout(timer);
  }
}

function normalizeCacheKey(cacheKey = '') {
  return String(cacheKey || '').trim().replace(/[^a-zA-Z0-9:_|.-]+/g, '_').slice(0, 240);
}

function byteLength(value) {
  return Buffer.byteLength(String(value || ''), 'utf8');
}

async function readTrackingCache(cacheKey, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !isTrackingCacheConfigured()) return null;

  try {
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}&select=cache_key,payload,version,updated_at,projects_count,payload_bytes,last_write_reason,refresh_started_at,refresh_lock_until,refresh_lock_owner&limit=1`, {
      method: 'GET',
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row?.payload) return null;

    const updatedAt = row.updated_at ? new Date(row.updated_at).getTime() : 0;
    const maxAgeMs = Math.max(0, Number(options.maxAgeMs || TRACKING_CACHE_MAX_AGE_MS));
    if (maxAgeMs > 0 && updatedAt > 0 && Date.now() - updatedAt > maxAgeMs) {
      return null;
    }

    return {
      payload: row.payload,
      updatedAt: row.updated_at || null,
      version: row.version || row.payload?.meta?.version || null,
      projectsCount: Number(row.projects_count || 0),
      payloadBytes: Number(row.payload_bytes || 0),
      lastWriteReason: row.last_write_reason || null,
      refreshStartedAt: row.refresh_started_at || null,
      refreshLockUntil: row.refresh_lock_until || null,
      refreshLockOwner: row.refresh_lock_owner || null,
    };
  } catch (error) {
    console.warn('[tracking-cache] Leitura ignorada:', error?.message || error);
    return null;
  }
}

async function shouldSkipWriteBySupabaseState(key, version, options = {}) {
  if (Number(options.minWriteIntervalMs || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS) <= 0) return false;

  const sinceLocalWrite = Date.now() - Number(lastWriteByKey[key] || 0);
  if (lastWriteByKey[key] && sinceLocalWrite < Number(options.minWriteIntervalMs || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS)) {
    return true;
  }

  try {
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}&select=cache_key,version,updated_at&limit=1`, {
      method: 'GET',
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row?.updated_at) return false;

    const updatedAtMs = new Date(row.updated_at).getTime();
    const recentlyWritten = updatedAtMs > 0 && Date.now() - updatedAtMs < Number(options.minWriteIntervalMs || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS);
    const sameVersion = version && row.version && String(row.version) === String(version);

    // No plano Free, evita regravar a mesma versão repetidamente.
    return Boolean(recentlyWritten && sameVersion);
  } catch (error) {
    // Se a checagem falhar, não bloqueia o sistema. A escrita tentará e, se falhar, será ignorada.
    console.warn('[tracking-cache] Checagem de intervalo ignorada:', error?.message || error);
    return false;
  }
}

async function writeTrackingCache(cacheKey, payload, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !payload || !isTrackingCacheConfigured()) return false;

  try {
    const projects = Array.isArray(payload.projects) ? payload.projects : [];
    const version = String(options.version || payload?.meta?.version || '').slice(0, 120);
    const serializedPayload = JSON.stringify(payload);
    const payloadBytes = byteLength(serializedPayload);
    const maxPayloadBytes = Number(options.maxPayloadBytes || TRACKING_CACHE_MAX_PAYLOAD_BYTES);

    if (payloadBytes > maxPayloadBytes) {
      console.warn(`[tracking-cache] Escrita ignorada: payload ${payloadBytes} bytes acima do limite ${maxPayloadBytes} bytes.`);
      return false;
    }

    // v37.19: atualização manual deve refletir a checagem no cache mesmo se a versão da sheet não mudou.
    // O intervalo/lock continua protegendo contra excesso, mas forceWrite bypassa o skip por mesma versão.
    if (!options.forceWrite && await shouldSkipWriteBySupabaseState(key, version, options)) {
      console.info('[tracking-cache] Escrita ignorada: cache recente com a mesma versão.');
      return false;
    }

    const body = {
      cache_key: key,
      scope: String(options.scope || '').slice(0, 220),
      source: String(options.source || 'projects-api').slice(0, 120),
      version,
      projects_count: projects.length,
      payload_bytes: payloadBytes,
      last_write_reason: String(options.reason || options.source || 'cache-refresh').slice(0, 220),
      payload,
      updated_at: new Date().toISOString(),
    };

    await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?on_conflict=cache_key`, {
      method: 'POST',
      headers: getHeaders('resolution=merge-duplicates,return=minimal'),
      body: JSON.stringify(body),
      timeoutMs: options.timeoutMs,
    });
    lastWriteByKey[key] = Date.now();
    return true;
  } catch (error) {
    console.warn('[tracking-cache] Escrita ignorada:', error?.message || error);
    return false;
  }
}


async function getTrackingCacheMeta(cacheKey, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !isTrackingCacheConfigured()) return null;

  try {
    const select = [
      'cache_key',
      'version',
      'updated_at',
      'projects_count',
      'payload_bytes',
      'last_write_reason',
      'refresh_lock_until',
      'refresh_lock_owner',
      'refresh_started_at',
    ].join(',');
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}&select=${select}&limit=1`, {
      method: 'GET',
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row) return null;
    const updatedAtMs = row.updated_at ? new Date(row.updated_at).getTime() : 0;
    const ageMs = updatedAtMs > 0 ? Math.max(0, Date.now() - updatedAtMs) : null;
    return {
      cacheKey: row.cache_key || key,
      version: row.version || null,
      updatedAt: row.updated_at || null,
      updatedAtMs,
      ageMs,
      projectsCount: Number(row.projects_count || 0),
      payloadBytes: Number(row.payload_bytes || 0),
      lastWriteReason: row.last_write_reason || null,
      refreshLockUntil: row.refresh_lock_until || null,
      refreshLockOwner: row.refresh_lock_owner || null,
      refreshStartedAt: row.refresh_started_at || null,
    };
  } catch (error) {
    console.warn('[tracking-cache] Metadados ignorados:', error?.message || error);
    return null;
  }
}

async function ensureTrackingCachePlaceholderRow(key, owner, lockUntilIso, nowIso, options = {}) {
  const body = {
    cache_key: key,
    scope: 'single-current-cache',
    source: 'refresh-lock-bootstrap',
    version: '',
    projects_count: 0,
    payload_bytes: 0,
    last_write_reason: 'refresh-lock-bootstrap',
    payload: {
      ok: false,
      projects: [],
      meta: {
        initializing: true,
        reason: 'refresh-lock-bootstrap',
        createdAt: nowIso,
      },
    },
    updated_at: nowIso,
    refresh_lock_until: lockUntilIso,
    refresh_lock_owner: owner,
    refresh_started_at: nowIso,
  };

  try {
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?on_conflict=cache_key`, {
      method: 'POST',
      headers: getHeaders('resolution=ignore-duplicates,return=representation'),
      body: JSON.stringify(body),
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    return Boolean(row && row.cache_key === key);
  } catch (error) {
    console.warn('[tracking-cache] Bootstrap de lock ignorado:', error?.message || error);
    return false;
  }
}

async function acquireTrackingRefreshLock(cacheKey, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !isTrackingCacheConfigured()) {
    return { acquired: false, reason: 'cache-not-configured' };
  }

  const ttlMs = Math.max(60 * 1000, Number(options.ttlMs || TRACKING_CACHE_REFRESH_LOCK_TTL_MS));
  const now = Date.now();
  const nowIso = new Date(now).toISOString();
  const lockUntilIso = new Date(now + ttlMs).toISOString();
  const owner = String(options.owner || `${process.env.SITE_NAME || 'step'}-${process.pid}-${now}-${crypto.randomBytes(4).toString('hex')}`).slice(0, 180);

  const localLock = refreshLockMemory[key];
  if (localLock?.until && Number(localLock.until) > now) {
    return {
      acquired: false,
      reason: 'local-lock-held',
      lockedUntil: new Date(Number(localLock.until)).toISOString(),
      owner: localLock.owner || null,
    };
  }

  const current = await getTrackingCacheMeta(key, options);
  if (!current) {
    const bootstrapped = await ensureTrackingCachePlaceholderRow(key, owner, lockUntilIso, nowIso, options);
    if (bootstrapped) {
      refreshLockMemory[key] = { until: now + ttlMs, owner };
      return { acquired: true, owner, lockedUntil: lockUntilIso, bootstrapped: true };
    }
  } else if (current.refreshLockUntil) {
    const lockUntilMs = new Date(current.refreshLockUntil).getTime();
    const startedAtMs = current.refreshStartedAt ? new Date(current.refreshStartedAt).getTime() : 0;
    const staleByStartedAt = startedAtMs > 0 && now - startedAtMs > TRACKING_CACHE_MAX_LOCK_HOLD_MS;
    if (lockUntilMs > now && !staleByStartedAt) {
      return {
        acquired: false,
        reason: 'supabase-lock-held',
        lockedUntil: current.refreshLockUntil,
        owner: current.refreshLockOwner || null,
      };
    }
    if (lockUntilMs > now && staleByStartedAt) {
      try {
        await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}`, {
          method: 'PATCH',
          headers: getHeaders('return=minimal'),
          body: JSON.stringify({
            refresh_lock_until: null,
            refresh_lock_owner: null,
            last_write_reason: 'stale-refresh-lock-cleared',
          }),
          timeoutMs: options.timeoutMs,
        });
        if (refreshLockMemory[key]) delete refreshLockMemory[key];
      } catch (error) {
        console.warn('[tracking-cache] Limpeza de lock antigo ignorada:', error?.message || error);
        return { acquired: false, reason: 'stale-lock-clear-error', error: error?.message || String(error) };
      }
    }
  }

  try {
    const filter = [
      `cache_key=eq.${encodeURIComponent(key)}`,
      `or=(refresh_lock_until.is.null,refresh_lock_until.lt.${encodeURIComponent(nowIso)})`,
    ].join('&');
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?${filter}`, {
      method: 'PATCH',
      headers: getHeaders('return=representation'),
      body: JSON.stringify({
        refresh_lock_until: lockUntilIso,
        refresh_lock_owner: owner,
        refresh_started_at: nowIso,
        last_write_reason: 'refresh-lock-acquired',
      }),
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row) {
      return { acquired: false, reason: 'lock-not-acquired' };
    }
    refreshLockMemory[key] = { until: now + ttlMs, owner };
    return { acquired: true, owner, lockedUntil: lockUntilIso };
  } catch (error) {
    console.warn('[tracking-cache] Lock não adquirido:', error?.message || error);
    return { acquired: false, reason: 'lock-error', error: error?.message || String(error) };
  }
}

async function clearTrackingRefreshLock(cacheKey, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !isTrackingCacheConfigured()) return false;

  try {
    await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}`, {
      method: 'PATCH',
      headers: getHeaders('return=minimal'),
      body: JSON.stringify({
        refresh_lock_until: null,
        refresh_lock_owner: null,
        refresh_started_at: null,
        last_write_reason: String(options.reason || 'refresh-lock-cleared').slice(0, 220),
      }),
      timeoutMs: options.timeoutMs,
    });
    if (refreshLockMemory[key]) delete refreshLockMemory[key];
    return true;
  } catch (error) {
    console.warn('[tracking-cache] Limpeza manual de lock ignorada:', error?.message || error);
    return false;
  }
}


async function releaseTrackingRefreshLock(cacheKey, owner, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !owner || !isTrackingCacheConfigured()) return false;

  try {
    await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}&refresh_lock_owner=eq.${encodeURIComponent(owner)}`, {
      method: 'PATCH',
      headers: getHeaders('return=minimal'),
      body: JSON.stringify({
        refresh_lock_until: null,
        refresh_lock_owner: null,
        last_write_reason: options.reason ? String(options.reason).slice(0, 220) : undefined,
      }),
      timeoutMs: options.timeoutMs,
    });
    if (refreshLockMemory[key]?.owner === owner) delete refreshLockMemory[key];
    return true;
  } catch (error) {
    console.warn('[tracking-cache] Release de lock ignorado:', error?.message || error);
    return false;
  }
}


async function touchTrackingCache(cacheKey, options = {}) {
  const key = normalizeCacheKey(cacheKey);
  if (!key || !isTrackingCacheConfigured()) return null;

  const nowIso = new Date().toISOString();
  const body = {
    updated_at: nowIso,
    last_write_reason: String(options.reason || 'cache-version-checked-no-change').slice(0, 220),
  };

  if (options.version != null && String(options.version || '').trim()) {
    body.version = String(options.version || '').slice(0, 120);
  }
  if (options.source != null && String(options.source || '').trim()) {
    body.source = String(options.source || '').slice(0, 120);
  }
  if (options.scope != null && String(options.scope || '').trim()) {
    body.scope = String(options.scope || '').slice(0, 220);
  }

  try {
    const rows = await supabaseCacheFetch(`/rest/v1/${encodeURIComponent(TRACKING_CACHE_TABLE)}?cache_key=eq.${encodeURIComponent(key)}`, {
      method: 'PATCH',
      headers: getHeaders('return=representation'),
      body: JSON.stringify(body),
      timeoutMs: options.timeoutMs,
    });
    const row = Array.isArray(rows) ? rows[0] : null;
    if (!row) return null;
    lastWriteByKey[key] = Date.now();
    return {
      updatedAt: row.updated_at || nowIso,
      version: row.version || body.version || null,
      projectsCount: Number(row.projects_count || 0),
      payloadBytes: Number(row.payload_bytes || 0),
      lastWriteReason: row.last_write_reason || body.last_write_reason,
    };
  } catch (error) {
    console.warn('[tracking-cache] Touch ignorado:', error?.message || error);
    return null;
  }
}

module.exports = {
  TRACKING_CACHE_MAX_AGE_MS,
  TRACKING_CACHE_MIN_WRITE_INTERVAL_MS,
  TRACKING_CACHE_MAX_PAYLOAD_BYTES,
  TRACKING_CACHE_REFRESH_LOCK_TTL_MS,
  TRACKING_CACHE_MAX_LOCK_HOLD_MS,
  isTrackingCacheConfigured,
  readTrackingCache,
  writeTrackingCache,
  getTrackingCacheMeta,
  acquireTrackingRefreshLock,
  releaseTrackingRefreshLock,
  clearTrackingRefreshLock,
  touchTrackingCache,
};
