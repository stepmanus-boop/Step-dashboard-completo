const { jsonResponse, requireSession } = require('./_auth');
const { buildPayload } = require('./projects');
const {
  acquireTrackingRefreshLock,
  releaseTrackingRefreshLock,
  clearTrackingRefreshLock,
  TRACKING_CACHE_REFRESH_LOCK_TTL_MS,
} = require('./_trackingCache');

const OPERATION_REGION = 'BR';
const TRACKING_CACHE_KEY = `projects:${OPERATION_REGION}:current`;
const BACKGROUND_LOCK_TTL_MS = Math.max(
  5 * 60 * 1000,
  Number(process.env.TRACKING_CACHE_BACKGROUND_LOCK_TTL_MS || TRACKING_CACHE_REFRESH_LOCK_TTL_MS || 10 * 60 * 1000)
);
const BACKGROUND_FULL_SHEET_TIMEOUT_MS = Math.max(
  30 * 1000,
  Number(process.env.TRACKING_CACHE_BACKGROUND_FULL_SHEET_TIMEOUT_MS || 120 * 1000)
);
const BACKGROUND_WIP_PO_TIMEOUT_MS = Math.max(
  1500,
  Number(process.env.TRACKING_CACHE_BACKGROUND_WIP_PO_TIMEOUT_MS || 15 * 1000)
);

function parseBoolean(value) {
  return ['1', 'true', 'yes', 'sim'].includes(String(value || '').trim().toLowerCase());
}

exports.handler = async (event = {}) => {
  if (event.httpMethod && event.httpMethod !== 'POST' && event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Metodo nao permitido.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  const query = event.queryStringParameters || {};
  const manual = parseBoolean(query.manual);
  const force = parseBoolean(query.force);
  const owner = `${manual ? 'manual-background' : 'background'}-${auth.session?.sub || auth.session?.username || 'user'}-${Date.now()}`;
  let lock = null;

  try {
    if (manual && force) {
      await clearTrackingRefreshLock(TRACKING_CACHE_KEY, {
        reason: 'manual-background-cleared-refresh-lock',
        timeoutMs: 7000,
      }).catch(() => false);
    }

    lock = await acquireTrackingRefreshLock(TRACKING_CACHE_KEY, {
      ttlMs: BACKGROUND_LOCK_TTL_MS,
      owner,
      timeoutMs: 7000,
    });

    if (!lock?.acquired) {
      return jsonResponse(200, {
        ok: true,
        background: true,
        synced: false,
        skipped: true,
        reason: lock?.reason || 'lock-held',
        lockedUntil: lock?.lockedUntil || null,
        lockOwner: lock?.owner || null,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    const payload = await buildPayload({
      force: true,
      preferCache: false,
      waitForPersistentCacheWrite: true,
      forcePersistentCacheWrite: true,
      persistentWriteReason: manual ? 'manual-background-full-reconcile-tracking-cache' : 'background-full-reconcile-tracking-cache',
      persistentWriteSource: manual ? 'manual-background-refresh' : 'sync-tracking-cache-background',
      fullSheetTimeoutMs: BACKGROUND_FULL_SHEET_TIMEOUT_MS,
      wipPoTimeoutMs: BACKGROUND_WIP_PO_TIMEOUT_MS,
    });

    return jsonResponse(200, {
      ok: true,
      background: true,
      synced: true,
      skipped: false,
      projectsCount: Array.isArray(payload.projects) ? payload.projects.length : 0,
      version: payload.meta?.version || null,
      lastSync: payload.meta?.lastSync || new Date().toISOString(),
      cacheUpdatedAt: payload.meta?.lastSync || null,
    }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    return jsonResponse(500, {
      ok: false,
      background: true,
      synced: false,
      error: error.message || 'Falha na atualizacao em background.',
    }, { headers: { 'cache-control': 'no-store' } });
  } finally {
    if (lock?.acquired && lock.owner) {
      await releaseTrackingRefreshLock(TRACKING_CACHE_KEY, lock.owner, {
        reason: 'background-refresh-lock-released',
        timeoutMs: 7000,
      }).catch(() => null);
    }
  }
};
