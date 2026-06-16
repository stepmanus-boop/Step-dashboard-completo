const { jsonResponse } = require('./_auth');
const { buildPayload, fetchCurrentTrackingSheetVersion } = require('./projects');
const {
  getTrackingCacheMeta,
  acquireTrackingRefreshLock,
  releaseTrackingRefreshLock,
  touchTrackingCache,
  TRACKING_CACHE_MIN_WRITE_INTERVAL_MS,
  TRACKING_CACHE_REFRESH_LOCK_TTL_MS,
  TRACKING_CACHE_MAX_LOCK_HOLD_MS,
} = require('./_trackingCache');

const OPERATION_REGION = 'BR';
const TRACKING_CACHE_KEY = `projects:${OPERATION_REGION}:current`;
const SCHEDULE_ENABLED = String(process.env.TRACKING_CACHE_SCHEDULE_ENABLED || '1') !== '0';
const AUTO_REFRESH_AFTER_MS = Math.max(
  60 * 1000,
  Number(process.env.TRACKING_CACHE_AUTO_REFRESH_AFTER_MS || process.env.TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || 15 * 60 * 1000)
);
const LOCK_TTL_MS = Math.max(
  60 * 1000,
  Math.min(
    Number(process.env.TRACKING_CACHE_REFRESH_LOCK_TTL_MS || TRACKING_CACHE_REFRESH_LOCK_TTL_MS || 5 * 60 * 1000),
    Number(process.env.TRACKING_CACHE_MAX_LOCK_HOLD_MS || TRACKING_CACHE_MAX_LOCK_HOLD_MS || 90 * 1000)
  )
);

const VERSION_CHECK_TIMEOUT_MS = Math.max(1500, Number(process.env.TRACKING_CACHE_VERSION_CHECK_TIMEOUT_MS || 4500));
const SYNC_FULL_SHEET_TIMEOUT_MS = Math.max(3500, Number(process.env.TRACKING_CACHE_SYNC_FULL_SHEET_TIMEOUT_MS || 20000));
const SYNC_WIP_PO_TIMEOUT_MS = Math.max(1500, Number(process.env.TRACKING_CACHE_SYNC_WIP_PO_TIMEOUT_MS || 4500));
const FORCED_FULL_REFRESH_INTERVAL_MS = Math.max(
  15 * 60 * 1000,
  Number(process.env.TRACKING_CACHE_FORCED_FULL_REFRESH_INTERVAL_MS || 60 * 60 * 1000)
);

function shortError(error) {
  return String(error?.message || error || 'Falha desconhecida').slice(0, 280);
}

function keptCacheResponse({ meta, metaAfterLock, metaBefore, startedAt, reason, error, lock }) {
  const bestMeta = meta || metaAfterLock || metaBefore || null;
  return jsonResponse(200, {
    ok: true,
    scheduled: true,
    synced: false,
    skipped: true,
    staleCacheKept: true,
    reason,
    warning: error ? shortError(error) : null,
    cacheKey: TRACKING_CACHE_KEY,
    durationMs: Date.now() - startedAt,
    projectsCount: Number(bestMeta?.projectsCount || 0),
    version: bestMeta?.version || null,
    cacheUpdatedAt: bestMeta?.updatedAt || null,
    cacheAgeMs: bestMeta?.ageMs ?? null,
    refreshStartedAt: bestMeta?.refreshStartedAt || null,
    lockedUntil: bestMeta?.refreshLockUntil || null,
    lockOwner: lock?.owner || bestMeta?.refreshLockOwner || null,
  }, { headers: { 'cache-control': 'no-store' } });
}

/**
 * v37.22: rotina independente do usuário, protegida contra timeout e lock preso.
 *
 * A rotina só faz trabalho pesado se a versão do Smartsheet mudou. Se a checagem leve
 * falhar, ela mantém o cache atual, libera o lock e tenta novamente na próxima rodada.
 */
exports.handler = async (event = {}) => {
  const startedAt = Date.now();
  let lock = null;
  let metaBeforeSnapshot = null;

  if (!SCHEDULE_ENABLED) {
    return jsonResponse(200, {
      ok: true,
      scheduled: true,
      synced: false,
      skipped: true,
      reason: 'schedule-disabled',
      cacheKey: TRACKING_CACHE_KEY,
    }, { headers: { 'cache-control': 'no-store' } });
  }

  try {
    const metaBefore = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 });
    metaBeforeSnapshot = metaBefore;
    const ageMs = metaBefore?.ageMs == null ? null : Number(metaBefore.ageMs);
    const hasUsableCache = Boolean(metaBefore?.updatedAt && Number(metaBefore.projectsCount || 0) > 0);

    if (hasUsableCache && ageMs != null && ageMs < AUTO_REFRESH_AFTER_MS) {
      return jsonResponse(200, {
        ok: true,
        scheduled: true,
        synced: false,
        skipped: true,
        reason: 'cache-fresh',
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaBefore.updatedAt,
        cacheAgeMs: ageMs,
        nextRefreshInMs: Math.max(0, AUTO_REFRESH_AFTER_MS - ageMs),
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    lock = await acquireTrackingRefreshLock(TRACKING_CACHE_KEY, {
      ttlMs: LOCK_TTL_MS,
      owner: `scheduled-${process.env.SITE_NAME || 'step'}-${Date.now()}`,
    });

    if (!lock?.acquired) {
      return jsonResponse(200, {
        ok: true,
        scheduled: true,
        synced: false,
        skipped: true,
        reason: lock?.reason || 'lock-held',
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaBefore?.updatedAt || null,
        cacheAgeMs: ageMs,
        lockedUntil: lock?.lockedUntil || null,
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    const metaAfterLock = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 });
    const ageAfterLockMs = metaAfterLock?.ageMs == null ? null : Number(metaAfterLock.ageMs);
    const freshAfterLock = Boolean(metaAfterLock?.updatedAt && Number(metaAfterLock.projectsCount || 0) > 0 && ageAfterLockMs != null && ageAfterLockMs < AUTO_REFRESH_AFTER_MS);
    if (freshAfterLock) {
      return jsonResponse(200, {
        ok: true,
        scheduled: true,
        synced: false,
        skipped: true,
        reason: 'cache-became-fresh',
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaAfterLock.updatedAt,
        cacheAgeMs: ageAfterLockMs,
        nextRefreshInMs: Math.max(0, AUTO_REFRESH_AFTER_MS - ageAfterLockMs),
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    // v37.25: a rotina agendada registra a checagem no Supabase logo após pegar o lock.
    // Assim o painel não fica com idade acima de 15 minutos quando a versão não mudou
    // ou quando a checagem do Smartsheet falha/expira.
    const heartbeatTouch = await touchTrackingCache(TRACKING_CACHE_KEY, {
      version: metaAfterLock?.version || metaBefore?.version || '',
      reason: 'scheduled-refresh-started-cache-touched',
      source: 'scheduled-sync-tracking-cache',
      scope: 'single-current-cache',
      timeoutMs: 7000,
    }).catch(() => null);
    const metaAfterHeartbeat = heartbeatTouch
      ? await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null)
      : null;

    let currentVersionInfo = null;
    try {
      currentVersionInfo = await fetchCurrentTrackingSheetVersion({ timeoutMs: VERSION_CHECK_TIMEOUT_MS, retries: 1 });
    } catch (error) {
      currentVersionInfo = null;
    }

    const currentVersion = String(currentVersionInfo?.version || '');
    const previousVersion = String(metaAfterLock?.version || metaBefore?.version || '');
    const hasPersistentPayload = Boolean((metaAfterLock?.projectsCount || metaBefore?.projectsCount || 0) > 0);
    const fullRefreshDue = hasPersistentPayload && ageAfterLockMs != null && ageAfterLockMs >= FORCED_FULL_REFRESH_INTERVAL_MS;

    if (currentVersion && previousVersion && currentVersion === previousVersion && hasPersistentPayload && !fullRefreshDue) {
      const touched = await touchTrackingCache(TRACKING_CACHE_KEY, {
        version: currentVersion,
        reason: 'scheduled-refresh-version-unchanged',
        source: 'scheduled-sync-tracking-cache',
        scope: 'single-current-cache',
      });
      const metaAfterTouch = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);
      if (!touched && !metaAfterTouch?.updatedAt && !metaAfterHeartbeat?.updatedAt) {
        return keptCacheResponse({
          meta: metaAfterTouch || metaAfterHeartbeat || metaAfterLock,
          metaAfterLock,
          metaBefore,
          startedAt,
          reason: 'touch-failed-cache-kept',
          error: 'Supabase não confirmou o touch do cache.',
          lock,
        });
      }
      return jsonResponse(200, {
        ok: true,
        scheduled: true,
        synced: true,
        skipped: false,
        touched: true,
        versionChanged: false,
        reason: 'version-unchanged-cache-touched',
        cacheKey: TRACKING_CACHE_KEY,
        durationMs: Date.now() - startedAt,
        projectsCount: Number(metaAfterTouch?.projectsCount || metaAfterLock?.projectsCount || metaBefore?.projectsCount || 0),
        version: currentVersion,
        lastSync: metaAfterTouch?.updatedAt || touched?.updatedAt || metaAfterHeartbeat?.updatedAt || new Date().toISOString(),
        cacheUpdatedAt: metaAfterTouch?.updatedAt || touched?.updatedAt || metaAfterHeartbeat?.updatedAt || null,
        cacheAgeMs: metaAfterTouch?.ageMs ?? metaAfterHeartbeat?.ageMs ?? 0,
        lockOwner: lock.owner || null,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    let payload;
    try {
      payload = await buildPayload({
        force: true,
        preferCache: false,
        waitForPersistentCacheWrite: true,
        forcePersistentCacheWrite: true,
        persistentWriteReason: 'scheduled-full-reconcile-tracking-cache',
        persistentWriteSource: 'scheduled-sync-tracking-cache',
        fullSheetTimeoutMs: SYNC_FULL_SHEET_TIMEOUT_MS,
        wipPoTimeoutMs: SYNC_WIP_PO_TIMEOUT_MS,
      });
    } catch (error) {
      const metaAfterFailure = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => metaAfterHeartbeat || metaAfterLock || metaBefore);
      return keptCacheResponse({
        meta: metaAfterFailure || metaAfterHeartbeat,
        metaAfterLock,
        metaBefore,
        startedAt,
        reason: 'full-refresh-failed-cache-kept',
        error,
        lock,
      });
    }

    const metaAfterWrite = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);

    return jsonResponse(200, {
      ok: true,
      scheduled: true,
      synced: true,
      skipped: false,
      touched: false,
      versionChanged: currentVersion && previousVersion ? currentVersion !== previousVersion : null,
      cacheKey: TRACKING_CACHE_KEY,
      durationMs: Date.now() - startedAt,
      projectsCount: Array.isArray(payload.projects) ? payload.projects.length : 0,
      version: payload.meta?.version || currentVersion || null,
      lastSync: payload.meta?.lastSync || new Date().toISOString(),
      cacheUpdatedAt: metaAfterWrite?.updatedAt || metaAfterHeartbeat?.updatedAt || null,
      cacheAgeMs: metaAfterWrite?.ageMs ?? metaAfterHeartbeat?.ageMs ?? null,
      lockOwner: lock.owner || null,
    }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    const metaAfterError = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);
    const bestMeta = metaAfterError || metaBeforeSnapshot || null;
    if (bestMeta?.updatedAt) {
      return keptCacheResponse({
        meta: bestMeta,
        startedAt,
        reason: 'unexpected-error-cache-kept',
        error,
        lock,
      });
    }
    return jsonResponse(500, {
      ok: false,
      scheduled: true,
      synced: false,
      error: error.message || 'Falha ao sincronizar cache agendado.',
    }, { headers: { 'cache-control': 'no-store' } });
  } finally {
    if (lock?.acquired && lock.owner) {
      await releaseTrackingRefreshLock(TRACKING_CACHE_KEY, lock.owner, { reason: 'refresh-lock-released' }).catch(() => null);
    }
  }
};
