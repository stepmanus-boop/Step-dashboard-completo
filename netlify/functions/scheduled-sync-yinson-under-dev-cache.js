const { jsonResponse } = require('./_auth');
const {
  getTrackingCacheMeta,
  TRACKING_CACHE_MIN_WRITE_INTERVAL_MS,
} = require('./_trackingCache');
const {
  refreshYinsonUnderDevCache,
  YINSON_UNDER_DEV_CACHE_KEY,
} = require('./client-under-dev');

const SCHEDULE_ENABLED = String(process.env.YINSON_UNDER_DEV_SCHEDULE_ENABLED || process.env.TRACKING_CACHE_SCHEDULE_ENABLED || '1') !== '0';
const AUTO_REFRESH_AFTER_MS = Math.max(
  60 * 1000,
  Number(process.env.YINSON_UNDER_DEV_AUTO_REFRESH_AFTER_MS || process.env.TRACKING_CACHE_AUTO_REFRESH_AFTER_MS || TRACKING_CACHE_MIN_WRITE_INTERVAL_MS || 15 * 60 * 1000)
);

function shortError(error) {
  return String(error?.message || error || 'Falha desconhecida').slice(0, 500);
}

exports.handler = async (event = {}) => {
  const startedAt = Date.now();

  if (!SCHEDULE_ENABLED) {
    return jsonResponse(200, {
      ok: true,
      scheduled: true,
      skipped: true,
      reason: 'schedule-disabled',
      cacheKey: YINSON_UNDER_DEV_CACHE_KEY,
    }, { headers: { 'cache-control': 'no-store' } });
  }

  const metaBefore = await getTrackingCacheMeta(YINSON_UNDER_DEV_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);
  const ageMs = metaBefore?.ageMs == null ? null : Number(metaBefore.ageMs);
  const hasPayload = Boolean(metaBefore?.updatedAt && Number(metaBefore.projectsCount || 0) > 0);

  if (hasPayload && ageMs != null && ageMs < AUTO_REFRESH_AFTER_MS) {
    return jsonResponse(200, {
      ok: true,
      scheduled: true,
      skipped: true,
      reason: 'cache-fresh',
      cacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      cacheUpdatedAt: metaBefore.updatedAt,
      cacheAgeMs: ageMs,
      nextRefreshInMs: Math.max(0, AUTO_REFRESH_AFTER_MS - ageMs),
    }, { headers: { 'cache-control': 'no-store' } });
  }

  try {
    const result = await refreshYinsonUnderDevCache({
      queryStringParameters: {
        force: '1',
        scheduled: '1',
      },
    }, {
      source: 'scheduled-sync-yinson-under-dev-cache',
      reason: 'scheduled-refresh-yinson-under-dev',
      owner: `scheduled-yinson-${process.env.SITE_NAME || 'step'}-${Date.now()}`,
    });

    const metaAfter = await getTrackingCacheMeta(YINSON_UNDER_DEV_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);

    return jsonResponse(200, {
      ok: true,
      scheduled: true,
      skipped: false,
      synced: Boolean(result?.refreshed),
      cacheFallback: Boolean(result?.cacheFallback),
      cacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      durationMs: Date.now() - startedAt,
      rows: Array.isArray(result?.payload?.data) ? result.payload.data.length : 0,
      cacheUpdatedAt: metaAfter?.updatedAt || result?.payload?.meta?.lastSync || null,
      cacheAgeMs: metaAfter?.ageMs ?? null,
      warning: result?.error ? shortError(result.error) : null,
    }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    const metaAfter = await getTrackingCacheMeta(YINSON_UNDER_DEV_CACHE_KEY, { timeoutMs: 7000 }).catch(() => metaBefore);
    return jsonResponse(metaAfter?.updatedAt ? 200 : 500, {
      ok: Boolean(metaAfter?.updatedAt),
      scheduled: true,
      skipped: Boolean(metaAfter?.updatedAt),
      staleCacheKept: Boolean(metaAfter?.updatedAt),
      cacheKey: YINSON_UNDER_DEV_CACHE_KEY,
      durationMs: Date.now() - startedAt,
      cacheUpdatedAt: metaAfter?.updatedAt || null,
      cacheAgeMs: metaAfter?.ageMs ?? null,
      error: shortError(error),
    }, { headers: { 'cache-control': 'no-store' } });
  }
};
