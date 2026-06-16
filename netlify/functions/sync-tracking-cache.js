const { jsonResponse, requireSession } = require('./_auth');
const { buildPayload, fetchCurrentTrackingSheetVersion } = require('./projects');
const {
  getTrackingCacheMeta,
  acquireTrackingRefreshLock,
  releaseTrackingRefreshLock,
  clearTrackingRefreshLock,
  touchTrackingCache,
  TRACKING_CACHE_MIN_WRITE_INTERVAL_MS,
  TRACKING_CACHE_REFRESH_LOCK_TTL_MS,
  TRACKING_CACHE_MAX_LOCK_HOLD_MS,
} = require('./_trackingCache');

const OPERATION_REGION = 'BR';
const TRACKING_CACHE_KEY = `projects:${OPERATION_REGION}:current`;
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
const MANUAL_REFRESH_MIN_AGE_MS = Math.max(
  0,
  Number(process.env.TRACKING_CACHE_MANUAL_REFRESH_MIN_AGE_MS || 2 * 60 * 1000)
);

// v37.25: mantém qualquer operação manual/agendada abaixo do limite prático da Function.
// Regra nova: botão manual e rotina agendada devem tocar o updated_at do cache
// assim que pegam o lock. A UI passa a refletir a última checagem confirmada no Supabase,
// mesmo quando a versão da sheet não mudou ou o Smartsheet demora.
const VERSION_CHECK_TIMEOUT_MS = Math.max(1500, Number(process.env.TRACKING_CACHE_VERSION_CHECK_TIMEOUT_MS || 4500));
const SYNC_FULL_SHEET_TIMEOUT_MS = Math.max(3500, Number(process.env.TRACKING_CACHE_SYNC_FULL_SHEET_TIMEOUT_MS || 20000));
const SYNC_WIP_PO_TIMEOUT_MS = Math.max(1500, Number(process.env.TRACKING_CACHE_SYNC_WIP_PO_TIMEOUT_MS || 4500));
const FORCED_FULL_REFRESH_INTERVAL_MS = Math.max(
  15 * 60 * 1000,
  Number(process.env.TRACKING_CACHE_FORCED_FULL_REFRESH_INTERVAL_MS || 60 * 60 * 1000)
);

function parseBoolean(value) {
  return ['1', 'true', 'yes', 'sim'].includes(String(value || '').trim().toLowerCase());
}

function shortError(error) {
  return String(error?.message || error || 'Falha desconhecida').slice(0, 280);
}

function responseWithKeptCache({ meta, metaAfterLock, metaBefore, mode, manual, startedAt, reason, error, lock }) {
  const bestMeta = meta || metaAfterLock || metaBefore || null;
  return jsonResponse(200, {
    ok: true,
    synced: false,
    skipped: true,
    staleCacheKept: true,
    reason,
    warning: error ? shortError(error) : null,
    mode,
    manual,
    cacheKey: TRACKING_CACHE_KEY,
    durationMs: Date.now() - startedAt,
    projectsCount: Number(bestMeta?.projectsCount || 0),
    version: bestMeta?.version || null,
    cacheUpdatedAt: bestMeta?.updatedAt || null,
    cacheAgeMs: bestMeta?.ageMs ?? null,
    refreshStartedAt: bestMeta?.refreshStartedAt || null,
    lockedUntil: bestMeta?.refreshLockUntil || null,
    lockOwner: lock?.owner || bestMeta?.refreshLockOwner || null,
  }, {
    headers: { 'cache-control': 'no-store' },
  });
}

/**
 * v37.22: sincronizador manual/controlado do cache operacional.
 *
 * Regra operacional:
 * - Login e abertura continuam cache-only pelo Supabase.
 * - Botão Atualizar tenta checar a versão do Smartsheet de forma leve.
 * - Se a versão não mudou, só atualiza o timestamp do cache.
 * - Se a versão mudou, tenta baixar a sheet completa dentro de um orçamento curto.
 * - Se o Smartsheet não responder, mantém o cache atual e não deixa lock preso.
 */
exports.handler = async (event) => {
  if (event.httpMethod && event.httpMethod !== 'POST' && event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  const query = event.queryStringParameters || {};
  const force = parseBoolean(query.force);
  const manual = parseBoolean(query.manual);
  const mode = String(query.mode || (force ? (manual ? 'manual' : 'force') : 'auto')).toLowerCase();
  const startedAt = Date.now();

  let lock = null;
  let metaBeforeSnapshot = null;

  try {
    const metaBefore = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 });
    metaBeforeSnapshot = metaBefore;
    const ageMs = metaBefore?.ageMs == null ? null : Number(metaBefore.ageMs);
    const hasUsableCache = Boolean(metaBefore?.updatedAt && Number(metaBefore.projectsCount || 0) > 0);

    // v37.25: clique manual deve SEMPRE registrar uma nova checagem no Supabase.
    // Não bloqueamos mais por idade mínima aqui; o lock continua evitando congestionamento.
    // A variável TRACKING_CACHE_MANUAL_REFRESH_MIN_AGE_MS fica ignorada para o botão Atualizar.

    if (!force && hasUsableCache && ageMs != null && ageMs < AUTO_REFRESH_AFTER_MS) {
      return jsonResponse(200, {
        ok: true,
        synced: false,
        skipped: true,
        reason: 'cache-fresh',
        mode,
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaBefore.updatedAt,
        cacheAgeMs: ageMs,
        nextRefreshInMs: Math.max(0, AUTO_REFRESH_AFTER_MS - ageMs),
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
        manual,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    lock = await acquireTrackingRefreshLock(TRACKING_CACHE_KEY, {
      ttlMs: LOCK_TTL_MS,
      owner: `${manual ? 'manual' : 'user'}-${auth.session?.sub || auth.session?.username || 'user'}-${Date.now()}`,
    });

    if (!lock?.acquired && manual && force && ['supabase-lock-held', 'lock-not-acquired'].includes(String(lock?.reason || ''))) {
      await clearTrackingRefreshLock(TRACKING_CACHE_KEY, {
        reason: 'manual-force-cleared-refresh-lock',
        timeoutMs: 7000,
      }).catch(() => false);
      lock = await acquireTrackingRefreshLock(TRACKING_CACHE_KEY, {
        ttlMs: LOCK_TTL_MS,
        owner: `manual-force-${auth.session?.sub || auth.session?.username || 'user'}-${Date.now()}`,
      });
    }

    if (!lock?.acquired) {
      return jsonResponse(200, {
        ok: true,
        synced: false,
        skipped: true,
        reason: lock?.reason || 'lock-held',
        mode,
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaBefore?.updatedAt || null,
        cacheAgeMs: ageMs,
        lockedUntil: lock?.lockedUntil || null,
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
        manual,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    const metaAfterLock = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 });
    const ageAfterLockMs = metaAfterLock?.ageMs == null ? null : Number(metaAfterLock.ageMs);
    const freshAfterLock = Boolean(!force && metaAfterLock?.updatedAt && Number(metaAfterLock.projectsCount || 0) > 0 && ageAfterLockMs != null && ageAfterLockMs < AUTO_REFRESH_AFTER_MS);
    if (freshAfterLock) {
      return jsonResponse(200, {
        ok: true,
        synced: false,
        skipped: true,
        reason: 'cache-became-fresh',
        mode,
        cacheKey: TRACKING_CACHE_KEY,
        cacheUpdatedAt: metaAfterLock.updatedAt,
        cacheAgeMs: ageAfterLockMs,
        nextRefreshInMs: Math.max(0, AUTO_REFRESH_AFTER_MS - ageAfterLockMs),
        autoRefreshAfterMs: AUTO_REFRESH_AFTER_MS,
        manual,
      }, { headers: { 'cache-control': 'no-store' } });
    }

    // v37.25: assim que o lock é adquirido, registra uma checagem real no Supabase.
    // Isso garante que o campo updated_at e o painel não fiquem acima de 15 minutos
    // quando a sheet não mudou ou quando a checagem do Smartsheet demora/falha.
    const heartbeatTouch = await touchTrackingCache(TRACKING_CACHE_KEY, {
      version: metaAfterLock?.version || metaBefore?.version || '',
      reason: manual ? 'manual-refresh-started-cache-touched' : 'refresh-started-cache-touched',
      source: manual ? 'manual-refresh' : 'sync-tracking-cache',
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
    const fullRefreshDue = !force && hasPersistentPayload && ageAfterLockMs != null && ageAfterLockMs >= FORCED_FULL_REFRESH_INTERVAL_MS;

    if (!force && currentVersion && previousVersion && currentVersion === previousVersion && hasPersistentPayload && !fullRefreshDue) {
      const touched = await touchTrackingCache(TRACKING_CACHE_KEY, {
        version: currentVersion,
        reason: manual ? 'manual-refresh-version-unchanged' : 'refresh-version-unchanged',
        source: manual ? 'manual-refresh' : 'sync-tracking-cache',
        scope: 'single-current-cache',
      });
      const metaAfterTouch = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);
      if (!touched && !metaAfterTouch?.updatedAt && !metaAfterHeartbeat?.updatedAt) {
        return responseWithKeptCache({
          meta: metaAfterTouch || metaAfterHeartbeat || metaAfterLock,
          metaAfterLock,
          metaBefore,
          mode,
          manual,
          startedAt,
          reason: 'touch-failed-cache-kept',
          error: 'Supabase não confirmou a atualização do timestamp.',
          lock,
        });
      }
      return jsonResponse(200, {
        ok: true,
        synced: true,
        skipped: false,
        touched: true,
        versionChanged: false,
        reason: 'version-unchanged-cache-touched',
        mode,
        manual,
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
        persistentWriteReason: manual ? 'manual-full-reconcile-tracking-cache' : 'full-reconcile-tracking-cache',
        persistentWriteSource: manual ? 'manual-refresh' : 'sync-tracking-cache',
        fullSheetTimeoutMs: SYNC_FULL_SHEET_TIMEOUT_MS,
        wipPoTimeoutMs: SYNC_WIP_PO_TIMEOUT_MS,
      });
    } catch (error) {
      const metaAfterFailure = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => metaAfterHeartbeat || metaAfterLock || metaBefore);
      return responseWithKeptCache({
        meta: metaAfterFailure || metaAfterHeartbeat,
        metaAfterLock,
        metaBefore,
        mode,
        manual,
        startedAt,
        reason: 'full-refresh-failed-cache-kept',
        error,
        lock,
      });
    }

    const metaAfterWrite = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);

    return jsonResponse(200, {
      ok: true,
      synced: true,
      skipped: false,
      touched: false,
      versionChanged: currentVersion && previousVersion ? currentVersion !== previousVersion : null,
      mode,
      manual,
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
    // v37.23: evita popup falso quando o cache existe mas uma leitura curta do Supabase/Smartsheet falhou.
    // O painel deve manter a base atual visível e informar que a tentativa não foi concluída.
    const metaAfterError = await getTrackingCacheMeta(TRACKING_CACHE_KEY, { timeoutMs: 7000 }).catch(() => null);
    const bestMeta = metaAfterError || metaBeforeSnapshot || null;
    if (bestMeta?.updatedAt) {
      return responseWithKeptCache({
        meta: bestMeta,
        mode,
        manual,
        startedAt,
        reason: 'unexpected-error-cache-kept',
        error,
        lock,
      });
    }
    return jsonResponse(500, {
      ok: false,
      synced: false,
      error: error.message || 'Falha ao sincronizar cache.',
    }, { headers: { 'cache-control': 'no-store' } });
  } finally {
    if (lock?.acquired && lock.owner) {
      await releaseTrackingRefreshLock(TRACKING_CACHE_KEY, lock.owner, { reason: 'refresh-lock-released' }).catch(() => null);
    }
  }
};
