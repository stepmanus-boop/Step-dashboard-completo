const { jsonResponse, getSession } = require('./_auth');
const { isSupabaseConfigured, getUserById } = require('./_supabase');

const AUTH_ME_CACHE_MS = Number(process.env.AUTH_ME_CACHE_MS || 5 * 60 * 1000);
const authMeUserCache = global.__AUTH_ME_USER_CACHE__ || {};
global.__AUTH_ME_USER_CACHE__ = authMeUserCache;

function getCachedFreshUser(sub) {
  const cacheKey = String(sub || '');
  if (!cacheKey) return null;
  const cached = authMeUserCache[cacheKey];
  if (!cached?.user) return null;
  if (Date.now() - Number(cached.savedAt || 0) > AUTH_ME_CACHE_MS) return null;
  return cached.user;
}

exports.handler = async (event) => {
  const session = getSession(event);
  if (!session) {
    return jsonResponse(200, { ok: true, authenticated: false, publicAccess: true, githubSyncEnabled: true, supabaseEnabled: isSupabaseConfigured() });
  }

  let freshUser = null;
  if (isSupabaseConfigured() && session.sub) {
    freshUser = getCachedFreshUser(session.sub);
    if (!freshUser) {
      try {
        freshUser = await getUserById(session.sub);
        if (freshUser) authMeUserCache[String(session.sub)] = { savedAt: Date.now(), user: freshUser };
      } catch (_) {
        freshUser = null;
      }
    }
  }

  const user = freshUser || session;
  return jsonResponse(200, {
    ok: true,
    authenticated: true,
    githubSyncEnabled: true,
    supabaseEnabled: isSupabaseConfigured(),
    user: {
      id: user.id || user.sub,
      name: user.name,
      username: user.username,
      role: user.role,
      sector: user.sector,
      alertSectors: Array.isArray(user.alertSectors) ? user.alertSectors : (user.sector && user.sector !== 'all' ? [user.sector] : []),
      projectPmAliases: Array.isArray(user.projectPmAliases) ? user.projectPmAliases : [],
      qualityCompetencies: Array.isArray(user.qualityCompetencies) ? user.qualityCompetencies : [],
      clientKey: user.clientKey || '',
      operationRegion: user.operationRegion || 'BR',
      siteKey: user.siteKey || user.operationRegion || 'BR',
      portalSite: user.portalSite || user.siteKey || user.operationRegion || 'BR',
      clientName: user.clientName || '',
      clientLogoUrl: user.clientLogoUrl || '',
      clientPlatformImageUrl: user.clientPlatformImageUrl || '',
      clientPlatformImages: user.clientPlatformImages || {},
      allowedClients: Array.isArray(user.allowedClients) ? user.allowedClients : [],
    },
  });
};
