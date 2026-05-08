const { jsonResponse, getSession } = require('./_auth');
const { isSupabaseConfigured, getUserById } = require('./_supabase');

exports.handler = async (event) => {
  const session = getSession(event);
  if (!session) {
    return jsonResponse(200, { ok: true, authenticated: false, publicAccess: true, githubSyncEnabled: true, supabaseEnabled: isSupabaseConfigured() });
  }

  let freshUser = null;
  if (isSupabaseConfigured() && session.sub) {
    try { freshUser = await getUserById(session.sub); } catch (_) { freshUser = null; }
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
      clientKey: user.role === 'client' ? (user.clientKey || user.clientName || user.name || user.username || '') : (user.clientKey || ''),
      clientName: user.role === 'client' ? (user.clientName || user.clientKey || user.name || user.username || '') : (user.clientName || ''),
      clientLogoUrl: user.clientLogoUrl || '',
      clientPlatformImageUrl: user.clientPlatformImageUrl || '',
      clientPlatformImages: user.clientPlatformImages || {},
      allowedClients: Array.isArray(user.allowedClients) ? user.allowedClients : [],
    },
  });
};
