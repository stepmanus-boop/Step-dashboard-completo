const { jsonResponse, getSession } = require('./_auth');
const { isSupabaseConfigured } = require('./_supabase');

exports.handler = async (event) => {
  const session = getSession(event);
  if (!session) {
    return jsonResponse(200, { ok: true, authenticated: false, publicAccess: true, githubSyncEnabled: true, supabaseEnabled: isSupabaseConfigured() });
  }

  return jsonResponse(200, {
    ok: true,
    authenticated: true,
    githubSyncEnabled: true,
    supabaseEnabled: isSupabaseConfigured(),
    user: {
      id: session.sub,
      name: session.name,
      username: session.username,
      role: session.role,
      sector: session.sector,
      alertSectors: Array.isArray(session.alertSectors) ? session.alertSectors : (session.sector && session.sector !== 'all' ? [session.sector] : []),
      supervisedUsers: Array.isArray(session.supervisedUsers) ? session.supervisedUsers : [],
    },
  });
};
