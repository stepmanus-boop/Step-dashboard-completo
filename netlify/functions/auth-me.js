const { jsonResponse, getSession } = require('./_auth');
const { isSupabaseConfigured, getUserById } = require('./_supabase');

exports.handler = async (event) => {
  const session = getSession(event);
  if (!session) {
    return jsonResponse(200, { ok: true, authenticated: false, publicAccess: true, githubSyncEnabled: true, supabaseEnabled: isSupabaseConfigured() });
  }

  let currentUser = null;
  if (isSupabaseConfigured() && session.sub) {
    try {
      currentUser = await getUserById(session.sub);
    } catch (_) {
      currentUser = null;
    }
  }

  const source = currentUser && currentUser.active !== false ? currentUser : {
    id: session.sub,
    name: session.name,
    username: session.username,
    role: session.role,
    sector: session.sector,
    alertSectors: Array.isArray(session.alertSectors) ? session.alertSectors : [],
    supervisedUsers: [],
  };

  return jsonResponse(200, {
    ok: true,
    authenticated: true,
    githubSyncEnabled: true,
    supabaseEnabled: isSupabaseConfigured(),
    user: {
      id: source.id || session.sub,
      name: source.name || session.name,
      username: source.username || session.username,
      role: source.role || session.role,
      sector: source.sector || session.sector,
      alertSectors: Array.isArray(source.alertSectors) ? source.alertSectors : (source.sector && source.sector !== 'all' ? [source.sector] : []),
      supervisedUsers: Array.isArray(source.supervisedUsers) ? source.supervisedUsers : [],
    },
  });
};
