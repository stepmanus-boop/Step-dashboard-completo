const { jsonResponse, clearSessionCookie, getSession } = require("./_auth");
const { isSupabaseConfigured, markUserPresenceOffline } = require('./_supabase');

exports.handler = async (event) => {
  const session = getSession(event);
  if (session?.sub && isSupabaseConfigured()) {
    await markUserPresenceOffline(session.sub, {
      lastViewName: 'Logout do sistema',
      lastViewUrl: '/',
      lastViewTitle: 'STEP - Painel Operacional',
    }).catch(() => null);
  }
  return jsonResponse(200, { ok: true }, { headers: { "set-cookie": clearSessionCookie() } });
};
