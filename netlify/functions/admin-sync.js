const { jsonResponse, requireAdmin } = require('./_auth');
const { isSupabaseConfigured } = require('./_supabase');

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  return jsonResponse(200, {
    ok: true,
    githubSyncEnabled: true,
    supabaseEnabled: isSupabaseConfigured(),
    message: 'Os dados operacionais agora são gravados direto no Supabase. Nenhum deploy foi disparado.',
    users: 0,
    alerts: 0,
  });
};
