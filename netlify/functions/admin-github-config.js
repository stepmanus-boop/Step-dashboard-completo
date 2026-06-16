const { jsonResponse, requireAdmin } = require('./_auth');
const { isSupabaseConfigured } = require('./_supabase');

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (event.httpMethod === 'GET') {
    return jsonResponse(200, {
      ok: true,
      configured: true,
      source: 'supabase',
      repo: 'Supabase ativo',
      branch: 'runtime',
      tokenMasked: 'supabase',
      message: 'O armazenamento operacional está usando Supabase.',
    });
  }

  if (event.httpMethod === 'POST') {
    return jsonResponse(200, {
      ok: true,
      configured: true,
      message: 'O armazenamento operacional está usando Supabase. Nenhuma configuração de GitHub é necessária para alertas e usuários.',
      supabaseEnabled: isSupabaseConfigured(),
    });
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
