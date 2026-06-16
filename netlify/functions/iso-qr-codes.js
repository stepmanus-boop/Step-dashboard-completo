const { jsonResponse, requireSession } = require('./_auth');
const { listIsoQrCodes } = require('./_isoQrCodes');

function parseLimit(value) {
  const number = Number(value || 100);
  if (!Number.isFinite(number)) return 100;
  return Math.min(500, Math.max(1, Math.floor(number)));
}

exports.handler = async (event = {}) => {
  if (event.httpMethod && event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  // O módulo é operacional. Clientes externos não consultam a base completa de QR Codes.
  if (auth.session?.role === 'client') {
    return jsonResponse(403, { ok: false, error: 'Acesso restrito ao painel operacional.' });
  }

  const query = event.queryStringParameters || {};
  try {
    const items = await listIsoQrCodes({
      query: query.q || query.search || '',
      limit: parseLimit(query.limit),
      region: 'BR',
    });
    return jsonResponse(200, { ok: true, items, count: items.length }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    return jsonResponse(500, {
      ok: false,
      error: error.message || 'Falha ao consultar QR Codes.',
      hint: 'Verifique se o SQL v37.64 de QR Code foi executado no Supabase.',
    }, { headers: { 'cache-control': 'no-store' } });
  }
};
