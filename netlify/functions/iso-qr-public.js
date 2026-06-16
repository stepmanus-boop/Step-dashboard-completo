const { jsonResponse } = require('./_auth');
const { getIsoQrByToken } = require('./_isoQrCodes');

exports.handler = async (event = {}) => {
  if (event.httpMethod && event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }
  const token = String(event.queryStringParameters?.token || '').trim();
  if (!token) return jsonResponse(400, { ok: false, error: 'Token do QR Code não informado.' });

  try {
    const item = await getIsoQrByToken(token);
    if (!item) return jsonResponse(404, { ok: false, error: 'QR Code não encontrado.' });
    return jsonResponse(200, { ok: true, item }, { headers: { 'cache-control': 'no-store' } });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao abrir rastreamento do QR Code.' }, { headers: { 'cache-control': 'no-store' } });
  }
};
