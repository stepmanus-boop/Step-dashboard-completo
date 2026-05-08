const { jsonResponse, requireSession } = require('./_auth');
const { isSupabaseConfigured, listClientUnitImages } = require('./_supabase');

exports.handler = async (event) => {
  const session = requireSession(event);
  if (!session.ok) return session.response;

  if (event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  try {
    const params = event.queryStringParameters || {};
    const clientName = String(params.clientName || params.client || '').trim();
    if (!clientName) return jsonResponse(200, { ok: true, clientName, images: [] });
    const images = await listClientUnitImages(clientName);
    return jsonResponse(200, { ok: true, clientName, images });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao carregar imagens.' });
  }
};
