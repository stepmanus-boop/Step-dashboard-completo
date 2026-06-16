const { jsonResponse, requireSession } = require('./_auth');
const { isSupabaseConfigured, upsertPushSubscription, removePushSubscription, listPushSubscriptions } = require('./_supabase');

exports.handler = async (event) => {
  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  const session = auth.session;

  if (event.httpMethod === 'GET') {
    const rows = await listPushSubscriptions(session.sub);
    return jsonResponse(200, { ok: true, subscribed: rows.length > 0, count: rows.length, vapidPublicKey: process.env.VAPID_PUBLIC_KEY || '' });
  }

  if (event.httpMethod === 'POST') {
    try {
      const body = JSON.parse(event.body || '{}');
      const subscription = body.subscription;
      if (!subscription?.endpoint) {
        return jsonResponse(400, { ok: false, error: 'Subscription inválida.' });
      }
      await upsertPushSubscription({
        userId: session.sub,
        username: session.username,
        sector: session.sector,
        endpoint: subscription.endpoint,
        subscription,
        active: true,
      });
      return jsonResponse(200, { ok: true });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao salvar push.' });
    }
  }

  if (event.httpMethod === 'DELETE') {
    try {
      const body = JSON.parse(event.body || '{}');
      const endpoint = String(body.endpoint || '').trim();
      if (!endpoint) return jsonResponse(400, { ok: false, error: 'Endpoint não informado.' });
      await removePushSubscription(endpoint);
      return jsonResponse(200, { ok: true });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao remover push.' });
    }
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
