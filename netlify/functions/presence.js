const { jsonResponse, requireSession, requireAdmin } = require('./_auth');
const { isSupabaseConfigured, upsertUserPresence, listUserPresence } = require('./_supabase');

function getClientIp(event) {
  const headers = event?.headers || {};
  return String(headers['x-forwarded-for'] || headers['X-Forwarded-For'] || headers['client-ip'] || '')
    .split(',')[0]
    .trim();
}

function getUserAgent(event) {
  const headers = event?.headers || {};
  return String(headers['user-agent'] || headers['User-Agent'] || '').trim();
}

exports.handler = async (event) => {
  if (!isSupabaseConfigured()) {
    return jsonResponse(200, { ok: false, error: 'Supabase não configurado.', presence: [] });
  }

  if (event.httpMethod === 'GET') {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    const presence = await listUserPresence();
    return jsonResponse(200, { ok: true, presence });
  }

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  const sessionResult = requireSession(event);
  if (!sessionResult.ok) return sessionResult.response;

  try {
    const body = JSON.parse(event.body || '{}');
    const session = sessionResult.session;
    const presence = await upsertUserPresence({
      userId: session.sub,
      username: session.username,
      name: session.name,
      role: session.role,
      sector: session.sector,
      alertSectors: Array.isArray(session.alertSectors) ? session.alertSectors : [],
      status: 'online',
      lastViewName: body.viewName || '',
      lastViewUrl: body.viewUrl || '',
      lastViewTitle: body.viewTitle || '',
      userAgent: getUserAgent(event),
      ipAddress: getClientIp(event),
    });
    return jsonResponse(200, { ok: true, presence });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao atualizar presença.' });
  }
};
