const { jsonResponse, requireSession, requireAdmin } = require('./_auth');
const { listManualAlerts, listAlertResponses, createAlertResponse, updateAlertResponse, isSupabaseConfigured } = require('./_supabase');
const { handler: sectorAlertsHandler } = require('./sector-alerts');

async function loadVisibleAlerts(event) {
  const response = await sectorAlertsHandler({ ...event, httpMethod: 'GET' });
  const data = JSON.parse(response.body || '{}');
  if (!response || response.statusCode >= 400 || !data?.ok) {
    throw new Error(data?.error || 'Falha ao validar alerta.');
  }
  return Array.isArray(data.alerts) ? data.alerts : [];
}

exports.handler = async (event) => {
  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  if (event.httpMethod === 'GET') {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    try {
      const alertId = String(event.queryStringParameters?.alertId || '').trim();
      const [responses, manualAlerts] = await Promise.all([
        listAlertResponses(alertId),
        listManualAlerts(),
      ]);
      const alertMap = new Map((Array.isArray(manualAlerts) ? manualAlerts : []).map((item) => [String(item.id), item]));
      const enriched = (Array.isArray(responses) ? responses : []).map((item) => {
        const alert = alertMap.get(String(item.alertId)) || null;
        return {
          ...item,
          alertTitle: alert?.title || '',
          alertMessage: alert?.message || '',
          alertSector: alert?.sector || '',
          alertPriority: alert?.priority || 'normal',
          alertCreatedAt: alert?.createdAt || null,
        };
      });
      return jsonResponse(200, { ok: true, responses: enriched });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao carregar respostas.' });
    }
  }

  if (event.httpMethod === 'POST') {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const alertId = String(body.alertId || '').trim();
      const responseText = String(body.responseText || '').trim();
      if (!alertId || !responseText) {
        return jsonResponse(400, { ok: false, error: 'Informe o alerta e a resposta.' });
      }
      const visibleAlerts = await loadVisibleAlerts(event);
      const alert = visibleAlerts.find((item) => String(item.id) === alertId);
      if (!alert) {
        return jsonResponse(404, { ok: false, error: 'Alerta não encontrado para este usuário.' });
      }
      const saved = await createAlertResponse({
        alertId,
        userId: auth.session.sub,
        username: auth.session.name || auth.session.username,
        userEmail: auth.session.username || '',
        sector: auth.session.sector || '',
        responseText,
        status: 'enviado',
      });
      return jsonResponse(200, { ok: true, response: saved });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao enviar resposta.' });
    }
  }

  if (event.httpMethod === 'PATCH') {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const responseId = String(body.responseId || '').trim();
      if (!responseId) {
        return jsonResponse(400, { ok: false, error: 'Resposta não informada.' });
      }
      const saved = await updateAlertResponse(responseId, {
        adminReply: String(body.adminReply || '').trim(),
        status: String(body.status || 'respondido').trim() || 'respondido',
      });
      return jsonResponse(200, { ok: true, response: saved });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao responder pelo admin.' });
    }
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
