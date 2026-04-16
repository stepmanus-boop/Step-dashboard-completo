const { jsonResponse, requireSession, requireAdmin, normalizeSectorList, normalizeText } = require('./_auth');
const { listManualAlerts, listAcknowledgements, createManualAlert, addAcknowledgement, findAcknowledgement, isSupabaseConfigured } = require('./_supabase');

function alertVisibleToUser(alert, session) {
  if (!alert || alert.active === false) return false;
  if (session.role === 'admin') return true;
  const allowedSectors = normalizeSectorList('', session.alertSectors);
  return allowedSectors.includes(normalizeText(alert.sector));
}

function getUserAlertExpiration(acknowledgements, session, alert) {
  if (!Array.isArray(acknowledgements) || !session || session.role === 'admin') return null;
  const selfAck = acknowledgements
    .filter((item) => item.userId === session.sub)
    .sort((a, b) => new Date(b.acknowledgedAt || 0).getTime() - new Date(a.acknowledgedAt || 0).getTime())[0];
  if (!selfAck?.acknowledgedAt) return null;
  const ackTime = new Date(selfAck.acknowledgedAt).getTime();
  if (!Number.isFinite(ackTime)) return null;
  const hours = Number(alert?.expiresAfterReadHours || 24);
  return new Date(ackTime + hours * 60 * 60 * 1000).toISOString();
}

exports.handler = async (event) => {
  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  if (event.httpMethod === 'GET') {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    const alerts = await listManualAlerts();
    const acks = await listAcknowledgements();
    const visible = alerts
      .filter((alert) => alertVisibleToUser(alert, auth.session))
      .map((alert) => {
        const acknowledgements = acks
          .filter((item) => item.alertId === alert.id)
          .sort((a, b) => new Date(b.acknowledgedAt || 0).getTime() - new Date(a.acknowledgedAt || 0).getTime());
        const acked = acknowledgements.some((item) => item.userId === auth.session.sub);
        const expiresAt = getUserAlertExpiration(acknowledgements, auth.session, alert);
        const expiredForUser = auth.session.role !== 'admin' && expiresAt ? new Date(expiresAt).getTime() <= Date.now() : false;
        return {
          ...alert,
          acknowledged: acked,
          ackCount: acknowledgements.length,
          lastAckAt: acknowledgements[0]?.acknowledgedAt || null,
          expiresAt,
          expiredForUser,
          acknowledgements: auth.session.role === 'admin' ? acknowledgements : undefined,
        };
      })
      .filter((alert) => auth.session.role === 'admin' || !alert.expiredForUser)
      .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());

    return jsonResponse(200, { ok: true, githubSyncEnabled: true, alerts: visible });
  }

  if (event.httpMethod === 'POST') {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const title = String(body.title || '').trim();
      const message = String(body.message || '').trim();
      const sector = normalizeText(body.sector);
      const priority = String(body.priority || 'normal').trim().toLowerCase();
      const requiresAck = body.requiresAck !== false;

      if (!title || !message || !sector) {
        return jsonResponse(400, { ok: false, error: 'Informe setor, título e mensagem.' });
      }

      const alert = await createManualAlert({
        title,
        message,
        sector,
        priority: ['low', 'normal', 'high', 'urgent'].includes(priority) ? priority : 'normal',
        requiresAck,
        createdBy: admin.session.username,
        expiresAfterReadHours: 24,
      });
      return jsonResponse(200, { ok: true, alert });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao criar alerta.' });
    }
  }

  if (event.httpMethod === 'PATCH') {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const alertId = String(body.alertId || '').trim();
      if (!alertId) {
        return jsonResponse(400, { ok: false, error: 'Alerta não informado.' });
      }

      const alerts = await listManualAlerts();
      const alert = alerts.find((item) => item.id === alertId);
      if (!alert || !alertVisibleToUser(alert, auth.session)) {
        return jsonResponse(404, { ok: false, error: 'Alerta não encontrado.' });
      }

      const existing = await findAcknowledgement(alertId, auth.session.sub);
      if (!existing) {
        await addAcknowledgement({
          alertId,
          userId: auth.session.sub,
          username: auth.session.username,
          sector: auth.session.sector,
        });
      }

      return jsonResponse(200, { ok: true });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao confirmar alerta.' });
    }
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
