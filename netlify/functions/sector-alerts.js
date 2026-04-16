const { jsonResponse, requireSession, requireAdmin, normalizeSectorList, normalizeText, normalizeSectorValue } = require('./_auth');
const { listManualAlerts, listAcknowledgements, createManualAlert, addAcknowledgement, findAcknowledgement, isSupabaseConfigured, getUserById, getUserByUsername } = require('./_supabase');

function alertVisibleToUser(alert, session) {
  if (!alert || alert.active === false) return false;
  if (session.role === 'admin') return true;
  const allowedSectors = normalizeSectorList(session.sector, session.alertSectors);
  const alertSector = normalizeSectorValue(alert.sector);
  return alertSector === 'all' || allowedSectors.includes(alertSector);
}


async function getEffectiveSession(session) {
  if (!session || session.role === 'admin') return session;
  try {
    const freshUser = (session.sub && await getUserById(session.sub)) || (session.username && await getUserByUsername(session.username)) || null;
    if (!freshUser) return session;
    return {
      ...session,
      role: freshUser.role || session.role,
      sector: normalizeSectorValue(freshUser.sector || session.sector),
      alertSectors: normalizeSectorList(freshUser.sector || session.sector, freshUser.alertSectors || session.alertSectors),
      name: freshUser.name || session.name,
      username: freshUser.username || session.username,
      sub: freshUser.id || session.sub,
    };
  } catch (error) {
    console.warn('Falha ao recarregar dados atualizados do usuário para alertas. Usando sessão atual.', error);
    return session;
  }
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
    const effectiveSession = await getEffectiveSession(auth.session);
    const alerts = await listManualAlerts();
    let acks = [];
    try {
      acks = await listAcknowledgements();
    } catch (error) {
      console.warn('Falha ao carregar confirmações de leitura dos alertas. Seguindo sem acknowledgements.', error);
      acks = [];
    }
    const visible = alerts
      .filter((alert) => alertVisibleToUser(alert, effectiveSession))
      .map((alert) => {
        const acknowledgements = acks
          .filter((item) => item.alertId === alert.id)
          .sort((a, b) => new Date(b.acknowledgedAt || 0).getTime() - new Date(a.acknowledgedAt || 0).getTime());
        const acked = acknowledgements.some((item) => item.userId === effectiveSession.sub);
        const expiresAt = getUserAlertExpiration(acknowledgements, effectiveSession, alert);
        const expiredForUser = effectiveSession.role !== 'admin' && expiresAt ? new Date(expiresAt).getTime() <= Date.now() : false;
        return {
          ...alert,
          acknowledged: acked,
          ackCount: acknowledgements.length,
          lastAckAt: acknowledgements[0]?.acknowledgedAt || null,
          expiresAt,
          expiredForUser,
          acknowledgements: effectiveSession.role === 'admin' ? acknowledgements : undefined,
        };
      })
      .filter((alert) => effectiveSession.role === 'admin' || !alert.expiredForUser)
      .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());

    return jsonResponse(200, { ok: true, githubSyncEnabled: true, alerts: visible, userSector: effectiveSession.sector, userAlertSectors: effectiveSession.alertSectors || [] });
  }

  if (event.httpMethod === 'POST') {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const title = String(body.title || '').trim();
      const message = String(body.message || '').trim();
      const sector = normalizeSectorValue(body.sector);
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

      const effectiveSession = await getEffectiveSession(auth.session);
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
