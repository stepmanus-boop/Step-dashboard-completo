const webpush = require('web-push');
const { jsonResponse, requireSession, requireAdmin, normalizeSectorList, normalizeText, normalizeSectorValue } = require('./_auth');
const { listManualAlerts, listAcknowledgements, createManualAlert, addAcknowledgement, findAcknowledgement, isSupabaseConfigured, getUserById, getUserByUsername, listUsers, listPushSubscriptions, removePushSubscription } = require('./_supabase');

function alertVisibleToUser(alert, session) {
  if (!alert || alert.active === false) return false;
  if (session.role === 'admin') return true;
  const allowedSectors = normalizeSectorList(session.sector, session.alertSectors);
  const alertSector = normalizeSectorValue(alert.sector);
  return alertSector === 'all' || allowedSectors.includes(alertSector);
}



function configureWebPush() {
  const publicKey = String(process.env.VAPID_PUBLIC_KEY || '');
  const privateKey = String(process.env.VAPID_PRIVATE_KEY || '');
  const subject = String(process.env.VAPID_SUBJECT || 'mailto:admin@example.com');
  if (!publicKey || !privateKey) return false;
  webpush.setVapidDetails(subject, publicKey, privateKey);
  return true;
}

async function notifySectorPushUsers(alert) {
  if (!configureWebPush()) return { sent: 0, skipped: true };
  const users = await listUsers();
  const subs = await listPushSubscriptions();
  if (!Array.isArray(users) || !Array.isArray(subs) || !subs.length) return { sent: 0 };

  const recipients = new Set(
    users
      .filter((user) => user.role !== 'admin')
      .filter((user) => normalizeSectorList(user.sector, user.alertSectors).includes(normalizeSectorValue(alert.sector)))
      .map((user) => String(user.id))
  );

  const targetSubs = subs.filter((item) => recipients.has(String(item.user_id || item.userId || '')));
  let sent = 0;
  await Promise.all(targetSubs.map(async (item) => {
    try {
      await webpush.sendNotification(item.subscription_json || item.subscriptionJson || item.subscription, JSON.stringify({
        title: alert.title || 'Novo alerta operacional',
        body: alert.message || 'Você recebeu um novo alerta para o seu setor.',
        tag: `manual-alert-${alert.id}`,
        url: '/',
      }));
      sent += 1;
    } catch (error) {
      if (error?.statusCode === 404 || error?.statusCode === 410) {
        try { await removePushSubscription(item.endpoint); } catch (_) {}
      }
      console.warn('Falha ao enviar web push', error?.message || error);
    }
  }));
  return { sent };
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
      const push = await notifySectorPushUsers(alert);
      return jsonResponse(200, { ok: true, alert, push });
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
      if (!alert || !alertVisibleToUser(alert, effectiveSession)) {
        return jsonResponse(404, { ok: false, error: 'Alerta não encontrado.' });
      }

      const existing = await findAcknowledgement(alertId, effectiveSession.sub);
      if (!existing) {
        await addAcknowledgement({
          alertId,
          userId: effectiveSession.sub,
          username: effectiveSession.username,
          sector: effectiveSession.sector,
        });
      }

      return jsonResponse(200, { ok: true });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao confirmar alerta.' });
    }
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
