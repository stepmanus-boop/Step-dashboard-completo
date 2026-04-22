const webpush = require('web-push');
const { jsonResponse, requireSession, normalizeSectorList, normalizeText, normalizeSectorValue } = require('./_auth');
const { listManualAlerts, listAcknowledgements, createManualAlert, deleteManualAlert, deleteManyManualAlerts, addAcknowledgement, findAcknowledgement, isSupabaseConfigured, getUserById, getUserByUsername, listUsers, listPushSubscriptions, removePushSubscription } = require('./_supabase');
const { buildPayload } = require('./projects');

function alertVisibleToUser(alert, session) {
  if (!alert || alert.active === false || alert.deleted === true || String(alert.status || '').toLowerCase() === 'deleted') return false;
  if (session.role === 'admin') return true;
  const allowedSectors = normalizeSectorList(session.sector, session.alertSectors);
  const alertSector = normalizeSectorValue(alert.sector);
  const isCreator = String(alert.createdBy || '').trim().toLowerCase() === String(session.username || '').trim().toLowerCase();
  return isCreator || alertSector === 'all' || allowedSectors.includes(alertSector);
}




function tokenizeNormalizedNames(values = []) {
  const set = new Set();
  const source = Array.isArray(values) ? values : [values];
  for (const value of source) {
    const normalized = normalizeText(value).trim();
    if (!normalized) continue;
    set.add(normalized);
    for (const part of normalized.split(/[^a-z0-9]+/)) {
      if (part) set.add(part);
    }
  }
  return set;
}

function projectBelongsToUser(project, user) {
  if (!project || !user) return false;
  const pmValue = String(project.pm || '').trim();
  if (!pmValue) return false;
  const candidates = tokenizeNormalizedNames([user.name, user.username, String(user.username || '').split('@')[0]]);
  if (!candidates.size) return false;
  const normalizedPm = normalizeText(pmValue).trim();
  const pmTokens = tokenizeNormalizedNames(pmValue.split(/[;,|/]+/));
  for (const candidate of candidates) {
    if (normalizedPm === candidate || normalizedPm.includes(candidate)) return true;
    if (pmTokens.has(candidate)) return true;
  }
  return false;
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

    const projectSignals = alerts
      .filter((alert) => normalizeSectorValue(alert.sector) === 'pcp')
      .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());

    return jsonResponse(200, { ok: true, githubSyncEnabled: true, alerts: visible, projectSignals, userSector: effectiveSession.sector, userAlertSectors: effectiveSession.alertSectors || [] });
  }

  if (event.httpMethod === 'POST') {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    try {
      const body = JSON.parse(event.body || '{}');
      const title = String(body.title || '').trim();
      const message = String(body.message || '').trim();
      const sector = normalizeSectorValue(body.sector);
      const priority = String(body.priority || 'normal').trim().toLowerCase();
      const requiresAck = body.requiresAck !== false;
      const projectRowId = Number(body.projectRowId || 0);
      const effectiveSession = await getEffectiveSession(auth.session);
      const userSector = normalizeSectorValue(effectiveSession.sector);
      const allowedProjectSectors = normalizeSectorList(effectiveSession.sector, effectiveSession.alertSectors);
      const canCreateGeneralAlert = effectiveSession.role === 'admin';
      const canCreateProjectSignal = userSector === 'projetos' || allowedProjectSectors.includes('projetos');
      const isProjectSignalToPcp = sector === 'pcp';

      if (!title || !message || !sector) {
        return jsonResponse(400, { ok: false, error: 'Informe setor, título e mensagem.' });
      }

      if (!canCreateGeneralAlert && !(canCreateProjectSignal && isProjectSignalToPcp)) {
        return jsonResponse(403, { ok: false, error: 'Apenas administradores podem criar alertas gerais. Usuários de Projetos podem enviar sinalizações ao PCP.' });
      }

      if (!canCreateGeneralAlert && canCreateProjectSignal) {
        if (!projectRowId) {
          return jsonResponse(400, { ok: false, error: 'BSP não informada para validação.' });
        }
        const payload = await buildPayload();
        const projects = Array.isArray(payload?.projects) ? payload.projects : [];
        const project = projects.find((item) => Number(item?.rowId || 0) === projectRowId);
        if (!project) {
          return jsonResponse(404, { ok: false, error: 'BSP não encontrada para validação.' });
        }
        if (!projectBelongsToUser(project, effectiveSession)) {
          return jsonResponse(403, { ok: false, error: 'Você só pode enviar notificações das BSPs que estejam no seu nome.' });
        }
      }

      const alert = await createManualAlert({
        title,
        message,
        sector,
        priority: ['low', 'normal', 'high', 'urgent'].includes(priority) ? priority : 'normal',
        requiresAck,
        createdBy: effectiveSession.username,
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


  if (event.httpMethod === 'DELETE') {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    try {
      const effectiveSession = await getEffectiveSession(auth.session);
      if (effectiveSession.role !== 'admin') {
        return jsonResponse(403, { ok: false, error: 'Apenas administradores podem apagar alertas.' });
      }
      const body = JSON.parse(event.body || '{}');
      const ids = Array.isArray(body.ids) ? body.ids : [];
      const id = String(body.id || '').trim();
      if (ids.length) {
        const deleted = await deleteManyManualAlerts(ids, effectiveSession.username);
        return jsonResponse(200, { ok: true, count: deleted.length });
      }
      if (!id) {
        return jsonResponse(400, { ok: false, error: 'Informe o alerta para apagar.' });
      }
      const deleted = await deleteManualAlert(id, effectiveSession.username);
      return jsonResponse(200, { ok: true, alert: deleted });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao apagar alerta.' });
    }
  }

  return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
};
