const crypto = require('crypto');
const { jsonResponse, requireSession } = require('./_auth');
const {
  isSupabaseConfigured,
  getUserById,
  listClientApiKeysForUser,
  createClientApiKey,
  revokeClientApiKey,
} = require('./_supabase');

function sha256(value) {
  return crypto.createHash('sha256').update(String(value || '')).digest('hex');
}

function makeApiToken() {
  return `step_${crypto.randomBytes(32).toString('base64url')}`;
}

function safeKey(key, includeRaw = false, rawToken = '') {
  return {
    id: key.id,
    name: key.name,
    clientKey: key.clientKey,
    clientName: key.clientName,
    tokenPreview: `${key.tokenPrefix || 'step_'}••••${key.tokenLast4 || ''}`,
    scopes: key.scopes || ['read:projects'],
    active: key.active !== false,
    expiresAt: key.expiresAt || null,
    lastUsedAt: key.lastUsedAt || null,
    createdAt: key.createdAt || null,
    revokedAt: key.revokedAt || null,
    ...(includeRaw ? { token: rawToken } : {}),
  };
}

async function hydrateSessionUser(session = {}) {
  if (!session?.sub) return session;
  if (!isSupabaseConfigured()) return session;
  try {
    const user = await getUserById(session.sub);
    return user ? { ...session, ...user, sub: session.sub } : session;
  } catch (_) {
    return session;
  }
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') {
    return jsonResponse(200, { ok: true });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, {
      ok: false,
      error: 'Supabase não configurado. Configure SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY no Netlify.',
    });
  }

  const user = await hydrateSessionUser(auth.session);
  if (!user || user.role !== 'client') {
    return jsonResponse(403, { ok: false, error: 'A geração de API está disponível apenas para usuários do Portal do Cliente.' });
  }

  try {
    if (event.httpMethod === 'GET') {
      const keys = await listClientApiKeysForUser(user.sub || user.id);
      return jsonResponse(200, { ok: true, keys: keys.map((key) => safeKey(key)) });
    }

    if (event.httpMethod === 'POST') {
      const body = JSON.parse(event.body || '{}');
      const token = makeApiToken();
      const tokenHash = sha256(token);
      const tokenPrefix = token.slice(0, 9);
      const tokenLast4 = token.slice(-4);
      const key = await createClientApiKey({
        userId: user.sub || user.id,
        username: user.username || '',
        clientKey: user.clientKey || '',
        clientName: user.clientName || user.clientKey || user.name || 'Cliente',
        allowedClients: Array.isArray(user.allowedClients) ? user.allowedClients : [],
        tokenHash,
        tokenPrefix,
        tokenLast4,
        name: String(body.name || 'API do Portal do Cliente').trim().slice(0, 120),
        scopes: ['read:projects'],
        createdBy: user.username || '',
        createdByName: user.name || user.clientName || '',
      });
      return jsonResponse(201, { ok: true, key: safeKey(key, true, token) });
    }

    if (event.httpMethod === 'DELETE') {
      const body = JSON.parse(event.body || '{}');
      const id = body.id || event.queryStringParameters?.id || '';
      if (!id) return jsonResponse(400, { ok: false, error: 'Informe a chave que será revogada.' });
      const revoked = await revokeClientApiKey(id, user.sub || user.id);
      return jsonResponse(200, { ok: true, key: safeKey(revoked || { id, active: false }) });
    }

    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  } catch (error) {
    const message = String(error?.message || 'Falha ao gerenciar API do cliente.');
    const missingTable = message.includes('client_api_keys');
    return jsonResponse(missingTable ? 500 : 400, {
      ok: false,
      error: missingTable
        ? 'Tabela client_api_keys não encontrada. Execute o SQL de criação no Supabase antes de gerar chaves.'
        : message,
    });
  }
};
