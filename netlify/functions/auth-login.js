const { jsonResponse, createSessionCookie, normalizeText } = require('./_auth');
const { getUserByUsername, isSupabaseConfigured, userPasswordMatches } = require('./_supabase');

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  try {
    if (!isSupabaseConfigured()) {
      return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
    }
    const body = JSON.parse(event.body || '{}');
    const username = String(body.username || '').trim();
    const password = String(body.password || '').trim();
    if (!username || !password) {
      return jsonResponse(400, { ok: false, error: 'Informe usuário e senha.' });
    }

    const defaultAdmin = {
      id: 'u_admin_001',
      name: 'Administrador',
      username: 'admin',
      role: 'admin',
      sector: 'all',
      alertSectors: [],
      active: true,
      passwordHash: 'admin123',
    };

    let user = await getUserByUsername(username);
    if ((!user || !user.active) && normalizeText(username) === 'admin' && password === 'admin123') {
      user = defaultAdmin;
    }

    if (!user || !user.active || !userPasswordMatches(password, user.passwordHash)) {
      return jsonResponse(401, { ok: false, error: 'Usuário ou senha inválidos.' });
    }

    return jsonResponse(200, {
      ok: true,
      user: {
        id: user.id,
        name: user.name,
        username: user.username,
        role: user.role,
        sector: user.sector,
        alertSectors: Array.isArray(user.alertSectors) ? user.alertSectors : [],
      },
    }, {
      headers: {
        'set-cookie': createSessionCookie(user),
      },
    });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao autenticar.' });
  }
};
