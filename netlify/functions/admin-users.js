const crypto = require('crypto');
const { jsonResponse, requireAdmin, hashPassword, normalizeText, normalizeSectorList, normalizeSectorValue } = require('./_auth');
const { listUsers, insertUser, updateUser, isSupabaseConfigured } = require('./_supabase');

function normalizeProjectPmAliases(input) {
  const values = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const seen = new Set();
  const aliases = [];
  for (const value of values) {
    const item = String(value || '').trim();
    if (!item) continue;
    const key = normalizeText(item);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    aliases.push(item);
  }
  return aliases;
}

function isProjectsUser(role, sector, alertSectors = []) {
  if (role === 'admin') return false;
  return normalizeSectorValue(sector) === 'projetos' || normalizeSectorList('', alertSectors).includes('projetos');
}

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  if (event.httpMethod === 'GET') {
    const users = await listUsers();
    return jsonResponse(200, {
      ok: true,
      githubSyncEnabled: true,
      users: users.map((user) => ({
        id: user.id,
        name: user.name,
        username: user.username,
        role: user.role,
        sector: user.sector,
        alertSectors: normalizeSectorList('', user.alertSectors),
        projectPmAliases: Array.isArray(user.projectPmAliases) ? user.projectPmAliases : [],
        active: Boolean(user.active),
        createdAt: user.createdAt || null,
      })),
    });
  }

  if (event.httpMethod === 'PUT') {
    try {
      const body = JSON.parse(event.body || '{}');
      const userId = String(body.userId || '').trim();
      if (!userId) {
        return jsonResponse(400, { ok: false, error: 'Usuário não informado.' });
      }

      const users = await listUsers();
      const current = users.find((user) => user.id === userId);
      if (!current) {
        return jsonResponse(404, { ok: false, error: 'Usuário não encontrado.' });
      }

      const name = String(body.name || '').trim();
      const username = String(body.username || '').trim();
      const password = String(body.password || '');
      const role = body.role === 'admin' ? 'admin' : 'sector';
      const sector = role === 'admin' ? 'all' : normalizeSectorValue(body.sector);
      const alertSectors = role === 'admin' ? [] : normalizeSectorList('', body.alertSectors);
      const projectPmAliases = isProjectsUser(role, sector, alertSectors) ? normalizeProjectPmAliases(body.projectPmAliases) : [];

      if (!name || !username) {
        return jsonResponse(400, { ok: false, error: 'Preencha nome e usuário.' });
      }
      if (role !== 'admin' && !sector && !alertSectors.length) {
        return jsonResponse(400, { ok: false, error: 'Selecione ao menos um setor monitorado ou setor principal.' });
      }

      const exists = users.some((user) => user.id !== userId && normalizeText(user.username) === normalizeText(username));
      if (exists) {
        return jsonResponse(409, { ok: false, error: 'Já existe um usuário com esse login.' });
      }
      if (current.id === admin.session.sub && role !== 'admin') {
        return jsonResponse(400, { ok: false, error: 'O admin atual não pode remover o próprio acesso.' });
      }

      const saved = await updateUser(userId, {
        name,
        username,
        role,
        sector,
        alertSectors: role === 'admin' ? [] : alertSectors,
        projectPmAliases,
        active: body.active === false ? false : true,
        ...(password ? { passwordHash: hashPassword(password) } : {}),
      });

      return jsonResponse(200, { ok: true, user: saved });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao editar usuário.' });
    }
  }

  if (event.httpMethod === 'PATCH') {
    try {
      const body = JSON.parse(event.body || '{}');
      const userId = String(body.userId || '').trim();
      const nextRole = body.role === 'admin' ? 'admin' : 'sector';
      if (!userId) {
        return jsonResponse(400, { ok: false, error: 'Usuário não informado.' });
      }
      const users = await listUsers();
      const current = users.find((user) => user.id === userId);
      if (!current) {
        return jsonResponse(404, { ok: false, error: 'Usuário não encontrado.' });
      }
      if (current.id === admin.session.sub && nextRole !== 'admin') {
        return jsonResponse(400, { ok: false, error: 'O admin atual não pode remover o próprio acesso.' });
      }
      const saved = await updateUser(userId, {
        role: nextRole,
        sector: nextRole === 'admin' ? 'all' : (current.sector && current.sector !== 'all' ? current.sector : ''),
        alertSectors: nextRole === 'admin' ? [] : normalizeSectorList('', current.alertSectors),
        projectPmAliases: nextRole === 'admin' ? [] : (isProjectsUser(nextRole, current.sector, current.alertSectors) ? normalizeProjectPmAliases(current.projectPmAliases) : []),
      });
      return jsonResponse(200, { ok: true, user: saved });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao atualizar perfil.' });
    }
  }

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  try {
    const body = JSON.parse(event.body || '{}');
    const name = String(body.name || '').trim();
    const username = String(body.username || '').trim();
    const password = String(body.password || '');
    const role = body.role === 'admin' ? 'admin' : 'sector';
    const sector = role === 'admin' ? 'all' : normalizeSectorValue(body.sector);
    const alertSectors = role === 'admin' ? [] : normalizeSectorList('', body.alertSectors);
    const projectPmAliases = isProjectsUser(role, sector, alertSectors) ? normalizeProjectPmAliases(body.projectPmAliases) : [];

    if (!name || !username || !password) {
      return jsonResponse(400, { ok: false, error: 'Preencha nome, usuário e senha.' });
    }
    if (role !== 'admin' && !sector && !alertSectors.length) {
      return jsonResponse(400, { ok: false, error: 'Selecione ao menos um setor monitorado ou setor principal.' });
    }

    const users = await listUsers();
    const exists = users.some((user) => normalizeText(user.username) === normalizeText(username));
    if (exists) {
      return jsonResponse(409, { ok: false, error: 'Já existe um usuário com esse login.' });
    }

    const saved = await insertUser({
      id: `u_${crypto.randomBytes(6).toString('hex')}`,
      name,
      username,
      passwordHash: hashPassword(password),
      role,
      sector,
      alertSectors,
      projectPmAliases,
      active: true,
    });

    return jsonResponse(200, { ok: true, user: saved });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao criar usuário.' });
  }
};
