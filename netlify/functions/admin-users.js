const crypto = require('crypto');
const { jsonResponse, requireAdmin, hashPassword, normalizeText, normalizeSectorList, normalizeSectorValue } = require('./_auth');
const { listUsers, insertUser, updateUser, isSupabaseConfigured, listUserPresence } = require('./_supabase');

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

function normalizeQualityCompetencies(input) {
  const rawValues = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const allowed = new Set(['dimensional_inicial', 'dimensional_final', 'nde', 'th', 'final_inspection_qc']);
  const aliases = {
    inicial: 'dimensional_inicial',
    dimensional_inicial: 'dimensional_inicial',
    initial_dimensional: 'dimensional_inicial',
    dimensional_final: 'dimensional_final',
    final_dimensional: 'dimensional_final',
    nde: 'nde',
    end: 'nde',
    th: 'th',
    hydro: 'th',
    hydro_test: 'th',
    final_inspection: 'final_inspection_qc',
    final_inspection_qc: 'final_inspection_qc',
  };
  const seen = new Set();
  const values = [];
  for (const value of rawValues) {
    const key = normalizeText(value).replace(/[\s-]+/g, '_').replace(/__+/g, '_');
    const normalized = aliases[key] || key;
    if (!allowed.has(normalized) || seen.has(normalized)) continue;
    seen.add(normalized);
    values.push(normalized);
  }
  return values;
}


function normalizeClientKey(input) {
  const value = String(input || '').trim();
  if (!value) return '';
  return value.toUpperCase().replace(/\s+/g, ' ');
}

function normalizeClientLogoUrl(input) {
  return String(input || '').trim();
}


function normalizeClientPlatformImageUrl(input) {
  return String(input || '').trim();
}

function normalizeClientPlatformImages(input) {
  if (!input) return {};
  let source = input;
  if (typeof source === 'string') {
    const text = source.trim();
    if (!text) return {};
    if (text.startsWith('{')) {
      try { source = JSON.parse(text); } catch (_) { source = {}; }
    } else {
      const mapped = {};
      for (const line of text.split(/\n+/)) {
        const raw = String(line || '').trim();
        if (!raw) continue;
        const separator = raw.includes('=') ? '=' : (raw.includes('|') ? '|' : ':');
        const parts = raw.split(separator);
        const key = String(parts.shift() || '').trim();
        const value = parts.join(separator).trim();
        if (key && value) mapped[key] = value;
      }
      source = mapped;
    }
  }
  if (!source || typeof source !== 'object' || Array.isArray(source)) return {};
  return Object.entries(source).reduce((acc, [key, value]) => {
    const cleanKey = String(key || '').trim();
    const cleanValue = String(value || '').trim();
    if (cleanKey && cleanValue) acc[cleanKey] = cleanValue;
    return acc;
  }, {});
}

function normalizeAllowedClients(primary, input) {
  const values = Array.isArray(input) ? input : String(input || '').split(/[\n;,|]+/);
  const seen = new Set();
  const result = [];
  for (const value of [primary, ...values]) {
    const item = normalizeClientKey(value);
    const key = normalizeText(item);
    if (!item || seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }
  return result;
}

function normalizeUserRole(input) {
  if (input === 'admin') return 'admin';
  if (input === 'client') return 'client';
  return 'sector';
}

function isQualityUser(role, sector, alertSectors = []) {
  if (role === 'admin') return false;
  return normalizeSectorValue(sector) === 'inspecao' || normalizeSectorList('', alertSectors).includes('inspecao');
}

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  if (event.httpMethod === 'GET') {
    const users = await listUsers();
    const presenceRows = await listUserPresence();
    const presenceByUserId = new Map(presenceRows.map((item) => [String(item.userId || ''), item]));
    return jsonResponse(200, {
      ok: true,
      githubSyncEnabled: true,
      presence: presenceRows,
      users: users.map((user) => {
        const presence = presenceByUserId.get(String(user.id || '')) || null;
        return {
          id: user.id,
          name: user.name,
          username: user.username,
          role: user.role,
          sector: user.sector,
          alertSectors: normalizeSectorList('', user.alertSectors),
          projectPmAliases: Array.isArray(user.projectPmAliases) ? user.projectPmAliases : [],
          qualityCompetencies: Array.isArray(user.qualityCompetencies) ? user.qualityCompetencies : [],
          clientKey: user.clientKey || '',
          clientName: user.clientName || '',
          clientLogoUrl: user.clientLogoUrl || '',
          clientPlatformImageUrl: user.clientPlatformImageUrl || '',
          clientPlatformImages: user.clientPlatformImages || {},
          allowedClients: Array.isArray(user.allowedClients) ? user.allowedClients : [],
          active: Boolean(user.active),
          createdAt: user.createdAt || null,
          presence,
          online: Boolean(presence?.online),
          lastSeenAt: presence?.lastSeenAt || null,
          lastLoginAt: presence?.lastLoginAt || null,
          lastLogoutAt: presence?.lastLogoutAt || null,
          lastViewAt: presence?.lastViewAt || null,
          lastViewName: presence?.lastViewName || '',
          lastViewUrl: presence?.lastViewUrl || '',
          lastViewTitle: presence?.lastViewTitle || '',
        };
      }),
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
      const role = normalizeUserRole(body.role);
      const sector = role === 'admin' ? 'all' : (role === 'client' ? 'all' : normalizeSectorValue(body.sector));
      const alertSectors = role !== 'sector' ? [] : normalizeSectorList('', body.alertSectors);
      const projectPmAliases = isProjectsUser(role, sector, alertSectors) ? normalizeProjectPmAliases(body.projectPmAliases) : [];
      const qualityCompetencies = isQualityUser(role, sector, alertSectors) ? normalizeQualityCompetencies(body.qualityCompetencies) : [];
      const clientKey = role === 'client' ? normalizeClientKey(body.clientKey || body.clientName || username) : '';
      const clientName = role === 'client' ? String(body.clientName || clientKey).trim() : '';
      const clientLogoUrl = role === 'client' ? normalizeClientLogoUrl(body.clientLogoUrl) : '';
      const clientPlatformImageUrl = role === 'client' ? normalizeClientPlatformImageUrl(body.clientPlatformImageUrl) : '';
      const clientPlatformImages = role === 'client' ? normalizeClientPlatformImages(body.clientPlatformImages) : {};
      const allowedClients = role === 'client' ? normalizeAllowedClients(clientKey, body.allowedClients) : [];

      if (!name || !username) {
        return jsonResponse(400, { ok: false, error: 'Preencha nome e usuário.' });
      }
      if (role === 'sector' && !sector && !alertSectors.length) {
        return jsonResponse(400, { ok: false, error: 'Selecione ao menos um setor monitorado ou setor principal.' });
      }
      if (role === 'client' && !clientKey) {
        return jsonResponse(400, { ok: false, error: 'Informe o cliente vinculado ao Portal do Cliente.' });
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
        qualityCompetencies,
        clientKey,
        clientName,
        clientLogoUrl,
        clientPlatformImageUrl,
        clientPlatformImages,
        allowedClients,
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
      const nextRole = normalizeUserRole(body.role);
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
        sector: nextRole === 'admin' ? 'all' : (nextRole === 'client' ? 'all' : (current.sector && current.sector !== 'all' ? current.sector : '')),
        alertSectors: nextRole !== 'sector' ? [] : normalizeSectorList('', current.alertSectors),
        projectPmAliases: nextRole !== 'sector' ? [] : (isProjectsUser(nextRole, current.sector, current.alertSectors) ? normalizeProjectPmAliases(current.projectPmAliases) : []),
        qualityCompetencies: nextRole !== 'sector' ? [] : (isQualityUser(nextRole, current.sector, current.alertSectors) ? normalizeQualityCompetencies(current.qualityCompetencies) : []),
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
    const role = normalizeUserRole(body.role);
    const sector = role === 'admin' ? 'all' : (role === 'client' ? 'all' : normalizeSectorValue(body.sector));
    const alertSectors = role !== 'sector' ? [] : normalizeSectorList('', body.alertSectors);
    const projectPmAliases = isProjectsUser(role, sector, alertSectors) ? normalizeProjectPmAliases(body.projectPmAliases) : [];
    const qualityCompetencies = isQualityUser(role, sector, alertSectors) ? normalizeQualityCompetencies(body.qualityCompetencies) : [];
    const clientKey = role === 'client' ? normalizeClientKey(body.clientKey || body.clientName || username) : '';
    const clientName = role === 'client' ? String(body.clientName || clientKey).trim() : '';
    const clientLogoUrl = role === 'client' ? normalizeClientLogoUrl(body.clientLogoUrl) : '';
    const clientPlatformImageUrl = role === 'client' ? normalizeClientPlatformImageUrl(body.clientPlatformImageUrl) : '';
    const clientPlatformImages = role === 'client' ? normalizeClientPlatformImages(body.clientPlatformImages) : {};
    const allowedClients = role === 'client' ? normalizeAllowedClients(clientKey, body.allowedClients) : [];

    if (!name || !username || !password) {
      return jsonResponse(400, { ok: false, error: 'Preencha nome, usuário e senha.' });
    }
    if (role === 'sector' && !sector && !alertSectors.length) {
      return jsonResponse(400, { ok: false, error: 'Selecione ao menos um setor monitorado ou setor principal.' });
    }
    if (role === 'client' && !clientKey) {
      return jsonResponse(400, { ok: false, error: 'Informe o cliente vinculado ao Portal do Cliente.' });
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
      qualityCompetencies,
      clientKey,
      clientName,
      clientLogoUrl,
      clientPlatformImageUrl,
      clientPlatformImages,
      allowedClients,
      active: true,
    });

    return jsonResponse(200, { ok: true, user: saved });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao criar usuário.' });
  }
};
