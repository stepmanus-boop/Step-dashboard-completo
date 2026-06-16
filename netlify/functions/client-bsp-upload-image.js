// Netlify function: upload an image from the portal directly to a Smartsheet row.
// The browser sends JSON with rowId, optional sheetId, fileName, mimeType and base64.
// The function keeps the Smartsheet token server-side and performs a simple upload
// to /sheets/{sheetId}/rows/{rowId}/attachments.
// Access rule: only admin or the PM linked to the BSP may upload images.

const { requireSession, normalizeText, normalizeSectorValue } = require('./_auth');
const { isSupabaseConfigured, getUserById } = require('./_supabase');
const { loadProjectPayload } = require('./_projectLookup');

const DEFAULT_TRACKING_SHEET_ID = '3612820139992964';

function getSmartsheetToken() {
  return process.env.SMARTSHEET_API_KEY
    || process.env.SMARTSHEET_TOKEN
    || process.env.SMARTSHEET_ACCESS_TOKEN
    || process.env.SMARTSHEET_API_TOKEN
    || process.env.SMARTSHEET_BEARER_TOKEN
    || process.env.SMARTSHEET_PAT
    || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN
    || '';
}
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024; // keep payload safe for Netlify/browser flow

function json(statusCode, payload) {
  return {
    statusCode,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify(payload),
  };
}

function sanitizeFileName(name) {
  const fallback = `bsp-image-${Date.now()}.jpg`;
  const raw = String(name || fallback).trim() || fallback;
  return raw
    .replace(/[\\/\r\n\t]/g, '-')
    .replace(/[<>:"|?*]/g, '-')
    .slice(0, 140);
}

function normalizeBase64(value) {
  const raw = String(value || '');
  const comma = raw.indexOf(',');
  return comma >= 0 ? raw.slice(comma + 1) : raw;
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

function userHasProjectsScope(user = {}) {
  if (!user || user.role === 'admin') return false;
  const sectors = Array.isArray(user.alertSectors) ? user.alertSectors : [];
  return normalizeSectorValue(user.sector) === 'projetos' || sectors.map(normalizeSectorValue).includes('projetos');
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

function projectBelongsToUser(project = {}, user = {}) {
  if (!project || !userHasProjectsScope(user)) return false;
  const pmValue = String(project.pm || '').trim();
  if (!pmValue) return false;
  const candidates = tokenizeNormalizedNames([
    user.name,
    user.username,
    String(user.username || '').split('@')[0],
    ...(Array.isArray(user.projectPmAliases) ? user.projectPmAliases : []),
  ]);
  if (!candidates.size) return false;
  const normalizedPm = normalizeText(pmValue).trim();
  const pmTokens = tokenizeNormalizedNames(pmValue.split(/[;,|/]+/));
  for (const candidate of candidates) {
    if (normalizedPm === candidate || normalizedPm.includes(candidate)) return true;
    if (pmTokens.has(candidate)) return true;
  }
  return false;
}

function canUploadImageToProject(project = {}, user = {}) {
  if (!user) return false;
  if (user.role === 'admin') return true;
  return projectBelongsToUser(project, user);
}

async function findProject(projectRowId) {
  const payload = await loadProjectPayload({ allowFallback: true });
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  const target = String(projectRowId || '').trim();
  return projects.find((project) => String(project.rowId || '') === target || String(project.rowNumber || '') === target) || null;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return json(405, { ok: false, error: 'Método não permitido.' });
  }

  const auth = requireSession(event);
  if (!auth.ok) return auth.response;
  const user = await hydrateSessionUser(auth.session);

  const API_KEY = getSmartsheetToken();
  if (!API_KEY) {
    return json(500, { ok: false, error: 'Token Smartsheet não configurado. Configure SMARTSHEET_API_KEY ou SMARTSHEET_TOKEN no Netlify.' });
  }

  let payload;
  try {
    payload = JSON.parse(event.body || '{}');
  } catch (err) {
    return json(400, { ok: false, error: 'JSON inválido.' });
  }

  const rowId = String(payload.rowId || '').trim();
  const sheetId = String(payload.sheetId || process.env.SMARTSHEET_TRACKING_SHEET_ID || process.env.SMARTSHEET_SHEET_ID || DEFAULT_TRACKING_SHEET_ID).trim();
  const mimeType = String(payload.mimeType || 'image/jpeg').trim().toLowerCase();
  const fileName = sanitizeFileName(payload.fileName);
  const base64 = normalizeBase64(payload.base64);

  if (!rowId || !/^\d+$/.test(rowId)) {
    return json(400, { ok: false, error: 'rowId obrigatório ou inválido.' });
  }

  let project = null;
  try {
    project = await findProject(rowId);
  } catch (err) {
    return json(503, { ok: false, error: 'Não foi possível validar a permissão da BSP no Tracking.' });
  }
  if (!project) {
    return json(404, { ok: false, error: 'BSP não encontrada para validar a permissão de importação.' });
  }
  if (!canUploadImageToProject(project, user)) {
    return json(403, { ok: false, error: 'Apenas administrador ou PM responsável pela BSP pode importar imagem.' });
  }

  if (!sheetId || !/^\d+$/.test(sheetId)) {
    return json(400, { ok: false, error: 'sheetId obrigatório ou inválido.' });
  }
  if (!base64) {
    return json(400, { ok: false, error: 'Imagem não enviada.' });
  }
  if (!mimeType.startsWith('image/')) {
    return json(400, { ok: false, error: 'Apenas arquivos de imagem são permitidos.' });
  }

  let buffer;
  try {
    buffer = Buffer.from(base64, 'base64');
  } catch (err) {
    return json(400, { ok: false, error: 'Base64 da imagem inválido.' });
  }

  if (!buffer.length) {
    return json(400, { ok: false, error: 'Imagem vazia.' });
  }
  if (buffer.length > MAX_UPLOAD_BYTES) {
    return json(413, {
      ok: false,
      error: `Imagem muito grande após compressão (${Math.round(buffer.length / 1024 / 1024)} MB). Reduza a imagem e tente novamente.`,
    });
  }

  const encodedName = encodeURIComponent(fileName);
  try {
    const smartsheetRes = await fetch(`https://api.smartsheet.com/2.0/sheets/${sheetId}/rows/${rowId}/attachments`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${API_KEY}`,
        'Content-Disposition': `attachment; filename="${encodedName}"`,
        'Content-Type': mimeType,
        'Content-Length': String(buffer.length),
      },
      body: buffer,
    });

    const text = await smartsheetRes.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch (_) { data = text; }

    if (!smartsheetRes.ok) {
      return json(smartsheetRes.status, {
        ok: false,
        error: 'Falha ao importar imagem no Smartsheet.',
        smartsheet: data,
      });
    }

    return json(200, {
      ok: true,
      message: 'Imagem importada com sucesso.',
      attachment: data?.result || data,
    });
  } catch (err) {
    return json(500, { ok: false, error: err.message || 'Erro interno ao importar imagem.' });
  }
};
