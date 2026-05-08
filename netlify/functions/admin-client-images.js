const { jsonResponse, requireAdmin } = require('./_auth');
const {
  isSupabaseConfigured,
  listClientUnitImages,
  upsertClientUnitImage,
  updateUserClientImages,
} = require('./_supabase');

const SUPABASE_URL = String(process.env.SUPABASE_URL || '').replace(/\/$/, '');
const SUPABASE_SERVICE_ROLE_KEY = String(process.env.SUPABASE_SERVICE_ROLE_KEY || '');

function slug(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'sem-nome';
}

function parseDataUrl(dataUrl) {
  const raw = String(dataUrl || '');
  const match = raw.match(/^data:([^;]+);base64,(.+)$/);
  if (!match) {
    return { contentType: 'application/octet-stream', buffer: Buffer.from(raw, 'base64') };
  }
  return { contentType: match[1], buffer: Buffer.from(match[2], 'base64') };
}

function publicStorageUrl(bucket, objectPath) {
  return `${SUPABASE_URL}/storage/v1/object/public/${bucket}/${objectPath.split('/').map(encodeURIComponent).join('/')}`;
}

async function uploadToStorage(bucket, objectPath, dataUrl, contentType = '') {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    throw new Error('Supabase não configurado para upload de imagens.');
  }
  const parsed = parseDataUrl(dataUrl);
  const finalContentType = contentType || parsed.contentType || 'application/octet-stream';
  const url = `${SUPABASE_URL}/storage/v1/object/${bucket}/${objectPath.split('/').map(encodeURIComponent).join('/')}`;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      'Content-Type': finalContentType,
      'x-upsert': 'true',
    },
    body: parsed.buffer,
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Falha no upload Supabase Storage ${response.status}: ${text}`);
  }

  return publicStorageUrl(bucket, objectPath);
}

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (!isSupabaseConfigured()) {
    return jsonResponse(500, { ok: false, error: 'Supabase não configurado no Netlify.' });
  }

  if (event.httpMethod === 'GET') {
    try {
      const params = event.queryStringParameters || {};
      const clientName = String(params.clientName || params.client || '').trim();
      const images = await listClientUnitImages(clientName);
      return jsonResponse(200, { ok: true, clientName, images });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || 'Falha ao carregar imagens.' });
    }
  }

  if (event.httpMethod !== 'POST') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  try {
    const body = JSON.parse(event.body || '{}');
    const action = String(body.action || '').trim();
    const clientName = String(body.clientName || '').trim();
    const unitName = String(body.unitName || '').trim();
    const userId = String(body.userId || '').trim();
    const fileName = String(body.fileName || 'imagem.png').trim();
    const ext = (fileName.split('.').pop() || 'png').replace(/[^a-z0-9]/gi, '').toLowerCase() || 'png';

    if (action === 'uploadLogo') {
      if (!clientName || !body.dataUrl) return jsonResponse(400, { ok: false, error: 'Cliente e arquivo são obrigatórios.' });
      const objectPath = `${slug(clientName)}/logo-${Date.now()}.${ext}`;
      const imageUrl = await uploadToStorage('client-logos', objectPath, body.dataUrl, body.contentType);
      if (userId) await updateUserClientImages(userId, { clientLogoUrl: imageUrl });
      return jsonResponse(200, { ok: true, imageUrl, objectPath });
    }

    if (action === 'uploadUnitImage') {
      if (!clientName || !unitName || !body.dataUrl) return jsonResponse(400, { ok: false, error: 'Cliente, unidade e arquivo são obrigatórios.' });
      const objectPath = `${slug(clientName)}/${slug(unitName)}-${Date.now()}.${ext}`;
      const imageUrl = await uploadToStorage('client-unit-images', objectPath, body.dataUrl, body.contentType);
      if (unitName === '__default__') {
        if (userId) await updateUserClientImages(userId, { clientPlatformImageUrl: imageUrl });
      } else {
        await upsertClientUnitImage({ clientName, unitName, imageUrl });
      }
      return jsonResponse(200, { ok: true, imageUrl, objectPath, unitName });
    }

    if (action === 'saveLogoUrl') {
      if (!userId) return jsonResponse(400, { ok: false, error: 'Usuário não informado.' });
      const imageUrl = String(body.imageUrl || '').trim();
      const user = await updateUserClientImages(userId, { clientLogoUrl: imageUrl });
      return jsonResponse(200, { ok: true, imageUrl, user });
    }

    if (action === 'savePlatformFallbackUrl') {
      if (!userId) return jsonResponse(400, { ok: false, error: 'Usuário não informado.' });
      const imageUrl = String(body.imageUrl || '').trim();
      const user = await updateUserClientImages(userId, { clientPlatformImageUrl: imageUrl });
      return jsonResponse(200, { ok: true, imageUrl, user });
    }

    if (action === 'saveUnitImageUrl') {
      if (!clientName || !unitName) return jsonResponse(400, { ok: false, error: 'Cliente e unidade são obrigatórios.' });
      const imageUrl = String(body.imageUrl || '').trim();
      const row = await upsertClientUnitImage({ clientName, unitName, imageUrl });
      return jsonResponse(200, { ok: true, imageUrl, row });
    }

    return jsonResponse(400, { ok: false, error: 'Ação inválida.' });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || 'Falha ao salvar imagem.' });
  }
};
