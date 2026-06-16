// Netlify function: proxy das imagens do Smartsheet para Projetos em Desenvolvimento – Yinson.
// Aceita attachment id direto OU rowId. Quando recebe rowId, localiza os anexos de imagem da linha
// e redireciona para a URL temporária do Smartsheet. O token nunca vai para o navegador.

const API_BASE = process.env.SMARTSHEET_API_BASE || 'https://api.smartsheet.com/2.0';
const REQUEST_TIMEOUT_MS = Number(process.env.SMARTSHEET_REQUEST_TIMEOUT_MS || process.env.SMARTSHEET_FETCH_TIMEOUT_MS || 20000);

function getToken() {
  return process.env.SMARTSHEET_API_KEY
    || process.env.SMARTSHEET_TOKEN
    || process.env.SMARTSHEET_ACCESS_TOKEN
    || process.env.SMARTSHEET_API_TOKEN
    || process.env.SMARTSHEET_BEARER_TOKEN
    || process.env.SMARTSHEET_PAT
    || process.env.SMARTSHEET_PERSONAL_ACCESS_TOKEN
    || '';
}

function json(statusCode, payload) {
  return {
    statusCode,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store, no-cache, must-revalidate',
    },
    body: JSON.stringify(payload),
  };
}

function unique(values) {
  return Array.from(new Set((values || []).flatMap((item) => String(item || '').split(',')).map((v) => v.trim()).filter(Boolean)));
}

function getCandidateSheetIds(event) {
  const qs = event.queryStringParameters || {};
  return unique([
    qs.sheetId,
    qs.sheetIds,
    qs.sheetId || '7523600180400004',
    process.env.SMARTSHEET_YINSON_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_CLIENT_UNDER_DEV_SHEET_ID,
    process.env.SMARTSHEET_YINSON_SHEET_ID,
    process.env.SMARTSHEET_SHEET_ID_YINSON,
    process.env.SMARTSHEET_SHEET_IDS_LIST,
    process.env.SMARTSHEET_SHEET_ID_BR,
    process.env.SMARTSHEET_SHEET_ID,
    '7523600180400004',
    '3612820139992964',
  ]);
}

async function fetchJson(url, options = {}) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    const text = await response.text();
    let payload = null;
    try { payload = text ? JSON.parse(text) : null; } catch (_) { payload = text; }
    if (!response.ok) {
      const error = new Error(`Smartsheet HTTP ${response.status}`);
      error.status = response.status;
      error.payload = payload;
      throw error;
    }
    return payload;
  } finally {
    clearTimeout(timeout);
  }
}

function isImageAttachment(att) {
  const mime = String(att?.mimeType || '').toLowerCase();
  const name = String(att?.name || '').toLowerCase();
  return mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp|heic|heif)$/i.test(name);
}

function sanitizeDownloadFilename(value, fallback = 'imagem.jpg') {
  const cleaned = String(value || fallback)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9._-]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 120);
  return cleaned || fallback;
}

function contentTypeFromName(name, fallback = 'application/octet-stream') {
  const lower = String(name || '').toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.gif')) return 'image/gif';
  if (lower.endsWith('.bmp')) return 'image/bmp';
  if (lower.endsWith('.heic')) return 'image/heic';
  if (lower.endsWith('.heif')) return 'image/heif';
  return fallback;
}

async function getAttachmentDetail(headers, attachmentId, sheetIds) {
  for (const sheetId of sheetIds) {
    try {
      const detail = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/attachments/${encodeURIComponent(attachmentId)}`, { headers });
      if (detail?.url) return detail;
    } catch (_) {}
  }
  return fetchJson(`${API_BASE}/attachments/${encodeURIComponent(attachmentId)}`, { headers });
}

async function getRowImageAttachment(headers, rowId, sheetIds, imageIndex) {
  for (const sheetId of sheetIds) {
    try {
      const payload = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/rows/${encodeURIComponent(rowId)}/attachments`, { headers });
      const images = (payload?.data || []).filter(isImageAttachment);
      if (images.length) {
        const safeIndex = Math.max(0, Math.min(Number(imageIndex || 0), images.length - 1));
        return { attachment: images[safeIndex], sheetId };
      }
    } catch (_) {}
  }
  return null;
}

async function getRowImageAttachments(headers, rowId, sheetIds) {
  for (const sheetId of sheetIds) {
    try {
      const payload = await fetchJson(`${API_BASE}/sheets/${encodeURIComponent(sheetId)}/rows/${encodeURIComponent(rowId)}/attachments`, { headers });
      const images = (payload?.data || []).filter(isImageAttachment);
      if (images.length) {
        return images.map((attachment) => ({ attachment, sheetId }));
      }
      // Se a chamada foi válida mas sem imagens, não tenta outras sheets para não gastar API.
      return [];
    } catch (_) {}
  }
  return [];
}

async function mapWithConcurrency(items, limit, mapper) {
  const results = new Array(items.length);
  let cursor = 0;
  const workers = Array.from({ length: Math.min(limit, items.length || 1) }, async () => {
    while (cursor < items.length) {
      const index = cursor;
      cursor += 1;
      results[index] = await mapper(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}

function parseRowIds(value) {
  return unique(String(value || '').split(/[\s,;|]+/)).filter((id) => /^\d+$/.test(id));
}

exports.handler = async (event) => {
  const API_KEY = getToken();
  if (!API_KEY) {
    return { statusCode: 500, body: 'Token Smartsheet não configurado. Configure SMARTSHEET_API_KEY ou SMARTSHEET_TOKEN no Netlify.' };
  }

  const qs = event.queryStringParameters || {};
  const headers = { Authorization: `Bearer ${API_KEY}` };
  const sheetIds = getCandidateSheetIds(event);

  // Modo leve para a tela Yinson: consulta metadados de anexos por linha em lotes.
  // Isso evita carregar attachments no endpoint principal e não estoura o limite de 30s da Netlify.
  const listMode = qs.list === '1' || qs.mode === 'list' || qs.action === 'list';
  if (listMode) {
    const maxRows = Math.max(1, Math.min(Number(qs.limit || 35), 60));
    const concurrency = Math.max(1, Math.min(Number(qs.concurrency || 6), 10));
    const rowIds = parseRowIds(qs.rowIds || qs.rows).slice(0, maxRows);
    if (!rowIds.length) return json(400, { ok: false, error: 'Nenhum rowId informado.' });
    try {
      const entries = await mapWithConcurrency(rowIds, concurrency, async (rid) => {
        const found = await getRowImageAttachments(headers, rid, sheetIds);
        return [rid, found.map(({ attachment, sheetId }) => ({
          id: attachment.id,
          name: attachment.name || `Imagem ${rid}`,
          mimeType: attachment.mimeType || '',
          rowId: rid,
          sheetId,
          source: 'row-attachment',
        }))];
      });
      const rowImages = Object.fromEntries(entries);
      const totalImages = Object.values(rowImages).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), 0);
      return json(200, {
        ok: true,
        sheetIds,
        requestedRows: rowIds.length,
        totalImages,
        rowImages,
      });
    } catch (err) {
      return json(err.status || 500, { ok: false, error: err.message || 'Erro ao consultar imagens das linhas.' });
    }
  }

  const id = qs.id;
  const rowId = qs.rowId;
  if (!id && !rowId) {
    return { statusCode: 400, body: 'Missing attachment id or rowId' };
  }

  try {
    let detail = null;
    if (id) {
      detail = await getAttachmentDetail(headers, id, sheetIds);
    } else if (rowId) {
      const rowImage = await getRowImageAttachment(headers, rowId, sheetIds, qs.imageIndex);
      if (!rowImage?.attachment?.id) {
        return { statusCode: 404, body: 'Nenhuma imagem encontrada para esta linha.' };
      }
      detail = await getAttachmentDetail(headers, rowImage.attachment.id, [rowImage.sheetId, ...sheetIds]);
    }

    if (detail?.url) {
      const wantsDownload = qs.download === '1' || qs.download === 'true';
      if (wantsDownload) {
        const response = await fetch(detail.url);
        if (!response.ok) {
          return { statusCode: response.status || 502, body: 'Falha ao baixar a imagem temporária do Smartsheet.' };
        }
        const arrayBuffer = await response.arrayBuffer();
        const requestedName = qs.filename || detail.name || `imagem-${id || rowId || Date.now()}.jpg`;
        const filename = sanitizeDownloadFilename(requestedName);
        const contentType = response.headers.get('content-type') || detail.mimeType || contentTypeFromName(filename);
        return {
          statusCode: 200,
          isBase64Encoded: true,
          headers: {
            'Content-Type': contentType,
            'Content-Disposition': `attachment; filename="${filename}"`,
            'Cache-Control': 'no-cache, no-store, must-revalidate',
          },
          body: Buffer.from(arrayBuffer).toString('base64'),
        };
      }
      return {
        statusCode: 302,
        headers: {
          Location: detail.url,
          'Cache-Control': 'no-cache, no-store, must-revalidate',
        },
        body: '',
      };
    }
    return { statusCode: 404, body: 'URL not found for attachment' };
  } catch (err) {
    return { statusCode: err.status || 500, body: err.message || 'Internal server error' };
  }
};
