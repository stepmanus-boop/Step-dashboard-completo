// Netlify function: fetch all image evidence for a BSP from the Tracking sheet.
// It searches the BSP summary row and all spool rows passed by rowIds/rowNumbers.
// Sources supported:
// 1) row attachments;
// 2) discussion attachments when returned by the sheet API;
// 3) cell images through the Smartsheet /imageurls endpoint.

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

function getTrackingSheetId(query = {}) {
  return String(
    query.sheetId
      || process.env.SMARTSHEET_TRACKING_SHEET_ID
      || process.env.SMARTSHEET_SHEET_ID
      || DEFAULT_TRACKING_SHEET_ID
  ).trim() || DEFAULT_TRACKING_SHEET_ID;
}

function splitIds(value) {
  if (!value) return [];
  return String(value)
    .split(',')
    .map((v) => v.trim())
    .filter((v) => /^\d+$/.test(v));
}

function uniq(list) {
  return Array.from(new Set((list || []).filter(Boolean).map(String)));
}

function imageNameFromCell(cell, image) {
  const candidates = [
    image && image.altText,
    cell && cell.displayValue,
    cell && cell.value,
    image && image.id ? `Imagem ${image.id}` : 'Imagem',
  ];
  const found = candidates.find((v) => v != null && String(v).trim());
  return String(found || 'Imagem').trim();
}

async function fetchJson(url, options) {
  const res = await fetch(url, options);
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    const err = new Error(`HTTP ${res.status} ${text}`);
    err.statusCode = res.status;
    throw err;
  }
  return res.json();
}

exports.handler = async (event) => {
  const API_KEY = getSmartsheetToken();
  if (!API_KEY) {
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({
        images: [],
        error: 'Token Smartsheet não configurado. Configure SMARTSHEET_API_KEY ou SMARTSHEET_TOKEN no Netlify.',
      }),
    };
  }
  const headers = { Authorization: `Bearer ${API_KEY}` };
  const query = event.queryStringParameters || {};

  // Executive BSP images must be searched in the Tracking sheet only.
  // A sheetId query parameter is still accepted for controlled tests, but the frontend sends 3612820139992964.
  const sheetId = getTrackingSheetId(query);

  const rowId = query.rowId ? String(query.rowId).trim() : '';
  const explicitRowIds = splitIds(query.rowIds);
  const rowNumbers = splitIds(query.rowNumbers);
  const directRowIds = uniq([rowId, ...explicitRowIds]);

  if (!directRowIds.length && !rowNumbers.length) {
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ images: [], warning: 'Missing rowId, rowIds or rowNumbers parameter' }),
    };
  }

  try {
    const rowsById = new Map();
    const addRows = (rows) => {
      for (const row of rows || []) {
        if (row && row.id) rowsById.set(String(row.id), row);
      }
    };

    // Fetch rows in batches using Get Sheet. This is faster and less error-prone than
    // making one attachments request and one row request for every spool.
    async function fetchRowsByParam(paramName, values) {
      const clean = uniq(values);
      if (!clean.length) return;
      const batchSize = paramName === 'rowIds' ? 80 : 150;
      for (let i = 0; i < clean.length; i += batchSize) {
        const batch = clean.slice(i, i + batchSize);
        const params = new URLSearchParams();
        params.set('include', 'attachments,discussions,objectValue');
        params.set('pageSize', '10000');
        params.set(paramName, batch.join(','));
        const url = `https://api.smartsheet.com/2.0/sheets/${sheetId}?${params.toString()}`;
        try {
          const json = await fetchJson(url, { headers });
          addRows(Array.isArray(json.rows) ? json.rows : []);
        } catch (err) {
          // If a rowIds batch fails because one ID is no longer valid, retry one by one.
          if (paramName === 'rowIds' && batch.length > 1) {
            for (const single of batch) {
              try {
                const paramsSingle = new URLSearchParams();
                paramsSingle.set('include', 'attachments,discussions,objectValue');
                paramsSingle.set('pageSize', '10000');
                paramsSingle.set('rowIds', single);
                const jsonSingle = await fetchJson(`https://api.smartsheet.com/2.0/sheets/${sheetId}?${paramsSingle.toString()}`, { headers });
                addRows(Array.isArray(jsonSingle.rows) ? jsonSingle.rows : []);
              } catch (_) {
                // Ignore invalid row in this sheet.
              }
            }
          }
          // For rowNumbers, non-existent row numbers are ignored by Smartsheet, so a failure is skipped.
        }
      }
    }

    await fetchRowsByParam('rowIds', directRowIds);
    await fetchRowsByParam('rowNumbers', rowNumbers);

    const attachmentMap = new Map();
    const cellImageMap = new Map();

    for (const row of rowsById.values()) {
      // Row-level attachments returned by include=attachments.
      for (const att of Array.isArray(row.attachments) ? row.attachments : []) {
        if (!att || !att.id) continue;
        const mime = String(att.mimeType || '').toLowerCase();
        const name = String(att.name || '').toLowerCase();
        const looksImage = mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp|heic|heif)$/i.test(name);
        if (looksImage && !attachmentMap.has(String(att.id))) {
          attachmentMap.set(String(att.id), { ...att, _sheetId: sheetId });
        }
      }

      // Discussion/comment attachments if included.
      for (const discussion of Array.isArray(row.discussions) ? row.discussions : []) {
        const comments = Array.isArray(discussion.comments) ? discussion.comments : [];
        const discussionAttachments = Array.isArray(discussion.attachments) ? discussion.attachments : [];
        const allDiscussionAttachments = [...discussionAttachments];
        for (const comment of comments) {
          if (Array.isArray(comment.attachments)) allDiscussionAttachments.push(...comment.attachments);
        }
        for (const att of allDiscussionAttachments) {
          if (!att || !att.id) continue;
          const mime = String(att.mimeType || '').toLowerCase();
          const name = String(att.name || '').toLowerCase();
          const looksImage = mime.startsWith('image/') || /\.(png|jpe?g|webp|gif|bmp|heic|heif)$/i.test(name);
          if (looksImage && !attachmentMap.has(String(att.id))) {
            attachmentMap.set(String(att.id), { ...att, _sheetId: sheetId });
          }
        }
      }

      // Cell images.
      for (const cell of Array.isArray(row.cells) ? row.cells : []) {
        const image = cell && (cell.image || (cell.objectValue && cell.objectValue.image));
        if (image && image.id && !cellImageMap.has(String(image.id))) {
          cellImageMap.set(String(image.id), {
            imageId: String(image.id),
            altText: imageNameFromCell(cell, image),
            height: Number(image.height) || undefined,
            width: Number(image.width) || undefined,
            rowId: row.id,
            sheetId,
          });
        }
      }
    }

    const images = [];

    // Resolve attachment URLs. Limit concurrency so we do not trigger rate spikes.
    const attachments = Array.from(attachmentMap.values());
    for (const att of attachments) {
      let url = att.url || '';
      if (!url) {
        try {
          let detailRes = await fetch(`https://api.smartsheet.com/2.0/sheets/${sheetId}/attachments/${att.id}`, { headers });
          if (!detailRes.ok) detailRes = await fetch(`https://api.smartsheet.com/2.0/attachments/${att.id}`, { headers });
          if (detailRes.ok) {
            const detail = await detailRes.json();
            url = detail.url || '';
          }
        } catch (_) {
          url = '';
        }
      }
      images.push({
        id: String(att.id),
        name: att.name || `Anexo ${att.id}`,
        mimeType: att.mimeType || 'image',
        url,
        sheetId,
        source: 'attachment',
      });
    }

    // Resolve cell image URLs in batches.
    const cellImages = Array.from(cellImageMap.values());
    for (let i = 0; i < cellImages.length; i += 100) {
      const batch = cellImages.slice(i, i + 100);
      const payload = batch.map((img) => {
        const item = { imageId: img.imageId };
        if (img.height) item.height = img.height;
        if (img.width) item.width = img.width;
        return item;
      });
      try {
        const res = await fetch('https://api.smartsheet.com/2.0/imageurls', {
          method: 'POST',
          headers: { ...headers, 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        if (res.ok) {
          const json = await res.json();
          const urls = Array.isArray(json.imageUrls) ? json.imageUrls : [];
          for (const u of urls) {
            const meta = batch.find((img) => img.imageId === u.imageId);
            if (!meta || !u.url) continue;
            images.push({
              id: meta.imageId,
              name: meta.altText || `Imagem ${meta.imageId}`,
              mimeType: 'image/cell',
              url: u.url,
              sheetId,
              source: 'cell-image',
            });
          }
        }
      } catch (_) {
        // Ignore cell image URL errors; attachments may still be returned.
      }
    }

    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({ images, rowsChecked: rowsById.size }),
    };
  } catch (err) {
    // Keep the frontend from failing hard; return a safe empty result with diagnostics.
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: JSON.stringify({ images: [], error: err.message || 'Erro ao carregar imagens' }),
    };
  }
};
