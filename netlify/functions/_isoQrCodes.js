const crypto = require('crypto');
const {
  isSupabaseConfigured,
  getSupabaseHeaders,
  supabaseFetch,
} = require('./_supabase');

const ISO_QR_TABLE = 'iso_qr_codes';
const OPERATION_REGION = 'BR';
const MAX_AUTO_QR_BATCH_SIZE = Math.max(25, Number(process.env.ISO_QR_AUTO_BATCH_SIZE || 500));

function cleanText(value = '') {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function cleanLower(value = '') {
  return cleanText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function safeColumnText(value = '', max = 500) {
  return cleanText(value).slice(0, max);
}

function buildBaseUrl(event = null) {
  const envUrl = cleanText(
    process.env.URL ||
    process.env.DEPLOY_PRIME_URL ||
    process.env.SITE_URL ||
    process.env.PUBLIC_SITE_URL ||
    ''
  ).replace(/\/$/, '');
  if (/^https?:\/\//i.test(envUrl)) return envUrl;

  const headers = event?.headers || {};
  const proto = headers['x-forwarded-proto'] || headers['X-Forwarded-Proto'] || 'https';
  const host = headers.host || headers.Host || '';
  if (host) return `${proto}://${host}`.replace(/\/$/, '');
  return '';
}

function buildQrUrl(token, options = {}) {
  const base = cleanText(options.baseUrl || buildBaseUrl(options.event)).replace(/\/$/, '');
  const path = `/qr-tracking.html?token=${encodeURIComponent(String(token || ''))}`;
  return base ? `${base}${path}` : path;
}

function getProjectBsp(project = {}) {
  return safeColumnText(project.projectDisplay || project.projectNumber || project.projectCode || project.project || project.rawProject || '', 120);
}

function getProjectWorkOrder(project = {}) {
  return safeColumnText(project.customerPoDisplay || project.customerPo || (Array.isArray(project.customerPoList) ? project.customerPoList.join(' / ') : '') || '', 180);
}

function getProjectTagNumber(project = {}, spool = {}) {
  return safeColumnText(spool.tagNumber || spool.tag || spool.lineNumber || project.lineNumber || '', 180);
}

function getIsoFullName(spool = {}, fallback = '') {
  return safeColumnText(spool.iso || spool.drawing || fallback || '', 260);
}

function normalizeIsoKey(value = '') {
  return cleanLower(value).replace(/[^a-z0-9]+/g, ' ').trim();
}

function projectToCandidate(project = {}, spool = {}, options = {}) {
  const iso = getIsoFullName(spool, project.summaryDrawing);
  if (!iso || cleanLower(iso) === 'iso') return null;
  const bsp = getProjectBsp(project);
  const client = safeColumnText(spool.client || project.client || '', 160);
  const token = crypto.randomUUID();
  return {
    region: OPERATION_REGION,
    client,
    client_key: cleanLower(client),
    bsp,
    bsp_key: cleanLower(bsp),
    work_order: getProjectWorkOrder(project),
    vessel: safeColumnText(spool.vessel || project.vessel || '', 160),
    tag_number: getProjectTagNumber(project, spool),
    iso,
    iso_key: normalizeIsoKey(iso),
    iso_full_name: iso,
    qr_token: token,
    qr_url: buildQrUrl(token, options),
    status: safeColumnText(spool.currentStatus || spool.stage || project.currentStatus || project.currentStage || project.projectStatus || '', 220),
    progress: Number(spool.overallProgress ?? spool.individualProgress ?? spool.stagePercent ?? project.overallProgress ?? 0) || 0,
    source: safeColumnText(options.source || 'tracking-cache-auto', 120),
    updated_at: new Date().toISOString(),
  };
}

function collectIsoQrCandidates(payload = {}, options = {}) {
  const seen = new Set();
  const candidates = [];
  const projects = Array.isArray(payload.projects) ? payload.projects : [];

  for (const project of projects) {
    const spools = Array.isArray(project?.spools) ? project.spools : [];
    const sourceRows = spools.length ? spools : (project?.summaryDrawing && cleanLower(project.summaryDrawing) !== 'iso' ? [{ iso: project.summaryDrawing }] : []);
    for (const spool of sourceRows) {
      const candidate = projectToCandidate(project, spool, options);
      if (!candidate?.iso) continue;
      const key = [candidate.region, cleanLower(candidate.client), cleanLower(candidate.bsp), normalizeIsoKey(candidate.iso)].join('::');
      if (seen.has(key)) continue;
      seen.add(key);
      candidates.push(candidate);
    }
  }

  return candidates;
}

function chunkArray(items = [], size = MAX_AUTO_QR_BATCH_SIZE) {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) chunks.push(items.slice(index, index + size));
  return chunks;
}

async function ensureIsoQrCodesForPayload(payload = {}, options = {}) {
  if (!isSupabaseConfigured()) {
    return { ok: false, skipped: true, reason: 'supabase-not-configured', totalCandidates: 0, inserted: 0 };
  }

  const candidates = collectIsoQrCandidates(payload, options);
  if (!candidates.length) {
    return { ok: true, skipped: true, reason: 'no-iso-candidates', totalCandidates: 0, inserted: 0 };
  }

  let inserted = 0;
  let attempted = 0;
  const errors = [];

  for (const chunk of chunkArray(candidates)) {
    attempted += chunk.length;
    try {
      const rows = await supabaseFetch(`/rest/v1/${ISO_QR_TABLE}?on_conflict=region,client_key,bsp_key,iso_key&select=id,iso,qr_token`, {
        method: 'POST',
        headers: getSupabaseHeaders('resolution=ignore-duplicates,return=representation'),
        body: JSON.stringify(chunk),
        timeoutMs: Number(options.timeoutMs || process.env.ISO_QR_SUPABASE_TIMEOUT_MS || 9000),
        retries: 1,
      });
      inserted += Array.isArray(rows) ? rows.length : 0;
    } catch (error) {
      errors.push(String(error?.message || error).slice(0, 260));
    }
  }

  return {
    ok: errors.length === 0,
    totalCandidates: candidates.length,
    attempted,
    inserted,
    skippedExisting: Math.max(0, attempted - inserted),
    errors,
  };
}

function escapePostgrestLike(value = '') {
  return cleanText(value).replace(/[,*()]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
}

function mapIsoQrRow(row = {}) {
  if (!row) return null;
  return {
    id: row.id,
    region: row.region || '',
    client: row.client || '',
    bsp: row.bsp || '',
    workOrder: row.work_order || '',
    vessel: row.vessel || '',
    tagNumber: row.tag_number || '',
    iso: row.iso || row.iso_full_name || '',
    isoFullName: row.iso_full_name || row.iso || '',
    qrToken: row.qr_token || '',
    qrUrl: row.qr_url || (row.qr_token ? buildQrUrl(row.qr_token) : ''),
    status: row.status || '',
    progress: Number(row.progress || 0),
    createdAt: row.created_at || null,
    updatedAt: row.updated_at || null,
  };
}

async function listIsoQrCodes(options = {}) {
  if (!isSupabaseConfigured()) throw new Error('Supabase não configurado para consultar QR Codes.');
  const limit = Math.min(500, Math.max(1, Number(options.limit || 100)));
  const params = [
    'select=id,region,client,bsp,work_order,vessel,tag_number,iso,iso_full_name,qr_token,qr_url,status,progress,created_at,updated_at',
    `region=eq.${encodeURIComponent(cleanText(options.region || OPERATION_REGION))}`,
    'order=bsp.asc,iso.asc',
    `limit=${limit}`,
  ];

  const query = escapePostgrestLike(options.query || '');
  if (query) {
    const pattern = `*${query}*`;
    const or = [
      `iso.ilike.${pattern}`,
      `iso_full_name.ilike.${pattern}`,
      `bsp.ilike.${pattern}`,
      `work_order.ilike.${pattern}`,
      `vessel.ilike.${pattern}`,
      `tag_number.ilike.${pattern}`,
      `client.ilike.${pattern}`,
    ].join(',');
    params.push(`or=${encodeURIComponent(`(${or})`)}`);
  }

  const rows = await supabaseFetch(`/rest/v1/${ISO_QR_TABLE}?${params.join('&')}`, {
    method: 'GET',
    headers: getSupabaseHeaders(),
    timeoutMs: Number(options.timeoutMs || 9000),
    retries: 1,
  });
  return (Array.isArray(rows) ? rows : []).map(mapIsoQrRow).filter(Boolean);
}

async function getIsoQrByToken(token, options = {}) {
  const cleanToken = cleanText(token);
  if (!cleanToken) return null;
  if (!isSupabaseConfigured()) throw new Error('Supabase não configurado para consultar QR Code.');
  const rows = await supabaseFetch(`/rest/v1/${ISO_QR_TABLE}?qr_token=eq.${encodeURIComponent(cleanToken)}&select=id,region,client,bsp,work_order,vessel,tag_number,iso,iso_full_name,qr_token,qr_url,status,progress,created_at,updated_at&limit=1`, {
    method: 'GET',
    headers: getSupabaseHeaders(),
    timeoutMs: Number(options.timeoutMs || 8000),
    retries: 1,
  });
  return mapIsoQrRow(Array.isArray(rows) ? rows[0] : null);
}

module.exports = {
  ISO_QR_TABLE,
  OPERATION_REGION,
  buildBaseUrl,
  buildQrUrl,
  collectIsoQrCandidates,
  ensureIsoQrCodesForPayload,
  listIsoQrCodes,
  getIsoQrByToken,
};
