const fs = require('fs/promises');
const path = require('path');
const { readLocalJson } = require('./_auth');
const { buildPayload } = require('./projects');

function withLookupTimeout(promise, timeoutMs) {
  const ms = Math.max(1000, Number(timeoutMs || 0));
  if (!ms) return promise;
  let timeoutId = null;
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => {
      const err = new Error(`Consulta do Tracking excedeu ${ms}ms.`);
      err.code = 'PROJECT_LOOKUP_TIMEOUT';
      reject(err);
    }, ms);
  });
  return Promise.race([promise, timeout]).finally(() => {
    if (timeoutId) clearTimeout(timeoutId);
  });
}

async function loadBundledFallbackPayload() {
  const emptyFallback = { ok: true, projects: [], meta: { fallbackMissing: true } };

  const candidates = [
    path.resolve(__dirname, '..', 'data', 'fallback-projects.json'),
    path.resolve(__dirname, '..', '..', 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'netlify', 'data', 'fallback-projects.json'),
    path.resolve(process.cwd(), 'data', 'fallback-projects.json'),
  ];

  for (const filePath of candidates) {
    try {
      const raw = await fs.readFile(filePath, 'utf8');
      const parsed = JSON.parse(raw || '{}');
      if (Array.isArray(parsed?.projects)) return parsed;
    } catch (error) {
      if (String(error?.code || '') !== 'ENOENT') {
        console.warn('Fallback de projetos encontrado, mas não pôde ser lido:', filePath, error?.message || error);
      }
    }
  }

  return readLocalJson('netlify/data/fallback-projects.json', emptyFallback);
}

async function loadProjectPayload(options = {}) {
  const allowFallback = options.allowFallback !== false;
  const preferCache = options.preferCache !== false;
  const timeoutMs = Number(options.timeoutMs || process.env.PROJECT_LOOKUP_TIMEOUT_MS || (allowFallback ? 4500 : 15000));
  try {
    return await withLookupTimeout(buildPayload({ preferCache }), timeoutMs);
  } catch (error) {
    if (!allowFallback) {
      const err = new Error('Não foi possível consultar o Smartsheet/Tracking em tempo real. O apontamento foi bloqueado para evitar registro em dados desatualizados.');
      err.statusCode = 503;
      err.cause = error;
      throw err;
    }
    const fallback = await loadBundledFallbackPayload();
    return {
      ...(fallback || {}),
      ok: true,
      meta: {
        ...(fallback?.meta || {}),
        servedFromFallbackForStageUpdate: true,
        fallbackReason: error?.message || 'lookup-failed',
      },
    };
  }
}

async function findProjectAndSpool(projectRowId, spoolIso, options = {}) {
  const payload = await loadProjectPayload(options);
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  const normalizedProjectId = String(projectRowId ?? '').trim();
  const project = projects.find((item) => {
    const rowId = String(item?.rowId ?? '').trim();
    const rowNumber = String(item?.rowNumber ?? '').trim();
    return (rowId && rowId === normalizedProjectId) || (rowNumber && rowNumber === normalizedProjectId);
  });
  if (!project) return { project: null, spool: null, payload };
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const normalizedSpoolIso = String(spoolIso || '').trim().toLowerCase();
  const spool = spools.find((item) => String(item?.iso || '').trim().toLowerCase() === normalizedSpoolIso);
  return { project, spool: spool || null, payload };
}

module.exports = {
  loadProjectPayload,
  findProjectAndSpool,
};
