
const { readLocalJson } = require('./_auth');
const { buildPayload } = require('./projects');

async function loadProjectPayload() {
  try {
    return await buildPayload();
  } catch (_) {
    return readLocalJson('netlify/data/fallback-projects.json', { ok: true, projects: [], meta: {} });
  }
}

function normalizeSpoolIso(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

async function findProjectAndSpool(projectRowId, spoolIso) {
  const payload = await loadProjectPayload();
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  const project = projects.find((item) => Number(item?.rowNumber || item?.rowId || 0) === Number(projectRowId || 0));
  if (!project) return { project: null, spool: null, payload };
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const targetIso = normalizeSpoolIso(spoolIso);
  const spool = spools.find((item) => normalizeSpoolIso(item?.iso) === targetIso);
  return { project, spool: spool || null, payload };
}

module.exports = {
  loadProjectPayload,
  findProjectAndSpool,
};
