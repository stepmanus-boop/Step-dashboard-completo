
const { readLocalJson } = require('./_auth');
const { buildPayload } = require('./projects');

async function loadProjectPayload() {
  try {
    return await buildPayload();
  } catch (_) {
    return readLocalJson('netlify/data/fallback-projects.json', { ok: true, projects: [], meta: {} });
  }
}

async function findProjectAndSpool(projectRowId, spoolIso) {
  const payload = await loadProjectPayload();
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
