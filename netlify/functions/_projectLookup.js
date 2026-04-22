
const { readLocalJson } = require('./_auth');
const { buildPayload } = require('./projects');

async function loadProjectPayload() {
  try {
    return await buildPayload();
  } catch (_) {
    return readLocalJson('netlify/data/fallback-projects.json', { ok: true, projects: [], meta: {} });
  }
}

function normalizeLookupValue(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function findProjectByRow(projects, projectRowId) {
  const target = Number(projectRowId || 0);
  if (!target) return null;
  return projects.find((item) => Number(item?.rowId || 0) === target || Number(item?.rowNumber || 0) === target) || null;
}

function findSpoolInProject(project, spoolIso) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const targetRaw = String(spoolIso || '').trim();
  const target = normalizeLookupValue(targetRaw);
  if (!target) return null;

  return spools.find((item) => {
    const iso = String(item?.iso || '').trim();
    const drawing = String(item?.drawing || '').trim();
    const normalizedIso = normalizeLookupValue(iso);
    const normalizedDrawing = normalizeLookupValue(drawing);
    return normalizedIso === target
      || normalizedDrawing === target
      || normalizedIso.includes(target)
      || target.includes(normalizedIso)
      || normalizedDrawing.includes(target)
      || target.includes(normalizedDrawing);
  }) || null;
}

async function findProjectAndSpool(projectRowId, spoolIso) {
  const payload = await loadProjectPayload();
  const projects = Array.isArray(payload?.projects) ? payload.projects : [];
  let project = findProjectByRow(projects, projectRowId);
  let spool = project ? findSpoolInProject(project, spoolIso) : null;

  if (!project && spoolIso) {
    for (const item of projects) {
      const candidate = findSpoolInProject(item, spoolIso);
      if (candidate) {
        project = item;
        spool = candidate;
        break;
      }
    }
  }

  if (project && !spool) {
    spool = findSpoolInProject(project, spoolIso);
  }

  return { project: project || null, spool: spool || null, payload };
}

module.exports = {
  loadProjectPayload,
  findProjectAndSpool,
};
