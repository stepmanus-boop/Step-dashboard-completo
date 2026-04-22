
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

function buildLookupCandidates(value, project = null) {
  const raw = String(value || '').trim();
  if (!raw) return [];

  const normalizedRaw = normalizeLookupValue(raw);
  const compactRaw = raw.replace(/\s+/g, ' ').trim();
  const projectNumber = String(project?.projectNumber || project?.projectDisplay || '').trim();
  const normalizedProjectNumber = normalizeLookupValue(projectNumber);
  const suffixRaw = projectNumber && raw.toLowerCase().startsWith(projectNumber.toLowerCase())
    ? raw.slice(projectNumber.length).trim()
    : raw;
  const normalizedSuffix = normalizeLookupValue(suffixRaw);
  const tokens = raw.split(/\s+/g).map((item) => normalizeLookupValue(item)).filter(Boolean);

  const candidates = new Set([normalizedRaw, normalizedSuffix, ...tokens]);

  if (normalizedProjectNumber && normalizedRaw.startsWith(normalizedProjectNumber)) {
    candidates.add(normalizedRaw.slice(normalizedProjectNumber.length));
  }

  const isoMatch = raw.match(/ISO[-\s_]*([A-Z0-9]+)/i);
  const splMatch = raw.match(/SPL[-\s_]*([A-Z0-9]+)/i);
  if (isoMatch) {
    candidates.add(normalizeLookupValue(`ISO-${isoMatch[1]}`));
    candidates.add(normalizeLookupValue(`ISO ${isoMatch[1]}`));
  }
  if (splMatch) {
    candidates.add(normalizeLookupValue(`SPL-${splMatch[1]}`));
    candidates.add(normalizeLookupValue(`SPL ${splMatch[1]}`));
  }
  if (isoMatch && splMatch) {
    candidates.add(normalizeLookupValue(`ISO-${isoMatch[1]}-SPL-${splMatch[1]}`));
    candidates.add(normalizeLookupValue(`ISO ${isoMatch[1]} SPL ${splMatch[1]}`));
  }

  return Array.from(candidates).filter(Boolean);
}

function findSpoolInProject(project, spoolIso) {
  const spools = Array.isArray(project?.spools) ? project.spools : [];
  const targets = buildLookupCandidates(spoolIso, project);
  if (!targets.length) return null;

  return spools.find((item) => {
    const variants = [
      String(item?.iso || '').trim(),
      String(item?.drawing || '').trim(),
      String(item?.description || '').trim(),
    ].filter(Boolean);
    const normalizedVariants = variants.map((value) => normalizeLookupValue(value)).filter(Boolean);

    return targets.some((target) => normalizedVariants.some((candidate) => (
      candidate === target
      || candidate.includes(target)
      || target.includes(candidate)
    )));
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
