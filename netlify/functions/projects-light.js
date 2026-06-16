const { jsonResponse } = require('./_auth');
const { handler: projectsHandler } = require('./projects');

function toLightProject(project = {}) {
  return {
    rowId: project.rowId,
    rowNumber: project.rowNumber,
    projectNumber: project.projectNumber,
    projectDisplay: project.projectDisplay,
    client: project.client,
    vessel: project.vessel,
    unit: project.unit,
    customerPo: project.customerPo,
    customerPoDisplay: project.customerPoDisplay,
    itemCount: project.itemCount,
    tagsCount: project.tagsCount,
    weight: project.weight,
    weldedWeight: project.weldedWeight,
    m2: project.m2,
    status: project.status,
    currentStage: project.currentStage,
    overallProgress: project.overallProgress,
    individualProgress: project.individualProgress,
    finishDate: project.finishDate,
  };
}

/**
 * v11 performance: devolve a lista leve para abertura rápida da tela.
 * Internamente reaproveita /api/projects para manter autenticação, escopo de cliente e cache.
 */
exports.handler = async (event) => {
  event.queryStringParameters = { ...(event.queryStringParameters || {}), preferCache: event.queryStringParameters?.preferCache || '1' };
  const response = await projectsHandler(event);
  if (response.statusCode !== 200) return response;
  const payload = JSON.parse(response.body || '{}');
  return jsonResponse(200, {
    ok: true,
    meta: { ...(payload.meta || {}), module: 'projects-light' },
    stats: payload.stats || {},
    alerts: Array.isArray(payload.alerts) ? payload.alerts : [],
    projects: Array.isArray(payload.projects) ? payload.projects.map(toLightProject) : [],
  }, {
    headers: { 'cache-control': 'private, max-age=60, stale-while-revalidate=120' },
  });
};
