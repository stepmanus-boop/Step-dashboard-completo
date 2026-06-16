const { jsonResponse } = require('./_auth');
const { handler: projectsHandler } = require('./projects');

function normalize(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

/**
 * v11 performance: módulo sob demanda para abrir uma BSP específica.
 * Mantém o painel leve porque detalhes pesados podem ser buscados apenas no clique.
 */
exports.handler = async (event) => {
  const rowId = String(event.queryStringParameters?.rowId || '').trim();
  const projectNumber = String(event.queryStringParameters?.project || event.queryStringParameters?.bsp || '').trim();
  if (!rowId && !projectNumber) {
    return jsonResponse(400, { ok: false, error: 'Informe rowId ou project.' });
  }

  event.queryStringParameters = { ...(event.queryStringParameters || {}), preferCache: event.queryStringParameters?.preferCache || '1' };
  const response = await projectsHandler(event);
  if (response.statusCode !== 200) return response;
  const payload = JSON.parse(response.body || '{}');
  const projects = Array.isArray(payload.projects) ? payload.projects : [];
  const target = projects.find((project) => {
    if (rowId && String(project.rowId) === rowId) return true;
    if (projectNumber) {
      const wanted = normalize(projectNumber);
      return normalize(project.projectNumber) === wanted || normalize(project.projectDisplay) === wanted || normalize(project.clientDisplayCode) === wanted;
    }
    return false;
  });

  if (!target) {
    return jsonResponse(404, { ok: false, error: 'BSP não encontrada no escopo do usuário.' });
  }

  return jsonResponse(200, {
    ok: true,
    meta: { ...(payload.meta || {}), module: 'project-detail' },
    project: target,
  }, {
    headers: { 'cache-control': 'private, max-age=60, stale-while-revalidate=120' },
  });
};
