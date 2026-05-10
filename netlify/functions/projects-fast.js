const { jsonResponse, requireSession } = require('./_auth');
const { buildPayload } = require('./projects');

/**
 * Endpoint rápido v32: delega para buildPayload com preferCache=1.
 * Isso garante que o login use o cache em memória ou snapshot em disco
 * sem tocar no Smartsheet.
 */
exports.handler = async (event) => {
  const auth = requireSession(event);
  if (!auth.ok) {
    return jsonResponse(401, { ok: false, error: 'Faça login para visualizar o painel.' });
  }

  try {
    // Sempre prefere cache neste endpoint
    const payload = await buildPayload({ force: false, preferCache: true });
    
    // Nota: Este endpoint não faz o scoping por cliente (scopePayloadForSession)
    // para ser o mais rápido possível. O frontend lida com o filtro se necessário,
    // ou o projects.js original pode ser usado se o scoping for obrigatório no servidor.
    // Mas para o Portal do Cliente ser instantâneo, vamos manter a lógica de scoping aqui também.
    
    // Importamos as funções de scoping do projects.js se necessário, 
    // mas como elas não estão exportadas, vamos apenas delegar para o handler principal
    // se quisermos o comportamento completo, ou duplicar a lógica mínima.
    
    // Decisão v32: Para garantir consistência total (incluindo scoping de cliente),
    // vamos apenas chamar o handler de projects.js injetando preferCache=1.
    
    const projectsHandler = require('./projects').handler;
    event.queryStringParameters = { ...event.queryStringParameters, preferCache: '1' };
    return await projectsHandler(event);
    
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message });
  }
};
