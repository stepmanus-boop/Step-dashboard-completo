const { jsonResponse } = require('./_auth');
const { buildPayload } = require('./projects');

/**
 * Endpoint de Pré-aquecimento v32
 * Objetivo: Forçar a leitura completa do Smartsheet e atualizar o cache em memória + snapshot em disco.
 * Deve ser chamado por um cron job externo (ex: UptimeRobot) a cada 10-15 min.
 */
exports.handler = async (event) => {
  // Proteção por token secreto
  const secret = event.queryStringParameters?.secret;
  const WARMUP_SECRET = process.env.WARMUP_SECRET;

  if (!WARMUP_SECRET || secret !== WARMUP_SECRET) {
    return jsonResponse(403, { ok: false, error: 'Não autorizado. Secret inválido.' });
  }

  if (event.httpMethod !== 'GET') {
    return jsonResponse(405, { ok: false, error: 'Método não permitido.' });
  }

  try {
    console.log('[v32-warmup] Iniciando pré-aquecimento do cache...');
    const startTime = Date.now();
    
    // Força leitura completa do Smartsheet
    const payload = await buildPayload({ force: true });
    
    const duration = Date.now() - startTime;
    const projectsCount = payload.projects?.length || 0;

    console.log(`[v32-warmup] Sucesso: ${projectsCount} projetos carregados em ${duration}ms.`);

    return jsonResponse(200, {
      ok: true,
      message: 'Cache aquecido com sucesso.',
      stats: {
        projectsCount,
        durationMs: duration,
        timestamp: new Date().toISOString(),
        version: payload.meta?.version
      }
    });
  } catch (error) {
    console.error('[v32-warmup] Falha no pré-aquecimento:', error.message);
    return jsonResponse(500, { ok: false, error: error.message });
  }
};
