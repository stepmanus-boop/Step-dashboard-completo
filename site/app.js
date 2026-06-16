/* STEP Dashboard v37.18-cache-age-visible - carregador de compatibilidade.
   O app principal foi dividido em chunks em site/js/ para reduzir bloqueio de carregamento e melhorar cache. */
(function(){
  if (window.__STEP_APP_CHUNKS_LOADED__) return;
  window.__STEP_APP_CHUNKS_LOADED__ = true;
  var chunks = [
    './js/app-01-core.js?v=37.18-cache-age-visible',
    './js/app-02-client-portal.js?v=37.18-cache-age-visible',
    './js/app-03-dashboard-render.js?v=37.18-cache-age-visible',
    './js/app-04-data-auth-admin.js?v=37.61-background-sync-confirmado',
    './js/app-05-stage-login-init.js?v=37.18-cache-age-visible'
  ];
  function loadNext(index){
    if (index >= chunks.length) return;
    var script = document.createElement('script');
    script.src = chunks[index];
    script.async = false;
    script.onload = function(){ loadNext(index + 1); };
    script.onerror = function(){
      console.error('[STEP] Falha ao carregar módulo:', chunks[index]);
      var body = document.getElementById('projects-body');
      if (body) body.innerHTML = '<tr><td colspan="18">Falha ao carregar módulos do painel. Atualize a página.</td></tr>';
    };
    document.head.appendChild(script);
  }
  loadNext(0);
})();
