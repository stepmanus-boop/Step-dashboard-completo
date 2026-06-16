const CACHE_NAME = 'step-dashboard-37.74-zd230-zpl-label-print-qr';
// - v37.02: revisa performance, otimiza i18n dinâmico/Yinson e preserva apontamentos rápidos
// Versão v28: preserva cache local entre logouts e usa caminho rápido pós-login
// - Mantém assets do app shell versionados para liberar app.js corrigido imediatamente
// - API continua sem cache no Service Worker; dados operacionais usam cache local validado pelo app
// - Interceptação de /api restrita à própria origem para evitar efeitos colaterais
// - v36.45: evita liberar Portal do Cliente com fallback antigo sem PO
// - v36.46: adiciona coluna Report no cronograma do cliente e exportação Excel por BSP
// - v36.47: remove a coluna Report da lista e exibe o report dentro da visão executiva principal
// - v36.48: Excel passa a exportar o painel executivo completo e remove destaque amarelo do report
// - v36.51: corrige carregamento dos apontamentos PCP sem derrubar a tela quando Tracking/Smartsheet demora
// - v36.52: corrige localização de Tracking por BSP/spool mesmo quando rowId antigo ou campo de avanço está vazio
// - v36.54: separa carregamento da Validação PCP da consulta Smartsheet para evitar timeout de 30s
// - v36.59: remove fallback estático antigo e força nova chave de cache do painel Brasil
// - v36.59: alinha carregamento de projetos ao painel Portugal estável, usando Smartsheet BR
// - v36.69: adiciona permissão por usuário para visualizar Painel do Cliente
// - v36.71: adiciona botão discreto de atualização no Portal do Cliente e renova cache local
// - v36.72: editor do Painel do Cliente passa a ajustar Drawings e Procurement antes da fabricação
// - v36.74: painel individual da TAG/ISO busca Drawing Documentation Control sob demanda
// - v36.79: restaura app/projetos da v36.76 e mantém Drawing leve isolado
// - v36.80: acelera login do Portal do Cliente; POs atualizam em segundo plano sem travar a abertura
// - v36.83: restaura carregamento estável do Portal do Cliente e mantém favicon STEP
// - v36.84: Portal do Cliente carrega somente linhas do próprio cliente e PO leve
// - v36.85: login aparece imediatamente; painel fica oculto até autenticar
// - v36.86: aumenta contraste do aviso Em Tratativa no Painel do Cliente
// - v36.68: remove bordas brancas e força a logo do cliente a preencher completamente o card
// - v36.96: adiciona seletor de idioma PT/EN/ES e inclui i18n.js no cache do app
// - v36.97: logout limpava Cache Storage e Service Worker
// - v36.98: logout limpa sessão sem remover cache operacional/app shell, mantendo login rápido
// - v37.06: corrige abertura das imagens da BSP no Painel Brasil usando aliases seguros do token Smartsheet
// - v37.08: cache persistente seguro para Supabase Free, sem histórico e com linha única
// - v37.17-cache-only-login-scheduled-manual: cache atualiza por rotina agendada independente do usuário
// - v37.18-cache-age-visible: exibe há quanto tempo o cache Supabase foi atualizado
// - v37.22: atualização manual/agendada checa versão antes de baixar sheet completa e toca timestamp se não houve mudança
// - v37.28: botão Atualizar e agendador tocam updated_at ao pegar lock; sem idade >15 min
// - v37.30: corrige cards PCP visíveis e adiciona lembrete popup para BSPs a vencer em 7 dias
// - v37.31: popup PCP mostra apenas clientes no primeiro nível e expande BSPs ao clicar
// - v36.99: amplia a cobertura da tradução PT/EN/ES no Portal do Cliente e áreas dinâmicas
// - v37.00: destaca botão Yinson de Projetos em Desenvolvimento e ajusta colunas internas
// - v37.48: aplica cores suaves na linha inteira da tabela Yinson conforme status
// - v37.49: adiciona cor laranja suave para linhas e cards com status ONGOING
// - v37.50: remove a faixa de cards/resumo de status da página Yinson
// - v37.62: fixa topo/menu da página Yinson e normaliza cores/status em Projetos em Desenvolvimento
// - v37.63: adiciona coluna PRIORITY LEVEL na segunda posição da tabela Yinson
// - v37.64: corrige heartbeat do cache 15 min e adiciona Pesquisa QR Code automática por ISO
// - v37.65: adiciona watchdog no painel e wrappers modernos de Scheduled Function
// - v37.66: corrige contagem exclusiva dos cards e arredondamento do pendente de solda
// - v37.67: Total de Projetos passa a usar a mesma base do Excel/tabela; iniciados/não iniciados/on hold ficam como recortes operacionais
// - v37.68: cards passam a usar sempre a base filtrada real do Smartsheet/tabela; Total não exclui On Hold/Não iniciado/Enviado
// - v37.70: cards principais passam a usar classificação exclusiva pela linha raiz do Smartsheet, sem duplicar BSP com ISOs mistos
// - v37.71: coluna Status do painel principal passa a exibir o status real do Tracking/Smartsheet, não o estado genérico Em produção
// - v37.72: amplia a tradução dinâmica PT/EN do Portal do Cliente, etapas operacionais, QR Code e relógios
// - v37.73: oculta informações internas do cliente, corrige data ISO individual e impressão Zebra fixa 3 quadros
// - v37.74: ajusta impressão Zebra ZD230-203dpi ZPL usando tamanho Custom do driver e espera as imagens carregarem
const APP_SHELL = [
  "/",
  "/app.css",
  "/app.js",
  "/js/app-01-core.js",
  "/js/app-02-client-portal.js",
  "/js/app-03-dashboard-render.js",
  "/js/app-04-data-auth-admin.js",
  "/js/app-05-stage-login-init.js",
  "/js/app-06-qr-codes.js",
  "/i18n.js",
  "/client-under-development.html",
  "/qr-tracking.html",
  "/manifest.webmanifest",
  "/assets/favicon.ico",
  "/assets/favicon-32x32.png",
  "/assets/favicon-16x16.png",
  "/favicon.ico",
  "/assets/step-logo.png",
  "/assets/icon-192.png",
  "/assets/icon-512.png",
  "/assets/apple-touch-icon.png",
];

const CORE_ASSETS = new Set([
  "/",
  "/index.html",
  "/app.css",
  "/app.js",
  "/js/app-01-core.js",
  "/js/app-02-client-portal.js",
  "/js/app-03-dashboard-render.js",
  "/js/app-04-data-auth-admin.js",
  "/js/app-05-stage-login-init.js",
  "/js/app-06-qr-codes.js",
  "/i18n.js",
  "/client-under-development.html",
  "/qr-tracking.html",
  "/manifest.webmanifest",
  "/assets/favicon.ico",
  "/assets/favicon-32x32.png",
  "/assets/favicon-16x16.png",
  "/favicon.ico",
  "/sw.js",
]);

function toUrl(input) {
  try {
    return new URL(input, self.location.origin);
  } catch {
    return null;
  }
}

function isCoreAsset(request) {
  const url = toUrl(request.url);
  if (!url || url.origin !== self.location.origin) return false;
  return CORE_ASSETS.has(url.pathname);
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = toUrl(request.url);
  if (url && url.origin === self.location.origin && url.pathname.startsWith("/api/")) {
    event.respondWith(
      fetch(request, { cache: "no-store" }).catch(() =>
        new Response(JSON.stringify({ ok: false, offline: true, error: "Sem conexão no momento." }), {
          headers: { "Content-Type": "application/json" },
        })
      )
    );
    return;
  }

  if (isCoreAsset(request)) {
    event.respondWith(
      fetch(request, { cache: "no-store" })
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => {});
          return response;
        })
        .catch(() => caches.match(request).then((cached) => cached || caches.match("/")))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then((cached) => {
      return cached || fetch(request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => {});
        return response;
      }).catch(() => caches.match("/"));
    })
  );
});


self.addEventListener("push", (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (_) {}
  const title = data.title || 'Novo alerta';
  const options = {
    body: data.body || 'Você recebeu uma nova notificação.',
    icon: '/assets/icon-192.png',
    badge: '/assets/icon-192.png',
    tag: data.tag || 'step-alert',
    data: { url: data.url || '/' },
    requireInteraction: true,
    renotify: true,
    vibrate: [220, 120, 220],
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = event.notification?.data?.url || '/';
  event.waitUntil(clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
    for (const client of clientList) {
      if ('focus' in client) {
        client.navigate(targetUrl).catch(() => {});
        return client.focus();
      }
    }
    if (clients.openWindow) return clients.openWindow(targetUrl);
  }));
});

// - v37.32: fixa ID da sheet Yinson/Projetos em Desenvolvimento informado pelo usuário e amplia aliases de colunas

// - v37.33: corrige mapeamento Yinson por ID fixo 7523600180400004, aceita Project/Client WO/Line Nº/OBSERVATIONS.

// - v37.35: Yinson usa leitura paginada leve da sheet antes de tentar modos pesados, evitando HTTP 500.
// - v37.37: Yinson passa a carregar primeiro do cache Supabase próprio e sincroniza em rotina/manual, sem depender do Smartsheet ao vivo.

// - v37.38: corrige lock Yinson preso na primeira sincronização e ignora cache bootstrap inválido.
