# Relatório técnico de otimização pós-login — STEP Dashboard Tracking

**Autor:** Manus AI  
**Data:** 2026-05-09  
**Projeto analisado:** `step_dashboard_tracking_CLIENTE_CORRIGIDO_v24`

## 1. Sumário executivo

A lentidão percebida após o login vinha principalmente de uma combinação de **requisições em cascata**, **uso bloqueante do cache local**, **hidratação repetida de sessão no Supabase** e **releituras sequenciais no backend em cenário sem cache**. O fluxo anterior validava o login, chamava `/api/auth-me`, carregava `/api/projects` e, somente depois, executava alertas, respostas, apontamentos e dados administrativos em série. Para o Portal do Cliente, a rota `/api/projects` ainda passava por `hydrateClientSession()`, que podia consultar o Supabase em todas as chamadas, somando latência antes de filtrar os projetos do cliente.

As correções aplicadas tornam o carregamento **cache-first**, com **stale-while-revalidate** no frontend. Quando há cache local aproveitável, o painel renderiza imediatamente e a atualização da API acontece em background. Em paralelo, o login agora inicia a reidratação de sessão e o carregamento de projetos ao mesmo tempo; as cargas auxiliares foram deslocadas para execução paralela e não bloqueiam a primeira visualização dos dados. No backend, foi adicionado cache curto para reidratação de usuários cliente e para `/api/auth-me`, além de paralelização da leitura complementar de PO com a leitura principal do Smartsheet.

## 2. Gargalos encontrados

O primeiro gargalo estava no **waterfall pós-login**. Após o `/api/auth-login`, o frontend chamava `bootstrapSession()`, depois `loadProjects()`, depois `loadManualAlerts()`, depois `loadAlertResponses()`, depois `loadStageUpdates()` e, para administrador, `loadAdminData()`. Esse encadeamento fazia com que tarefas secundárias competissem com a exibição do painel e estendessem a sensação de tela carregando.

O segundo gargalo estava no **cache local subutilizado**. A função `loadProjects()` já lia `localStorage`, mas, quando o cache existia e estava vencido, ela renderizava os dados antigos e continuava aguardando a API antes de liberar o fluxo chamador. Na prática, o usuário podia ver parte da UI, mas a sequência pós-login permanecia bloqueada. Além disso, o TTL local era de 2 minutos, abaixo da janela aceitável para o Portal do Cliente informada no contexto.

O terceiro gargalo estava na **hidratação de sessão cliente**. O backend já possuía cache global de projetos, mas a função `hydrateClientSession()` podia chamar o Supabase a cada `/api/projects` para reidratar `clientKey`, `clientName`, `allowedClients` e metadados do cliente. Essa consulta era especialmente sensível para o usuário cliente `PRIO`, porque ocorria antes da filtragem e montagem do payload do portal.

O quarto gargalo estava na **rota `/api/auth-me`**. Em sessões persistidas, o frontend aguardava `bootstrapSession()` antes de iniciar `loadProjects()`, e `auth-me.js` também podia consultar o Supabase a cada chamada. Esse comportamento atrasava a primeira renderização em reloads, mesmo quando havia cache local válido.

O quinto gargalo estava no **backend sem cache quente**. Em cache miss ou atualização forçada, `projects.js` buscava primeiro a base complementar de PO e depois a sheet principal. Essas duas leituras externas eram independentes e podiam ser executadas em paralelo.

## 3. Alterações implementadas

| Área | Antes | Depois | Arquivos |
|---|---|---|---|
| Cache local de projetos | Cache válido retornava rápido, mas cache vencido podia bloquear o fluxo enquanto aguardava a API. TTL de 2 minutos. | Cache aproveitável renderiza imediatamente; a API revalida em background. TTL local ajustado para 5 minutos. | `site/app.js` |
| Fluxo pós-login | `/api/auth-me`, `/api/projects`, alertas, respostas, apontamentos e admin eram executados em série. | Reidratação de sessão e projetos começam em paralelo; tarefas auxiliares rodam em background via `Promise.allSettled`. | `site/app.js` |
| Feedback visual | O usuário podia perceber tela vazia enquanto a primeira chamada ainda não havia retornado. | Foi criado estado de carregamento imediato para tabela, detalhe e metadados quando ainda não há projetos em memória. | `site/app.js` |
| Hidratação cliente em `/api/projects` | `hydrateClientSession()` podia consultar o Supabase a cada requisição do Portal do Cliente. | Cache global de hidratação por usuário, com TTL padrão de 5 minutos e fallback para a última hidratação válida em caso de instabilidade. | `netlify/functions/projects.js` |
| Cache backend de projetos | Cache rápido padrão de 45 segundos. | Cache rápido padrão ampliado para 120 segundos, mantendo `force=1` para atualização manual. | `netlify/functions/projects.js` |
| Leitura Smartsheet sem cache | Base complementar de PO e sheet principal eram lidas em sequência. | Leitura complementar de PO e sheet principal executam em paralelo com `Promise.all`. | `netlify/functions/projects.js` |
| Sessão persistida | `/api/auth-me` podia bater no Supabase em todo reload antes de liberar o painel. | Cache curto por `sub` também em `auth-me.js`, com TTL padrão de 5 minutos. | `netlify/functions/auth-me.js` |
| Service Worker | Cache de assets versão v24 e interceptação genérica de URLs contendo `/api/`. | Cache de assets versão v25 para publicar `app.js` novo; interceptação de API restrita à própria origem e continua sem cache de dados. | `site/sw.js` |

## 4. Pontos técnicos alterados

A principal mudança no frontend está em `loadProjects()`. A função agora diferencia o carregamento interativo do polling em background. No login ou reload, se existir cache local não rejeitado, `applyProjectsPayload()` renderiza imediatamente e `revalidateProjectsInBackground()` agenda a busca fresca sem bloquear o fluxo. Em polling, cache fresco evita tráfego; cache vencido prossegue para a API, preservando a atualização periódica.

Também foram adicionados `setProjectsLoadingState()` e `startPostSessionBackgroundLoads()`. O primeiro evita tela em branco quando ainda não há dados em memória; o segundo concentra as cargas não críticas, como alertas, respostas, push subscription, apontamentos e dados administrativos, executando-as em paralelo e sem segurar a primeira exibição dos projetos.

No backend, `projects.js` ganhou `SESSION_HYDRATION_CACHE_MS`, `sessionHydrationCache`, `mergeHydratedClientSession()` e poda simples do cache. Assim, o Portal do Cliente continua protegido e compatível com sessões antigas, mas deixa de pagar o custo de Supabase a cada chamada de `/api/projects` dentro da janela de cache. Em caso de falha do Supabase, a última hidratação válida pode ser reutilizada, evitando bloquear o painel.

Em `auth-me.js`, foi aplicado cache curto semelhante para reduzir a latência do `bootstrapSession()` em reloads e visitas subsequentes. Isso é importante porque, quando há sessão persistida, o frontend precisa saber quem é o usuário antes de escolher a chave correta do cache local de projetos.

## 5. Validações executadas

| Validação | Resultado |
|---|---|
| `node --check site/app.js` | OK |
| `node --check netlify/functions/projects.js` | OK |
| `node --check netlify/functions/auth-me.js` | OK |
| `node --check site/sw.js` | OK |
| `npm run validate` | Executado sem erro reportado pelo shell |

Não foi possível executar uma medição real com as credenciais em ambiente Netlify/produção dentro do sandbox, porque isso dependeria das variáveis remotas, cookies e APIs externas efetivamente configuradas. Ainda assim, as validações sintáticas confirmam que os arquivos modificados são parseáveis pelo Node.js.

## 6. Estimativa de melhoria esperada

| Cenário | Antes provável | Depois esperado | Justificativa |
|---|---:|---:|---|
| Login com cache local válido | 2 s a 8 s percebidos, dependendo de Supabase, Smartsheet e alertas auxiliares | Menos de 1 s percebido | O payload local é aplicado imediatamente e a API roda em background. |
| Login com cache local vencido, mas aproveitável | 3 s a 10 s de fluxo bloqueado | Menos de 1 s para exibir dados antigos; atualização fresca em background | O cache vencido não é descartado se não for inválido; ele é usado como primeira pintura. |
| Login sem cache local, mas backend com cache quente | 1,5 s a 5 s | Aproximadamente 1 s a 3 s | `/api/projects` evita Supabase repetido por usuário cliente e usa cache backend mais longo. |
| Login sem cache local e sem cache backend | 4 s a 12 s, conforme Smartsheet | Aproximadamente 3 s a 8 s | Leituras independentes do Smartsheet/PO foram paralelizadas; ainda depende da API externa. |
| Reload com sessão persistida e cache local válido | Bloqueado por `/api/auth-me` antes de usar cache | Normalmente abaixo de 1 s após primeiro reload quente | `/api/auth-me` usa cache curto de Supabase e o frontend aplica cache logo após autenticar a sessão. |

A melhoria mais perceptível ocorrerá no **Portal do Cliente PRIO**, porque a experiência passa a ser orientada por cache local imediato e porque a reidratação cliente no backend deixa de consultar o Supabase em todas as chamadas de projetos dentro da janela de 5 minutos.

## 7. Recomendações pós-deploy

Após publicar a versão corrigida, recomendo abrir o DevTools em uma sessão anônima e registrar duas medições para cada usuário (`douglas@pcp` e `PRIO`): primeiro login sem cache local e segundo login/reload com cache local. A métrica principal deve ser o tempo entre o retorno de `/api/auth-login` e a primeira renderização de linhas/cards de projeto. Também é importante confirmar que o service worker ativo passou para `step-gerencia-pwa-v25-performance-cache-first`, pois isso garante que o navegador buscou o `app.js` novo.

Se for necessário reduzir ainda mais o tempo de primeiro login sem cache, o próximo candidato seria aplicar cache curto semelhante em `sector-alerts.js`, porque alertas e respostas ainda podem reidratar sessão e chamar Supabase; porém, como essas cargas agora estão fora do caminho crítico, elas não devem impedir a exibição inicial dos projetos.
