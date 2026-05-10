# Ajuste adicional v26 — redução da demora pós-login

**Autor:** Manus AI  
**Data:** 2026-05-09  
**Projeto:** `step_dashboard_tracking_CLIENTE_CORRIGIDO_v25_PERFORMANCE`

## Diagnóstico após o retorno do teste

O retorno do teste mostrou que a exibição dos dados não estava quebrada; os dados apareciam, porém a espera continuava alta. Esse comportamento é compatível com dois gargalos remanescentes. Primeiro, o logout apagava o cache local de projetos, o que fazia o login seguinte depender novamente de chamadas remotas ao Supabase e ao Smartsheet. Segundo, mesmo com cache global em memória no backend, a rota `/api/projects` ainda podia fazer checagem de versão antes de devolver dados quando a janela curta expirava.

## Correções adicionais aplicadas

| Arquivo | Mudança | Efeito esperado |
|---|---|---|
| `site/app.js` | O logout deixou de apagar o cache local de projetos. | O próximo login do mesmo usuário pode reutilizar o cache escopado por papel, usuário e cliente, reduzindo a espera percebida. |
| `site/app.js` | A chamada principal de projetos após login passou a usar `/api/projects?preferCache=1`. | Quando o backend já tem payload em memória, ele responde sem bloquear em checagem de versão do Smartsheet. |
| `netlify/functions/projects.js` | `hydrateClientSession()` e `buildPayload()` agora rodam em paralelo no handler. | O Portal do Cliente não espera a hidratação Supabase terminar antes de iniciar a carga de projetos. |
| `netlify/functions/projects.js` | O cache rápido backend subiu para 5 minutos e ganhou modo `preferCache`. | Reabre/login fica mais rápido em instância Netlify quente, mantendo `force=1` para atualização manual. |
| `site/sw.js` | Versão do service worker atualizada para `v26`. | O navegador é forçado a baixar o `app.js` corrigido depois do deploy. |

## Observação importante para teste

Para verificar a melhoria real, é importante publicar esta versão e testar duas situações distintas. No primeiro acesso sem cache local e com função fria, ainda pode haver demora porque o sistema precisa buscar dados do Smartsheet. A melhora mais forte deve aparecer no segundo login, no reload da página ou em logins subsequentes, pois agora o cache local não é apagado no logout e o backend usa o caminho rápido quando já possui payload em memória.

## Validação

| Validação | Resultado |
|---|---|
| `node --check site/app.js` | OK |
| `node --check netlify/functions/projects.js` | OK |
| `node --check netlify/functions/auth-me.js` | OK |
| `node --check site/sw.js` | OK |
| `npm run validate` | Executado sem erro reportado pelo shell |

## Resultado esperado

Com cache local já existente, a renderização deve ocorrer em menos de 1 segundo após o login, porque os dados ficam preservados entre logout e novo login. Sem cache local, mas com cache backend quente, a resposta deve ser perceptivelmente mais rápida porque o handler paraleliza sessão e payload e pode devolver a última base em memória usando `preferCache=1`. Em primeira carga totalmente fria, a latência continuará dependente do Smartsheet, mas o app exibirá estado de carregamento imediatamente e reutilizará cache nas próximas visitas.
