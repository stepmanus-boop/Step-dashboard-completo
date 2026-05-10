# Relatório v28 — Trava para liberar o painel somente com dados visíveis

A versão v28 corrige o comportamento observado em que a barra de progresso chegava a 100%, mas o Portal do Cliente era aberto ainda com cartões marcados como `--`. O problema não era apenas o carregamento visual; faltava uma validação explícita de que o payload havia sido aplicado ao estado e de que os elementos do painel já estavam renderizados na tela.

## O que foi corrigido

| Ponto | Antes | Depois |
|---|---|---|
| Conclusão da barra | A barra fechava após `loadProjects` terminar, mesmo que a tela ainda estivesse com indicadores `--`. | A barra só chega a 100% depois que `hasDashboardDataReady()` confirma projetos no estado e dados renderizados no painel. |
| Portal do Cliente | Podia abrir com cards `BSPs`, `Tags`, `Peso` e `Progresso` ainda vazios. | Para cliente, a liberação exige BSPs diferentes de `--` e `0`, além de cards de unidade/vessel visíveis. |
| Requisição de login | `loadProjects` podia resolver sem garantir payload operacional não vazio. | No fluxo de login, `loadProjects({ requireData: true })` rejeita payload vazio e mantém o carregamento/retentativa. |
| Retentativa | Não havia etapa dedicada para conferir se a renderização realmente aconteceu. | Foi adicionada a etapa “Conferindo dados na tela...”, com novas tentativas de carregamento antes de liberar a interface. |
| Cache do navegador | O service worker anterior poderia manter arquivos antigos. | O service worker foi atualizado para v28 para forçar a troca do `app.js`. |

## Como testar

Após publicar o pacote v28, limpe ou atualize o navegador uma vez para garantir que o service worker novo seja instalado. Em seguida, faça login com o usuário PRIO. A barra deve passar pelas etapas de BSPs, POs e dashboards, mas agora ela não deve fechar enquanto os cards continuarem com `--`. Se a API ainda não tiver retornado dados válidos, a tela continuará em carregamento e exibirá a etapa de conferência dos dados.

O comportamento esperado é que, ao chegar em 100%, o Portal do Cliente já esteja com BSPs, tags, pesos, progresso e unidades preenchidos. Se os dados não forem recebidos, o painel não será liberado vazio.
