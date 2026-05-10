# Relatório v27 — Barra de progresso no carregamento pós-login

## Objetivo da alteração

Esta versão adiciona uma **tela de carregamento progressivo após o login**, com barra percentual, mensagens por etapa e abertura do painel somente depois que o fluxo visual chega a **100%**. A intenção é evitar a percepção de tela parada durante a busca dos dados e deixar explícito para o usuário que o sistema está carregando BSPs, POs, indicadores e dashboards.

## Arquivos modificados

| Arquivo | Alteração realizada |
|---|---|
| `site/index.html` | Inserido o overlay global `login-progress-overlay`, com título, mensagem, barra, percentual e texto de detalhe. |
| `site/app.css` | Adicionados os estilos da tela de progresso, cartão central, barra animada e responsividade mobile. |
| `site/app.js` | Integrado o progresso ao `handleLoginSubmit`, iniciando na validação de acesso e avançando durante autenticação, sessão, `loadProjects` e renderização. |
| `site/sw.js` | Atualizado para **v27** para forçar o navegador a baixar os assets novos após o deploy. |

## Etapas exibidas ao usuário

| Progresso | Frase principal | Finalidade |
|---:|---|---|
| 6%–20% | **Validando acesso...** | Mostra que as credenciais estão sendo conferidas. |
| 34% | **Carregando BSPs...** | Indica que os projetos/BSPs estão sendo organizados por cliente e unidade. |
| 55% | **Carregando POs...** | Indica preparação de POs, demandas e referências de fabricação. |
| 72% | **Atualizando indicadores...** | Indica cálculo de status, pesos, alertas e pendências. |
| 88% | **Definindo dashboards...** | Indica montagem visual final antes de liberar o painel. |
| 100% | **Tudo pronto.** | Confirma que os dados terminaram de carregar e abre o dashboard. |

## Garantia funcional aplicada

O painel só fecha a tela de progresso quando a chamada `loadProjects({ preferServerCache: true })` termina. Isso significa que o overlay não é apenas decorativo: ele acompanha o ciclo real de carregamento dos projetos e só chega ao fechamento final depois que os dados foram recebidos/aplicados pelo frontend.

## Como testar

Depois do deploy, recomenda-se limpar ou atualizar o service worker se o navegador ainda estiver servindo versão antiga. Em seguida, faça login com `PRIO` e com `douglas@pcp`. A tela progressiva deve aparecer imediatamente após clicar em entrar, exibir as mensagens de BSPs, POs e dashboards, chegar a **100%** e então liberar a visualização dos dados.

## Validação

A sintaxe dos arquivos JavaScript alterados foi validada com `node --check`, e o ZIP final foi testado com `unzip -t`.
