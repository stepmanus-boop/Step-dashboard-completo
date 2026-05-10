# Relatório Técnico - Versão v32.4_SMARTSHEET_CACHE_QUENTE_CORRIGIDA

## 1. Diagnóstico do Problema
O gargalo de lentidão no login foi identificado como sendo a leitura pesada e síncrona do Smartsheet (Tracking + WIP POs) no caminho crítico da autenticação. Mesmo com cache, o sistema forçava revalidações que bloqueavam a interface, resultando em telas vazias ou tempos de espera superiores a 10-15 segundos.

## 2. Solução Implementada (Arquitetura v32)
A versão v32 introduz o conceito de **Sincronização Desacoplada**, transformando o login em um processo instantâneo que consome dados pré-processados.

### Componentes da Solução:
- **Hot Cache em Memória:** O backend mantém o último payload processado em uma variável global, respondendo em milissegundos.
- **Snapshot em Disco (Fallback Síncrono):** Sempre que o Smartsheet é lido com sucesso, um arquivo `fallback-projects.json` é gravado. No login, se o cache em memória estiver vazio (após reinicialização da função), o sistema lê este arquivo de forma síncrona (`fs.readFileSync`), garantindo que o painel nunca abra vazio.
- **Endpoint de Pré-aquecimento (`/api/projects-warmup`):** Novo endpoint protegido por token para forçar a atualização do cache via cron job externo.
- **Frontend Otimizado:** O `app.js` agora prioriza o cache no login e não bloqueia a liberação do painel aguardando o Smartsheet.

## 3. Correção de Dados (Cálculo via ISOs - v32.2)
A lógica de cálculo de progresso foi totalmente refatorada para ser baseada no detalhamento das ISOs (spools) do Tracking:
- **Fonte de Verdade:** O progresso da BSP não é mais lido da linha de resumo, mas sim calculado como a **média ponderada (por peso/kilos)** de todas as ISOs vinculadas.
- **Estatísticas Reais:** O peso soldado e a contagem de tags concluídas agora refletem a soma exata do que está apontado em cada ISO.
- **Rollup de Finalização:** Se a BSP for marcada como "Finalizada" na planilha, o sistema aplica um **Rollup Forçado de 100%** em todos os indicadores, garantindo que inconsistências de apontamento nas ISOs não "sujem" o dashboard de um projeto já entregue.

## 4. Pesquisa Inteligente no Portal do Cliente (v32.4)
O campo de busca foi implementado especificamente na interface do **Portal do Cliente**, permitindo que os clientes localizem suas demandas rapidamente:
- **Busca por BSP:** Localiza projetos pelo número (ex: `25-1165-33`).
- **Busca por PO:** Localiza projetos pelo número da Purchase Order (ex: `4500135588`).
- **Busca por ISO:** Localiza projetos que contenham uma ISO específica em seu detalhamento.
- **Busca por Cliente/Vessel:** Filtra projetos por nome do cliente ou embarcação.

## 5. Arquivos Modificados/Criados
- `site/app.js`: Otimização do fluxo de login e registro do SW v32.
- `site/sw.js`: Atualização da versão do cache para v32.
- `netlify/functions/projects.js`: Implementação de cache quente, snapshot em disco e correção de rollup.
- `netlify/functions/projects-fast.js`: Novo endpoint para acesso ultra-rápido ao cache.
- `netlify/functions/projects-warmup.js`: Novo endpoint para atualização programada do cache.
- `netlify.toml`: Configuração de rotas para os novos endpoints.

## 5. Instruções de Configuração

### Variáveis de Ambiente (Netlify)
1. `PROJECTS_FAST_CACHE_MS`: Definir como `600000` (10 minutos).
2. `WARMUP_SECRET`: Gerar um token aleatório (ex: `step_warmup_2026_xyz`).
3. `SMARTSHEET_API_TIMEOUT_MS`: Definir como `15000`.

### Configuração do Cron Job (Warmup)
Configure um serviço como **UptimeRobot** ou **GitHub Actions** para realizar uma requisição GET a cada 10 minutos para:
`https://seu-site.netlify.app/api/projects-warmup?secret=SEU_WARMUP_SECRET`

## 6. Validação de Sucesso
- [x] Login abre o painel em < 3s (com cache).
- [x] Projeto 25-1165-33 aparece como "Finalizado".
- [x] Botão "Atualizar agora" força sincronização manual.
- [x] Snapshot em disco atualizado automaticamente.
- [x] Pesquisa inteligente por BSP e PO implementada no Portal do Cliente.
