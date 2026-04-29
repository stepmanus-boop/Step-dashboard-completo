# STEP Progress Tracking Dashboard

Painel online para TV corporativa, preparado para **GitHub + Netlify** com leitura da sheet do **Smartsheet** por **Netlify Functions**.

## O que esse projeto já faz

- lê a planilha do Smartsheet pelo backend
- separa **linha principal do projeto** e **linhas dos spools/ISOs**
- usa o **total oficial do projeto** na linha principal
- expande os detalhes automaticamente no painel
- faz **rotação automática** dos projetos
- pagina automaticamente os spools do projeto em destaque
- mostra:
  - projeto
  - quantidade de spools
  - peso total em kg
  - área de painting em m²
  - etapa atual
  - % individual
  - % geral
  - status
  - datas fixas do processo
- mostra relógio do **Brasil** e de **Portugal**
- pronto para subir no **GitHub** e sincronizar com **Netlify**

## Estrutura

```text
site/
  index.html
  app.css
  app.js
  assets/step-logo.png

netlify/
  functions/
    projects.js

scripts/
  validate-config.js

.env.example
.gitignore
netlify.toml
package.json
```

## Segurança

A token do Smartsheet **não foi embutida no projeto**.

Use variáveis de ambiente no Netlify:

- `SMARTSHEET_TOKEN`
- `SMARTSHEET_SHEET_NAME`
- `SMARTSHEET_SHEET_ID` (opcional, mas recomendado)
- `SMARTSHEET_API_BASE` (opcional)

## Como subir no GitHub

1. Extraia esta pasta.
2. Entre nela no terminal.
3. Rode:
   ```bash
   git init
   git add .
   git commit -m "Initial STEP progress dashboard"
   ```
4. Crie o repositório no GitHub.
5. Conecte o remoto:
   ```bash
   git remote add origin SEU_REPOSITORIO_GITHUB
   git branch -M main
   git push -u origin main
   ```

## Como sincronizar com o Netlify

1. No Netlify, clique em **Add new site**.
2. Escolha **Import from Git**.
3. Conecte ao seu repositório do GitHub.
4. Configure as variáveis de ambiente:
   - `SMARTSHEET_TOKEN`
   - `SMARTSHEET_SHEET_NAME=Progress Tracking Sheet - Piping Fabrication`
   - `SMARTSHEET_SHEET_ID` se já souber o ID da sheet
5. Deploy.

O `netlify.toml` já está configurado com:

- `publish = "site"`
- `functions = "netlify/functions"`

## Desenvolvimento local

Instale dependências:

```bash
npm install
```

Rode em modo local Netlify:

```bash
npm run dev
```

## Regras de leitura já aplicadas

- o projeto é identificado pelo **campo Project**
- a leitura interna usa a **numeração do projeto**
- o prefixo exibido respeita o que vier na base, como:
  - `BSP`
  - `BEP`
  - `BEB`
  - outros
- a linha principal fornece o total oficial:
  - `Quantity Spools`
  - `Kilos`
  - `M2 Painting`
- as linhas filhas são usadas para:
  - ISO
  - descrição
  - peso individual
  - painting individual
  - etapa atual por spool

## Lógica de progresso

A etapa atual segue a ordem que você definiu, inclusive com:

- `Initial Dimensional Inspection/3D` antes de `Full welding execution`
- campos de data fixados quando existirem
- campos opcionais de paint ignorados se vierem vazios
- etapa em andamento destacada automaticamente

## Observação importante

Eu deixei o projeto preparado para o **Modo 1**:
- polling pelo front-end
- leitura por Netlify Function
- cache em memória da function quando a instância estiver quente
- atualização automática na tela

## Próximo passo recomendado

Depois de colocar sua nova token no Netlify, faça um primeiro deploy e confira se o nome da sheet e o `sheetId` estão corretos.


## Módulo de login por setor

O projeto agora inclui:

- login simples por usuário e senha
- perfil `admin` e perfil `sector`
- criação de usuários por setor
- alerta operacional por setor
- confirmação de leitura dos alertas
- persistência automática em arquivos JSON versionados no repositório

### Login inicial
- usuário: `admin`
- senha: `admin123`

### Arquivos usados
- `data/users.json`
- `data/manual-alerts.json`
- `data/alert-acks.json`

### Variáveis de ambiente necessárias para gravar direto no GitHub
- `SESSION_SECRET`
- `GITHUB_REPO`
- `GITHUB_BRANCH`
- `GITHUB_TOKEN`

Sem as variáveis do GitHub, o login funciona, mas a gravação persistente em produção depende do token do repositório.


## PWA e mobile

Esta versão inclui:
- instalação como app no Chrome/Edge e Android
- suporte a tela inicial no iPhone/iPad via Safari
- manifest e service worker
- ícones para app
- ajustes de layout responsivo para mobile

### Instalação
- Desktop/Android: use o botão **Instalar app** ou o menu do navegador.
- iPhone/iPad: abra no Safari e use **Adicionar à Tela de Início**.


## Ajustes desta versão

- visualização pública liberada sem login
- botão de login opcional para acesso setorial ou admin
- admin pode promover outros usuários para admin
- leitura dos arquivos locais (`data/*.json`) corrigida para deploy empacotado
- API real do Smartsheet fica dentro do projeto em `netlify/functions/projects.js`
- removida a token de teste embutida; agora a integração usa apenas `SMARTSHEET_TOKEN` configurado no ambiente

### Importante sobre a API
O projeto já vem com a API local integrada no próprio repositório:
- `GET /api/projects` → leitura dos projetos
- `POST /api/auth-login` → login
- `GET /api/auth-me` → sessão
- `GET/POST/PATCH /api/sector-alerts` → alertas por setor
- `GET/POST/PATCH /api/admin-users` → usuários e promoção para admin

Para usar a base real, configure no Netlify:
- `SMARTSHEET_TOKEN`
- `SMARTSHEET_SHEET_NAME`
- opcionalmente `SMARTSHEET_SHEET_ID`

## Configuração rápida
Veja o arquivo `GUIA-CONFIGURACAO-GITHUB-NETLIFY.txt` para ativar persistência via GitHub e API real do Smartsheet.

## Validação PCP integrada ao Smartsheet/Tracking

Esta versão inclui o fluxo de validação PCP com atualização direta do Smartsheet pela Netlify Function `/api/stage-updates`.

Principais pontos implementados:

- O botão **Validação PCP** abre a tela em nova aba usando `?stageWorkspace=1#stage-validation`.
- O PCP/Admin pode atualizar ou regravar o Tracking individualmente ou em lote.
- Percentuais são enviados ao Smartsheet como número decimal: `0.25`, `0.5`, `0.75` e `1`.
- O sistema não retrocede avanço maior já existente no Tracking.
- Apontamentos de 25%, 50% e 75% atualizam somente percentual.
- Apontamentos de 100% atualizam percentual e a data correspondente.
- Pintura 100% também alimenta `Final Inspection` e `Package and Delivered` com 25%, sem reduzir valores maiores.
- Duplicados de spool/drawing são atualizados juntos dentro do mesmo projeto/BSP.
- A área **Pendências de datas do histórico** confere somente apontamentos 100% registrados pelo app e já validados.

Antes do deploy, confirme no Netlify as variáveis:

```env
SMARTSHEET_TOKEN=...
SMARTSHEET_SHEET_ID=...
SUPABASE_URL=...
SUPABASE_SERVICE_ROLE_KEY=...
```

Se a tabela `stage_updates` já existir no Supabase, execute também:

```text
supabase/SQL-MIGRACAO-STAGE-UPDATES-VALIDACAO-PCP.sql
```


## Ajuste de compatibilidade Smartsheet

A função de Validação PCP agora usa a mesma configuração Smartsheet já existente no projeto, além de aceitar variáveis de ambiente (SMARTSHEET_TOKEN, SMARTSHEET_ACCESS_TOKEN, SMARTSHEET_API_TOKEN, SMARTSHEET_BEARER_TOKEN, SMARTSHEET_PAT, SMARTSHEET_PERSONAL_ACCESS_TOKEN). Isso evita erro de token ausente quando o dashboard já funcionava com a configuração interna do ZIP.
