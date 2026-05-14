VERSÃO v37.0 - MULTIAMBIENTE BRASIL / PORTUGAL

Planilhas configuradas:
BR Tracking: Progress Tracking Sheet - Piping Fabrication
BR WIP: Work in Progress - STEP

PT Tracking: Progress Tracking Sheet - Piping Fabrication PT
PT WIP: WORK-IN-PROGRESS -PT

Como funciona:
1. O sistema continua usando o mesmo Supabase.
2. Cada usuário pode ter Ambiente Operacional: Brasil ou Portugal.
3. O frontend envia ?region=BR ou ?region=PT para /api/projects.
4. A função projects.js resolve as planilhas pelo nome usando a mesma API Smartsheet.
5. Se quiser fixar ambiente pelo Netlify, use:
   SMARTSHEET_DEFAULT_REGION=PT

Variáveis opcionais para travar por ID:
SMARTSHEET_SHEET_ID_PT
SMARTSHEET_WIP_STEP_SHEET_ID_PT
SMARTSHEET_SHEET_ID_BR
SMARTSHEET_WIP_STEP_SHEET_ID_BR

Nomes padrão Portugal:
Progress Tracking Sheet - Piping Fabrication PT
WORK-IN-PROGRESS -PT
