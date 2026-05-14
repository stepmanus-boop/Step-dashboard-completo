VERSÃO v37.1 - PORTUGAL ONLY

Esta versão foi ajustada para carregar somente Portugal.

Planilhas usadas:
Tracking PT: Progress Tracking Sheet - Piping Fabrication PT
WIP PT: WORK-IN-PROGRESS -PT

O Supabase continua sendo o mesmo, centralizado.

O sistema força region=PT no frontend e no backend, evitando consulta das planilhas do Brasil.

Variáveis opcionais para melhor performance:
SMARTSHEET_SHEET_ID_PT
SMARTSHEET_WIP_STEP_SHEET_ID_PT

Caso os IDs sejam configurados no Netlify, o sistema não precisa procurar pelo nome da planilha e carrega mais rápido.
