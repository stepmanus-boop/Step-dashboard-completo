v37.3 - SITE PORTUGAL COM MESMO SUPABASE

Objetivo:
- O site de Portugal usa o mesmo banco Supabase.
- O cadastro fica separado por operation_region/site_key.
- Clientes com o mesmo nome comercial podem ter cadastros separados:
  SBM_BR
  SBM_PT

O que foi adicionado:
- users.operation_region
- users.site_key
- clientKey continua sendo o identificador operacional do cliente.

No site Portugal:
- operationRegion padrão = PT
- siteKey padrão = PT
- clientKey sugerido: NOME_PT

Execute primeiro:
supabase_portugal_site_v37_3.sql

Depois publique o ZIP normalmente.
