v37.6 - ADMIN E PCP UNIVERSAIS

Regra implementada:
- Administrador: universal.
- PCP: universal.
- Demais usuários: separados por país/ambiente/unidade.

Exemplo:
- admin pode entrar no site BR e PT.
- pcp pode entrar no site BR e PT.
- login sbm pode existir em BR e PT, separado por operationRegion.
- usuário de Pintura/Solda/Cliente fica preso ao ambiente cadastrado.

Execute no Supabase:
supabase_v37_6_admin_pcp_universal.sql
