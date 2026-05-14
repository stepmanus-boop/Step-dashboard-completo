v37.5 - LOGIN POR AMBIENTE

Correção:
- Agora o mesmo login pode existir no Brasil e em Portugal.
- A verificação de duplicidade usa username + operationRegion.

Exemplo:
Login: sbm / BR / clientKey SBM_BR
Login: sbm / PT / clientKey SBM_PT

Importante:
Execute o SQL:
supabase_v37_5_login_por_ambiente.sql

Isso remove a unicidade global antiga de username e cria uma unicidade por ambiente.
