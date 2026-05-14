v37.7 - Correção do login duplicado em cadastros legados

Problema:
- Usuários antigos do Brasil estavam sem operation_region.
- O sistema interpretava como PT e bloqueava o cadastro SBM Portugal.

Correção:
- Cadastros antigos sem região passam a ser tratados como BR.
- Admin e PCP continuam GLOBAL.
- Demais usuários são únicos por username + operation_region.

Execute:
supabase_v37_7_corrige_login_legado.sql
