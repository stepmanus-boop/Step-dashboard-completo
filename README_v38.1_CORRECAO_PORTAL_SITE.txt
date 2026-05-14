v38.1 - Correção portal_site BR/PT

Problema:
- Supabase ainda tinha portal_site='global'.
- Existia índice único username + portal_site.
- Isso bloqueava cadastro SBM PT quando SBM BR já existia.

Correção:
- portal_site agora segue BR/PT.
- Código grava portal_site junto com operation_region/site_key.
- SQL remove índice antigo global e cria regra correta.

Execute:
supabase_v38_1_corrige_portal_site_br_pt.sql
