v38.3 - Login oculto por região

Implementação:
- Cliente pode usar o mesmo login visível em BR e PT.
- Quando houver duplicidade, o sistema salva internamente:
  SBM__BR
  SBM__PT
- No painel e no login, o usuário continua vendo/digitando:
  SBM

Motivo:
- O painel BR e PT são links separados.
- Cada site envia seu ambiente ao backend.
- O backend resolve automaticamente qual cadastro usar.

Execute:
supabase_v38_3_login_oculto_br_pt.sql
