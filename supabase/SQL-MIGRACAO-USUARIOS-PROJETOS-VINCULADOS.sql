-- Permite vincular usuários do setor Projetos para ampliar a visão em "Meus projetos".
-- Exemplo: usuário Rodrigo vê as próprias BSPs + BSPs do Álvaro e Thales.

alter table public.users
  add column if not exists supervised_users jsonb not null default '[]'::jsonb;

comment on column public.users.supervised_users is
  'Lista de usuários de Projetos vinculados ao usuário. Usado para filtrar Meus projetos por PM.';
