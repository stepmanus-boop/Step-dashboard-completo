-- Permite que usuários do setor Projetos visualizem BSPs de PMs adicionais.
-- Rode este script no Supabase SQL Editor antes de publicar o código alterado.

alter table public.users
add column if not exists project_pm_aliases text[] not null default '{}';

comment on column public.users.project_pm_aliases is
'Lista de nomes de PMs adicionais, conforme coluna PM do Tracking, que o usuário de Projetos pode visualizar em Meus projetos.';

-- Opcional, mas ajuda em consultas futuras se essa coluna for filtrada pelo banco.
create index if not exists users_project_pm_aliases_gin_idx
on public.users using gin (project_pm_aliases);
