-- Migração para liberar perfil Supervisor de Projetos
-- Execute uma vez no SQL Editor do Supabase antes de cadastrar/editar supervisores.

alter table if exists public.users
  add column if not exists supervised_users jsonb not null default '[]'::jsonb;

comment on column public.users.supervised_users is
  'Lista JSON de usuários de Projetos supervisionados: [{"id":"...","name":"...","username":"..."}]';

-- Caso exista algum CHECK antigo restringindo role apenas a admin/sector, remova o check manualmente
-- e recrie com supervisor. Como o nome do constraint pode variar, o bloco abaixo tenta localizar automaticamente.
do $$
declare
  constraint_name text;
begin
  select con.conname into constraint_name
  from pg_constraint con
  join pg_class rel on rel.oid = con.conrelid
  join pg_namespace nsp on nsp.oid = rel.relnamespace
  where nsp.nspname = 'public'
    and rel.relname = 'users'
    and con.contype = 'c'
    and pg_get_constraintdef(con.oid) ilike '%role%'
    and pg_get_constraintdef(con.oid) ilike '%admin%'
    and pg_get_constraintdef(con.oid) ilike '%sector%'
  limit 1;

  if constraint_name is not null then
    execute format('alter table public.users drop constraint %I', constraint_name);
  end if;
end $$;

alter table if exists public.users
  add constraint users_role_check
  check (role in ('admin', 'sector', 'supervisor'));
