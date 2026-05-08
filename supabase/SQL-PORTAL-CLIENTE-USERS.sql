-- Portal do Cliente - campos adicionais em users
-- Rode este SQL no Supabase antes de criar usuários do tipo client.

alter table if exists public.users
  add column if not exists client_key text,
  add column if not exists client_name text,
  add column if not exists client_logo_url text,
  add column if not exists client_platform_image_url text,
  add column if not exists client_platform_images jsonb default '{}'::jsonb,
  add column if not exists allowed_clients text[] default '{}';

-- Ajusta a constraint do campo role para aceitar o perfil client.
do $$
declare c record;
begin
  for c in
    select conname
    from pg_constraint
    where conrelid = 'public.users'::regclass
      and contype = 'c'
      and pg_get_constraintdef(oid) ilike '%role%'
  loop
    execute format('alter table public.users drop constraint if exists %I', c.conname);
  end loop;
end $$;

alter table public.users
  add constraint users_role_check check (role in ('admin', 'sector', 'client'));
