-- Portal do Cliente - campos adicionais em public.users
-- Rode este SQL no Supabase para habilitar cliente, logo e fotos por plataforma.
-- Depois de executar, aguarde alguns segundos ou faça reload do schema cache da API.

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

-- Força o PostgREST/Supabase a recarregar o schema cache, evitando erro PGRST204.
notify pgrst, 'reload schema';
