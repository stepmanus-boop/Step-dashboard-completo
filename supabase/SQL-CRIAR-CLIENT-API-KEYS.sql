-- STEP | Portal do Cliente - Chaves de API internas
-- Execute este SQL no Supabase antes de usar o botão "Gerar API" no painel do cliente.

create extension if not exists pgcrypto;

create table if not exists public.client_api_keys (
  id uuid primary key default gen_random_uuid(),
  user_id text not null,
  username text default '',
  client_key text default '',
  client_name text default '',
  allowed_clients jsonb default '[]'::jsonb,
  token_hash text not null unique,
  token_prefix text default '',
  token_last4 text default '',
  name text default 'API do cliente',
  scopes jsonb default '["read:projects"]'::jsonb,
  active boolean default true,
  expires_at timestamptz null,
  last_used_at timestamptz null,
  created_by text default '',
  created_by_name text default '',
  created_at timestamptz default now(),
  revoked_at timestamptz null
);

create index if not exists idx_client_api_keys_user_id
on public.client_api_keys(user_id);

create index if not exists idx_client_api_keys_token_hash
on public.client_api_keys(token_hash)
where active = true;

create index if not exists idx_client_api_keys_client_key
on public.client_api_keys(client_key);

-- Blindagem de permissão: todas as chaves desta tabela são SOMENTE LEITURA.
-- Mesmo que alguém tente gravar outro escopo por engano, o banco força read:projects.
alter table public.client_api_keys
add column if not exists access_mode text default 'read_only';

update public.client_api_keys
set scopes = '["read:projects"]'::jsonb,
    access_mode = 'read_only'
where scopes is distinct from '["read:projects"]'::jsonb
   or access_mode is distinct from 'read_only';

do $$
begin
  if not exists (
    select 1 from pg_constraint where conname = 'client_api_keys_read_only_scope_chk'
  ) then
    alter table public.client_api_keys
    add constraint client_api_keys_read_only_scope_chk
    check (
      access_mode = 'read_only'
      and jsonb_typeof(scopes) = 'array'
      and scopes = '["read:projects"]'::jsonb
    );
  end if;
end $$;

alter table public.client_api_keys enable row level security;

-- A aplicação usa SERVICE_ROLE_KEY nas Netlify Functions. Esta policy deixa a tabela fechada para anon/auth direto.
-- user_id fica como text para ser compatível tanto com IDs locais (ex.: u_admin_001) quanto UUID do Supabase.
drop policy if exists "client_api_keys_no_direct_select" on public.client_api_keys;
create policy "client_api_keys_no_direct_select"
on public.client_api_keys
for select
using (false);

notify pgrst, 'reload schema';
