-- v37.07 - Cache persistente para acelerar login e proteger o Portal do Cliente
-- Execute no SQL Editor do Supabase antes/depois do deploy. Se a tabela não existir,
-- o sistema continua funcionando pelo Smartsheet, mas sem o ganho definitivo de cache persistente.

create table if not exists public.step_tracking_cache (
  cache_key text primary key,
  scope text,
  source text,
  version text,
  projects_count integer default 0,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

create index if not exists idx_step_tracking_cache_updated_at
  on public.step_tracking_cache (updated_at desc);

alter table public.step_tracking_cache enable row level security;

-- Segurança: o cache operacional deve ser acessado somente pelo backend com SERVICE_ROLE_KEY.
do $$
begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'step_tracking_cache'
      and policyname = 'step_tracking_cache_no_public_access'
  ) then
    create policy step_tracking_cache_no_public_access
      on public.step_tracking_cache
      for all
      using (false)
      with check (false);
  end if;
end $$;
