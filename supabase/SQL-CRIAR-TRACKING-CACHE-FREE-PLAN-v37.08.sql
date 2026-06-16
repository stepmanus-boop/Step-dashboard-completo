-- v37.08 - Cache persistente seguro para Supabase Free
-- Objetivo: acelerar o login e evitar falhas do painel quando o Smartsheet estiver lento,
-- sem transformar o Supabase em histórico pesado.
--
-- Características desta versão:
-- 1) Mantém somente UMA linha de cache: projects:BR:current
-- 2) Remove caches antigos por cliente e snapshots extras
-- 3) Adiciona coluna payload_bytes para controle de tamanho
-- 4) Bloqueia acesso público por RLS; somente backend com SERVICE_ROLE_KEY acessa
-- 5) Não salva imagens, não cria histórico e não usa realtime

create table if not exists public.step_tracking_cache (
  cache_key text primary key,
  scope text,
  source text,
  version text,
  projects_count integer default 0,
  payload_bytes integer default 0,
  last_write_reason text,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.step_tracking_cache
  add column if not exists payload_bytes integer default 0;

alter table public.step_tracking_cache
  add column if not exists last_write_reason text;

alter table public.step_tracking_cache
  add column if not exists projects_count integer default 0;

-- Se você executou a v37.07 antes, ela podia ter criado cache full e/ou por cliente.
-- Esta migração preserva o cache full mais útil como linha única atual.
insert into public.step_tracking_cache (
  cache_key,
  scope,
  source,
  version,
  projects_count,
  payload_bytes,
  last_write_reason,
  payload,
  updated_at
)
select
  'projects:BR:current' as cache_key,
  'single-current-cache' as scope,
  coalesce(source, 'migration-v37.08') as source,
  version,
  coalesce(projects_count, 0) as projects_count,
  coalesce(nullif(payload_bytes, 0), length(payload::text)) as payload_bytes,
  'migration-from-v37.07' as last_write_reason,
  payload,
  updated_at
from public.step_tracking_cache
where cache_key in ('projects:BR:full', 'projects:BR:current')
order by
  case when cache_key = 'projects:BR:current' then 0 else 1 end,
  updated_at desc
limit 1
on conflict (cache_key) do update set
  scope = excluded.scope,
  source = excluded.source,
  version = excluded.version,
  projects_count = excluded.projects_count,
  payload_bytes = excluded.payload_bytes,
  last_write_reason = excluded.last_write_reason,
  payload = excluded.payload,
  updated_at = greatest(public.step_tracking_cache.updated_at, excluded.updated_at);

-- Mantém o banco enxuto: remove caches por cliente e qualquer snapshot extra.
delete from public.step_tracking_cache
where cache_key <> 'projects:BR:current';

create index if not exists idx_step_tracking_cache_updated_at
  on public.step_tracking_cache (updated_at desc);

create index if not exists idx_step_tracking_cache_version
  on public.step_tracking_cache (version);

-- Garante que novas gravações não criem histórico nem linhas por cliente.
do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'step_tracking_cache_single_key_only'
      and conrelid = 'public.step_tracking_cache'::regclass
  ) then
    alter table public.step_tracking_cache
      add constraint step_tracking_cache_single_key_only
      check (cache_key = 'projects:BR:current') not valid;
  end if;
end $$;

-- Limite de segurança para evitar payload gigante por acidente no plano Free.
-- O código também bloqueia antes de gravar. Este check protege o banco.
do $$
begin
  if not exists (
    select 1
    from pg_constraint
    where conname = 'step_tracking_cache_payload_bytes_limit'
      and conrelid = 'public.step_tracking_cache'::regclass
  ) then
    alter table public.step_tracking_cache
      add constraint step_tracking_cache_payload_bytes_limit
      check (coalesce(payload_bytes, 0) <= 8388608) not valid;
  end if;
end $$;

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

revoke all on table public.step_tracking_cache from anon;
revoke all on table public.step_tracking_cache from authenticated;

comment on table public.step_tracking_cache is
  'Cache operacional STEP. v37.08: uso seguro para Supabase Free, somente uma linha projects:BR:current, sem histórico e sem acesso público.';

comment on column public.step_tracking_cache.payload is
  'Última base válida do tracking em JSONB. Não salvar imagens/base64 aqui.';

comment on column public.step_tracking_cache.payload_bytes is
  'Tamanho aproximado do payload salvo. Limite de segurança do plano Free: 8MB.';
