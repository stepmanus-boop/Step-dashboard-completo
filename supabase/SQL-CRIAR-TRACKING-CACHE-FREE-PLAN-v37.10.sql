-- v37.10: mantém a mesma tabela de cache único; a correção principal está no código para bloquear fallback empacotado antigo.
-- v37.10 - Cache persistente seguro para Supabase Free + fallback sem tela vazia
-- Execute no Supabase SQL Editor.
--
-- Objetivo:
-- - Manter somente UMA linha de cache: projects:BR:current
-- - Não criar histórico, não criar cache por cliente, não salvar imagens/base64
-- - Permitir que o backend grave a última base válida sem estourar o plano Free
-- - Atualizar o limite de payload para 24MB, mantendo o banco protegido

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

-- Preserva a linha correta, caso já exista cache anterior.
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
  coalesce(source, 'migration-v37.10') as source,
  version,
  coalesce(projects_count, 0) as projects_count,
  coalesce(nullif(payload_bytes, 0), length(payload::text)) as payload_bytes,
  'migration-to-v37.10' as last_write_reason,
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

-- Remove qualquer linha que não seja o cache atual do Brasil.
delete from public.step_tracking_cache
where cache_key <> 'projects:BR:current';

create index if not exists idx_step_tracking_cache_updated_at
  on public.step_tracking_cache (updated_at desc);

create index if not exists idx_step_tracking_cache_version
  on public.step_tracking_cache (version);

-- Garante linha única. Se já existir constraint antiga, preserva.
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

-- Atualiza o limite de segurança de 8MB da v37.08 para 24MB.
-- O código da Function também bloqueia payload acima desse limite antes de gravar.
do $$
begin
  if exists (
    select 1
    from pg_constraint
    where conname = 'step_tracking_cache_payload_bytes_limit'
      and conrelid = 'public.step_tracking_cache'::regclass
  ) then
    alter table public.step_tracking_cache
      drop constraint step_tracking_cache_payload_bytes_limit;
  end if;

  alter table public.step_tracking_cache
    add constraint step_tracking_cache_payload_bytes_limit
    check (coalesce(payload_bytes, 0) <= 25165824) not valid;
end $$;

alter table public.step_tracking_cache enable row level security;

-- Segurança: nenhum acesso direto pelo browser. A Function usa SERVICE_ROLE_KEY e ignora RLS.
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

grant all on table public.step_tracking_cache to service_role;

comment on table public.step_tracking_cache is
  'Cache operacional STEP. v37.10: Supabase Free seguro, uma linha projects:BR:current, sem histórico, sem acesso público e com fallback do app quando vazio.';

comment on column public.step_tracking_cache.payload is
  'Última base válida do tracking em JSONB. Não salvar imagens/base64 aqui.';

comment on column public.step_tracking_cache.payload_bytes is
  'Tamanho aproximado do payload salvo. Limite de segurança do plano Free nesta versão: 24MB.';
