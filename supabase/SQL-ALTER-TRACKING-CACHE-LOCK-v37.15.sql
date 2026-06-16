-- v37.15 - Organização de atualização do cache com lock no Supabase
-- Execute no Supabase SQL Editor.
--
-- Objetivo:
-- - Manter o modelo Free seguro com uma única linha: projects:BR:current
-- - Adicionar colunas de lock na própria linha do cache
-- - Impedir que vários usuários disparem atualização do Smartsheet ao mesmo tempo
-- - O frontend abre pelo cache do Supabase e só atualiza quando passar do intervalo configurado

alter table public.step_tracking_cache
  add column if not exists refresh_lock_until timestamptz;

alter table public.step_tracking_cache
  add column if not exists refresh_lock_owner text;

alter table public.step_tracking_cache
  add column if not exists refresh_started_at timestamptz;

create index if not exists idx_step_tracking_cache_refresh_lock_until
  on public.step_tracking_cache (refresh_lock_until);

-- Mantém a regra de uma linha operacional. Não cria cache por usuário e não cria histórico.
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

-- Limpa lock vencido/antigo se existir.
update public.step_tracking_cache
set
  refresh_lock_until = null,
  refresh_lock_owner = null
where cache_key = 'projects:BR:current'
  and refresh_lock_until is not null
  and refresh_lock_until < now();

comment on column public.step_tracking_cache.refresh_lock_until is
  'v37.15: validade do lock de sincronização. Enquanto estiver no futuro, outro usuário não atualiza o Smartsheet.';

comment on column public.step_tracking_cache.refresh_lock_owner is
  'v37.15: identificador temporário de quem iniciou a sincronização controlada.';

comment on column public.step_tracking_cache.refresh_started_at is
  'v37.15: horário em que a sincronização controlada foi iniciada.';
