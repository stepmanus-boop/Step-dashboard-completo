-- v37.64 - QR Code automático por ISO + ajuste de chaves de cache
-- Execute no Supabase SQL Editor antes de subir o deploy desta versão.
--
-- O que este SQL faz:
-- 1) Permite que o cache principal e o cache Yinson convivam na tabela step_tracking_cache.
-- 2) Cria a tabela iso_qr_codes para armazenar um QR único por ISO.
-- 3) Mantém RLS fechado para o navegador; as Netlify Functions usam SERVICE_ROLE_KEY.

-- =========================================================
-- 1) Ajuste seguro da regra de cache
-- =========================================================

alter table public.step_tracking_cache
  drop constraint if exists step_tracking_cache_single_key_only;

alter table public.step_tracking_cache
  drop constraint if exists step_tracking_cache_allowed_keys;

alter table public.step_tracking_cache
  add constraint step_tracking_cache_allowed_keys
  check (
    cache_key in (
      'projects:BR:current',
      'yinson:under-dev:current'
    )
  ) not valid;

-- =========================================================
-- 2) Tabela de QR Code por ISO
-- =========================================================

create table if not exists public.iso_qr_codes (
  id uuid primary key default gen_random_uuid(),
  region text not null default 'BR',
  client text not null default '',
  client_key text not null default '',
  bsp text not null default '',
  bsp_key text not null default '',
  work_order text not null default '',
  vessel text not null default '',
  tag_number text not null default '',
  iso text not null,
  iso_key text not null,
  iso_full_name text not null,
  qr_token uuid not null default gen_random_uuid(),
  qr_url text not null,
  status text not null default '',
  progress numeric not null default 0,
  source text not null default 'tracking-cache-auto',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint iso_qr_codes_unique_iso unique (region, client_key, bsp_key, iso_key),
  constraint iso_qr_codes_unique_token unique (qr_token)
);

-- Compatibilidade caso a tabela já exista de uma tentativa anterior.
alter table public.iso_qr_codes add column if not exists client_key text not null default '';
alter table public.iso_qr_codes add column if not exists bsp_key text not null default '';
alter table public.iso_qr_codes add column if not exists iso_key text not null default '';
alter table public.iso_qr_codes add column if not exists source text not null default 'tracking-cache-auto';

-- Garante a regra atual de unicidade sem depender de maiúsculas/minúsculas visualmente diferentes.
alter table public.iso_qr_codes drop constraint if exists iso_qr_codes_unique_iso;
alter table public.iso_qr_codes
  add constraint iso_qr_codes_unique_iso unique (region, client_key, bsp_key, iso_key);


create index if not exists idx_iso_qr_codes_search_iso
  on public.iso_qr_codes (iso);

create index if not exists idx_iso_qr_codes_search_bsp
  on public.iso_qr_codes (bsp);

create index if not exists idx_iso_qr_codes_keys
  on public.iso_qr_codes (region, client_key, bsp_key, iso_key);

create index if not exists idx_iso_qr_codes_client_bsp
  on public.iso_qr_codes (client, bsp);

create index if not exists idx_iso_qr_codes_updated_at
  on public.iso_qr_codes (updated_at desc);

create index if not exists idx_iso_qr_codes_qr_token
  on public.iso_qr_codes (qr_token);

-- =========================================================
-- 3) Segurança
-- =========================================================

alter table public.iso_qr_codes enable row level security;

revoke all on table public.iso_qr_codes from anon;
revoke all on table public.iso_qr_codes from authenticated;
grant all on table public.iso_qr_codes to service_role;

-- Bloqueia acesso direto pelo browser. A consulta deve passar pelas Functions.
do $$
begin
  if not exists (
    select 1 from pg_policies
    where schemaname = 'public'
      and tablename = 'iso_qr_codes'
      and policyname = 'iso_qr_codes_no_public_access'
  ) then
    create policy iso_qr_codes_no_public_access
      on public.iso_qr_codes
      for all
      using (false)
      with check (false);
  end if;
end $$;

comment on table public.iso_qr_codes is
  'v37.64: QR Codes únicos por ISO, criados automaticamente durante atualização do cache do Tracking.';

comment on column public.iso_qr_codes.iso_full_name is
  'Nome completo do ISO exibido na etiqueta minimalista abaixo do QR Code.';

comment on column public.iso_qr_codes.qr_url is
  'URL de rastreamento embutida no QR Code. A etiqueta impressa mostra somente o QR e o ISO completo.';
