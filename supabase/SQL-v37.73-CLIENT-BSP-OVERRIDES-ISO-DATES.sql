-- STEP Dashboard v37.73 | Habilitar ajustes executivos e datas individuais por ISO
-- Execute no Supabase SQL Editor caso apareça erro relacionado a client_bsp_overrides
-- ou caso as alterações de início/término individual da ISO não fiquem salvas.

create extension if not exists pgcrypto;

create table if not exists public.client_bsp_overrides (
  id uuid primary key default gen_random_uuid(),
  region text default 'BR',

  project_row_id text not null,
  project_number text default '',
  project_display text default '',
  client_key text default '',
  client_name text default '',
  vessel text default '',
  pm text default '',

  fabrication_start_override date null,
  boilermaker_finish_override date null,
  welding_finish_override date null,
  inspection_finish_override date null,
  th_finish_override date null,
  coating_finish_override date null,
  project_finish_override date null,

  executive_status text default '',
  executive_note text default '',
  delay_reason text default '',
  custom_fields jsonb default '{}'::jsonb,
  visible_to_client boolean default true,

  created_by text default '',
  created_by_name text default '',
  created_at timestamptz default now(),
  updated_by text default '',
  updated_by_name text default '',
  updated_at timestamptz default now(),

  constraint client_bsp_overrides_project_row_id_key unique(project_row_id)
);

alter table public.client_bsp_overrides
  add column if not exists region text default 'BR',
  add column if not exists project_number text default '',
  add column if not exists project_display text default '',
  add column if not exists client_key text default '',
  add column if not exists client_name text default '',
  add column if not exists vessel text default '',
  add column if not exists pm text default '',
  add column if not exists fabrication_start_override date null,
  add column if not exists boilermaker_finish_override date null,
  add column if not exists welding_finish_override date null,
  add column if not exists inspection_finish_override date null,
  add column if not exists th_finish_override date null,
  add column if not exists coating_finish_override date null,
  add column if not exists project_finish_override date null,
  add column if not exists executive_status text default '',
  add column if not exists executive_note text default '',
  add column if not exists delay_reason text default '',
  add column if not exists custom_fields jsonb default '{}'::jsonb,
  add column if not exists visible_to_client boolean default true,
  add column if not exists created_by text default '',
  add column if not exists created_by_name text default '',
  add column if not exists created_at timestamptz default now(),
  add column if not exists updated_by text default '',
  add column if not exists updated_by_name text default '',
  add column if not exists updated_at timestamptz default now();

create unique index if not exists client_bsp_overrides_project_row_id_key
on public.client_bsp_overrides(project_row_id);

create index if not exists idx_client_bsp_overrides_project_row_id
on public.client_bsp_overrides(project_row_id);

create index if not exists idx_client_bsp_overrides_project_number
on public.client_bsp_overrides(project_number);

create index if not exists idx_client_bsp_overrides_client_name
on public.client_bsp_overrides(client_name);

create index if not exists idx_client_bsp_overrides_pm
on public.client_bsp_overrides(pm);

alter table public.client_bsp_overrides enable row level security;

drop policy if exists "client_bsp_overrides_no_direct_select" on public.client_bsp_overrides;
drop policy if exists "client_bsp_overrides_no_direct_insert" on public.client_bsp_overrides;
drop policy if exists "client_bsp_overrides_no_direct_update" on public.client_bsp_overrides;
drop policy if exists "client_bsp_overrides_no_direct_delete" on public.client_bsp_overrides;

create policy "client_bsp_overrides_no_direct_select"
on public.client_bsp_overrides
for select
using (false);

create policy "client_bsp_overrides_no_direct_insert"
on public.client_bsp_overrides
for insert
with check (false);

create policy "client_bsp_overrides_no_direct_update"
on public.client_bsp_overrides
for update
using (false)
with check (false);

create policy "client_bsp_overrides_no_direct_delete"
on public.client_bsp_overrides
for delete
using (false);

notify pgrst, 'reload schema';
