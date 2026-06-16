create extension if not exists pgcrypto;

create table if not exists public.stage_updates (
  id text primary key,
  project_row_id bigint not null,
  project_number text,
  project_display text,
  client text,
  spool_iso text not null,
  spool_description text,
  sector text not null,
  progress integer not null check (progress in (25, 50, 75, 100)),
  completion_date date,
  note text,
  status text not null default 'pending_advance',
  created_by text,
  created_by_name text,
  created_at timestamptz not null default now(),
  resolved_by text,
  resolved_by_name text,
  resolved_at timestamptz,
  resolution_note text,
  updated_at timestamptz not null default now()
);

alter table public.stage_updates
  drop constraint if exists stage_updates_status_check;

alter table public.stage_updates
  add constraint stage_updates_status_check
  check (status in ('pending', 'resolved', 'pending_advance', 'pending_review', 'resolved_advance', 'resolved_review'));

alter table public.stage_updates
  alter column status set default 'pending_advance';

create index if not exists idx_stage_updates_project_row_id on public.stage_updates(project_row_id);
create index if not exists idx_stage_updates_spool_sector_status on public.stage_updates(spool_iso, sector, status);
create index if not exists idx_stage_updates_created_at on public.stage_updates(created_at desc);
create index if not exists idx_stage_updates_resolved_progress on public.stage_updates(status, progress, resolved_at desc);

create or replace function public.set_stage_updates_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_stage_updates_updated_at on public.stage_updates;
create trigger trg_stage_updates_updated_at
before update on public.stage_updates
for each row
execute function public.set_stage_updates_updated_at();
