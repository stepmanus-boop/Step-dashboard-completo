create table if not exists public.stage_updates (
  id text primary key,
  project_row_id bigint not null,
  project_number text,
  project_display text,
  client text,
  spool_iso text not null,
  spool_description text,
  sector text not null,
  progress integer not null,
  completion_date date,
  note text default '',
  status text not null default 'pending',
  created_by text,
  created_by_name text,
  created_at timestamptz not null default now(),
  resolved_by text,
  resolved_by_name text,
  resolved_at timestamptz,
  resolution_note text default ''
);

create index if not exists idx_stage_updates_project_sector_status on public.stage_updates (project_row_id, sector, status);
create index if not exists idx_stage_updates_created_at on public.stage_updates (created_at desc);
