create extension if not exists pgcrypto;

create table if not exists public.alert_responses (
  id uuid primary key default gen_random_uuid(),
  alert_id uuid not null references public.manual_alerts(id) on delete cascade,
  user_id uuid,
  username text,
  user_email text,
  sector text,
  response_text text not null,
  admin_reply text,
  status text default 'enviado',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_alert_responses_alert_id on public.alert_responses(alert_id);
create index if not exists idx_alert_responses_created_at on public.alert_responses(created_at desc);
