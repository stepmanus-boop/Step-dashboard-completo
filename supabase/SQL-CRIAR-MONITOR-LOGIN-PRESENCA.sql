-- Monitor de login/presença dos usuários do dashboard STEP.
-- Rode no Supabase SQL Editor antes de publicar esta versão.

create table if not exists public.user_presence (
  user_id text primary key,
  username text not null default '',
  name text not null default '',
  role text not null default 'sector',
  sector text not null default '',
  alert_sectors text[] not null default '{}',
  status text not null default 'offline',
  last_seen_at timestamptz,
  last_login_at timestamptz,
  last_logout_at timestamptz,
  last_view_at timestamptz,
  last_view_name text not null default '',
  last_view_url text not null default '',
  last_view_title text not null default '',
  user_agent text not null default '',
  ip_address text not null default '',
  updated_at timestamptz not null default now(),
  constraint user_presence_status_check check (status in ('online', 'offline'))
);

comment on table public.user_presence is
'Guarda o último sinal de atividade dos usuários para o monitor de login ao vivo do painel admin.';

comment on column public.user_presence.last_seen_at is
'Último heartbeat recebido do navegador. O sistema considera online quando status=online e o último sinal é recente.';

comment on column public.user_presence.last_view_name is
'Última área/tela visualizada pelo usuário no dashboard.';

create index if not exists user_presence_status_seen_idx
on public.user_presence (status, last_seen_at desc);

create index if not exists user_presence_username_idx
on public.user_presence (username);
