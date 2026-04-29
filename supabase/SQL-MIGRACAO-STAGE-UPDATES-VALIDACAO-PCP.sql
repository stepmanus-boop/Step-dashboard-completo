-- Execute no Supabase caso a tabela stage_updates já exista com o CHECK antigo de status.
-- Necessário para o fluxo de Validação PCP integrado ao Smartsheet/Tracking.

alter table public.stage_updates
  drop constraint if exists stage_updates_status_check;

alter table public.stage_updates
  add constraint stage_updates_status_check
  check (status in ('pending', 'resolved', 'pending_advance', 'pending_review', 'resolved_advance', 'resolved_review'));

alter table public.stage_updates
  alter column status set default 'pending_advance';

create index if not exists idx_stage_updates_resolved_progress
  on public.stage_updates(status, progress, resolved_at desc);
