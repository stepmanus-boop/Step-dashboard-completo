-- ============================================================
-- STEP DASHBOARD v36.69
-- Permissão para visualizar Painel do Cliente por usuário
-- Execute no Supabase uma vez antes de usar o novo botão no Admin.
-- ============================================================

alter table public.users
add column if not exists can_view_client_panel boolean not null default false;

-- Admins e clientes já têm acesso pelo perfil, mas a flag pode ficar registrada.
update public.users
set can_view_client_panel = true,
    updated_at = now()
where role in ('admin', 'client')
  and can_view_client_panel is distinct from true;

notify pgrst, 'reload schema';

select
  username,
  name,
  role,
  sector,
  can_view_client_panel
from public.users
order by username;
