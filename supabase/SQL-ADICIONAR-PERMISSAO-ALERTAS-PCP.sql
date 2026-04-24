-- Permissão individual para usuário poder enviar sinalizações/alertas ao PCP.
-- Execute este SQL no Supabase antes ou junto do deploy desta versão.

alter table public.users
add column if not exists can_send_pcp_alerts boolean not null default false;

-- Mantém compatibilidade com o comportamento antigo:
-- usuários já cadastrados no setor Projetos continuam podendo enviar alertas ao PCP.
update public.users
set can_send_pcp_alerts = true
where lower(coalesce(sector, '')) in ('projetos', 'projeto', 'projects', 'project', 'pm')
  and can_send_pcp_alerts = false;
