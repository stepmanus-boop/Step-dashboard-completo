-- Nenhuma alteração estrutural é obrigatória no Supabase para este recurso.
-- O botão de mudar senha reutiliza a tabela public.users já existente
-- e apenas atualiza a coluna password_hash do usuário autenticado.

-- SQL opcional: garantir que a coluna updated_at seja preenchida automaticamente
-- caso sua tabela users ainda não esteja atualizando esse campo.

alter table public.users
add column if not exists updated_at timestamptz default now();

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_users_updated_at on public.users;
create trigger trg_users_updated_at
before update on public.users
for each row execute function public.set_updated_at();
