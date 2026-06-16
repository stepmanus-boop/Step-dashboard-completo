-- Competências internas para usuários do setor Qualidade/Inspeção.
-- Rode este script no Supabase SQL Editor antes de publicar esta versão.

alter table public.users
add column if not exists quality_competencies text[] not null default '{}';

comment on column public.users.quality_competencies is
'Competências internas da Qualidade que o usuário pode visualizar/apontar: dimensional_inicial, dimensional_final, nde, th, final_inspection_qc. Vazio mantém acesso a todas as competências da Qualidade.';

create index if not exists users_quality_competencies_gin_idx
on public.users using gin (quality_competencies);
