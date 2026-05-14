-- Diagnóstico exato de conflito de login por ambiente

-- 1) Ver todos os registros de um login específico
-- Troque 'sbm' pelo login que está tentando cadastrar.
SELECT
  id,
  name,
  username,
  role,
  sector,
  client_key,
  client_name,
  operation_region,
  site_key,
  portal_site,
  created_at,
  updated_at
FROM public.users
WHERE lower(username) = lower('sbm')
ORDER BY operation_region, updated_at DESC;

-- 2) Ver duplicidade por login + ambiente
SELECT
  lower(username) AS login,
  operation_region,
  count(*) AS qtd
FROM public.users
GROUP BY lower(username), operation_region
HAVING count(*) > 1
ORDER BY qtd DESC, login;

-- 3) Ver usuários com campos incoerentes BR/PT
SELECT
  id,
  name,
  username,
  role,
  sector,
  client_key,
  operation_region,
  site_key,
  portal_site
FROM public.users
WHERE operation_region NOT IN ('BR','PT')
   OR site_key NOT IN ('BR','PT')
   OR portal_site NOT IN ('BR','PT')
   OR operation_region IS NULL
   OR site_key IS NULL
   OR portal_site IS NULL
ORDER BY username;
