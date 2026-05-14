-- v38.3 - Login visível igual com sufixo oculto por país
-- Objetivo:
-- Cliente usa "SBM" no BR e "SBM" no PT.
-- Internamente, quando necessário, o banco pode salvar "SBM__PT" ou "SBM__BR".

-- 1) Garantir campos BR/PT
ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS portal_site text DEFAULT 'PT';

-- 2) Normalizar portal/ambiente
UPDATE public.users
SET operation_region = CASE WHEN upper(coalesce(operation_region, '')) = 'BR' THEN 'BR' ELSE 'PT' END,
    site_key = CASE WHEN upper(coalesce(site_key, '')) = 'BR' THEN 'BR' ELSE 'PT' END,
    portal_site = CASE WHEN upper(coalesce(portal_site, '')) = 'BR' THEN 'BR' ELSE 'PT' END;

-- 3) Ajustar por client_key
UPDATE public.users
SET operation_region = 'BR', site_key = 'BR', portal_site = 'BR'
WHERE upper(coalesce(client_key, '')) LIKE '%_BR';

UPDATE public.users
SET operation_region = 'PT', site_key = 'PT', portal_site = 'PT'
WHERE upper(coalesce(client_key, '')) LIKE '%_PT';

-- 4) Remove restrições antigas globais
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_key;
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_portal_site_key;
DROP INDEX IF EXISTS public.users_username_portal_site_unique_idx;
DROP INDEX IF EXISTS public.idx_users_username_portal_site_unique;
DROP INDEX IF EXISTS public.idx_users_username_unique;
DROP INDEX IF EXISTS public.users_username_unique;
DROP INDEX IF EXISTS public.idx_users_username_operation_region_unique;
DROP INDEX IF EXISTS public.idx_users_username_universal_unique;
DROP INDEX IF EXISTS public.idx_users_username_region_non_universal_unique;

-- 5) Adiciona sufixo oculto quando existir mesmo username em BR e PT
WITH ranked AS (
  SELECT
    id,
    username,
    operation_region,
    lower(regexp_replace(username, '__(BR|PT)$', '', 'i')) AS visible_username,
    COUNT(*) OVER (
      PARTITION BY lower(regexp_replace(username, '__(BR|PT)$', '', 'i'))
    ) AS total_same_visible
  FROM public.users
  WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp')
)
UPDATE public.users u
SET username = regexp_replace(u.username, '__(BR|PT)$', '', 'i') || '__' || u.operation_region
FROM ranked r
WHERE u.id = r.id
  AND r.total_same_visible > 1
  AND u.username !~* '__(BR|PT)$';

-- 6) Índices
-- Admin/PCP continuam únicos de forma global.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_universal_unique
ON public.users (lower(username))
WHERE role = 'admin' OR lower(coalesce(sector, '')) = 'pcp';

-- Demais usuários ficam únicos pelo username interno + região.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_region_non_universal_unique
ON public.users (lower(username), operation_region)
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');

CREATE INDEX IF NOT EXISTS idx_users_client_key_region
ON public.users (client_key, operation_region);

CREATE INDEX IF NOT EXISTS idx_users_portal_site
ON public.users (portal_site);

-- Conferência:
-- SELECT username, operation_region, client_key, client_name FROM public.users ORDER BY lower(username), operation_region;
