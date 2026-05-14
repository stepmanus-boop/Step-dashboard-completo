-- v37.6 - Admin e PCP universais; demais usuários por país/ambiente

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT';

-- Admin e PCP ficam universais.
UPDATE public.users
SET operation_region = 'GLOBAL',
    site_key = 'GLOBAL'
WHERE role = 'admin'
   OR lower(coalesce(sector, '')) = 'pcp';

-- Demais usuários continuam separados por ambiente.
UPDATE public.users
SET operation_region = COALESCE(NULLIF(operation_region, ''), 'PT'),
    site_key = COALESCE(NULLIF(site_key, ''), COALESCE(NULLIF(operation_region, ''), 'PT'))
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');

-- Remove índices únicos antigos de username se existirem.
DROP INDEX IF EXISTS public.users_username_key;
DROP INDEX IF EXISTS public.idx_users_username_unique;
DROP INDEX IF EXISTS public.users_username_unique;
DROP INDEX IF EXISTS public.idx_users_username_operation_region_unique;

-- Admin/PCP: login único global.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_universal_unique
ON public.users (lower(username))
WHERE role = 'admin' OR lower(coalesce(sector, '')) = 'pcp';

-- Demais usuários: login único somente dentro do ambiente.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_region_non_universal_unique
ON public.users (lower(username), operation_region)
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');

CREATE INDEX IF NOT EXISTS idx_users_client_key_region
ON public.users (client_key, operation_region);
