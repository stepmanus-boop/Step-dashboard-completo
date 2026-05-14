-- v38.1 - Corrige portal_site para BR/PT e remove trava antiga global
-- Execute no Supabase SQL Editor.

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS portal_site text DEFAULT 'PT';

-- Remove constraint/index antigo que travava username + portal_site = global.
DROP INDEX IF EXISTS public.users_username_portal_site_unique_idx;
DROP INDEX IF EXISTS public.idx_users_username_portal_site_unique;
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_portal_site_unique_idx;
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_portal_site_key;

-- Normaliza usuários com base em client_key.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR',
    portal_site = 'BR'
WHERE upper(coalesce(client_key, '')) LIKE '%_BR';

UPDATE public.users
SET operation_region = 'PT',
    site_key = 'PT',
    portal_site = 'PT'
WHERE upper(coalesce(client_key, '')) LIKE '%_PT';

-- Quem ainda estiver vazio/global vira BR por segurança, pois é legado Brasil.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR',
    portal_site = 'BR'
WHERE operation_region IS NULL
   OR operation_region = ''
   OR operation_region = 'GLOBAL'
   OR site_key IS NULL
   OR site_key = ''
   OR site_key = 'GLOBAL'
   OR portal_site IS NULL
   OR portal_site = ''
   OR portal_site = 'global'
   OR portal_site = 'GLOBAL';

-- Garante apenas BR/PT.
UPDATE public.users
SET operation_region = CASE WHEN upper(operation_region) = 'BR' THEN 'BR' ELSE 'PT' END,
    site_key = CASE WHEN upper(site_key) = 'BR' THEN 'BR' ELSE 'PT' END,
    portal_site = CASE WHEN upper(portal_site) = 'BR' THEN 'BR' ELSE 'PT' END;

-- Remove unicidade antiga global de username.
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_key;
DROP INDEX IF EXISTS public.idx_users_username_unique;
DROP INDEX IF EXISTS public.users_username_unique;
DROP INDEX IF EXISTS public.idx_users_username_operation_region_unique;
DROP INDEX IF EXISTS public.idx_users_username_universal_unique;
DROP INDEX IF EXISTS public.idx_users_username_region_non_universal_unique;

-- Admin/PCP: login único universal pela regra.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_universal_unique
ON public.users (lower(username))
WHERE role = 'admin' OR lower(coalesce(sector, '')) = 'pcp';

-- Demais usuários: login separado por BR/PT.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_region_non_universal_unique
ON public.users (lower(username), operation_region)
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');

CREATE INDEX IF NOT EXISTS idx_users_client_key_region
ON public.users (client_key, operation_region);

CREATE INDEX IF NOT EXISTS idx_users_portal_site
ON public.users (portal_site);
