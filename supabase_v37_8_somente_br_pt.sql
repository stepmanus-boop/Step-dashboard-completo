-- v37.8 - Somente BR e PT no Supabase
-- Admin e PCP continuam universais na regra do sistema,
-- mas operation_region/site_key ficam apenas BR ou PT.

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT';

-- Converte GLOBAL ou vazio para BR por segurança nos usuários já existentes do Brasil.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR'
WHERE operation_region IS NULL
   OR operation_region = ''
   OR operation_region = 'GLOBAL'
   OR site_key IS NULL
   OR site_key = ''
   OR site_key = 'GLOBAL';

-- Se client_key indicar PT, força PT.
UPDATE public.users
SET operation_region = 'PT',
    site_key = 'PT'
WHERE upper(coalesce(client_key, '')) LIKE '%_PT';

-- Se client_key indicar BR, força BR.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR'
WHERE upper(coalesce(client_key, '')) LIKE '%_BR';

-- Garante que qualquer valor diferente vire PT ou BR.
UPDATE public.users
SET operation_region = CASE WHEN upper(operation_region) = 'BR' THEN 'BR' ELSE 'PT' END,
    site_key = CASE WHEN upper(site_key) = 'BR' THEN 'BR' ELSE 'PT' END;

-- Remove constraint/índices antigos globais.
ALTER TABLE public.users
DROP CONSTRAINT IF EXISTS users_username_key;

DROP INDEX IF EXISTS public.idx_users_username_unique;
DROP INDEX IF EXISTS public.users_username_unique;
DROP INDEX IF EXISTS public.idx_users_username_operation_region_unique;
DROP INDEX IF EXISTS public.idx_users_username_universal_unique;
DROP INDEX IF EXISTS public.idx_users_username_region_non_universal_unique;

-- Admin/PCP: login global único, independentemente de BR/PT.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_universal_unique
ON public.users (lower(username))
WHERE role = 'admin' OR lower(coalesce(sector, '')) = 'pcp';

-- Demais usuários: login separado por BR/PT.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_region_non_universal_unique
ON public.users (lower(username), operation_region)
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');

CREATE INDEX IF NOT EXISTS idx_users_operation_region ON public.users(operation_region);
CREATE INDEX IF NOT EXISTS idx_users_client_key_region ON public.users(client_key, operation_region);
