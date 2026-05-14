-- v37.7 - Corrige cadastros antigos sem região para não bloquear Portugal

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text,
ADD COLUMN IF NOT EXISTS site_key text;

-- Admin e PCP seguem universais.
UPDATE public.users
SET operation_region = 'GLOBAL',
    site_key = 'GLOBAL'
WHERE role = 'admin'
   OR lower(coalesce(sector, '')) = 'pcp';

-- Inferência por client_key.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR'
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp')
  AND (operation_region IS NULL OR operation_region = '')
  AND upper(coalesce(client_key, '')) LIKE '%_BR';

UPDATE public.users
SET operation_region = 'PT',
    site_key = 'PT'
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp')
  AND (operation_region IS NULL OR operation_region = '')
  AND upper(coalesce(client_key, '')) LIKE '%_PT';

-- Legados sem região/sufixo ficam como BR para não bloquear novos cadastros PT.
UPDATE public.users
SET operation_region = 'BR',
    site_key = 'BR'
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp')
  AND (operation_region IS NULL OR operation_region = '');

DROP INDEX IF EXISTS public.idx_users_username_operation_region_unique;
DROP INDEX IF EXISTS public.idx_users_username_universal_unique;
DROP INDEX IF EXISTS public.idx_users_username_region_non_universal_unique;
ALTER TABLE public.users DROP CONSTRAINT IF EXISTS users_username_key;

CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_universal_unique
ON public.users (lower(username))
WHERE role = 'admin' OR lower(coalesce(sector, '')) = 'pcp';

CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_region_non_universal_unique
ON public.users (lower(username), operation_region)
WHERE NOT (role = 'admin' OR lower(coalesce(sector, '')) = 'pcp');
