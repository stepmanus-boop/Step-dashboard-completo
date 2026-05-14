-- v37.4 - Campos de país/ambiente para cadastros separados no mesmo Supabase

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT';

UPDATE public.users
SET operation_region = COALESCE(NULLIF(operation_region, ''), 'PT'),
    site_key = COALESCE(NULLIF(site_key, ''), COALESCE(NULLIF(operation_region, ''), 'PT'));

CREATE INDEX IF NOT EXISTS idx_users_operation_region ON public.users(operation_region);
CREATE INDEX IF NOT EXISTS idx_users_site_key ON public.users(site_key);
CREATE INDEX IF NOT EXISTS idx_users_client_key_region ON public.users(client_key, operation_region);
