-- v37.3 - Supabase: separação por site/ambiente Portugal
-- Execute no SQL Editor do Supabase uma única vez.

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT';

-- Garante que cadastros antigos fiquem como Portugal neste site separado.
UPDATE public.users
SET operation_region = COALESCE(NULLIF(operation_region, ''), 'PT'),
    site_key = COALESCE(NULLIF(site_key, ''), COALESCE(NULLIF(operation_region, ''), 'PT'));

-- Índices para busca rápida por ambiente e client_key.
CREATE INDEX IF NOT EXISTS idx_users_operation_region ON public.users(operation_region);
CREATE INDEX IF NOT EXISTS idx_users_site_key ON public.users(site_key);
CREATE INDEX IF NOT EXISTS idx_users_client_key_region ON public.users(client_key, operation_region);

-- Exemplo de cadastro separado:
-- SBM Brasil:   client_key = 'SBM_BR', operation_region = 'BR'
-- SBM Portugal: client_key = 'SBM_PT', operation_region = 'PT'
