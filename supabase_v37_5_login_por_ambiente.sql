-- v37.5 - Login único por país/ambiente, não global
-- Execute no Supabase SQL Editor.

ALTER TABLE public.users
ADD COLUMN IF NOT EXISTS operation_region text DEFAULT 'PT',
ADD COLUMN IF NOT EXISTS site_key text DEFAULT 'PT';

UPDATE public.users
SET operation_region = COALESCE(NULLIF(operation_region, ''), 'PT'),
    site_key = COALESCE(NULLIF(site_key, ''), COALESCE(NULLIF(operation_region, ''), 'PT'));

-- Remove índice único antigo de username, se existir com nomes comuns.
DROP INDEX IF EXISTS public.users_username_key;
DROP INDEX IF EXISTS public.idx_users_username_unique;
DROP INDEX IF EXISTS public.users_username_unique;

-- Cria unicidade correta: mesmo username pode existir em BR e PT,
-- mas não pode repetir dentro do mesmo país/ambiente.
CREATE UNIQUE INDEX IF NOT EXISTS idx_users_username_operation_region_unique
ON public.users (lower(username), operation_region);

CREATE INDEX IF NOT EXISTS idx_users_client_key_region
ON public.users (client_key, operation_region);

-- Exemplo permitido:
-- username='sbm', operation_region='BR', client_key='SBM_BR'
-- username='sbm', operation_region='PT', client_key='SBM_PT'
