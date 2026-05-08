-- ============================================================
-- STEP - LOGO DO CLIENTE + IMAGENS INDIVIDUAIS POR UNIDADE
-- Rode no SQL Editor do Supabase antes de usar o upload.
-- ============================================================

alter table public.users
add column if not exists client_name text;

alter table public.users
add column if not exists portal_display_name text;

alter table public.users
add column if not exists client_logo_url text;

alter table public.users
add column if not exists client_platform_image_url text;

alter table if exists public.profiles
add column if not exists client_name text;

alter table if exists public.profiles
add column if not exists portal_display_name text;

alter table if exists public.profiles
add column if not exists client_logo_url text;

alter table if exists public.profiles
add column if not exists client_platform_image_url text;

create table if not exists public.client_unit_images (
  id uuid primary key default gen_random_uuid(),
  client_name text not null,
  unit_name text not null,
  image_url text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

create unique index if not exists client_unit_images_client_unit_unique
on public.client_unit_images (client_name, unit_name);

create or replace function public.set_client_unit_images_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_client_unit_images_updated_at on public.client_unit_images;

create trigger trg_client_unit_images_updated_at
before update on public.client_unit_images
for each row
execute function public.set_client_unit_images_updated_at();

alter table public.client_unit_images enable row level security;

drop policy if exists "client_unit_images_select_all" on public.client_unit_images;
create policy "client_unit_images_select_all"
on public.client_unit_images
for select
using (true);

drop policy if exists "client_unit_images_insert_authenticated" on public.client_unit_images;
create policy "client_unit_images_insert_authenticated"
on public.client_unit_images
for insert
to authenticated
with check (true);

drop policy if exists "client_unit_images_update_authenticated" on public.client_unit_images;
create policy "client_unit_images_update_authenticated"
on public.client_unit_images
for update
to authenticated
using (true)
with check (true);

drop policy if exists "client_unit_images_delete_authenticated" on public.client_unit_images;
create policy "client_unit_images_delete_authenticated"
on public.client_unit_images
for delete
to authenticated
using (true);

insert into storage.buckets (id, name, public)
values ('client-logos', 'client-logos', true)
on conflict (id) do update set public = true;

insert into storage.buckets (id, name, public)
values ('client-unit-images', 'client-unit-images', true)
on conflict (id) do update set public = true;

drop policy if exists "client_logos_public_read" on storage.objects;
create policy "client_logos_public_read"
on storage.objects
for select
using (bucket_id = 'client-logos');

drop policy if exists "client_logos_auth_insert" on storage.objects;
create policy "client_logos_auth_insert"
on storage.objects
for insert
to authenticated
with check (bucket_id = 'client-logos');

drop policy if exists "client_logos_auth_update" on storage.objects;
create policy "client_logos_auth_update"
on storage.objects
for update
to authenticated
using (bucket_id = 'client-logos')
with check (bucket_id = 'client-logos');

drop policy if exists "client_logos_auth_delete" on storage.objects;
create policy "client_logos_auth_delete"
on storage.objects
for delete
to authenticated
using (bucket_id = 'client-logos');

drop policy if exists "client_unit_images_public_read" on storage.objects;
create policy "client_unit_images_public_read"
on storage.objects
for select
using (bucket_id = 'client-unit-images');

drop policy if exists "client_unit_images_auth_insert" on storage.objects;
create policy "client_unit_images_auth_insert"
on storage.objects
for insert
to authenticated
with check (bucket_id = 'client-unit-images');

drop policy if exists "client_unit_images_auth_update" on storage.objects;
create policy "client_unit_images_auth_update"
on storage.objects
for update
to authenticated
using (bucket_id = 'client-unit-images')
with check (bucket_id = 'client-unit-images');

drop policy if exists "client_unit_images_auth_delete" on storage.objects;
create policy "client_unit_images_auth_delete"
on storage.objects
for delete
to authenticated
using (bucket_id = 'client-unit-images');

notify pgrst, 'reload schema';
