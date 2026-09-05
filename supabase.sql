create table profiles (id uuid primary key, name text not null, email text, role text not null check (role in ('parlamentario','asesor','admin')), created_at timestamptz default now());
create table areas (id bigint generated always as identity primary key, name text unique not null);
create table user_areas (user_id uuid references profiles(id) on delete cascade, area_id bigint references areas(id) on delete cascade, primary key(user_id,area_id));
create table tasks (id uuid primary key default gen_random_uuid(), title text not null, description text, owner_id uuid references profiles(id), area_id bigint references areas(id), priority text not null default 'pronto', status text not null default 'pendiente', due_date date, created_at timestamptz default now(), updated_at timestamptz default now());
create table agenda_events (id uuid primary key default gen_random_uuid(), owner_id uuid references profiles(id), title text not null, location text, starts_at timestamptz not null, created_at timestamptz default now());
create table deadlines (id uuid primary key default gen_random_uuid(), owner_id uuid references profiles(id), title text not null, due_date date not null, source_url text, created_at timestamptz default now());
-- Próximas fases: radar_items, initiatives, initiative_sources, notifications, activity_log.
