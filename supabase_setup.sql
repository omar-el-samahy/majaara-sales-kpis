create table if not exists public.employees (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  role text not null check (role in ('rep','leader')),
  sales_target numeric not null default 0,
  meetings_target numeric not null default 0,
  cycle_target_days numeric,
  active boolean not null default true,
  created_at timestamptz not null default now()
);

create table if not exists public.entries (
  id bigint generated always as identity primary key,
  employee_id uuid not null references public.employees(id) on delete cascade,
  entry_date date not null,
  sales numeric,
  meetings integer,
  deals integer,
  cycle_days numeric,
  crm_done boolean,
  csat numeric,
  team_sales numeric,
  leader_meetings integer,
  team_meetings integer,
  team_deals integer,
  members_ok integer,
  members_total integer,
  crm_status text,
  initiative_status text,
  notes text,
  created_at timestamptz not null default now()
);

create index if not exists entries_emp_date_idx on public.entries (employee_id, entry_date);

alter table public.employees enable row level security;
alter table public.entries enable row level security;

drop policy if exists "anon employees access" on public.employees;
create policy "anon employees access" on public.employees
  for all to anon using (true) with check (true);

drop policy if exists "anon entries access" on public.entries;
create policy "anon entries access" on public.entries
  for all to anon using (true) with check (true);
