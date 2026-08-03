-- Gantt Chart Helper — cloud schema
--
-- Run this once in your Supabase project: Dashboard → SQL Editor → New query
-- → paste → Run. It is idempotent, so re-running it is safe.

-- ── Workbooks: one row per user ─────────────────────────────────────────────

create table if not exists public.workbooks (
  user_id    uuid primary key references auth.users (id) on delete cascade,
  payload    jsonb       not null,
  updated_at timestamptz not null default now()
);

alter table public.workbooks enable row level security;

-- A user can only ever see or touch their own row.
drop policy if exists "own workbook: select" on public.workbooks;
create policy "own workbook: select" on public.workbooks
  for select using (auth.uid() = user_id);

drop policy if exists "own workbook: insert" on public.workbooks;
create policy "own workbook: insert" on public.workbooks
  for insert with check (auth.uid() = user_id);

drop policy if exists "own workbook: update" on public.workbooks;
create policy "own workbook: update" on public.workbooks
  for update using (auth.uid() = user_id) with check (auth.uid() = user_id);

drop policy if exists "own workbook: delete" on public.workbooks;
create policy "own workbook: delete" on public.workbooks
  for delete using (auth.uid() = user_id);

-- ── Shared sheets: a published snapshot, readable by link ───────────────────

create table if not exists public.shared_sheets (
  id         uuid primary key default gen_random_uuid(),
  owner_id   uuid        not null default auth.uid() references auth.users (id) on delete cascade,
  sheet_name text        not null,
  sheet      jsonb       not null,
  created_at timestamptz not null default now()
);

alter table public.shared_sheets enable row level security;

-- Deliberately NO public select policy. Readers go through the function below,
-- which requires knowing the exact id. A blanket "select using (true)" would
-- let anyone list every share link ever created.
drop policy if exists "own shares: select" on public.shared_sheets;
create policy "own shares: select" on public.shared_sheets
  for select using (auth.uid() = owner_id);

drop policy if exists "own shares: insert" on public.shared_sheets;
create policy "own shares: insert" on public.shared_sheets
  for insert with check (auth.uid() = owner_id);

drop policy if exists "own shares: delete" on public.shared_sheets;
create policy "own shares: delete" on public.shared_sheets
  for delete using (auth.uid() = owner_id);

-- Link holders read through this function only: it returns exactly one row for
-- an exact id and cannot be used to enumerate the table.
create or replace function public.get_shared_sheet(share_id uuid)
returns jsonb
language sql
security definer
set search_path = public, pg_temp
stable
as $$
  select s.sheet from public.shared_sheets s where s.id = share_id;
$$;

revoke all on function public.get_shared_sheet(uuid) from public;
grant execute on function public.get_shared_sheet(uuid) to anon, authenticated;

-- ── Housekeeping ────────────────────────────────────────────────────────────

create index if not exists shared_sheets_owner_idx
  on public.shared_sheets (owner_id, created_at desc);
