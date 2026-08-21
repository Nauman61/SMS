-- Run this once in your Supabase project's SQL Editor (Supabase dashboard
-- → SQL Editor → New query → paste this whole file → Run).
--
-- This creates a single table that stores all of the school app's data as
-- key/value pairs (one row per data set: staff, students, fees, etc.) —
-- matching exactly how the app already organizes its data. You do not need
-- to understand SQL to use this; just run it once and you're done.

create table if not exists kv_store (
  key text primary key,
  value text not null,
  updated_at timestamptz not null default now()
);

-- Row Level Security (RLS): required by Supabase before a table can be
-- accessed by client apps at all.
alter table kv_store enable row level security;

-- IMPORTANT SECURITY NOTE:
-- This policy allows anyone holding your Supabase URL + anon key (which
-- ship inside the built website's JavaScript) to read and write this
-- table directly, bypassing the app's own Admin/Staff login screen. For a
-- small private school tool this is a reasonable trade-off (the same way
-- the app's login was already just a shared password, not bank-grade
-- security) — but treat your Supabase URL/key with the same care as a
-- password, and don't publish them in a public GitHub repo. If you need
-- stronger security later (e.g. real per-user accounts enforced at the
-- database level), that requires a custom backend with server-side auth —
-- ask and we can plan that as a next step.
drop policy if exists "Allow all access to kv_store" on kv_store;
create policy "Allow all access to kv_store"
  on kv_store
  for all
  using (true)
  with check (true);
