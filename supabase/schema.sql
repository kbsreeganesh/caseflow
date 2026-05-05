-- Run this once in Supabase → SQL Editor
-- Creates the single key-value table CaseFlow uses

create table if not exists caseflow_store (
  key        text primary key,
  value      text not null,
  updated_at timestamptz default now()
);

-- Enable Row Level Security
alter table caseflow_store enable row level security;

-- Allow anonymous users to read and write
-- (OK for internal trial — tighten later for production)
create policy "anon_all" on caseflow_store
  for all
  using (true)
  with check (true);
