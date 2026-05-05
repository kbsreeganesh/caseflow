-- CaseFlow Supabase Setup — paste this entire block into SQL Editor and click Run

create table if not exists caseflow_store (
  key        text primary key,
  value      text not null,
  updated_at timestamptz default now()
);

alter table caseflow_store enable row level security;

drop policy if exists "caseflow public access" on caseflow_store;
create policy "caseflow public access"
  on caseflow_store for all
  using (true) with check (true);

select 'Setup complete — CaseFlow table ready' as status;
