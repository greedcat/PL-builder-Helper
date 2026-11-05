-- PL builder settings: one row per settings document.
-- id 'pl-builder-config' (Maple.PL_cols) and 'container-match-config' (match_helper.config).
create table if not exists settings (
  id text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now()
);
