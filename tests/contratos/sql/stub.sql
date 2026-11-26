do $$ begin create role anon; create role authenticated; create role service_role bypassrls; exception when duplicate_object then null; end $$;
create schema auth; create table auth.users(id uuid primary key);
create function auth.uid() returns uuid language sql stable as $$ select nullif(current_setting('test.uid', true),'')::uuid $$;
grant usage on schema auth to anon, authenticated, service_role; grant execute on function auth.uid() to public;
grant usage on schema public to anon, authenticated, service_role;
