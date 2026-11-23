-- ════════════════════════════════════════════════════════════════════════
-- Portal de Paineis de Mapeamento — estrutura do banco (Supabase)
-- Cole tudo no SQL Editor do projeto e clique em Run. Pode rodar de novo
-- sem problema: os comandos so criam o que ainda nao existe.
-- ════════════════════════════════════════════════════════════════════════

-- Empresas que tem painel (o "slug" e o nome da pasta do painel no Storage)
create table if not exists public.empresas (
  slug       text primary key,
  nome       text not null,
  ativo      boolean not null default true,
  criado_em  timestamptz not null default now()
);

-- Quem pode entrar no portal (1 linha por usuario do Supabase Auth)
create table if not exists public.perfis (
  id         uuid primary key references auth.users(id) on delete cascade,
  nome       text not null,
  email      text not null,
  papel      text not null default 'visualizador' check (papel in ('admin','visualizador')),
  ativo      boolean not null default true,
  criado_em  timestamptz not null default now()
);

-- Quais empresas cada usuario pode ver
create table if not exists public.acessos (
  usuario    uuid not null references public.perfis(id) on delete cascade,
  empresa    text not null references public.empresas(slug) on delete cascade on update cascade,
  criado_em  timestamptz not null default now(),
  primary key (usuario, empresa)
);

-- ── Funcoes de apoio (rodam com permissao do banco, sem expor as tabelas) ──
create or replace function public.eh_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from perfis where id = auth.uid() and papel = 'admin' and ativo);
$$;

create or replace function public.pode_ver(p_empresa text) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from perfis where id = auth.uid() and ativo)
     and (public.eh_admin()
          or exists (select 1 from acessos a join empresas e on e.slug = a.empresa
                     where a.usuario = auth.uid() and a.empresa = p_empresa and e.ativo));
$$;

revoke all on function public.eh_admin() from public, anon;
revoke all on function public.pode_ver(text) from public, anon;
grant execute on function public.eh_admin() to authenticated;
grant execute on function public.pode_ver(text) to authenticated;

-- ── Seguranca por linha (RLS) ──
alter table public.empresas enable row level security;
alter table public.perfis   enable row level security;
alter table public.acessos  enable row level security;

revoke all on public.empresas, public.perfis, public.acessos from anon;
grant select, insert, update, delete on public.empresas, public.perfis, public.acessos to authenticated;

drop policy if exists empresas_ver   on public.empresas;
drop policy if exists empresas_admin on public.empresas;
create policy empresas_ver   on public.empresas for select to authenticated using (public.pode_ver(slug));
create policy empresas_admin on public.empresas for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists perfis_ver   on public.perfis;
drop policy if exists perfis_admin on public.perfis;
create policy perfis_ver   on public.perfis for select to authenticated using (id = auth.uid() or public.eh_admin());
create policy perfis_admin on public.perfis for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists acessos_ver   on public.acessos;
drop policy if exists acessos_admin on public.acessos;
create policy acessos_ver   on public.acessos for select to authenticated using (usuario = auth.uid() or public.eh_admin());
create policy acessos_admin on public.acessos for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

-- ── Area privada dos paineis (Storage) ──
insert into storage.buckets (id, name, public)
values ('paineis', 'paineis', false)
on conflict (id) do update set public = false;

drop policy if exists paineis_ler on storage.objects;
create policy paineis_ler on storage.objects for select to authenticated
  using (bucket_id = 'paineis' and public.pode_ver((storage.foldername(name))[1]));
-- (so o robo grava, usando a chave secreta, que ignora estas regras)

-- ── Empresas iniciais ──
insert into public.empresas (slug, nome) values
  ('millena', 'Millena'),
  ('novatec', 'Novatec')
on conflict (slug) do nothing;

-- ── 22/11/2026: prazo do temporizador por empresa (editado na aba Empresas do Admin) ──
alter table public.empresas add column if not exists prazo date;
grant select, insert, update on public.empresas to service_role;

-- ── 22/11/2026: atualizacao automatica por empresa (aba Empresas do Admin) ──
alter table public.empresas add column if not exists intervalo_min integer not null default 60;
alter table public.empresas add column if not exists atualizacao_ativa boolean not null default true;

-- ── 22/11/2026: admin gera senha nova para um usuario (botoes "Nova senha" e "WhatsApp") ──
create extension if not exists pgcrypto with schema extensions;
create or replace function public.admin_definir_senha(p_usuario uuid, p_senha text) returns void
language plpgsql security definer set search_path = public, extensions, auth as $$
begin
  if not public.eh_admin() then raise exception 'apenas administradores podem trocar senhas'; end if;
  if length(coalesce(p_senha, '')) < 8 then raise exception 'a senha precisa ter pelo menos 8 caracteres'; end if;
  update auth.users set encrypted_password = extensions.crypt(p_senha, extensions.gen_salt('bf')), updated_at = now()
   where id = p_usuario;
  if not found then raise exception 'usuario nao encontrado'; end if;
end $$;
revoke all on function public.admin_definir_senha(uuid, text) from public, anon;
grant execute on function public.admin_definir_senha(uuid, text) to authenticated;

-- ── 23/11/2026: filtro do que cada usuario ve dentro do painel (setor, cargo, nome) ──
-- Ver o arquivo supabase/2026-11-23_filtros.sql (tabelas filtros e paineis_gerados,
-- funcao pode_ler_painel e a nova regra paineis_ler do Storage, que substitui a de cima).
