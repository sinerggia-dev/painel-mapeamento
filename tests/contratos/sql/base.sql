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


