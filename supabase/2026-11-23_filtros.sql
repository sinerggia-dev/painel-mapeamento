-- ════════════════════════════════════════════════════════════════════════
-- 23/11/2026: o que cada usuario ve DENTRO do painel (setor, cargo, nome ou todos)
-- Cole tudo no SQL Editor do Supabase e clique em Run. Pode rodar de novo sem problema.
--
-- Como funciona:
--  * sem linha em "filtros" para (usuario, empresa)  -> ve o painel completo (index.html)
--  * com linha                                        -> so ve o painel feito para ele
--    (paineis/<empresa>/u/<usuario>.html), que o robo gera so com quem estiver em
--    QUALQUER um dos setores, cargos ou nomes marcados. O painel completo fica bloqueado.
--  * quando o admin muda o filtro, o painel antigo do usuario fica bloqueado ate o robo
--    gerar o novo (alterado_em do filtro > base_em do painel gerado).
-- ════════════════════════════════════════════════════════════════════════

create table if not exists public.filtros (
  usuario      uuid not null references public.perfis(id) on delete cascade,
  empresa      text not null references public.empresas(slug) on delete cascade on update cascade,
  setores      text[] not null default '{}',
  cargos       text[] not null default '{}',
  funcionarios text[] not null default '{}',
  alterado_em  timestamptz not null default now(),
  primary key (usuario, empresa)
);

-- alterado_em e sempre a hora do banco (ninguem consegue "voltar" a data)
create or replace function public._filtros_carimbo() returns trigger
language plpgsql set search_path = public as $$
begin new.alterado_em := clock_timestamp(); return new; end $$;
drop trigger if exists filtros_carimbo on public.filtros;
create trigger filtros_carimbo before insert or update on public.filtros
  for each row execute function public._filtros_carimbo();

-- qual versao do filtro o robo usou para gerar o painel de cada usuario (so o robo grava)
create table if not exists public.paineis_gerados (
  usuario  uuid not null references public.perfis(id) on delete cascade,
  empresa  text not null references public.empresas(slug) on delete cascade on update cascade,
  base_em  timestamptz not null,
  primary key (usuario, empresa)
);

alter table public.filtros         enable row level security;
alter table public.paineis_gerados enable row level security;
revoke all on public.filtros, public.paineis_gerados from anon;
revoke all on public.paineis_gerados from authenticated;
grant select, insert, update, delete on public.filtros to authenticated;
grant select, insert, update, delete on public.filtros, public.paineis_gerados to service_role;
grant select on public.perfis, public.acessos to service_role;

drop policy if exists filtros_ver   on public.filtros;
drop policy if exists filtros_admin on public.filtros;
create policy filtros_ver   on public.filtros for select to authenticated using (usuario = auth.uid() or public.eh_admin());
create policy filtros_admin on public.filtros for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

-- quem pode baixar cada arquivo do bucket "paineis"
create or replace function public.pode_ler_painel(p_nome text) returns boolean
language plpgsql stable security definer set search_path = public as $$
declare
  a   text[] := string_to_array(p_nome, '/');
  uid uuid   := auth.uid();
begin
  if uid is null then return false; end if;
  if public.eh_admin() then return true; end if;                 -- admin ve tudo (inclusive _privado)
  if not public.pode_ver(a[1]) then return false; end if;        -- precisa ter acesso a empresa
  if array_length(a, 1) = 2 and a[2] = 'index.html' then         -- painel completo: so quem nao tem filtro
    return not exists (select 1 from filtros where usuario = uid and empresa = a[1]);
  end if;
  if array_length(a, 1) = 3 and a[2] = 'u' and a[3] = uid::text || '.html' then   -- o painel feito para ele
    return exists (select 1 from filtros f
                     join paineis_gerados g on g.usuario = f.usuario and g.empresa = f.empresa
                    where f.usuario = uid and f.empresa = a[1] and g.base_em >= f.alterado_em);
  end if;
  return false;
end $$;
revoke all on function public.pode_ler_painel(text) from public, anon;
grant execute on function public.pode_ler_painel(text) to authenticated;

drop policy if exists paineis_ler on storage.objects;
create policy paineis_ler on storage.objects for select to authenticated
  using (bucket_id = 'paineis' and public.pode_ler_painel(name));
