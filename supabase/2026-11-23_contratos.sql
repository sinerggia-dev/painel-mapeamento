-- ════════════════════════════════════════════════════════════════════════
-- GESTAO DE CONTRATOS — estrutura do banco (mesmo projeto Supabase do portal)
-- Cole tudo no SQL Editor do Supabase e clique em Run. Pode rodar de novo sem
-- problema: so cria o que ainda nao existe e refaz funcoes e regras.
--
-- Usa o mesmo login do portal (tabela perfis; admin = papel 'admin').
-- Tudo dos contratos comeca com "ctr_" e nao mexe nas tabelas do mapeamento.
-- Os ARQUIVOS ficam no SharePoint (site GestaoContratos); aqui so os dados.
--
-- Quem ve o que (funcao ctr_nivel):
--   admin do portal ........................ tudo (nivel 3)
--   usuario com linha em ctr_acessos ....... o nivel dele (1 ver, 2 ver e baixar,
--                                            3 editar e enviar), so nas filiais e
--                                            classes marcadas (lista vazia = todas)
--   contrato confidencial .................. so admin, quem criou e quem estiver
--                                            em ctr_liberados
-- ════════════════════════════════════════════════════════════════════════

-- ── Cadastros que vem do sistema do cliente (importados pelo Admin) ──
create table if not exists public.ctr_naturezas (
  empresa    text    not null references public.empresas(slug) on delete cascade on update cascade,
  codigo     integer not null,
  descricao  text    not null,
  classe     text,
  tipo       text,                 -- DESPESA / RECEITA
  aviso_dias integer,              -- aviso previo padrao para contratos desta natureza (opcional)
  primary key (empresa, codigo)
);

create table if not exists public.ctr_filiais (
  empresa text    not null references public.empresas(slug) on delete cascade on update cascade,
  numero  integer not null,
  nome    text    not null,
  primary key (empresa, numero)
);

create table if not exists public.ctr_fornecedores (
  empresa text not null references public.empresas(slug) on delete cascade on update cascade,
  codigo  text not null,           -- codigo no sistema do cliente (ex.: BFJ490)
  nome    text not null,
  primary key (empresa, codigo)
);

-- ── Quem pode usar os contratos de cada empresa ──
create table if not exists public.ctr_acessos (
  usuario   uuid     not null references public.perfis(id) on delete cascade,
  empresa   text     not null references public.empresas(slug) on delete cascade on update cascade,
  nivel     smallint not null default 1 check (nivel between 1 and 3),
  filiais   integer[] not null default '{}',   -- vazio = todas as filiais
  classes   text[]    not null default '{}',   -- vazio = todas as classes de natureza
  criado_em timestamptz not null default now(),
  primary key (usuario, empresa)
);

-- ── Contratos ──
create table if not exists public.ctr_contratos (
  id                   bigint generated always as identity primary key,
  numero               text unique,
  empresa              text not null references public.empresas(slug) on update cascade,
  filial               integer,
  fornecedor           text,                       -- codigo do fornecedor no sistema do cliente
  fornecedor_nome      text not null,
  natureza             integer,                    -- codigo da natureza (ctr_naturezas)
  titulo               text not null,
  valor_mensal         numeric(14,2),
  inicio               date,
  fim                  date,
  aviso_dias           integer not null default 30 check (aviso_dias between 0 and 3650),
  reajuste             text,
  renovacao_automatica boolean not null default false,
  responsavel          uuid references public.perfis(id) on delete set null,
  confidencial         boolean not null default false,
  situacao             text not null default 'ativo' check (situacao in ('ativo','encerrado')),
  origem               text not null default 'manual',   -- manual / importado
  observacao           text,
  criado_por           uuid references public.perfis(id) on delete set null,
  criado_em            timestamptz not null default now(),
  alterado_em          timestamptz not null default now(),
  check (fim is null or inicio is null or fim >= inicio)
);
create index if not exists ctr_contratos_empresa on public.ctr_contratos (empresa, fim);

create table if not exists public.ctr_liberados (
  contrato bigint not null references public.ctr_contratos(id) on delete cascade,
  usuario  uuid   not null references public.perfis(id) on delete cascade,
  primary key (contrato, usuario)
);

-- Documentos: so a funcao do servidor (ctr-arquivos) grava, depois de conferir no SharePoint
create table if not exists public.ctr_documentos (
  id          bigint generated always as identity primary key,
  contrato    bigint not null references public.ctr_contratos(id) on delete restrict,
  categoria   text   not null,
  nome        text   not null,          -- nome padronizado no SharePoint
  nome_original text,
  tamanho     bigint,
  tipo_mime   text,
  versao      integer not null default 1,
  sp_item     text   not null,          -- id do arquivo no SharePoint
  sp_caminho  text,
  enviado_por uuid references public.perfis(id) on delete set null,
  enviado_em  timestamptz not null default now()
);
create index if not exists ctr_documentos_contrato on public.ctr_documentos (contrato);
create unique index if not exists ctr_documentos_sp_item on public.ctr_documentos (sp_item);

-- Historico: so recebe linhas novas; ninguem altera nem apaga
create table if not exists public.ctr_historico (
  id       bigint generated always as identity primary key,
  contrato bigint references public.ctr_contratos(id) on delete cascade,
  empresa  text,
  usuario  uuid references public.perfis(id) on delete set null,
  acao     text not null,
  detalhe  text,
  em       timestamptz not null default now()
);
create index if not exists ctr_historico_contrato on public.ctr_historico (contrato, em desc);

-- ════════ Funcoes de permissao ════════
create or replace function public.ctr_nivel(p_empresa text, p_filial integer, p_natureza integer,
                                            p_confidencial boolean, p_criado_por uuid, p_id bigint)
returns integer language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := auth.uid();
  a   public.ctr_acessos;
  cls text;
begin
  if uid is null then return 0; end if;
  if not exists (select 1 from perfis where id = uid and ativo) then return 0; end if;
  if public.eh_admin() then return 3; end if;
  select * into a from ctr_acessos where usuario = uid and empresa = p_empresa;
  if not found then return 0; end if;
  if cardinality(a.filiais) > 0 and (p_filial is null or not (p_filial = any (a.filiais))) then return 0; end if;
  if cardinality(a.classes) > 0 then
    select classe into cls from ctr_naturezas where empresa = p_empresa and codigo = p_natureza;
    if cls is null or not (cls = any (a.classes)) then return 0; end if;
  end if;
  if p_confidencial and p_criado_por is distinct from uid
     and not exists (select 1 from ctr_liberados where contrato = p_id and usuario = uid) then
    return 0;
  end if;
  return a.nivel;
end $$;

-- nivel do usuario logado em um contrato (usado nas regras dos documentos e pela funcao do servidor)
create or replace function public.ctr_nivel_contrato(p_id bigint) returns integer
language sql stable security definer set search_path = public as $$
  select coalesce((select public.ctr_nivel(c.empresa, c.filial, c.natureza, c.confidencial, c.criado_por, c.id)
                     from ctr_contratos c where c.id = p_id), 0);
$$;

-- usuario logado pode usar os contratos desta empresa?
create or replace function public.ctr_usa_empresa(p_empresa text) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from perfis where id = auth.uid() and ativo)
     and (public.eh_admin() or exists (select 1 from ctr_acessos where usuario = auth.uid() and empresa = p_empresa));
$$;

-- nomes das pessoas que usam os contratos da empresa (para "responsavel"); so para quem tambem usa
create or replace function public.ctr_pessoas(p_empresa text) returns table (id uuid, nome text)
language sql stable security definer set search_path = public as $$
  select p.id, p.nome from perfis p
   where public.ctr_usa_empresa(p_empresa) and p.ativo
     and (p.papel = 'admin' or exists (select 1 from ctr_acessos a where a.usuario = p.id and a.empresa = p_empresa))
   order by p.nome;
$$;
revoke all on function public.ctr_pessoas(text) from public, anon;
grant execute on function public.ctr_pessoas(text) to authenticated;

revoke all on function public.ctr_nivel(text,integer,integer,boolean,uuid,bigint) from public, anon;
revoke all on function public.ctr_nivel_contrato(bigint) from public, anon;
revoke all on function public.ctr_usa_empresa(text) from public, anon;
grant execute on function public.ctr_nivel(text,integer,integer,boolean,uuid,bigint) to authenticated, service_role;
grant execute on function public.ctr_nivel_contrato(bigint) to authenticated, service_role;
grant execute on function public.ctr_usa_empresa(text) to authenticated, service_role;

-- ════════ Gatilhos dos contratos: numero, datas e historico automatico ════════
create or replace function public._ctr_contratos_antes() returns trigger
language plpgsql set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    new.numero      := 'CTR-' || to_char(clock_timestamp(), 'YYYY') || '-' || lpad(new.id::text, 5, '0');
    new.criado_por  := coalesce(auth.uid(), new.criado_por);
    new.criado_em   := clock_timestamp();
  else
    -- o que identifica o contrato nao muda
    new.numero := old.numero; new.empresa := old.empresa; new.origem := old.origem;
    new.criado_por := old.criado_por; new.criado_em := old.criado_em;
  end if;
  new.alterado_em := clock_timestamp();
  return new;
end $$;
drop trigger if exists ctr_contratos_antes on public.ctr_contratos;
create trigger ctr_contratos_antes before insert or update on public.ctr_contratos
  for each row execute function public._ctr_contratos_antes();

create or replace function public._ctr_contratos_historico() returns trigger
language plpgsql security definer set search_path = public as $$
declare
  mud text := '';
  k   text;
  novo jsonb := to_jsonb(new);
  velho jsonb;
begin
  if tg_op = 'INSERT' then
    insert into ctr_historico (contrato, empresa, usuario, acao, detalhe)
    values (new.id, new.empresa, auth.uid(), 'contrato cadastrado',
            case when new.origem = 'importado' then 'importado da planilha' end);
    return null;
  end if;
  velho := to_jsonb(old);
  foreach k in array array['titulo','fornecedor_nome','fornecedor','filial','natureza','valor_mensal','inicio','fim',
                           'aviso_dias','reajuste','renovacao_automatica','responsavel','confidencial','situacao','observacao'] loop
    if novo -> k is distinct from velho -> k then
      mud := mud || case when mud = '' then '' else '; ' end || k || ': '
             || coalesce(velho ->> k, '—') || ' → ' || coalesce(novo ->> k, '—');
    end if;
  end loop;
  if mud <> '' then
    insert into ctr_historico (contrato, empresa, usuario, acao, detalhe)
    values (new.id, new.empresa, auth.uid(), 'dados alterados', left(mud, 2000));
  end if;
  return null;
end $$;
drop trigger if exists ctr_contratos_historico on public.ctr_contratos;
create trigger ctr_contratos_historico after insert or update on public.ctr_contratos
  for each row execute function public._ctr_contratos_historico();

-- historico e documentos nao podem ser alterados nem apagados (nem pela chave secreta)
create or replace function public._ctr_imutavel() returns trigger
language plpgsql as $$
begin
  if tg_op = 'DELETE' and tg_table_name = 'ctr_historico'
     and not exists (select 1 from public.ctr_contratos where id = old.contrato) then
    return old;   -- deixa apagar so quando o proprio contrato foi apagado (cascade)
  end if;
  raise exception 'registro protegido: % nao pode ser alterado nem apagado', tg_table_name;
end $$;
drop trigger if exists ctr_historico_imutavel on public.ctr_historico;
create trigger ctr_historico_imutavel before update or delete on public.ctr_historico
  for each row execute function public._ctr_imutavel();
drop trigger if exists ctr_documentos_imutavel on public.ctr_documentos;
create trigger ctr_documentos_imutavel before update or delete on public.ctr_documentos
  for each row execute function public._ctr_imutavel();

-- ════════ Seguranca por linha (RLS) ════════
alter table public.ctr_naturezas    enable row level security;
alter table public.ctr_filiais      enable row level security;
alter table public.ctr_fornecedores enable row level security;
alter table public.ctr_acessos      enable row level security;
alter table public.ctr_contratos    enable row level security;
alter table public.ctr_liberados    enable row level security;
alter table public.ctr_documentos   enable row level security;
alter table public.ctr_historico    enable row level security;

revoke all on public.ctr_naturezas, public.ctr_filiais, public.ctr_fornecedores, public.ctr_acessos,
              public.ctr_contratos, public.ctr_liberados, public.ctr_documentos, public.ctr_historico from anon;
grant select, insert, update, delete on public.ctr_naturezas, public.ctr_filiais, public.ctr_fornecedores,
              public.ctr_acessos, public.ctr_contratos, public.ctr_liberados to authenticated;
grant select on public.ctr_documentos, public.ctr_historico to authenticated;
revoke insert, update, delete on public.ctr_documentos, public.ctr_historico from authenticated;
grant select, insert on public.ctr_documentos, public.ctr_historico to service_role;
grant select on public.ctr_contratos, public.ctr_naturezas, public.ctr_acessos, public.perfis, public.empresas to service_role;

-- cadastros do cliente: quem usa a empresa le; so admin grava
do $$
declare t text;
begin
  foreach t in array array['ctr_naturezas','ctr_filiais','ctr_fornecedores'] loop
    execute format('drop policy if exists %I on public.%I', t || '_ver', t);
    execute format('drop policy if exists %I on public.%I', t || '_admin', t);
    execute format('create policy %I on public.%I for select to authenticated using (public.ctr_usa_empresa(empresa))', t || '_ver', t);
    execute format('create policy %I on public.%I for all to authenticated using (public.eh_admin()) with check (public.eh_admin())', t || '_admin', t);
  end loop;
end $$;

drop policy if exists ctr_acessos_ver   on public.ctr_acessos;
drop policy if exists ctr_acessos_admin on public.ctr_acessos;
create policy ctr_acessos_ver   on public.ctr_acessos for select to authenticated using (usuario = auth.uid() or public.eh_admin());
create policy ctr_acessos_admin on public.ctr_acessos for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists ctr_liberados_ver   on public.ctr_liberados;
drop policy if exists ctr_liberados_admin on public.ctr_liberados;
create policy ctr_liberados_ver   on public.ctr_liberados for select to authenticated using (usuario = auth.uid() or public.eh_admin());
create policy ctr_liberados_admin on public.ctr_liberados for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists ctr_contratos_ver     on public.ctr_contratos;
drop policy if exists ctr_contratos_criar   on public.ctr_contratos;
drop policy if exists ctr_contratos_editar  on public.ctr_contratos;
drop policy if exists ctr_contratos_apagar  on public.ctr_contratos;
create policy ctr_contratos_ver on public.ctr_contratos for select to authenticated
  using (public.ctr_nivel(empresa, filial, natureza, confidencial, criado_por, id) >= 1);
create policy ctr_contratos_criar on public.ctr_contratos for insert to authenticated
  with check (public.ctr_nivel(empresa, filial, natureza, confidencial, auth.uid(), id) >= 3);
create policy ctr_contratos_editar on public.ctr_contratos for update to authenticated
  using      (public.ctr_nivel(empresa, filial, natureza, confidencial, criado_por, id) >= 3)
  with check (public.ctr_nivel(empresa, filial, natureza, confidencial, criado_por, id) >= 3);
create policy ctr_contratos_apagar on public.ctr_contratos for delete to authenticated
  using (public.eh_admin() and not exists (select 1 from public.ctr_documentos d where d.contrato = ctr_contratos.id));

drop policy if exists ctr_documentos_ver on public.ctr_documentos;
create policy ctr_documentos_ver on public.ctr_documentos for select to authenticated
  using (public.ctr_nivel_contrato(contrato) >= 1);

drop policy if exists ctr_historico_ver on public.ctr_historico;
create policy ctr_historico_ver on public.ctr_historico for select to authenticated
  using (case when contrato is null then public.eh_admin() else public.ctr_nivel_contrato(contrato) >= 1 end);
