-- ════════════════════════════════════════════════════════════════════════
-- GESTAO DE CONTRATOS — estrutura do banco (mesmo projeto Supabase do portal)
-- Cole tudo no SQL Editor do Supabase e clique em Run. Pode rodar de novo sem
-- problema: so cria o que ainda nao existe e refaz funcoes, regras e a lista
-- padrao de categorias.
--
-- Usa o mesmo login do portal (tabela perfis; admin = papel 'admin').
-- Tudo dos contratos comeca com "ctr_" e nao mexe nas tabelas do mapeamento.
-- Os ARQUIVOS ficam no SharePoint (site GestaoContratos); aqui so os dados.
--
-- Organizacao (igual para todo cliente, nao depende do sistema dele):
--   GRUPO (fixo, da Sinergia) > CATEGORIA (lista padrao + as que o admin criar
--   para a empresa) > CONTRATO. Cada categoria traz regras padrao: aviso previo,
--   indice de reajuste e documentos obrigatorios (com validade).
--   Os codigos do sistema do cliente (fornecedor, natureza) sao campos opcionais
--   do contrato: servem para conferir os pagamentos depois.
--
-- Quem ve o que (funcao ctr_nivel):
--   admin do portal ........................ tudo (nivel 3)
--   usuario com linha em ctr_acessos ....... o nivel dele (1 ver, 2 ver e baixar,
--                                            3 editar e enviar), so nas unidades e
--                                            grupos marcados (lista vazia = todos)
--   contrato confidencial .................. so admin, quem criou e quem estiver
--                                            em ctr_liberados
-- ════════════════════════════════════════════════════════════════════════

-- ── Grupos: fixos, so mudam por SQL (servem para relatorios e para comparar clientes) ──
create table if not exists public.ctr_grupos (
  codigo text primary key,
  nome   text not null,
  ordem  smallint not null default 0
);
insert into public.ctr_grupos (codigo, nome, ordem) values
  ('imoveis',       'Imóveis e ocupação',               1),
  ('terceirizacao', 'Mão de obra terceirizada',         2),
  ('manutencao',    'Manutenção e serviços prediais',   3),
  ('logistica',     'Logística e transporte',           4),
  ('tecnologia',    'Tecnologia e telecom',             5),
  ('marketing',     'Marketing e publicidade',          6),
  ('profissionais', 'Serviços profissionais',           7),
  ('veiculos',      'Veículos, máquinas e equipamentos', 8),
  ('financeiro',    'Seguros e serviços financeiros',   9),
  ('outros',        'Outros',                          10)
on conflict (codigo) do update set nome = excluded.nome, ordem = excluded.ordem;

-- ── Categorias: empresa vazia = lista padrao da Sinergia; com empresa = criada pelo admin ──
-- documentos: tipos de documento obrigatorios (alem do contrato assinado), ver TIPOS_DOC nas telas
create table if not exists public.ctr_categorias (
  id         bigint generated always as identity primary key,
  empresa    text references public.empresas(slug) on delete cascade on update cascade,
  grupo      text not null references public.ctr_grupos(codigo) on update cascade,
  nome       text not null check (length(trim(nome)) between 2 and 80),
  aviso_dias integer not null default 30 check (aviso_dias between 0 and 3650),
  reajuste   text check (reajuste in ('IPCA','IGP-M','INPC','Convenção coletiva','Fixo','Outro')),
  documentos text[] not null default '{}',
  dica       text,
  ativo      boolean not null default true
);
create unique index if not exists ctr_categorias_nome on public.ctr_categorias (coalesce(empresa, ''), lower(nome));

insert into public.ctr_categorias (empresa, grupo, nome, aviso_dias, reajuste, documentos, dica) values
  (null, 'imoveis', 'Aluguel de imóvel comercial', 240, 'IGP-M', '{garantia,avcb,alvara,seguro}',
   'Lei do Inquilinato (art. 51): a ação renovatória, que garante a renovação do ponto, só pode ser proposta entre 1 ano e 6 meses antes do fim do contrato. O aviso de 240 dias dá 2 meses para preparar.'),
  (null, 'imoveis', 'Aluguel de imóvel residencial', 60, 'IGP-M', '{garantia}', null),
  (null, 'imoveis', 'Estacionamento e garagem', 30, 'IPCA', '{}', null),
  (null, 'terceirizacao', 'Vigilância e segurança', 30, 'Convenção coletiva', '{cnd-federal,crf-fgts,cndt,autorizacao-pf}',
   'Se a terceirizada não pagar os funcionários, a sua empresa pode ter que pagar (Lei 6.019/74). Exija as certidões em dia. Vigilância armada ou patrimonial precisa de autorização da Polícia Federal.'),
  (null, 'terceirizacao', 'Limpeza e conservação', 30, 'Convenção coletiva', '{cnd-federal,crf-fgts,cndt}',
   'Se a terceirizada não pagar os funcionários, a sua empresa pode ter que pagar (Lei 6.019/74). Exija as certidões em dia.'),
  (null, 'terceirizacao', 'Portaria e recepção', 30, 'Convenção coletiva', '{cnd-federal,crf-fgts,cndt}',
   'Se a terceirizada não pagar os funcionários, a sua empresa pode ter que pagar (Lei 6.019/74). Exija as certidões em dia.'),
  (null, 'terceirizacao', 'Mão de obra temporária', 30, 'Convenção coletiva', '{cnd-federal,crf-fgts,cndt}', null),
  (null, 'manutencao', 'Manutenção predial', 30, 'IPCA', '{}', null),
  (null, 'manutencao', 'Ar-condicionado (PMOC)', 30, 'IPCA', '{}',
   'Prédios com ar-condicionado precisam de Plano de Manutenção (PMOC, Lei 13.589/2018). Guarde o plano junto com o contrato.'),
  (null, 'manutencao', 'Elevadores', 60, 'IPCA', '{}', null),
  (null, 'manutencao', 'Dedetização e controle de pragas', 30, 'IPCA', '{licenca}', null),
  (null, 'manutencao', 'Montagem de móveis', 30, 'Fixo', '{}',
   'Montador pessoa física que recebe todo mês, sempre igual e com horário definido, pode virar vínculo de emprego. Mantenha contrato escrito de prestação de serviço.'),
  (null, 'logistica', 'Frete e entregas', 30, 'IPCA', '{seguro}',
   'Freteiro autônomo pessoa física com pagamento fixo mensal é risco trabalhista. Prefira transportadora (empresa) ou contrato de transporte autônomo bem feito.'),
  (null, 'logistica', 'Transportadora', 30, 'IPCA', '{seguro}', null),
  (null, 'tecnologia', 'Sistema e licenças de software', 60, 'IPCA', '{}', null),
  (null, 'tecnologia', 'Internet e telefonia', 30, 'IPCA', '{}', 'Confira a multa de fidelidade antes de cancelar.'),
  (null, 'tecnologia', 'Suporte de informática', 30, 'IPCA', '{}', null),
  (null, 'marketing', 'Agência de publicidade', 30, 'IPCA', '{}', null),
  (null, 'marketing', 'Influenciador digital', 15, 'Fixo', '{}', null),
  (null, 'marketing', 'Carro de som e mídia local', 15, 'Fixo', '{}', null),
  (null, 'profissionais', 'Contabilidade', 60, 'IPCA', '{}', null),
  (null, 'profissionais', 'Assessoria jurídica', 60, 'IPCA', '{}', null),
  (null, 'profissionais', 'Consultoria', 30, 'IPCA', '{}', null),
  (null, 'profissionais', 'Medicina e segurança do trabalho', 30, 'IPCA', '{}', null),
  (null, 'veiculos', 'Locação de veículos', 30, 'IPCA', '{seguro}', null),
  (null, 'veiculos', 'Locação de máquinas e equipamentos', 30, 'IPCA', '{}', null),
  (null, 'veiculos', 'Manutenção de frota', 30, 'IPCA', '{}', null),
  (null, 'financeiro', 'Seguros', 30, 'Fixo', '{seguro}', 'Renove antes do fim da apólice para não ficar sem cobertura.'),
  (null, 'financeiro', 'Maquininha de cartão e serviços bancários', 30, 'Fixo', '{}', null),
  (null, 'outros', 'Outros contratos', 30, null, '{}', null)
on conflict (coalesce(empresa, ''), lower(nome)) do update
  set grupo = excluded.grupo, aviso_dias = excluded.aviso_dias, reajuste = excluded.reajuste,
      documentos = excluded.documentos, dica = excluded.dica;

-- ── Unidades (lojas, filiais, CDs) de cada empresa ──
create table if not exists public.ctr_unidades (
  empresa text not null references public.empresas(slug) on delete cascade on update cascade,
  codigo  text not null check (length(trim(codigo)) between 1 and 30),
  nome    text not null,
  primary key (empresa, codigo)
);

-- ── Identidade de cada empresa (o que a empresa ve: logo, cor e como chama as unidades) ──
-- Sem linha = identidade da Sinergia. A logo fica aqui mesmo (imagem pequena, ja reduzida pela tela).
create table if not exists public.ctr_config (
  empresa        text primary key references public.empresas(slug) on delete cascade on update cascade,
  nome_exibicao  text check (nome_exibicao is null or length(nome_exibicao) <= 60),
  cor            text check (cor is null or cor ~ '^#[0-9A-Fa-f]{6}$'),
  logo           text check (logo is null or (logo like 'data:image/%' and length(logo) <= 400000)),
  logo_fundo     boolean not null default false,       -- fundo branco atras da logo (logo escura)
  termo_unidade  text not null default 'Unidade'  check (length(trim(termo_unidade)) between 2 and 30),
  termo_unidades text not null default 'Unidades' check (length(trim(termo_unidades)) between 2 and 30),
  usa_natureza   boolean not null default false,       -- cadastro pela natureza do ERP da empresa (ver ctr_naturezas)
  menu_estilo    text not null default 'escuro' check (menu_estilo in ('escuro','claro','marca')),
  tema           text not null default 'auto'   check (tema in ('auto','claro','escuro')),
  alterado_em    timestamptz not null default now()
);
alter table public.ctr_config add column if not exists usa_natureza boolean not null default false;
alter table public.ctr_config add column if not exists menu_estilo text not null default 'escuro' check (menu_estilo in ('escuro','claro','marca'));
alter table public.ctr_config add column if not exists tema text not null default 'auto' check (tema in ('auto','claro','escuro'));

-- tela de entrada personalizada (link .../contratos/?e=<empresa>): so nome, logo e cor, antes do login
create or replace function public.ctr_marca_publica(p_empresa text)
returns table (nome text, logo text, cor text, logo_fundo boolean)
language sql stable security definer set search_path = public as $$
  select coalesce(c.nome_exibicao, e.nome), c.logo, c.cor, c.logo_fundo
    from ctr_config c join empresas e on e.slug = c.empresa
   where c.empresa = p_empresa and e.ativo;
$$;
revoke all on function public.ctr_marca_publica(text) from public;
grant execute on function public.ctr_marca_publica(text) to anon, authenticated;

-- ── Naturezas do ERP da empresa (ex.: Xadrex da Millena Moveis), importadas pelo admin ──
-- A empresa cadastra pela natureza (o padrao dela); cada natureza fica ligada a uma categoria da
-- Sinergia, que traz as regras (aviso, reajuste, documentos) e mantem os relatorios entre clientes.
create table if not exists public.ctr_naturezas (
  empresa      text not null references public.empresas(slug) on delete cascade on update cascade,
  codigo       text not null check (length(trim(codigo)) between 1 and 30),
  descricao    text not null,
  classe       text,
  tipo         text,                                -- DESPESA / RECEITA
  ativa        boolean not null default true,       -- desativada (no Xadrex, descricao com "#") nao aparece
  tem_contrato boolean not null default false,      -- aparece no cadastro de contratos
  categoria    bigint references public.ctr_categorias(id) on delete set null,
  revisada     boolean not null default false,      -- ligacao com a categoria conferida por uma pessoa
  primary key (empresa, codigo)
);

-- ── Quem pode usar os contratos de cada empresa ──
create table if not exists public.ctr_acessos (
  usuario   uuid     not null references public.perfis(id) on delete cascade,
  empresa   text     not null references public.empresas(slug) on delete cascade on update cascade,
  nivel     smallint not null default 1 check (nivel between 1 and 3),
  unidades  text[]   not null default '{}',   -- vazio = todas as unidades
  grupos    text[]   not null default '{}',   -- vazio = todos os grupos
  criado_em timestamptz not null default now(),
  primary key (usuario, empresa)
);

-- ── Contratos ──
create table if not exists public.ctr_contratos (
  id                   bigint generated always as identity primary key,
  numero               text unique,
  empresa              text not null references public.empresas(slug) on update cascade,
  unidade              text,
  categoria            bigint not null references public.ctr_categorias(id) on delete restrict,
  titulo               text not null,                    -- objeto do contrato
  fornecedor_nome      text not null,
  fornecedor_doc       text,                             -- CPF ou CNPJ, so letras e numeros
  fornecedor_codigo    text,                             -- codigo do fornecedor no sistema do cliente
  natureza_cliente     text,                             -- codigo(s) da natureza no sistema do cliente
  valor_mensal         numeric(14,2) check (valor_mensal is null or valor_mensal >= 0),
  inicio               date,
  fim                  date,                             -- vazio = prazo indeterminado
  aviso_dias           integer not null default 30 check (aviso_dias between 0 and 3650),
  reajuste             text check (reajuste in ('IPCA','IGP-M','INPC','Convenção coletiva','Fixo','Outro')),
  mes_reajuste         smallint check (mes_reajuste between 1 and 12),
  renovacao_automatica boolean not null default false,
  responsavel          uuid references public.perfis(id) on delete set null,
  confidencial         boolean not null default false,
  situacao             text not null default 'ativo' check (situacao in ('ativo','encerrado')),
  origem               text not null default 'manual',   -- manual / importado
  observacao           text,
  criado_por           uuid references public.perfis(id) on delete set null,
  criado_em            timestamptz not null default now(),
  alterado_em          timestamptz not null default now(),
  check (fim is null or inicio is null or fim >= inicio),
  foreign key (empresa, unidade) references public.ctr_unidades(empresa, codigo) on update cascade
);
create index if not exists ctr_contratos_empresa on public.ctr_contratos (empresa, fim);
create index if not exists ctr_contratos_categoria on public.ctr_contratos (categoria);

create table if not exists public.ctr_liberados (
  contrato bigint not null references public.ctr_contratos(id) on delete cascade,
  usuario  uuid   not null references public.perfis(id) on delete cascade,
  primary key (contrato, usuario)
);

-- Documentos: so a funcao do servidor (ctr-arquivos) grava, depois de conferir no SharePoint
create table if not exists public.ctr_documentos (
  id          bigint generated always as identity primary key,
  contrato    bigint not null references public.ctr_contratos(id) on delete restrict,
  categoria   text   not null,          -- tipo do documento (contrato-assinado, cndt, avcb...)
  validade    date,                     -- para certidoes, alvaras, seguros etc.
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
create or replace function public.ctr_nivel(p_empresa text, p_unidade text, p_categoria bigint,
                                            p_confidencial boolean, p_criado_por uuid, p_id bigint)
returns integer language plpgsql stable security definer set search_path = public as $$
declare
  uid uuid := auth.uid();
  a   public.ctr_acessos;
  grp text;
begin
  if uid is null then return 0; end if;
  if not exists (select 1 from perfis where id = uid and ativo) then return 0; end if;
  if public.eh_admin() then return 3; end if;
  select * into a from ctr_acessos where usuario = uid and empresa = p_empresa;
  if not found then return 0; end if;
  if cardinality(a.unidades) > 0 and (p_unidade is null or not (p_unidade = any (a.unidades))) then return 0; end if;
  if cardinality(a.grupos) > 0 then
    select grupo into grp from ctr_categorias where id = p_categoria;
    if grp is null or not (grp = any (a.grupos)) then return 0; end if;
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
  select coalesce((select public.ctr_nivel(c.empresa, c.unidade, c.categoria, c.confidencial, c.criado_por, c.id)
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

revoke all on function public.ctr_nivel(text,text,bigint,boolean,uuid,bigint) from public, anon;
revoke all on function public.ctr_nivel_contrato(bigint) from public, anon;
revoke all on function public.ctr_usa_empresa(text) from public, anon;
grant execute on function public.ctr_nivel(text,text,bigint,boolean,uuid,bigint) to authenticated, service_role;
grant execute on function public.ctr_nivel_contrato(bigint) to authenticated, service_role;
grant execute on function public.ctr_usa_empresa(text) to authenticated, service_role;

-- ════════ Gatilhos dos contratos: numero, categoria valida e historico automatico ════════
create or replace function public._ctr_contratos_antes() returns trigger
language plpgsql set search_path = public as $$
declare cat public.ctr_categorias; natcat bigint;
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
  -- empresa que cadastra pela natureza: a categoria vem da natureza (sempre coerentes)
  new.natureza_cliente := nullif(trim(coalesce(new.natureza_cliente, '')), '');
  if new.natureza_cliente is not null then
    select categoria into natcat from ctr_naturezas where empresa = new.empresa and codigo = new.natureza_cliente;
    if natcat is not null then new.categoria := natcat; end if;
  end if;
  -- a categoria tem que ser da lista padrao ou desta empresa, e estar ativa
  if tg_op = 'INSERT' or new.categoria is distinct from old.categoria then
    select * into cat from ctr_categorias where id = new.categoria;
    if not found or (cat.empresa is not null and cat.empresa <> new.empresa) then
      raise exception 'categoria invalida para esta empresa';
    end if;
    if not cat.ativo then raise exception 'categoria desativada: escolha outra'; end if;
  end if;
  new.fornecedor_doc := nullif(upper(regexp_replace(coalesce(new.fornecedor_doc, ''), '[^A-Za-z0-9]', '', 'g')), '');
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
            case when new.origem = 'importado' then 'importado da planilha modelo' end);
    return null;
  end if;
  velho := to_jsonb(old);
  foreach k in array array['titulo','fornecedor_nome','fornecedor_doc','fornecedor_codigo','unidade','categoria',
                           'natureza_cliente','valor_mensal','inicio','fim','aviso_dias','reajuste','mes_reajuste',
                           'renovacao_automatica','responsavel','confidencial','situacao','observacao'] loop
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

-- natureza: a categoria ligada tem que valer para a empresa; mudou a ligacao = contratos acompanham
create or replace function public._ctr_naturezas_antes() returns trigger
language plpgsql set search_path = public as $$
begin
  if new.categoria is not null and not exists (select 1 from ctr_categorias c where c.id = new.categoria and (c.empresa is null or c.empresa = new.empresa)) then
    raise exception 'categoria invalida para esta empresa';
  end if;
  return new;
end $$;
drop trigger if exists ctr_naturezas_antes on public.ctr_naturezas;
create trigger ctr_naturezas_antes before insert or update on public.ctr_naturezas
  for each row execute function public._ctr_naturezas_antes();

create or replace function public._ctr_naturezas_depois() returns trigger
language plpgsql set search_path = public as $$
begin
  if new.categoria is not null and new.categoria is distinct from old.categoria then
    update ctr_contratos set categoria = new.categoria
     where empresa = new.empresa and natureza_cliente = new.codigo and categoria is distinct from new.categoria;
  end if;
  return null;
end $$;
drop trigger if exists ctr_naturezas_depois on public.ctr_naturezas;
create trigger ctr_naturezas_depois after update on public.ctr_naturezas
  for each row execute function public._ctr_naturezas_depois();

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
alter table public.ctr_grupos       enable row level security;
alter table public.ctr_categorias   enable row level security;
alter table public.ctr_unidades     enable row level security;
alter table public.ctr_config       enable row level security;
alter table public.ctr_naturezas    enable row level security;
alter table public.ctr_acessos      enable row level security;
alter table public.ctr_contratos    enable row level security;
alter table public.ctr_liberados    enable row level security;
alter table public.ctr_documentos   enable row level security;
alter table public.ctr_historico    enable row level security;

revoke all on public.ctr_grupos, public.ctr_categorias, public.ctr_unidades, public.ctr_config, public.ctr_naturezas, public.ctr_acessos,
              public.ctr_contratos, public.ctr_liberados, public.ctr_documentos, public.ctr_historico from anon;
grant select on public.ctr_grupos to authenticated;
revoke insert, update, delete on public.ctr_grupos from authenticated;
grant select, insert, update, delete on public.ctr_categorias, public.ctr_unidades, public.ctr_config, public.ctr_naturezas,
              public.ctr_acessos, public.ctr_contratos, public.ctr_liberados to authenticated;
grant select on public.ctr_documentos, public.ctr_historico to authenticated;
revoke insert, update, delete on public.ctr_documentos, public.ctr_historico from authenticated;
grant select, insert on public.ctr_documentos, public.ctr_historico to service_role;
grant select on public.ctr_contratos, public.ctr_categorias, public.ctr_grupos, public.ctr_naturezas, public.ctr_acessos,
              public.perfis, public.empresas to service_role;

drop policy if exists ctr_grupos_ver on public.ctr_grupos;
create policy ctr_grupos_ver on public.ctr_grupos for select to authenticated using (true);

-- categorias: a lista padrao todos leem; as da empresa, quem usa a empresa. So admin cria/altera,
-- e so as da empresa (a lista padrao so muda por este SQL).
drop policy if exists ctr_categorias_ver   on public.ctr_categorias;
drop policy if exists ctr_categorias_admin on public.ctr_categorias;
create policy ctr_categorias_ver on public.ctr_categorias for select to authenticated
  using (empresa is null or public.ctr_usa_empresa(empresa));
create policy ctr_categorias_admin on public.ctr_categorias for all to authenticated
  using (public.eh_admin() and empresa is not null) with check (public.eh_admin() and empresa is not null);

drop policy if exists ctr_unidades_ver   on public.ctr_unidades;
drop policy if exists ctr_unidades_admin on public.ctr_unidades;
create policy ctr_unidades_ver   on public.ctr_unidades for select to authenticated using (public.ctr_usa_empresa(empresa));
create policy ctr_unidades_admin on public.ctr_unidades for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists ctr_naturezas_ver   on public.ctr_naturezas;
drop policy if exists ctr_naturezas_admin on public.ctr_naturezas;
create policy ctr_naturezas_ver   on public.ctr_naturezas for select to authenticated using (public.ctr_usa_empresa(empresa));
create policy ctr_naturezas_admin on public.ctr_naturezas for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

drop policy if exists ctr_config_ver   on public.ctr_config;
drop policy if exists ctr_config_admin on public.ctr_config;
create policy ctr_config_ver   on public.ctr_config for select to authenticated using (public.ctr_usa_empresa(empresa));
create policy ctr_config_admin on public.ctr_config for all    to authenticated using (public.eh_admin()) with check (public.eh_admin());

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
  using (public.ctr_nivel(empresa, unidade, categoria, confidencial, criado_por, id) >= 1);
create policy ctr_contratos_criar on public.ctr_contratos for insert to authenticated
  with check (public.ctr_nivel(empresa, unidade, categoria, confidencial, auth.uid(), id) >= 3);
create policy ctr_contratos_editar on public.ctr_contratos for update to authenticated
  using      (public.ctr_nivel(empresa, unidade, categoria, confidencial, criado_por, id) >= 3)
  with check (public.ctr_nivel(empresa, unidade, categoria, confidencial, criado_por, id) >= 3);
create policy ctr_contratos_apagar on public.ctr_contratos for delete to authenticated
  using (public.eh_admin() and not exists (select 1 from public.ctr_documentos d where d.contrato = ctr_contratos.id));

drop policy if exists ctr_documentos_ver on public.ctr_documentos;
create policy ctr_documentos_ver on public.ctr_documentos for select to authenticated
  using (public.ctr_nivel_contrato(contrato) >= 1);

drop policy if exists ctr_historico_ver on public.ctr_historico;
create policy ctr_historico_ver on public.ctr_historico for select to authenticated
  using (case when contrato is null then public.eh_admin() else public.ctr_nivel_contrato(contrato) >= 1 end);

-- ════════ Millena Moveis: nome, "Filial" e cadastro pela natureza do Xadrex ════════
-- (so cria se ainda nao existir; depois quem muda e o admin, em Identidade da empresa)
insert into public.ctr_config (empresa, nome_exibicao, termo_unidade, termo_unidades, usa_natureza)
select 'millena', 'Millena Móveis', 'Filial', 'Filiais', true
 where exists (select 1 from public.empresas where slug = 'millena')
on conflict (empresa) do nothing;

