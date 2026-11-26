\set A '''00000000-0000-0000-0000-00000000000a'''
\set B '''00000000-0000-0000-0000-00000000000b'''
\set C '''00000000-0000-0000-0000-00000000000c'''
\set D '''00000000-0000-0000-0000-00000000000d'''
select id as aluguel from ctr_categorias where nome='Aluguel de imóvel comercial' \gset
select id as seg from ctr_categorias where nome='Vigilância e segurança' \gset
set role authenticated;
\echo --- admin cria unidades, categoria propria, acessos
set test.uid = :A;
insert into ctr_unidades values ('millena','1','Matriz'),('millena','2','Loja 2'),('novatec','1','N1');
insert into ctr_categorias(empresa,grupo,nome,aviso_dias,reajuste,documentos) values ('millena','veiculos','Locação de empilhadeira',30,'IPCA','{seguro}') returning id as emp_cat \gset
insert into ctr_categorias(empresa,grupo,nome) values ('novatec','outros','So da Novatec') returning id as nov_cat \gset
insert into ctr_acessos(usuario,empresa,nivel,unidades,grupos) values (:B,'millena',3,'{1}','{}'),(:C,'millena',1,'{}','{imoveis}');
\echo esperado ERRO: admin altera categoria padrao
\set ON_ERROR_STOP 0
update ctr_categorias set aviso_dias=1 where id=:aluguel returning id;
\echo --- gestor (nivel 3, unidade 1)
set test.uid = :B;
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome,fornecedor_doc,valor_mensal) values ('millena','1',:aluguel,'Aluguel matriz','Locadora X','11.222.333/0001-81',1000) returning id, numero, fornecedor_doc;
\echo esperado ERRO (unidade 2):
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','2',:seg,'Seg loja 2','Vig');
\echo esperado ERRO (categoria da Novatec):
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','1',:nov_cat,'x','y');
\echo esperado ERRO (unidade inexistente):
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','9',:seg,'x','y');
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','1',:emp_cat,'Empilhadeira','Loc') returning id;
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','1',:seg,'Vigilancia matriz','Vig') returning id;
\echo esperado ERRO: gestor cria categoria
insert into ctr_categorias(empresa,grupo,nome) values ('millena','outros','Gestor cria');
select count(*) as cats_visiveis_gestor from ctr_categorias;
\echo --- leitor (so imoveis): ve 1
set test.uid = :C;
select titulo from ctr_contratos;
\echo esperado 0 linhas (nivel 1 nao edita):
update ctr_contratos set valor_mensal=1 returning id;
\echo --- fora: nada
set test.uid = :D;
select count(*) from ctr_contratos; select count(*) as cats_fora from ctr_categorias;
\echo --- admin: tudo; historico; desativar categoria
set test.uid = :A;
update ctr_contratos set valor_mensal=1500, mes_reajuste=3 where titulo='Aluguel matriz';
select acao, detalhe from ctr_historico order by id;
update ctr_categorias set ativo=false where id=:emp_cat;
\echo esperado ERRO: trocar para categoria desativada
update ctr_contratos set categoria=:emp_cat where titulo='Vigilancia matriz';
\echo esperado ERRO: apagar categoria em uso
delete from ctr_categorias where id=:emp_cat;
reset role;
\echo --- service role grava documento com validade
set role service_role;
insert into ctr_documentos(contrato,categoria,validade,nome,sp_item) select id,'cndt','2027-01-01','a.pdf','sp1' from ctr_contratos where titulo='Vigilancia matriz';
select categoria, validade from ctr_documentos;
