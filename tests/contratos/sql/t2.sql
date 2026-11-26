-- Testes do banco das partes novas: identidade, naturezas, contatos, contratos escolhidos e linha do tempo.
-- Rodar depois de sql/dados.sql e do SQL do sistema (o sql.sh faz isso). Banco limpo.
\set A '''00000000-0000-0000-0000-00000000000a'''
\set B '''00000000-0000-0000-0000-00000000000b'''
\set D '''00000000-0000-0000-0000-00000000000d'''
set role authenticated;
set test.uid = :A;
select id as alug from ctr_categorias where nome='Aluguel de imóvel comercial' \gset
select id as seg from ctr_categorias where nome='Vigilância e segurança' \gset
insert into ctr_unidades values ('millena','1','Matriz');

\echo --- identidade: Millena Moveis ja vem configurada; cor invalida e recusada
select nome_exibicao, usa_natureza, cor, tema from ctr_config where empresa='millena';
\set ON_ERROR_STOP 0
\echo esperado ERRO cor invalida
insert into ctr_config(empresa,cor) values ('novatec','vermelho');
\echo --- tela de entrada publica: so nome, logo, cor e tema
reset role; set role anon;
select nome, cor, tema from ctr_marca_publica('millena');
\echo esperado ERRO: anon nao le a tabela
select * from ctr_config;
reset role; set role authenticated; set test.uid = :A;

\echo --- naturezas: a categoria vem da natureza e os contratos acompanham
insert into ctr_naturezas(empresa,codigo,descricao,classe,tipo,tem_contrato,categoria) values ('millena','702','ALUGUEL DE IMOVEIS','7 - MANUTENCAO DE PREDIOS','DESPESA',true,:alug),('millena','721','SEGURANCA','7 - MANUTENCAO DE PREDIOS','DESPESA',true,null);
insert into ctr_contratos(empresa,unidade,natureza_cliente,titulo,fornecedor_nome,valor_mensal,inicio,fim,reajuste,mes_reajuste) values ('millena','1',' 702 ','Aluguel matriz','Loc',9000,'2024-03-01','2027-02-28','IGP-M',3) returning id as c1, categoria = :alug as categoria_da_natureza \gset
select :'categoria_da_natureza' as categoria_veio_da_natureza;
\echo esperado ERRO: natureza sem categoria e sem categoria informada
insert into ctr_contratos(empresa,unidade,natureza_cliente,titulo,fornecedor_nome) values ('millena','1','721','Seg','Vig');
update ctr_naturezas set categoria=:seg where codigo='702';
select categoria = :seg as contrato_acompanhou from ctr_contratos where id=:c1;
\echo esperado ERRO: categoria de outra empresa na natureza
insert into ctr_categorias(empresa,grupo,nome) values ('novatec','outros','So Novatec') returning id as nov \gset
update ctr_naturezas set categoria=:nov where codigo='721';

\echo --- linha do tempo: reajuste e aditivo atualizam o contrato; nada se apaga
insert into ctr_alteracoes(contrato,tipo,vigencia,valor_depois,indice,percentual,acumulado) values (:c1,'reajuste','2025-03-01',9378,'IGP-M',4.2,4.2) returning valor_antes, valor_depois;
insert into ctr_alteracoes(contrato,tipo,vigencia,fim_depois,observacao) values (:c1,'aditivo_prazo','2026-01-10','2029-02-28','Aditivo 1');
select valor_mensal, fim from ctr_contratos where id=:c1;
\echo esperado ERRO: reajuste sem percentual
insert into ctr_alteracoes(contrato,tipo,vigencia,valor_depois) values (:c1,'reajuste','2026-03-01',9500);
\echo esperado ERRO: alterar e apagar alteracao
update ctr_alteracoes set percentual=1;
reset role; delete from ctr_alteracoes; set role authenticated; set test.uid = :A;

\echo --- contatos e contratos escolhidos
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','1',:alug,'Outro','Y') returning id as c2 \gset
reset role;
insert into ctr_acessos(usuario,empresa,nivel,contratos) values (:B,'millena',3,array[:c1]::bigint[]);
insert into ctr_contatos(usuario,cargo,funcao,telefone,trocar_senha) values (:B,'Gerente','Loja 1','5581999990000',true);
\echo esperado ERRO telefone com letras
insert into ctr_contatos(usuario,telefone) values (:D,'81 9999');
set role authenticated; set test.uid = :B;
\echo gestor com so um contrato escolhido: ve 1
select count(*) as gestor_ve from ctr_contratos;
\echo esperado ERRO: com contratos escolhidos nao cria contrato novo
insert into ctr_contratos(empresa,unidade,categoria,titulo,fornecedor_nome) values ('millena','1',:alug,'Novo','W');
\echo le o proprio contato, nao muda o cargo (0 linhas), troca a senha
update ctr_contatos set cargo='Diretor' returning cargo;
select ctr_senha_trocada();
select cargo, trocar_senha from ctr_contatos;
\echo esperado ERRO: nivel 3 nao lanca alteracao em contrato que nao ve
insert into ctr_alteracoes(contrato,tipo,vigencia,valor_depois) values (:c2,'aditivo_valor','2026-05-01',1);
set test.uid = :D;
select count(*) as fora_ve_contatos from ctr_contatos;

\echo --- decisao no prazo de aviso: registra o ciclo, entra no historico e nao se apaga
set test.uid = :A;
insert into ctr_decisoes(contrato,decisao,responsavel,prazo,observacao) values (:c1,'renegociar',:B,'2026-12-10','Pedir 5% de desconto') returning decisao, ciclo_fim;
select acao, detalhe from ctr_historico where contrato=:c1 and acao='decisão registrada';
\echo esperado ERRO: alterar decisao
update ctr_decisoes set decisao='renovar';
\echo esperado ERRO: decisao invalida
insert into ctr_decisoes(contrato,decisao) values (:c1,'talvez');
set test.uid = :B;
\echo gestor (nivel 3, contrato escolhido) registra decisao no contrato dele
insert into ctr_decisoes(contrato,decisao) values (:c1,'em_analise') returning decisao;
\echo esperado ERRO: nao registra em contrato que nao ve
insert into ctr_decisoes(contrato,decisao) values (:c2,'renovar');
