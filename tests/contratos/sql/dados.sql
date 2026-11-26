insert into auth.users values ('00000000-0000-0000-0000-00000000000a'),('00000000-0000-0000-0000-00000000000b'),('00000000-0000-0000-0000-00000000000c'),('00000000-0000-0000-0000-00000000000d');
insert into empresas(slug,nome) values ('millena','Millena'),('novatec','Novatec');
insert into perfis values ('00000000-0000-0000-0000-00000000000a','Admin','a@x','admin',true),
 ('00000000-0000-0000-0000-00000000000b','Gestor','b@x','visualizador',true),
 ('00000000-0000-0000-0000-00000000000c','Leitor','c@x','visualizador',true),
 ('00000000-0000-0000-0000-00000000000d','Fora','d@x','visualizador',true);
grant select on empresas, perfis to authenticated;
