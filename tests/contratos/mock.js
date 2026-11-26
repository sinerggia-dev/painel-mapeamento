<script>(function(){
// ═══ PROTÓTIPO: banco e SharePoint simulados, dados inventados, nada é salvo ═══
var SEED=__SEED__;
var PAPEL='admin'; try{PAPEL=sessionStorage.getItem('proto-papel')||'admin'}catch(e){}
var LOGADO=false; try{LOGADO=sessionStorage.getItem('proto-logado')==='1'}catch(e){}
var USERS={admin:{id:'u-adm',nome:'Natanael Silva',email:'admin@sinergia.exemplo',papel:'admin'},
  gestor:{id:'u-ges',nome:'Carla Mendes',email:'carla@millena.exemplo',papel:'visualizador'},
  leitor:{id:'u-lei',nome:'Rafael Lima',email:'rafael@millena.exemplo',papel:'visualizador'},
  primeiro:{id:'u-pri',nome:'Bruna Costa',email:'bruna@millena.exemplo',papel:'visualizador'}};
var U=USERS[PAPEL]||USERS.admin;
var hoje=new Date(); function d(n){var x=new Date(hoje);x.setDate(x.getDate()+n);return x.getFullYear()+'-'+String(x.getMonth()+1).padStart(2,'0')+'-'+String(x.getDate()).padStart(2,'0')}
var catSeq=1;
var DB={
 perfis:[{id:'u-adm',nome:'Natanael Silva',email:'admin@sinergia.exemplo',papel:'admin',ativo:true},
  {id:'u-ges',nome:'Carla Mendes',email:'carla@millena.exemplo',papel:'visualizador',ativo:true},
  {id:'u-lei',nome:'Rafael Lima',email:'rafael@millena.exemplo',papel:'visualizador',ativo:true},
  {id:'u-jur',nome:'Juliana Rocha',email:'juliana@millena.exemplo',papel:'visualizador',ativo:true},
  {id:'u-fin',nome:'Paulo Freitas',email:'paulo@millena.exemplo',papel:'visualizador',ativo:false},
  {id:'u-pri',nome:'Bruna Costa',email:'bruna@millena.exemplo',papel:'visualizador',ativo:true}],
 empresas:[{slug:'millena',nome:'Millena Móveis'},{slug:'novatec',nome:'Novatec'}],
 ctr_grupos:SEED.grupos,
 ctr_categorias:SEED.cats.map(function(c){return Object.assign({id:catSeq++,empresa:null,ativo:true},c)}),
 ctr_acessos:[{usuario:'u-ges',empresa:'millena',nivel:3,unidades:['1','2','3','4','5','6'],grupos:[],contratos:[]},
  {usuario:'u-lei',empresa:'millena',nivel:1,unidades:[],grupos:['imoveis'],contratos:[]},
  {usuario:'u-jur',empresa:'millena',nivel:2,unidades:[],grupos:[],contratos:[]},
  {usuario:'u-pri',empresa:'millena',nivel:2,unidades:[],grupos:[],contratos:[5,6]}],
 ctr_contatos:[{usuario:'u-ges',cargo:'Gerente administrativa',funcao:'Contratos das filiais 1 a 6',telefone:'5581999990001',trocar_senha:false},
  {usuario:'u-lei',cargo:'Analista de manutenção',funcao:'Aluguéis',telefone:null,trocar_senha:false},
  {usuario:'u-pri',cargo:'Supervisora de segurança',funcao:'Vigilância e limpeza',telefone:'5581999990002',trocar_senha:true}],
 ctr_unidades:['Matriz (exemplo)','Loja Centro (exemplo)','Loja Shopping (exemplo)','Loja Norte (exemplo)','Loja Sul (exemplo)','CD Logística (exemplo)','Loja Litoral (exemplo)','Loja Agreste (exemplo)']
  .map(function(n,i){return {empresa:'millena',codigo:String(i+1),nome:n}}),
 ctr_contratos:[], ctr_documentos:[], ctr_historico:[], ctr_liberados:[], ctr_alteracoes:[],
 ctr_config:[{empresa:'millena',nome_exibicao:'Millena Móveis',cor:'#F8B900',logo:__LOGO__,logo_fundo:false,termo_unidade:'Filial',termo_unidades:'Filiais',usa_natureza:true,menu_estilo:'escuro',tema:'escuro'}],
 ctr_naturezas:[]
};
['Sede Novatec (exemplo)','Unidade Recife (exemplo)','Unidade Aracaju (exemplo)'].forEach(function(n,i){DB.ctr_unidades.push({empresa:'novatec',codigo:'N'+(i+1),nome:n})});
DB.ctr_categorias.push({id:catSeq++,empresa:'millena',grupo:'veiculos',nome:'Locação de empilhadeira',aviso_dias:30,reajuste:'IPCA',documentos:['seguro'],dica:'Categoria criada só para esta empresa (exemplo).',ativo:true});
function cat(nome){return (DB.ctr_categorias.find(function(c){return c.nome===nome})||{}).id}
[['702','ALUGUEL DE IMOVEIS','7 - MANUTENCAO DE PREDIOS','Aluguel de imóvel comercial',1],['703','# ALUGUEL DE IMOVEIS','7 - MANUTENCAO DE PREDIOS',null,0],
 ['706','ALUGUEL ESCRITORIO','7 - MANUTENCAO DE PREDIOS','Aluguel de imóvel comercial',0],['722','ALUGUEL CASA EXEMPLO','7 - MANUTENCAO DE PREDIOS','Aluguel de imóvel residencial',0],
 ['721','SEGURANCA','7 - MANUTENCAO DE PREDIOS','Vigilância e segurança',1],['727','# SEGURANCA','7 - MANUTENCAO DE PREDIOS',null,0],['742','CONDOMINIO','7 - MANUTENCAO DE PREDIOS',null,0],
 ['744','LIMPEZA E CONSERVACAO','7 - MANUTENCAO DE PREDIOS','Limpeza e conservação',1],['743','#SEGUROS','7 - MANUTENCAO DE PREDIOS',null,0],
 ['113','MONTAGEM','1 - FORNECEDORES','Montagem de móveis',1],['102','FRETE DE CLIENTES','1 - FORNECEDORES','Frete e entregas',1],
 ['500','TELEFONIA FIXA','5 - SISTEMA DE TELEFONIA','Internet e telefonia',0],['1001','MANUT. SISTEMA INFORMATICA','10 - SISTEMAS DE INFORMATICA','Sistema e licenças de software',1],
 ['1405','LOCACAO MAQ E EQUIPAMENTOS','14 - EQUIPAMENTOS E MAQ. DIVERSOS','Locação de máquinas e equipamentos',1],['1501','SERVICO ADVOGADO','15 - SERVICO DE PROFIS. LIBERAIS','Assessoria jurídica',1],
 ['319','EXAMES MEDICOS','3 - DESPESAS COM FUNCIONARIOS','Medicina e segurança do trabalho',1],['1600','SERVICO AG. DE PROPAGANDA','16 - PUBLICIDADE E PROPAGANDA','Agência de publicidade',0],
 ['1629','INFLUENCER DIGITAL','16 - PUBLICIDADE E PROPAGANDA','Influenciador digital',0],['744S','SEGURO DAS LOJAS','7 - MANUTENCAO DE PREDIOS','Seguros',1],
 ['808','SERVICO DE ELETRICA','8 - REPARO E MELHORIA DE PREDIOS',null,0],['212','PENSAO ALIMENTICIA','2 - FOLHA DE PAGAMENTO',null,0],['692','TAXA DE MEIO AMBIENTE','6 - IMPOSTOS E TAXAS',null,0],
 ['1830','VENDA CARTAO','18 - RECEITA DE VENDAS',null,0]]
 .forEach(function(n){var ativa=n[1].charAt(0)!=='#';DB.ctr_naturezas.push({empresa:'millena',codigo:n[0],descricao:n[1],classe:n[2],tipo:/RECEITA/.test(n[2])?'RECEITA':'DESPESA',ativa:ativa,tem_contrato:!!n[3]&&ativa,categoria:n[3]?cat(n[3]):null,revisada:!!n[4]})});
function natCat(emp,cod){var n=DB.ctr_naturezas.find(function(x){return x.empresa===emp&&x.codigo===cod});return n&&n.categoria}
var seq=1, docSeq=1, altSeq=1;
function slug(s){return String(s).normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^A-Za-z0-9]+/g,'-').replace(/^-|-$/g,'')}
function novoContrato(o){var id=seq++;var c=Object.assign({id:id,numero:'CTR-2026-'+String(id).padStart(5,'0'),empresa:'millena',aviso_dias:30,renovacao_automatica:false,tipo_contrato:'novo',confidencial:false,situacao:'ativo',origem:'manual',criado_em:new Date(Date.now()-86400000*(40-id)).toISOString(),criado_por:U.id,responsavel:null,inicio:null,fim:null,valor_mensal:null,reajuste:null,mes_reajuste:null,observacao:null,unidade:null,categoria:null,fornecedor_doc:null,fornecedor_codigo:null,natureza_cliente:null},o);
 if(c.fornecedor_doc)c.fornecedor_doc=String(c.fornecedor_doc).toUpperCase().replace(/[^A-Z0-9]/g,'')||null;
 if(c.natureza_cliente){c.natureza_cliente=String(c.natureza_cliente).trim();var nc=natCat(c.empresa,c.natureza_cliente);if(nc)c.categoria=nc}
 DB.ctr_contratos.push(c);DB.ctr_historico.push({contrato:id,acao:'contrato cadastrado',detalhe:c.origem==='importado'?'importado da planilha modelo':null,usuario:c.criado_por,em:c.criado_em});return c}
function catNome(id){return (DB.ctr_categorias.find(function(x){return x.id===id})||{nome:'Sem categoria'}).nome}
function doc(c,tipo,ext,tam,validade){DB.ctr_documentos.push({id:docSeq++,contrato:c.id,categoria:tipo,validade:validade||null,nome:c.numero+'_'+slug(catNome(c.categoria)).slice(0,30)+'_'+slug(c.fornecedor_nome)+'_'+tipo+'_v1.'+ext,tamanho:tam,versao:1,enviado_por:'u-ges',enviado_em:new Date(Date.now()-86400000*5).toISOString()})}
// [objeto, fornecedor, CNPJ/CPF (inventado e valido), codigo no sistema, unidade, categoria, valor, inicio, fim, reajuste, contrato assinado?]
var M=[
 ['Aluguel da loja Centro','Imobiliária Exemplo Ltda','11.222.333/0001-81','IMB001','2','Aluguel de imóvel comercial',9800,-700,35,'IGP-M',1],
 ['Aluguel da loja Shopping','Shopping Fictício S/A','11.444.777/0001-61','SHP002','3','Aluguel de imóvel comercial',15400,-300,420,'IPCA',1],
 ['Aluguel do CD de logística','Galpões Modelo Ltda','12.345.678/0001-95','GAL003','6','Aluguel de imóvel comercial',32000,-900,-12,'IGP-M',1],
 ['Aluguel da loja Norte','Locador Pessoa Física A','529.982.247-25','PF0041','4','Aluguel de imóvel comercial',4200,-400,150,'IGP-M',0],
 ['Vigilância patrimonial','Segurança Exemplar Ltda','33.000.167/0001-01','SEG010','1','Vigilância e segurança',21500,-200,58,'Convenção coletiva',1],
 ['Limpeza das lojas','Limpa Bem Serviços (exemplo)','60.701.190/0001-04','LIM011','1','Limpeza e conservação',12800,-150,215,'Convenção coletiva',1],
 ['ERP e licenças','SoftDemo Tecnologia','27.865.757/0001-02','SIS020','1','Sistema e licenças de software',5200,-250,80,'IPCA',1],
 ['Assessoria jurídica trabalhista','Advocacia Hipotética','04.252.011/0001-10','ADV030','1','Assessoria jurídica',7000,-60,300,'Fixo',1],
 ['Montagem de móveis – Loja Sul','Montador Autônomo Exemplo','111.444.777-35','MON040','5','Montagem de móveis',2600,-120,240,'Fixo',0],
 ['Exames ocupacionais','Clínica Ocupacional Teste','45.997.418/0001-53','MED050','1','Medicina e segurança do trabalho',1900,-330,30,'IPCA',1],
 ['Empilhadeira do CD','Locadora de Máquinas Teste','07.526.557/0001-00','EMP060','6','Locação de empilhadeira',3400,-90,275,'IPCA',1],
 ['Seguro das lojas','Seguradora Exemplo S/A','61.198.164/0001-60','SEGU70','1','Seguros',2100,-200,165,'Fixo',1],
 ['Frete de entregas – Litoral','Transportes Inventados Ltda','10.441.226/0001-07','FRT080','7','Frete e entregas',6400,-500,null,'IPCA',0]
];
M.forEach(function(m){var c=novoContrato({titulo:m[0],fornecedor_nome:m[1],fornecedor_doc:m[2],fornecedor_codigo:m[3],unidade:m[4],categoria:cat(m[5]),valor_mensal:m[6],inicio:d(m[7]),fim:m[8]===null?null:d(m[8]),reajuste:m[9],mes_reajuste:m[9]==='Fixo'?null:new Date(d(m[7])+'T12:00').getMonth()+1,aviso_dias:(DB.ctr_categorias.find(function(x){return x.nome===m[5]})||{}).aviso_dias||30,responsavel:'u-ges',criado_por:'u-ges'});if(m[10])doc(c,'contrato-assinado','pdf',1800000)});
[['Aluguel da sede','Imobiliária Norte (exemplo)','N1','Aluguel de imóvel comercial',18500,-500,95,'IGP-M'],
 ['Limpeza da sede','Conserva Tudo (exemplo)','N1','Limpeza e conservação',9400,-100,265,'Convenção coletiva'],
 ['Internet das unidades','Telecom Exemplo S/A','N2','Internet e telefonia',1800,-400,-5,'IPCA']]
 .forEach(function(m){var c=novoContrato({empresa:'novatec',titulo:m[0],fornecedor_nome:m[1],unidade:m[2],categoria:cat(m[3]),valor_mensal:m[4],inicio:d(m[5]),fim:d(m[6]),reajuste:m[7],criado_por:'u-adm'});doc(c,'contrato-assinado','pdf',900000)});
var C=DB.ctr_contratos.filter(function(c){return c.empresa==='millena'});
var NATC={'Aluguel de imóvel comercial':'702','Vigilância e segurança':'721','Limpeza e conservação':'744','Sistema e licenças de software':'1001','Assessoria jurídica':'1501','Montagem de móveis':'113','Medicina e segurança do trabalho':'319','Locação de empilhadeira':'1405','Seguros':'744S','Frete e entregas':'102'};
C.forEach(function(c){var nm=catNome(c.categoria);if(NATC[nm])c.natureza_cliente=NATC[nm];if(nm==='Locação de empilhadeira'){c.categoria=cat('Locação de máquinas e equipamentos')}});
doc(C[0],'garantia','pdf',300000,d(120)); doc(C[0],'avcb','pdf',500000,d(-20)); doc(C[0],'alvara','pdf',220000,d(200)); doc(C[0],'seguro','pdf',410000,d(18));
doc(C[1],'garantia','pdf',300000,d(400)); doc(C[1],'avcb','pdf',500000,d(300)); doc(C[1],'alvara','pdf',220000,d(250)); doc(C[1],'seguro','pdf',410000,d(90));
doc(C[4],'cnd-federal','pdf',90000,d(100)); doc(C[4],'crf-fgts','pdf',80000,d(-3)); doc(C[4],'cndt','pdf',85000,d(25)); doc(C[4],'autorizacao-pf','pdf',150000,d(210));
doc(C[5],'cnd-federal','pdf',90000,d(140)); doc(C[5],'crf-fgts','pdf',80000,d(12)); doc(C[5],'cndt','pdf',85000,d(160));
doc(C[10],'seguro','pdf',300000,d(60)); doc(C[11],'seguro','pdf',300000,d(165));
doc(C[0],'aditivo','pdf',420000); doc(C[1],'anexo','xlsx',96000); doc(C[4],'correspondencia','msg',210000);
C[7].confidencial=true;
C[1].tipo_contrato='renovacao';
// a mesma empresa de seguranca tambem cuida da loja 2: as certidoes enviadas no primeiro contrato valem para este
var cv=novoContrato({titulo:'Vigilância da loja Centro',fornecedor_nome:C[4].fornecedor_nome,fornecedor_doc:C[4].fornecedor_doc,fornecedor_codigo:C[4].fornecedor_codigo,unidade:'2',categoria:C[4].categoria,valor_mensal:9800,inicio:d(-100),fim:d(265),reajuste:'Convenção coletiva',aviso_dias:30,responsavel:'u-ges',criado_por:'u-ges',natureza_cliente:'721'});doc(cv,'contrato-assinado','pdf',900000);
DB.ctr_alteracoes.push({id:altSeq++,contrato:C[5].id,tipo:'aditivo_prazo',vigencia:d(-40),valor_antes:C[5].valor_mensal,valor_depois:null,fim_antes:d(150),fim_depois:C[5].fim,observacao:'Inclusão da limpeza da loja Sul e prorrogação',criado_por:'u-ges',criado_em:new Date().toISOString()});
doc(C[12],'distrato','pdf',150000);DB.ctr_alteracoes.push({id:altSeq++,contrato:C[12].id,tipo:'outro',vigencia:d(-10),valor_antes:C[12].valor_mensal,valor_depois:null,fim_antes:null,fim_depois:null,documento:docSeq-1,observacao:'Distrato: encerrado por acordo entre as partes',criado_por:'u-ges',criado_em:new Date().toISOString()});C[12].situacao='encerrado';
(function(){var c=C[0],ini=new Date(c.inicio+'T12:00');var vig=new Date(ini.getFullYear()+1,ini.getMonth(),ini.getDate());var v=vig.getFullYear()+'-'+String(vig.getMonth()+1).padStart(2,'0')+'-'+String(vig.getDate()).padStart(2,'0');DB.ctr_alteracoes.push({id:altSeq++,contrato:c.id,tipo:'reajuste',vigencia:v,valor_antes:9405,valor_depois:9800,fim_antes:c.fim,indice:'IGP-M',percentual:4.2,acumulado:4.2,criado_por:'u-ges',criado_em:new Date().toISOString()})})();
DB.ctr_historico.push({contrato:1,acao:'dados alterados',detalhe:'valor_mensal: 9200.00 → 9800.00; reajuste: — → IGP-M',usuario:'u-ges',em:new Date(Date.now()-86400000*3).toISOString()});
// regras de acesso simuladas (no sistema real quem faz isso é o banco)
function acesso(emp){return DB.ctr_acessos.find(function(a){return a.usuario===U.id&&a.empresa===(emp||'millena')})}
function grupo(c){var k=DB.ctr_categorias.find(function(x){return x.id===c.categoria});return k&&k.grupo}
function nivel(c){if(U.papel==='admin')return 3;var a=acesso(c.empresa);if(!a)return 0;
 if(a.contratos&&a.contratos.length&&a.contratos.indexOf(c.id)<0)return 0;
 if(a.unidades.length&&a.unidades.indexOf(c.unidade)<0)return 0;if(a.grupos.length&&a.grupos.indexOf(grupo(c))<0)return 0;
 if(c.confidencial&&c.criado_por!==U.id&&!DB.ctr_liberados.some(function(l){return l.contrato===c.id&&l.usuario===U.id}))return 0;return a.nivel}
function visivel(t,r){if(U.papel==='admin')return true;
 if(t==='ctr_contratos')return nivel(r)>0;
 if(t==='ctr_documentos'||t==='ctr_historico'||t==='ctr_alteracoes'){var c=DB.ctr_contratos.find(function(x){return x.id===r.contrato});return c&&nivel(c)>0}
 if(t==='ctr_acessos'||t==='ctr_liberados'||t==='ctr_contatos')return r.usuario===U.id;
 if(t==='perfis')return r.id===U.id;
 if(t==='ctr_config'||t==='ctr_unidades'||t==='ctr_naturezas')return !!acesso(r.empresa);return true}
function erro(m){return {data:null,error:{message:m}}}
function validarContrato(c){if(c.natureza_cliente){var nc=natCat(c.empresa,String(c.natureza_cliente).trim());if(nc)c.categoria=nc}var k=DB.ctr_categorias.find(function(x){return x.id===c.categoria});
 if(!k||(k.empresa&&k.empresa!==c.empresa))return 'categoria invalida para esta empresa';if(!k.ativo)return 'categoria desativada: escolha outra';
 if(c.unidade&&!DB.ctr_unidades.some(function(u){return u.empresa===c.empresa&&u.codigo===c.unidade}))return 'insert or update on table "ctr_contratos" violates foreign key constraint "ctr_contratos_empresa_unidade_fkey"';return null}
function Q(t){this.t=t;this.f=[];this.op='select';this.p=null;this.one=0}
Q.prototype.select=function(){return this};Q.prototype.eq=function(k,v){this.f.push(function(r){return r[k]===v});return this};
Q.prototype.in=function(k,v){this.f.push(function(r){return v.indexOf(r[k])>=0});return this};
Q.prototype.or=function(s){var m=/empresa\.eq\.([^,]+)/.exec(s);var e=m&&m[1];this.f.push(function(r){return r.empresa==null||r.empresa===e});return this};
Q.prototype.order=function(){return this};Q.prototype.limit=function(){return this};
Q.prototype.maybeSingle=function(){this.one=1;return this};Q.prototype.single=function(){this.one=2;return this};
Q.prototype.insert=function(p){this.op='insert';this.p=p;return this};Q.prototype.upsert=function(p,o){this.op='upsert';this.p=p;this.o=o;return this};
Q.prototype.update=function(p){this.op='update';this.p=p;return this};Q.prototype.delete=function(){this.op='delete';return this};
Q.prototype.then=function(a,b){var s=this;return new Promise(function(ok){setTimeout(function(){ok(s.run())},90)}).then(a,b)};
Q.prototype.run=function(){var T=DB[this.t],s=this,rows;
 var adminOnly=['ctr_categorias','ctr_unidades','ctr_acessos','ctr_liberados','ctr_grupos','ctr_config','ctr_naturezas','ctr_contatos'];
 if(this.op!=='select'&&U.papel!=='admin'&&(adminOnly.indexOf(this.t)>=0||this.t==='perfis'))return erro('new row violates row-level security policy');
 if(this.t==='ctr_categorias'&&this.op!=='select'&&this.op!=='insert'){var alvo=T.filter(function(r){return s.f.every(function(f){return f(r)})});if(alvo.some(function(r){return r.empresa==null}))return {data:[],error:null}}
 if(this.op==='insert'||this.op==='upsert'){var arr=[].concat(this.p);
  if(this.t==='ctr_contratos'){for(var i=0;i<arr.length;i++){var c0=Object.assign({criado_por:U.id},arr[i]);if(nivel(c0)<3)return erro('new row violates row-level security policy for table "ctr_contratos"');var ve=validarContrato(c0);if(ve)return erro(ve)}rows=arr.map(novoContrato)}
  else if(this.t==='ctr_alteracoes'){rows=arr.map(function(p){var c=DB.ctr_contratos.find(function(x){return x.id===p.contrato});if(!c||nivel(c)<3)return null;var n=Object.assign({id:altSeq++,valor_antes:c.valor_mensal,fim_antes:c.fim,criado_por:U.id,criado_em:new Date().toISOString()},p);T.push(n);
   if(p.valor_depois!=null)c.valor_mensal=p.valor_depois;if(p.fim_depois)c.fim=p.fim_depois;DB.ctr_historico.push({contrato:c.id,acao:p.tipo==='reajuste'?'reajuste registrado':'aditivo registrado',detalhe:'vale a partir de '+p.vigencia.split('-').reverse().join('/')+(p.percentual!=null?' · '+(p.indice||'')+' '+String(p.percentual).replace('.',',')+'%':'')+(p.valor_depois!=null?' · valor '+n.valor_antes+' → '+p.valor_depois:''),usuario:U.id,em:new Date().toISOString()});return n});if(rows.indexOf(null)>=0)return erro('new row violates row-level security policy for table "ctr_alteracoes"')}
  else if(this.t==='ctr_categorias'){rows=arr.map(function(p){if(T.some(function(r){return (r.empresa||'')===(p.empresa||'')&&r.nome.toLowerCase()===p.nome.toLowerCase()}))return null;var n=Object.assign({id:catSeq++,ativo:true,documentos:[]},p);T.push(n);return n});if(rows.indexOf(null)>=0)return erro('duplicate key value violates unique constraint "ctr_categorias_nome"')}
  else{var keys=(this.o&&this.o.onConflict||'').split(',').filter(Boolean);
   if(this.t==='ctr_acessos')keys=['usuario','empresa'];if(this.t==='ctr_contatos')keys=['usuario'];
   rows=arr.map(function(p){var ex=keys.length&&T.find(function(r){return keys.every(function(k){return r[k]===p[k]})});if(ex)Object.assign(ex,p);else T.push(Object.assign({},p));return p})}}
 else{rows=T.filter(function(r){return visivel(s.t,r)&&s.f.every(function(f){return f(r)})});
  if(this.op==='update'){if(this.t==='ctr_contratos'){rows=rows.filter(function(r){return nivel(r)>=3});if(s.p.categoria!=null){for(var j=0;j<rows.length;j++){if(rows[j].categoria!==s.p.categoria){var v2=validarContrato(Object.assign({},rows[j],s.p));if(v2)return erro(v2)}}}}
   rows.forEach(function(r){var mud=[];Object.keys(s.p).forEach(function(k){if(JSON.stringify(r[k]??null)!==JSON.stringify(s.p[k]??null))mud.push(k+': '+(r[k]??'—')+' → '+(s.p[k]??'—'))});Object.assign(r,s.p);
    if(s.t==='ctr_naturezas'&&s.p.categoria)DB.ctr_contratos.forEach(function(c){if(c.empresa===r.empresa&&c.natureza_cliente===r.codigo&&c.categoria!==s.p.categoria){DB.ctr_historico.push({contrato:c.id,acao:'dados alterados',detalhe:'categoria: '+c.categoria+' → '+s.p.categoria,usuario:U.id,em:new Date().toISOString()});c.categoria=s.p.categoria}});
    if(s.t==='ctr_contratos'&&s.p.natureza_cliente){var nc2=natCat(r.empresa,String(s.p.natureza_cliente).trim());if(nc2)r.categoria=nc2}
    if(s.t==='ctr_contratos'&&s.p.fornecedor_doc)r.fornecedor_doc=String(s.p.fornecedor_doc).toUpperCase().replace(/[^A-Z0-9]/g,'');
    if(s.t==='ctr_contratos'&&mud.length)DB.ctr_historico.push({contrato:r.id,acao:'dados alterados',detalhe:mud.join('; '),usuario:U.id,em:new Date().toISOString()})})}
  if(this.op==='delete'){if(this.t==='ctr_contratos')rows=rows.filter(function(r){return !DB.ctr_documentos.some(function(x){return x.contrato===r.id})});
   if(this.t==='ctr_categorias'&&rows.some(function(r){return DB.ctr_contratos.some(function(c){return c.categoria===r.id})}))return erro('violates foreign key constraint "ctr_contratos_categoria_fkey"');
   if(this.t==='ctr_unidades'&&rows.some(function(r){return DB.ctr_contratos.some(function(c){return c.unidade===r.codigo})}))return erro('update or delete on table "ctr_unidades" violates foreign key constraint "ctr_contratos_empresa_unidade_fkey"');
   rows.forEach(function(r){T.splice(T.indexOf(r),1)})}}
 rows=rows.map(function(r){return JSON.parse(JSON.stringify(r))});
 if(this.one)return {data:rows[0]||null,error:this.one===2&&!rows[0]?{message:'não encontrado'}:null};return {data:rows,error:null}};
var ouvinte=null;
window.supabase={createClient:function(){return {from:function(t){return new Q(t)},
 rpc:function(n,a){if(n==='ctr_senha_trocada'){DB.ctr_contatos.forEach(function(c){if(c.usuario===U.id)c.trocar_senha=false});return Promise.resolve({data:null,error:null})}
  if(n==='admin_definir_senha')return Promise.resolve(U.papel==='admin'?{data:null,error:null}:{data:null,error:{message:'apenas administradores podem trocar senhas'}});
  if(n==='ctr_marca_publica'){var cf=DB.ctr_config.find(function(x){return x.empresa===a.p_empresa});var em=DB.empresas.find(function(x){return x.slug===a.p_empresa});return Promise.resolve({data:cf?[{nome:cf.nome_exibicao||em.nome,logo:cf.logo,cor:cf.cor,logo_fundo:cf.logo_fundo,tema:cf.tema}]:[],error:null})}return Promise.resolve(n==='ctr_pessoas'?{data:DB.perfis.filter(function(p){return p.ativo&&(p.papel==='admin'||DB.ctr_acessos.some(function(x){return x.usuario===p.id}))}).map(function(p){return {id:p.id,nome:p.nome}}),error:null}:{data:U.papel==='admin',error:null})},
 auth:{getUser:function(){return Promise.resolve({data:{user:{id:U.id,email:U.email}}})},getSession:function(){return Promise.resolve({data:{session:{access_token:'proto'}}})},
  onAuthStateChange:function(cb){ouvinte=cb;setTimeout(function(){cb('INITIAL_SESSION',LOGADO?{user:{id:U.id}}:null)},30);return {data:{subscription:{unsubscribe:function(){}}}}},
  signInWithPassword:function(){try{sessionStorage.setItem('proto-logado','1')}catch(e){}setTimeout(function(){ouvinte&&ouvinte('SIGNED_IN',{user:{id:U.id}})},150);return Promise.resolve({error:null})},
  signOut:function(){try{sessionStorage.removeItem('proto-logado')}catch(e){}setTimeout(function(){ouvinte&&ouvinte('SIGNED_OUT',null)},50);return Promise.resolve({})},
  signUp:function(o){var id='u-'+Math.random().toString(36).slice(2,8);return Promise.resolve({data:{user:{id:id,identities:[{}]}},error:null})},
  resetPasswordForEmail:function(){return Promise.resolve({})},updateUser:function(){return Promise.resolve({})}}}}};
// funcao do SharePoint e envio simulados
var fetchReal=window.fetch;
window.fetch=function(url,op){url=String(url);
 if(url.indexOf('/functions/v1/')>=0){var j=JSON.parse(op.body);var r;
  if(j.acao==='indice'){var base={'IPCA':0.38,'IGP-M':0.30,'INPC':0.36}[j.indice];var hj=new Date(),lim=hj.getFullYear()+'-'+String(hj.getMonth()+1).padStart(2,'0');var ms=[],a=+j.de.slice(0,4),m=+j.de.slice(5,7);while(true){var ym=a+'-'+String(m).padStart(2,'0');if(ym>j.ate)break;if(ym<lim){var s=Math.sin(a*12+m)*(j.indice==='IGP-M'?0.7:0.25);ms.push({mes:ym,valor:Math.round((base+s)*100)/100})}m++;if(m>12){m=1;a++}}r={indice:j.indice,meses:ms,exemplo:true}}
  else if(j.acao==='testar')r={ok:true,pasta:'Contratos',webUrl:'https://sinerggiacombr.sharepoint.com/sites/GestaoContratos/Shared%20Documents',usado:52428800,total:1};
  else if(j.acao==='iniciar'){var c=DB.ctr_contratos.find(function(x){return x.id===j.contrato});var v=DB.ctr_documentos.filter(function(x){return x.contrato===c.id&&x.categoria===j.categoria}).length+1;
   var nome=c.numero+'_'+slug(catNome(c.categoria)).slice(0,30)+'_'+slug(c.fornecedor_nome)+'_'+j.categoria+'_v'+v+'.'+((j.nome_original.split('.').pop()||'pdf').toLowerCase());
   var nn=DB.ctr_naturezas.find(function(x){return x.empresa===c.empresa&&x.codigo===c.natureza_cliente});window.__ultimo={nome:nome,pasta:'Contratos/'+(c.empresa==='novatec'?'Novatec':'Millena Móveis')+'/'+(nn?c.natureza_cliente+' - '+nn.descricao.replace(/^#\s*/,''):catNome(c.categoria))+'/'+c.numero+' - '+c.fornecedor_nome,tam:j.tamanho};
   r={uploadUrl:'https://proto.upload/sessao',nome:nome,pasta:window.__ultimo.pasta,versao:v}}
  else if(j.acao==='registrar'){var u=window.__ultimo;DB.ctr_documentos.push({id:docSeq++,contrato:j.contrato,categoria:j.categoria,validade:j.validade||null,nome:u.nome,nome_original:j.nome_original,tamanho:u.tam,versao:1,enviado_por:U.id,enviado_em:new Date().toISOString()});
   DB.ctr_historico.push({contrato:j.contrato,acao:'arquivo enviado',detalhe:u.nome+(j.validade?' (válido até '+j.validade.split('-').reverse().join('/')+')':''),usuario:U.id,em:new Date().toISOString()});r={ok:true,id:docSeq-1,caminho:u.pasta+'/'+u.nome}}
  else{var dd=DB.ctr_documentos.find(function(x){return x.id===j.documento});DB.ctr_historico.push({contrato:dd.contrato,acao:j.acao==='baixar'?'arquivo baixado':'arquivo visualizado',detalhe:dd.nome,usuario:U.id,em:new Date().toISOString()});r={url:'about:blank',nome:dd.nome}}
  return new Promise(function(ok){setTimeout(function(){ok(new Response(JSON.stringify(r),{status:200,headers:{'Content-Type':'application/json'}}))},300)})}
 if(url.indexOf('https://proto.upload/')===0){var m=/-(\d+)\/(\d+)/.exec(op.headers['Content-Range']);var fim=Number(m[1])+1===Number(m[2]);
  return new Promise(function(ok){setTimeout(function(){ok(new Response(fim?'{"id":"item-'+Date.now()+'"}':'{}',{status:fim?201:202,headers:{'Content-Type':'application/json'}}))},250)})}
 return fetchReal.apply(this,arguments)};
window.__trocarPapel=function(p){try{sessionStorage.setItem('proto-papel',p);sessionStorage.setItem('proto-logado','1');localStorage.removeItem('ctr-emp')}catch(e){}location.reload()};
window.__papel=PAPEL;
})();
</script>
