# Monta o prototipo (index.html + Supabase simulado + dados inventados) em uma pasta de trabalho.
# Uso: python3 tests/contratos/build.py <pasta>   (a logo e opcional: <pasta>/logo.png)
import json,re,sys,random,os,base64
SP=sys.argv[1]; AQUI=os.path.dirname(os.path.abspath(__file__)); RAIZ=os.path.abspath(os.path.join(AQUI,'..','..'))
os.makedirs(SP,exist_ok=True)
# grupos e categorias padrao saem do proprio SQL (sempre iguais ao banco)
_sql=open(os.path.join(RAIZ,'supabase','2026-11-23_contratos.sql')).read()
_g=re.findall(r"\('(\w+)',\s+'([^']+)',\s+(\d+)\)",_sql[_sql.index('insert into public.ctr_grupos'):_sql.index('on conflict (codigo)')])
_blk=_sql[_sql.index('insert into public.ctr_categorias'):_sql.index('on conflict (coalesce(empresa')]
_cats=[]
for _m in re.finditer(r"\(null, '(\w+)', '([^']+)', (\d+), (null|'[^']*'), '\{([^}]*)\}',\s*(null|'(?:[^']|'')*')\)",_blk):
    _cats.append(dict(grupo=_m[1],nome=_m[2],aviso_dias=int(_m[3]),reajuste=None if _m[4]=='null' else _m[4][1:-1],documentos=[x for x in _m[5].split(',') if x],dica=None if _m[6]=='null' else _m[6][1:-1].replace("''","'")))
open(SP+'/seed.json','w').write(json.dumps(dict(grupos=[dict(codigo=a,nome=c2,ordem=int(o)) for a,c2,o in _g],cats=_cats),ensure_ascii=False))
_logo=SP+'/logo.png'
LOGO=("'data:image/png;base64,"+base64.b64encode(open(_logo,'rb').read()).decode()+"'") if os.path.exists(_logo) else 'null'
def cnpj(base):
    v=[int(c) for c in base]
    for p in ([5,4,3,2,9,8,7,6,5,4,3,2],[6,5,4,3,2,9,8,7,6,5,4,3,2]):
        s=sum(a*b for a,b in zip(v,p)); r=s%11; v.append(0 if r<2 else 11-r)
    t=''.join(map(str,v)); return f"{t[:2]}.{t[2:5]}.{t[5:8]}/{t[8:12]}-{t[12:]}"
mock=open(os.path.join(AQUI,'mock.js')).read().replace('__LOGO__',LOGO)
random.seed(7)
def fix(m):
    s=m.group(0)
    if s=="'11.222.333/0001-81'": return s
    return "'"+cnpj(''.join(str(random.randint(0,9)) for _ in range(8))+'0001')+"'"
mock=re.sub(r"'\d\d\.\d{3}\.\d{3}/0001-\d\d'",fix,mock)
mock=mock.replace('__SEED__',open(SP+'/seed.json').read().strip())
t=open(os.path.join(RAIZ,'contratos','index.html')).read()
t=t.replace('<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/dist/umd/supabase.min.js"></script>',mock.strip())
t=re.sub(r'<title>.*?</title>','<title>Contratos Sinergia</title>',t)
for tag in ['<link rel="apple-touch-icon" href="../icone-180.png">','<link rel="manifest" href="manifest.webmanifest">','<link rel="icon" href="../favicon.svg">']: t=t.replace(tag+'\n','')

R=[("const janela = window.open('', '_blank');   // abre ja no clique, para o navegador nao bloquear","const janela = null;"),
   ("if (janela) janela.location = r.url; else location.href = r.url;","toast(baixar ? 'No sistema real o arquivo é baixado do SharePoint. O download ficou registrado no histórico.' : 'No sistema real abre a visualização do arquivo no SharePoint. Ficou registrado no histórico.');"),
   ("  XLSX.writeFile(wb, `Contratos_${consolidado() ? 'todas' : EMP}_${new Date().toISOString().slice(0, 10)}.xlsx`);","  try { XLSX.writeFile(wb, `Contratos_${consolidado() ? 'todas' : EMP}_${new Date().toISOString().slice(0, 10)}.xlsx`); } catch (_) {}\n  toast(`Excel com ${linhas.length} contratos (os filtros aplicados). Se o download não começar, é limite do protótipo dentro do claude.ai.`);"),
   ("  XLSX.writeFile(wb, `Planilha modelo de contratos - ${emp}.xlsx`);","  try { XLSX.writeFile(wb, `Planilha modelo de contratos - ${emp}.xlsx`); } catch (_) {}\n  toast('Planilha modelo gerada. Se o download não começar, é limite do protótipo dentro do claude.ai: no sistema real ela baixa normalmente.', 6000);")]
R.append(('<div><span>Período do índice</span>', '<div class="nota">Protótipo: índices de exemplo, inventados. No sistema real eles vêm do Banco Central.</div><div><span>Período do índice</span>'))
for a,b in R:
    assert a in t, a[:50]; t=t.replace(a,b)
banner='''<style>
.proto{position:sticky;top:0;z-index:45;background:#16181d;color:#fff;font-size:12.5px;padding:7px max(16px,env(safe-area-inset-right,0px)) 7px max(16px,env(safe-area-inset-left,0px));display:flex;gap:10px;align-items:center;flex-wrap:wrap;border-bottom:2px solid #d12f33}
.proto b{color:#f3a3a5;letter-spacing:.08em}
.proto select{background:#22252c;border:1px solid #3a3e47;border-radius:6px;padding:3px 8px;color:#fff}
.proto label{display:flex;gap:6px;align-items:center;margin-left:auto}
.side{top:36px;height:calc(100vh - 36px)}
.login-hero{top:36px}
@media (max-width:820px){.side{top:0;height:auto}.proto{padding-top:calc(6px + env(safe-area-inset-top,0px))}.proto>span{display:none}.proto label{margin-left:auto;font-size:12px}.proto select{max-width:58vw}.main{padding-top:20px}}
</style>
<div class="proto"><b>PROTÓTIPO</b><span>Telas reais do sistema, com dados inventados. Nada é salvo.</span>
<label for="proto-papel">Entrar como:<select id="proto-papel" onchange="__trocarPapel(this.value)"><option value="admin">Admin Sinergia (Natanael)</option><option value="gestor">Usuária Millena: Carla – editar e enviar, filiais 1 a 6</option><option value="leitor">Usuário Millena: Rafael – só ver, só Imóveis</option><option value="primeiro">Usuária Millena: Bruna – primeiro acesso, 2 contratos</option></select></label></div>
<script>document.getElementById('proto-papel').value=window.__papel;document.addEventListener('DOMContentLoaded',function(){var e=document.getElementById('l-email'),s=document.getElementById('l-senha');if(e){e.value='admin@sinergia.exemplo';s.value='senha-de-exemplo'}});</script>
'''
t=t.replace('<body>\n','<body>\n'+banner,1)
open(SP+'/gestao-contratos.html','w').write(t)
print(len(t))
# modo escuro com as guardas de tema do claude.ai
import re as _re
_t=open(SP+'/gestao-contratos.html').read()
_m=_re.search(r'@media \(prefers-color-scheme: dark\)\{:root\{(.*?)\}\}',_t,_re.S)
if _m:
    _c=_m.group(1); _t=_t[:_m.start()]+'@media (prefers-color-scheme: dark){:root:not([data-theme="light"]){'+_c+'}}\n:root[data-theme="dark"]{'+_c+'}'+_t[_m.end():]
    open(SP+'/gestao-contratos.html','w').write(_t)
