"""
gerar_dashboard.py
Lê os Excel da Millena (Relatório, Usuários e Aprovações)
e gera Dashboard_Millena_DDMMAAAA.html. Copia do painel da Novatec: a Millena ocupa
o slot 'se' (antes Millena); os slots 'al' e 'rec' ficam vazios e escondidos no HTML.
"""
import openpyxl, json, webbrowser, unicodedata, zipfile
from pathlib import Path
from datetime import datetime


def _norm(nome):
    """Normaliza nome para comparação: strip, lower, sem acentos."""
    if not nome:
        return ''
    s = nome.strip().lower()
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('ascii')

BASE = Path(__file__).parent
NOW  = datetime.now()
HOJE = NOW.strftime('%d/%m/%Y')
HORA = NOW.strftime('%H:%M')
DATA_ARQUIVO = NOW.strftime('%d%m%Y')

# ── Grupo do painel (config.json na pasta do projeto; sem ele = Millena) ──────
try:
    _CFG = json.loads((BASE.parent / 'config.json').read_text(encoding='utf-8'))
except Exception:
    _CFG = {}
GRUPO_NOME = _CFG.get('nome') or 'Millena'          # nome exibido no painel
ARQ        = _CFG.get('arquivo') or 'Millena'       # prefixo dos arquivos (sem espacos)
LOGO_ARQ   = _CFG.get('logo') or 'logo_millena.png'


def arquivar_obsoletos():
    padroes = [
        f'Relatorio_{ARQ}_*.xlsx',
        f'Usuarios_{ARQ}_*.xlsx',
        f'Aprovacoes_{ARQ}_*.xlsx',
    ]
    for padrao in padroes:
        arquivos = sorted(
            [f for f in BASE.glob(padrao) if '_Obsoleto' not in f.stem],
            key=lambda f: f.stat().st_mtime,
            reverse=True
        )
        for antigo in arquivos[1:]:  # tudo exceto o mais recente
            novo = antigo.with_stem(antigo.stem + '_Obsoleto')
            if not novo.exists():
                antigo.rename(novo)
                print(f'  Obsoleto: {antigo.name}  ->  {novo.name}')


def find_latest(pattern):
    files = sorted(
        [f for f in BASE.glob(pattern) if '_Obsoleto' not in f.stem],
        key=lambda f: f.stat().st_mtime,
        reverse=True
    )
    return files[0] if files else None


def ler_excel(arquivo, aba=None):
    if not arquivo:
        return []
    wb = openpyxl.load_workbook(arquivo)
    ws = wb[aba] if aba and aba in wb.sheetnames else wb.active
    headers = [str(ws.cell(1, c).value or f'c{c}') for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        row = {headers[c - 1]: (str(ws.cell(r, c).value).strip() if ws.cell(r, c).value is not None else '')
               for c in range(1, ws.max_column + 1)}
        if any(row.values()):
            rows.append(row)
    return rows


def ler_resumo(arquivo):
    """Lê a aba Resumo (pares chave-valor) e retorna dicionário."""
    if not arquivo:
        return {}
    wb = openpyxl.load_workbook(arquivo, data_only=True)
    if 'Resumo' not in wb.sheetnames:
        return {}
    data = {}
    for row in wb['Resumo'].iter_rows(values_only=True):
        if row[0] and row[1] is not None:
            data[str(row[0])] = row[1]
    return data


# ── Prazo do Temporizador (config.json na pasta do projeto) ────────────────────
try:
    _prazo_cfg = json.loads((BASE.parent / 'config.json').read_text(encoding='utf-8')).get('prazo', '')
    _prazo_dt  = datetime.strptime(_prazo_cfg, '%d/%m/%Y')
    PRAZO_TXT  = _prazo_dt.strftime('%d/%m/%Y')
    PRAZO_ISO  = _prazo_dt.strftime('%Y-%m-%dT23:59:59')
except Exception:
    PRAZO_TXT, PRAZO_ISO = 'a definir', ''


# ── Gerenciamento de duplicatas ───────────────────────────────────────────────
arquivar_obsoletos()

# ── Leitura ───────────────────────────────────────────────────────────────────
_arq_rel_al  = None   # slot sem empresa
_arq_rel_se  = find_latest(f'Relatorio_{ARQ}_*.xlsx')
_arq_rel_rec = None   # slot sem empresa
if not _arq_rel_se:
    raise SystemExit(f'Relatorio_{ARQ}_*.xlsx nao encontrado em ' + str(BASE))
rel_al   = ler_excel(_arq_rel_al)
rel_se   = ler_excel(_arq_rel_se)
rel_rec  = ler_excel(_arq_rel_rec)
usr_al   = []
usr_se   = ler_excel(find_latest(f'Usuarios_{ARQ}_*.xlsx'))
usr_rec  = []
aprov_al  = []
aprov_se  = ler_excel(find_latest(f'Aprovacoes_{ARQ}_*.xlsx'))
aprov_rec = []
TEM_APROVACOES = bool(aprov_al or aprov_se or aprov_rec)

# Abas extras do Relatório
resumo_al   = ler_resumo(_arq_rel_al)
resumo_se   = ler_resumo(_arq_rel_se)
resumo_rec  = ler_resumo(_arq_rel_rec)
versoes_al  = ler_excel(_arq_rel_al, 'Versões & Parecer')
versoes_se  = ler_excel(_arq_rel_se, 'Versões & Parecer')
versoes_rec = ler_excel(_arq_rel_rec, 'Versões & Parecer')
versoes_al  = [v for v in versoes_al  if 'Nenhuma' not in v.get('Atividade', '')]
versoes_se  = [v for v in versoes_se  if 'Nenhuma' not in v.get('Atividade', '')]
versoes_rec = [v for v in versoes_rec if 'Nenhuma' not in v.get('Atividade', '')]

# ── Estado anterior (delta de envios) ─────────────────────────────────────────
ESTADO_FILE = BASE / "estado_anterior.json"
estado_anterior = {}
if ESTADO_FILE.exists():
    try:
        estado_anterior = json.loads(ESTADO_FILE.read_text(encoding='utf-8'))
    except Exception:
        pass


# ── Processamento ─────────────────────────────────────────────────────────────
def _txt(v):
    if v is None: return ''
    if isinstance(v, bool): return 'Sim' if v else 'Não'
    if isinstance(v, datetime): return v.strftime('%d/%m/%Y, %H:%M:%S')
    return v


def processar(rel, usr, empresa):
    # Remove entradas com nome iniciando em "Inativo devido ao e-mail"
    usr = [u for u in usr if not u.get('Nome', '').strip().lower().startswith('inativo devido ao e-mail')]
    # Remove duplicatas de nome (mantém o primeiro, preferindo Ativo)
    usr.sort(key=lambda u: (0 if u.get('Status','').strip()=='Ativo' else 1))
    _seen_nomes = set()
    usr_dedup = []
    for u in usr:
        n = _norm(u.get('Nome', ''))
        if n not in _seen_nomes:
            _seen_nomes.add(n)
            usr_dedup.append(u)
    usr = usr_dedup

    # Apenas Ativos em tudo — métricas, listas e totais
    usr = [u for u in usr if u.get('Status', 'Ativo').strip() == 'Ativo']
    usr_todos = usr

    func_subm_norm = set(_norm(r.get('Funcionário', '')) for r in rel if r.get('Funcionário'))

    # contagem de atividades por funcionario (chave normalizada, evita perda por
    # divergência de espaços/acentos entre Relatório e Usuários)
    ativ_count_norm = {}
    for r in rel:
        f = r.get('Funcionário', '')
        if f:
            ativ_count_norm[_norm(f)] = ativ_count_norm.get(_norm(f), 0) + 1

    # quem enviou — conta apenas usuários efetivamente cadastrados (ativos),
    # para que "concluídos" nunca ultrapasse "cadastrados" e o % de adesão
    # reflita exatamente os usuários cadastrados
    enviaram_list = []
    for u in usr:
        n_norm = _norm(u['Nome'])
        if n_norm in func_subm_norm:
            enviaram_list.append({
                'nome': u['Nome'],
                'cargo': u.get('Cargo', ''),
                'setor': u.get('Setor', ''),
                'perfil': u.get('Perfil', ''),
                'empresa': u.get('Empresa', ''),
                'atividades': ativ_count_norm.get(n_norm, 0)
            })
    enviaram_list.sort(key=lambda x: x['atividades'], reverse=True)

    # quem nao enviou — inclui ativos E inativos com campo status visível
    # métricas de total/pendentes usam apenas ativos (usr), mas a lista usa usr_todos
    nao_list = []
    for u in sorted(usr_todos, key=lambda x: x['Nome']):
        if _norm(u['Nome']) not in func_subm_norm:
            nao_list.append({
                'nome': u['Nome'],
                'cargo': u.get('Cargo', ''),
                'setor': u.get('Setor', ''),
                'perfil': u.get('Perfil', ''),
                'empresa': u.get('Empresa', ''),
                'status_usr': u.get('Status', 'Ativo').strip()
            })

    # por cargo — combina usuarios e atividades
    cargo_users = {}
    for u in usr:
        c = u.get('Cargo', '')
        if c:
            if c not in cargo_users:
                cargo_users[c] = {'total': 0, 'funcs_subm': set(), 'atividades': 0}
            cargo_users[c]['total'] += 1

    for r in rel:
        c = r.get('Cargo', '')
        f = r.get('Funcionário', '')
        if c and f:
            if c not in cargo_users:
                cargo_users[c] = {'total': 0, 'funcs_subm': set(), 'atividades': 0}
            cargo_users[c]['funcs_subm'].add(f)
            cargo_users[c]['atividades'] += 1

    cargo_data = []
    for cargo, d in cargo_users.items():
        env = len(d['funcs_subm'])
        tot = d['total']
        pct = round(env / max(tot, 1) * 100)
        cargo_data.append({
            'cargo': cargo,
            'total': tot,
            'enviaram': env,
            'atividades': d['atividades'],
            'pct': pct
        })
    cargo_data.sort(key=lambda x: x['atividades'], reverse=True)

    # atividades
    atividades = []
    for r in rel:
        if r.get('Funcionário'):
            atividades.append({
                'funcionario': r.get('Funcionário', ''),
                'cargo': r.get('Cargo', ''),
                'setor': r.get('Setor', ''),
                'nome': r.get('Nome', ''),
                'detalhamento': r.get('Detalhamento', ''),
                'periodicidade': r.get('Periodicidade', ''),
                'tempo': r.get('Tempo', ''),
                'status': r.get('Status', ''),
                'como_realiza': r.get('Como realiza', ''),
                'pontos_criticos': r.get('Pontos críticos', ''),
                'sugestoes': r.get('Sugestões de melhoria', ''),
                'criada_em': r.get('Criada em', ''),
                'empresa': _txt(r.get('Empresa', '')),
                'horario_fixo': _txt(r.get('Horário fixo', '')),
                'decisao_gestor': _txt(r.get('Decisão gestor', '')),
                'decisao_consultor': _txt(r.get('Decisão consultor', '')),
                'edicoes': _txt(r.get('Edições do consultor', '')),
                'parecer': _txt(r.get('Parecer do consultor', '')),
                'atualizada_em': _txt(r.get('Atualizada em', '')),
            })

    # Pendentes = total cadastrados - cadastrados que enviaram (garante consistência nos cartões)
    _nao_enviaram = len(usr) - len(enviaram_list)
    if _nao_enviaram < 0:
        _nao_enviaram = 0

    # Envios por data — tenta localizar coluna de data no relatório
    _date_col = next((c for c in (rel[0].keys() if rel else [])
                      if any(k in c.lower() for k in ['data', 'criada', 'criado', 'enviado', 'atualiz'])), None)
    _envios_por_data = {}
    if _date_col:
        _datas_func = {}
        for r in rel:
            f = r.get('Funcionário', '')
            v = r.get(_date_col, '').strip()
            if not f or not v or v == 'None':
                continue
            # Suporta: "DD/MM/YYYY, HH:MM:SS", "YYYY-MM-DD HH:MM:SS", "DD/MM/YYYY HH:MM"
            v = v.split(',')[0].split('T')[0].split(' ')[0].strip()
            parts = v.replace('-', '/').split('/')
            if len(parts) == 3:
                if len(parts[0]) == 4:
                    v = f'{parts[2]}/{parts[1]}/{parts[0]}'
                _datas_func.setdefault(v, set()).add(f)
        _envios_por_data = {k: len(v) for k, v in sorted(_datas_func.items())}

    # por setor (do cadastro de Usuarios) — cadastrados, enviaram, atividades
    _setores = {}
    for u in usr:
        st = (u.get('Setor') or '').strip() or '(sem setor)'
        d = _setores.setdefault(st, {'setor': st, 'total': 0, 'enviaram': 0, 'atividades': 0})
        d['total'] += 1
        n_norm = _norm(u['Nome'])
        if n_norm in func_subm_norm:
            d['enviaram'] += 1
            d['atividades'] += ativ_count_norm.get(n_norm, 0)
    setor_data = sorted(_setores.values(),
                        key=lambda d: (-d['enviaram'] / max(d['total'], 1), -d['total'], d['setor']))
    for d in setor_data:
        d['pct'] = round(d['enviaram'] / max(d['total'], 1) * 100)

    return {
        'empresa': empresa,
        'setor_data': setor_data,
        'total': len(usr),
        'enviaram': len(enviaram_list),
        'nao_enviaram': _nao_enviaram,
        'total_atividades': len(rel),
        'enviaram_list': enviaram_list,
        'nao_enviaram_list': nao_list,
        'cargo_data': cargo_data,
        'atividades': atividades,
        'envios_por_data': _envios_por_data,
    }


def agregar_analise(atividades):
    """Totais por coluna do Relatório (Periodicidade, Como realiza, Status,
    Pontos críticos, Sugestões de melhoria, Criada em) — exclui Horário fixo,
    Decisão gestor, Decisão consultor, Edições do consultor e Parecer do consultor."""
    from collections import Counter

    def _valor_ou_vazio(v):
        v = (v or '').strip()
        return v if v else '(vazio)'

    per  = Counter(_valor_ou_vazio(a.get('periodicidade')) for a in atividades)
    real = Counter(_valor_ou_vazio(a.get('como_realiza'))  for a in atividades)
    stat = Counter(_valor_ou_vazio(a.get('status'))        for a in atividades)
    pontos_count    = sum(1 for a in atividades if a.get('pontos_criticos', '').strip())
    sugestoes_count = sum(1 for a in atividades if a.get('sugestoes', '').strip())

    por_mes = Counter()
    for a in atividades:
        c = (a.get('criada_em') or '').strip()
        mes = '(vazio)'
        if c:
            d = c.split(',')[0].split('T')[0].split(' ')[0].strip()
            partes = d.replace('-', '/').split('/')
            if len(partes) == 3:
                mes = f'{partes[1]}/{partes[0]}' if len(partes[0]) == 4 else f'{partes[1]}/{partes[2]}'
        por_mes[mes] += 1

    def _chave_mes(item):
        if item[0] == '(vazio)':
            return ('9999', '99')
        mm, aaaa = item[0].split('/')
        return (aaaa, mm)

    return {
        'total': len(atividades),
        'periodicidade': dict(sorted(per.items(), key=lambda x: -x[1])),
        'como_realiza': dict(sorted(real.items(), key=lambda x: -x[1])),
        'status': dict(sorted(stat.items(), key=lambda x: -x[1])),
        'pontos_criticos_count': pontos_count,
        'sugestoes_count': sugestoes_count,
        'por_mes': dict(sorted(por_mes.items(), key=_chave_mes)),
    }


def processar_aprovacoes(aprov, rel):
    """Cruza Aprovações com Relatório. Retorna lista de itens enriquecidos."""
    if not aprov:
        return {'itens': [], 'aprovados': 0, 'pendentes': 0, 'reprovados': 0, 'total': 0}

    # Detecta colunas disponíveis (case-insensitive)
    cols = list(aprov[0].keys()) if aprov else []
    def col(keys):
        for k in keys:
            for c in cols:
                if k.lower() in c.lower():
                    return c
        return None

    c_func  = col(['funcionário', 'funcionario', 'nome', 'colaborador'])
    c_cargo = col(['cargo', 'função', 'funcao'])
    c_atv   = col(['atividade', 'tarefa', 'nome da atividade']) or ('Nome' if 'Nome' in cols else None)
    c_stat  = col(['status', 'situação', 'situacao'])
    c_aprov = col(['aprovador', 'gestor', 'aprovado por'])
    c_data  = col(['data', 'dt aprovação', 'dt_aprovacao', 'atualizada em', 'criada em'])

    # Mapa de atividades do Relatório por funcionário
    rel_map = {}
    for r in rel:
        f = r.get('Funcionário', '')
        if f:
            rel_map.setdefault(f, []).append(r.get('Nome', ''))

    itens = []
    for a in aprov:
        func  = a.get(c_func,  '') if c_func  else ''
        cargo = a.get(c_cargo, '') if c_cargo else ''
        atv   = a.get(c_atv,  '') if c_atv   else ''
        stat  = a.get(c_stat, '') if c_stat   else ''
        aprov_por = a.get(c_aprov, '') if c_aprov else ''
        data_ap   = a.get(c_data,  '') if c_data   else ''

        # Normaliza status
        sl = stat.lower()
        if any(w in sl for w in ['aprovad', 'concluíd', 'ok', 'finaliz']):
            status_norm = 'Aprovado'
        elif any(w in sl for w in ['reprovad', 'recusad', 'rejeit']):
            status_norm = 'Reprovado'
        else:
            status_norm = 'Pendente'

        # Verifica se tem atividade correspondente no Relatório
        atv_rel = rel_map.get(func, [])
        tem_relatorio = bool(atv_rel)

        itens.append({
            'funcionario': func,
            'cargo': cargo,
            'atividade': atv,
            'status': status_norm,
            'aprovador': aprov_por,
            'data': data_ap,
            'tem_relatorio': tem_relatorio,
        })

    aprovados  = sum(1 for i in itens if i['status'] == 'Aprovado')
    pendentes  = sum(1 for i in itens if i['status'] == 'Pendente')
    reprovados = sum(1 for i in itens if i['status'] == 'Reprovado')

    return {
        'itens': itens,
        'aprovados': aprovados,
        'pendentes': pendentes,
        'reprovados': reprovados,
        'total': len(itens),
    }


al  = processar(rel_al,  usr_al,  'Novatec Alagoas')
se  = processar(rel_se,  usr_se,  GRUPO_NOME)
rec = processar(rel_rec, usr_rec, 'Novatec Recife')
apr_al  = processar_aprovacoes(aprov_al,  rel_al)
apr_se  = processar_aprovacoes(aprov_se,  rel_se)
apr_rec = processar_aprovacoes(aprov_rec, rel_rec)
anl_al    = agregar_analise(al['atividades'])
anl_se    = agregar_analise(se['atividades'])
anl_rec   = agregar_analise(rec['atividades'])
anl_total = agregar_analise(al['atividades'] + se['atividades'] + rec['atividades'])


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


def _soma_resumo(chave):
    return _num(resumo_al.get(chave, 0)) + _num(resumo_se.get(chave, 0)) + _num(resumo_rec.get(chave, 0))


def _html_setores(dados):
    import html as _h
    linhas = []
    for d in dados.get('setor_data', []):
        cor = '#10b981' if d['pct'] >= 70 else ('#f59e0b' if d['pct'] >= 30 else '#ef4444')
        linhas.append(
            f'<div class="setor-row"><div class="setor-top">'
            f'<span class="setor-nome">{_h.escape(d["setor"])}</span>'
            f'<span class="setor-num"><strong>{d["enviaram"]}</strong>/{d["total"]}'
            f' · {d["atividades"]} ativ. · <strong style="color:{cor}">{d["pct"]}%</strong></span></div>'
            f'<div class="setor-track"><div class="setor-fill" style="width:{max(d["pct"], 2)}%;background:{cor}"></div></div></div>')
    return ''.join(linhas) or '<div class="setor-vazio">Sem setores no cadastro.</div>'

SETOR_HTML = _html_setores(se)

_logo = BASE.parent / LOGO_ARQ
LOGO_SRC = ('data:image/png;base64,' + __import__('base64').b64encode(_logo.read_bytes()).decode()) if _logo.exists() else ''

DATA = {
    'alagoas': al,
    'sergipe': se,
    'recife':  rec,
    'aprovacoes': {'alagoas': apr_al, 'sergipe': apr_se, 'recife': apr_rec},
    'analise': {'alagoas': anl_al, 'sergipe': anl_se, 'recife': anl_rec, 'total': anl_total},
    'resumo': {
        'alagoas': {
            'tempo_total':  resumo_al.get('Tempo total mapeado (horas)', 0),
            'aprovadas':    resumo_al.get('Aprovadas', 0),
            'ressalvas':    resumo_al.get('Ressalvas', 0),
            'ag_gestor':    resumo_al.get('Aguardando gestor', 0),
            'ag_consultor': resumo_al.get('Aguardando consultor', 0),
            'editadas':     resumo_al.get('Atividades editadas pelo consultor', 0),
            'com_parecer':  resumo_al.get('Atividades com parecer do consultor', 0),
        },
        'sergipe': {
            'tempo_total':  resumo_se.get('Tempo total mapeado (horas)', 0),
            'aprovadas':    resumo_se.get('Aprovadas', 0),
            'ressalvas':    resumo_se.get('Ressalvas', 0),
            'ag_gestor':    resumo_se.get('Aguardando gestor', 0),
            'ag_consultor': resumo_se.get('Aguardando consultor', 0),
            'editadas':     resumo_se.get('Atividades editadas pelo consultor', 0),
            'com_parecer':  resumo_se.get('Atividades com parecer do consultor', 0),
        },
        'recife': {
            'tempo_total':  resumo_rec.get('Tempo total mapeado (horas)', 0),
            'aprovadas':    resumo_rec.get('Aprovadas', 0),
            'ressalvas':    resumo_rec.get('Ressalvas', 0),
            'ag_gestor':    resumo_rec.get('Aguardando gestor', 0),
            'ag_consultor': resumo_rec.get('Aguardando consultor', 0),
            'editadas':     resumo_rec.get('Atividades editadas pelo consultor', 0),
            'com_parecer':  resumo_rec.get('Atividades com parecer do consultor', 0),
        },
        'total': {
            'tempo_total':  round(_soma_resumo('Tempo total mapeado (horas)'), 1),
            'aprovadas':    _soma_resumo('Aprovadas'),
            'ressalvas':    _soma_resumo('Ressalvas'),
            'ag_gestor':    _soma_resumo('Aguardando gestor'),
            'ag_consultor': _soma_resumo('Aguardando consultor'),
            'editadas':     _soma_resumo('Atividades editadas pelo consultor'),
            'com_parecer':  _soma_resumo('Atividades com parecer do consultor'),
        },
    },
    'versoes': {
        'alagoas': [{'atividade': v.get('Atividade',''), 'funcionario': v.get('Funcionário',''),
                     'setor': v.get('Setor',''), 'tipo': v.get('Tipo',''),
                     'autor': v.get('Autor',''), 'data': v.get('Data/hora',''),
                     'conteudo': v.get('Conteúdo','')} for v in versoes_al],
        'sergipe': [{'atividade': v.get('Atividade',''), 'funcionario': v.get('Funcionário',''),
                     'setor': v.get('Setor',''), 'tipo': v.get('Tipo',''),
                     'autor': v.get('Autor',''), 'data': v.get('Data/hora',''),
                     'conteudo': v.get('Conteúdo','')} for v in versoes_se],
        'recife':  [{'atividade': v.get('Atividade',''), 'funcionario': v.get('Funcionário',''),
                     'setor': v.get('Setor',''), 'tipo': v.get('Tipo',''),
                     'autor': v.get('Autor',''), 'data': v.get('Data/hora',''),
                     'conteudo': v.get('Conteúdo','')} for v in versoes_rec],
    },
}
DATA_JSON = json.dumps(DATA, ensure_ascii=False).replace('</', '<\\/')

# ── Métricas globais ──────────────────────────────────────────────────────────
total_users = al['total'] + se['total'] + rec['total']
total_env   = al['enviaram'] + se['enviaram'] + rec['enviaram']
total_nenv  = al['nao_enviaram'] + se['nao_enviaram'] + rec['nao_enviaram']
total_atv   = al['total_atividades'] + se['total_atividades'] + rec['total_atividades']
pct_global  = round(total_env / max(total_users, 1) * 100, 1)
pct_al      = round(al['enviaram']  / max(al['total'],  1) * 100, 1)
pct_se      = round(se['enviaram']  / max(se['total'],  1) * 100, 1)
pct_rec     = round(rec['enviaram'] / max(rec['total'], 1) * 100, 1)

def fmt(n): return str(n).replace('.', ',')

def _trophy(pct):
    if pct >= 100:
        return ' <span style="font-size:20px;vertical-align:middle;filter:drop-shadow(0 0 8px gold)">\U0001F3C6</span>'
    return ''

trophy_al  = _trophy(pct_al)
trophy_se  = _trophy(pct_se)
trophy_rec = _trophy(pct_rec)

n_cargos_al  = len(al['cargo_data'])
n_cargos_se  = len(se['cargo_data'])
n_cargos_rec = len(rec['cargo_data'])
apr_total    = apr_al['total']    + apr_se['total']    + apr_rec['total']
apr_aprov    = apr_al['aprovados']+ apr_se['aprovados']+ apr_rec['aprovados']
apr_pend     = apr_al['pendentes']+ apr_se['pendentes']+ apr_rec['pendentes']
apr_reprov   = apr_al['reprovados']+apr_se['reprovados']+apr_rec['reprovados']
APROV_TAB    = 'true' if TEM_APROVACOES else 'false'

# ── Delta do dia (acumula do início do dia até agora) ─────────────────────────
ultima_atualiz = estado_anterior.get('timestamp', '—')
mesma_data     = (estado_anterior.get('data_ref', '') == HOJE)

# Nomes atuais de quem enviou
al_nomes_atual  = set(p['nome'] for p in al['enviaram_list'])
se_nomes_atual  = set(p['nome'] for p in se['enviaram_list'])
rec_nomes_atual = set(p['nome'] for p in rec['enviaram_list'])

if not estado_anterior or not mesma_data:
    delta_al = delta_se = delta_rec = 0
    novos_usr_al = novos_usr_se = novos_usr_rec = set()
    al_env_dia   = al['enviaram']
    se_env_dia   = se['enviaram']
    rec_env_dia  = rec['enviaram']
    al_nomes_dia  = list(al_nomes_atual)
    se_nomes_dia  = list(se_nomes_atual)
    rec_nomes_dia = list(rec_nomes_atual)
else:
    al_env_dia   = estado_anterior.get('al_env_dia',  al['enviaram'])
    se_env_dia   = estado_anterior.get('se_env_dia',  se['enviaram'])
    rec_env_dia  = estado_anterior.get('rec_env_dia', rec['enviaram'])
    al_nomes_dia  = estado_anterior.get('al_nomes_dia',  [])
    se_nomes_dia  = estado_anterior.get('se_nomes_dia',  [])
    rec_nomes_dia = estado_anterior.get('rec_nomes_dia', [])
    delta_al  = al['enviaram']  - al_env_dia
    delta_se  = se['enviaram']  - se_env_dia
    delta_rec = rec['enviaram'] - rec_env_dia
    novos_usr_al  = al_nomes_atual  - set(al_nomes_dia)
    novos_usr_se  = se_nomes_atual  - set(se_nomes_dia)
    novos_usr_rec = rec_nomes_atual - set(rec_nomes_dia)

delta_total     = delta_al + delta_se + delta_rec
novos_usr_total = len(novos_usr_al) + len(novos_usr_se) + len(novos_usr_rec)

# ── Controle de alertas automáticos ───────────────────────────────────────────
# Estreantes: quais já foram notificados hoje (evita e-mail duplicado)
_estreantes_notif_prev = set(estado_anterior.get('estreantes_notificados', [])) if mesma_data else set()
if mesma_data and novos_usr_total > 0:
    _alerta_est_al  = novos_usr_al  - _estreantes_notif_prev
    _alerta_est_se  = novos_usr_se  - _estreantes_notif_prev
    _alerta_est_rec = novos_usr_rec - _estreantes_notif_prev
else:
    _alerta_est_al = _alerta_est_se = _alerta_est_rec = set()
_total_alerta_est    = len(_alerta_est_al) + len(_alerta_est_se) + len(_alerta_est_rec)
_estreantes_notif_atual = _estreantes_notif_prev | _alerta_est_al | _alerta_est_se | _alerta_est_rec

# 100% adesão: detecta unidades que atingiram 100% pela primeira vez
_unidades_100_prev = set(estado_anterior.get('unidades_100_notificadas', []))
_unidades_100_atuais = set()
if pct_al  >= 100.0 and al['total']  > 0: _unidades_100_atuais.add('al')
if pct_se  >= 100.0 and se['total']  > 0: _unidades_100_atuais.add('se')
if pct_rec >= 100.0 and rec['total'] > 0: _unidades_100_atuais.add('rec')
_unidades_100_novas  = _unidades_100_atuais - _unidades_100_prev
_unidades_100_salvar = (_unidades_100_prev & _unidades_100_atuais) | _unidades_100_novas

# Atividades: throttle de 1 hora entre e-mails
_ultima_email_atv = estado_anterior.get('ultima_email_atividade', '') if mesma_data else ''
_pode_email_atv   = False
if mesma_data and delta_total > 0:
    if not _ultima_email_atv:
        _pode_email_atv = True
    else:
        try:
            _dt_ult = datetime.strptime(_ultima_email_atv, '%d/%m/%Y %H:%M')
            _pode_email_atv = (NOW - _dt_ult).total_seconds() >= 3600
        except Exception:
            _pode_email_atv = True
_email_atv_ts = f'{HOJE} {HORA}' if _pode_email_atv else _ultima_email_atv

# Data e hora da última atualização — sempre reflete a execução atual
_hora_ref = f'{HOJE} {HORA}'

# ── Cartão NOVA ATIVIDADE (usa classe .card para uniformidade) ────────────────
if not mesma_data or delta_total == 0:
    _icone_atv = '<svg style="position:absolute;right:12px;top:40%;transform:translateY(-50%);opacity:.18;width:44px;height:44px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M19 3h-4.18C14.4 1.84 13.3 1 12 1c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm-2 14l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z"/></svg>'
    _hoje_badge = '<div style="position:absolute;right:10px;bottom:8px;font-size:10px;font-weight:700;color:#f59e0b;letter-spacing:.06em">HOJE</div>'
    novos_card = (
        '<div class="card" id="card-novas" role="button" tabindex="0" title="Ver as atividades enviadas hoje" data-total="0" data-al="0" data-se="0" data-rec="0">'
        + _icone_atv + _hoje_badge +
        '<div class="card-label">Novas Atividades</div>'
        '<div class="card-value" id="val-novas" style="color:#f59e0b;font-size:20px">0</div>'
        '<div class="card-detail" id="det-novas-se">nenhuma hoje</div>'
        '</div>'
    )
else:
    s_al  = ('+' if delta_al  >= 0 else '') + str(delta_al)
    s_se  = ('+' if delta_se  >= 0 else '') + str(delta_se)
    s_rec = ('+' if delta_rec >= 0 else '') + str(delta_rec)
    _icone_atv = '<svg style="position:absolute;right:12px;top:40%;transform:translateY(-50%);opacity:.18;width:44px;height:44px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M19 3h-4.18C14.4 1.84 13.3 1 12 1c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm-2 14l-4-4 1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z"/></svg>'
    _hoje_badge = '<div style="position:absolute;right:10px;bottom:8px;font-size:10px;font-weight:700;color:#f59e0b;letter-spacing:.06em">HOJE</div>'
    novos_card = (
        f'<div class="card pink" id="card-novas" role="button" tabindex="0" title="Ver as atividades enviadas hoje" data-total="+{delta_total}" data-al="{s_al}" data-se="{s_se}" data-rec="{s_rec}">'
        + _icone_atv + _hoje_badge +
        '<div class="card-label">🔔 Novas Atividades</div>'
        f'<div class="card-value" id="val-novas" style="color:#f59e0b;font-size:20px">+{delta_total}</div>'
        f'<div class="card-detail" id="det-novas-se">{GRUPO_NOME}: {s_se}</div>'
        '</div>'
    )

novos_usuarios_card = ''

# Salva estado: baseline do dia + snapshot atual + controle de alertas
ESTADO_FILE.write_text(json.dumps({
    'timestamp':              f'{HOJE} {HORA}',
    'data_ref':               HOJE,
    'al_env_dia':             al_env_dia,
    'se_env_dia':             se_env_dia,
    'rec_env_dia':            rec_env_dia,
    'al_nomes_dia':           al_nomes_dia,
    'se_nomes_dia':           se_nomes_dia,
    'rec_nomes_dia':          rec_nomes_dia,
    'al_env':                 al['enviaram'],
    'se_env':                 se['enviaram'],
    'rec_env':                rec['enviaram'],
    'al_total':               al['total'],
    'se_total':               se['total'],
    'rec_total':              rec['total'],
    'al_nomes':               list(al_nomes_atual),
    'se_nomes':               list(se_nomes_atual),
    'rec_nomes':              list(rec_nomes_atual),
    'estreantes_notificados': list(_estreantes_notif_atual),
    'ultima_email_atividade': _email_atv_ts,
    'unidades_100_notificadas': list(_unidades_100_salvar),
}, ensure_ascii=False, indent=2), encoding='utf-8')

# Flag: envia e-mail se houve qualquer alteração nos dados em relação ao ciclo anterior
_prev_al_env  = estado_anterior.get('al_env',   None)
_prev_se_env  = estado_anterior.get('se_env',   None)
_prev_rec_env = estado_anterior.get('rec_env',  None)
_prev_al_tot  = estado_anterior.get('al_total', None)
_prev_se_tot  = estado_anterior.get('se_total', None)
_prev_rec_tot = estado_anterior.get('rec_total',None)
_TEM_NOVIDADES = (
    _prev_al_env  is None or
    al['enviaram']  != _prev_al_env  or
    se['enviaram']  != _prev_se_env  or
    rec['enviaram'] != _prev_rec_env or
    al['total']     != _prev_al_tot  or
    se['total']     != _prev_se_tot  or
    rec['total']    != _prev_rec_tot
)
(BASE / 'tem_novidades.json').write_text(
    json.dumps({'tem_novidades': _TEM_NOVIDADES}), encoding='utf-8'
)

# ── Template HTML ─────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<meta name="theme-color" content="#161b22">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Painel GRUPO_NOME_PH">
<meta name="format-detection" content="telephone=no">
<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="icone-180.png">
<link rel="icon" type="image/png" sizes="192x192" href="icone-192.png">
<title>GRUPO_NOME_PH — Relatório de Atividades · DATE_HOJE</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
  :root {
    --bg:#0d1117;--surface:#161b22;--surface2:#1c2330;--border:#2a3441;
    --al:#2563eb;--al-glow:rgba(37,99,235,.18);
    --se:#7c3aed;--se-glow:rgba(124,58,237,.18);
    --re:#f97316;--re-glow:rgba(249,115,22,.18);
    --green:#10b981;--orange:#f59e0b;--red:#ef4444;
    --text:#e2e8f0;--muted:#94a3b8;--subtle:#cbd5e1;--white:#f8fafc;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--text);font-family:'Inter',sans-serif;font-size:14px;line-height:1.6;min-height:100vh}
  header{background:var(--surface);border-bottom:1px solid var(--border);padding:20px 40px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
  .logo{display:flex;align-items:center;gap:12px}
  .logo-icon{width:36px;height:36px;background:linear-gradient(135deg,var(--al),var(--se));border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px}
  .logo-text{font-size:20px;font-weight:700;color:var(--white)}
  .logo-sub{font-size:13px;color:var(--muted);margin-top:2px}
  .logo-millena{height:52px;width:auto;border-radius:8px;display:block;flex-shrink:0}
  @media(max-width:600px){.logo-millena{height:38px}header{flex-wrap:wrap;gap:8px}.logo-text{font-size:15px;line-height:1.25}.header-date{display:none}}
  .header-date{font-family:'DM Mono',monospace;font-size:12px;color:var(--muted);background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:6px 12px}
  #filtro-turno option{background:#1c2330;color:#f1f5f9}
  #filtro-turno option:checked{background:#334155;color:#fff}
  .nav{background:var(--surface);border-bottom:1px solid var(--border);padding:0 40px;display:flex;gap:4px;position:sticky;top:var(--hdr-h,0px);z-index:95}
  .nav-btn{background:none;border:none;cursor:pointer;padding:14px 22px;font-family:'Inter',sans-serif;font-size:16px;font-weight:500;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-1px;transition:all .18s}
  .nav-btn:hover{color:var(--text)}
  .nav-btn.active{color:var(--white);border-bottom-color:var(--al)}
  main{padding:32px 40px;max-width:none;margin:0}   /* ocupa a tela toda (antes parava em 1400px) */
  .tab-panel{display:none}
  .tab-panel.active{display:block}
  .section-title{font-size:24px;font-weight:700;color:var(--white);margin-bottom:6px}
  .section-sub{font-size:15px;color:var(--muted);margin-bottom:24px}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:28px}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:10px 14px;position:relative;overflow:hidden;transition:transform .15s,box-shadow .15s}
  .card:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(0,0,0,.35)}
  .card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px}
  .card.al::before{background:linear-gradient(90deg,var(--al),#60a5fa)}
  .card.se::before{background:linear-gradient(90deg,var(--se),#a78bfa)}
  .card.re::before{background:linear-gradient(90deg,var(--re),#fdba74)}
  .card.green::before{background:var(--green)}
  .card.red::before{background:var(--red)}
  .card.orange::before{background:var(--orange)}
  .card.teal::before{background:#14b8a6}
  .card.pink::before{background:#ffffff}
  .card-label{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:4px}
  .card-value{font-size:24px;font-weight:700;color:var(--white);line-height:1;margin-bottom:2px}
  .card-detail{font-size:12px;color:var(--muted);margin-top:2px;line-height:1.4}
  .empresa-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:24px;margin-bottom:32px}
  @media(max-width:1100px){.empresa-grid{grid-template-columns:1fr 1fr}}
  @media(max-width:700px){.empresa-grid{grid-template-columns:1fr}}
  .empresa-block{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden}
  .empresa-header{padding:18px 22px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border)}
  .empresa-header.al{background:var(--al-glow);border-left:3px solid var(--al)}
  .empresa-header.se{background:var(--se-glow);border-left:3px solid var(--se)}
  .empresa-header.re{background:var(--re-glow);border-left:3px solid var(--re)}
  .empresa-name{font-size:16px;font-weight:700;color:var(--white)}
  .empresa-badge{font-family:'DM Mono',monospace;font-size:12px;padding:4px 11px;border-radius:20px;font-weight:600}
  .badge-al{background:var(--al-glow);color:#60a5fa;border:1px solid #2563eb55}
  .badge-se{background:var(--se-glow);color:#a78bfa;border:1px solid #7c3aed55}
  .badge-re{background:var(--re-glow);color:#fdba74;border:1px solid #f9731655}
  .donut-wrap{padding:24px;display:flex;align-items:center;gap:24px}
  .donut-canvas{flex-shrink:0}
  .donut-legend{display:flex;flex-direction:column;gap:10px}
  .legend-item{display:flex;align-items:center;gap:8px;font-size:13px}
  .legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
  .progress-section{padding:0 22px 26px}
  .progress-label{display:flex;justify-content:space-between;align-items:center;font-size:15px;font-weight:500;color:var(--muted);margin-bottom:6px}
  .progress-track{height:8px;background:var(--surface2);border-radius:99px;overflow:hidden;border:1px solid var(--border)}
  .progress-fill{height:100%;border-radius:99px;transition:width .6s ease}
  .table-wrap{background:var(--surface);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:28px}
  .table-header{padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;gap:12px}
  .table-title{font-size:14px;font-weight:600;color:var(--white)}
  .table-count{font-size:11px;font-family:'DM Mono',monospace;background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:3px 10px;color:var(--subtle)}
  .search-input{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:7px 12px;font-size:13px;color:var(--text);font-family:'Inter',sans-serif;outline:none;width:220px;transition:border-color .15s}
  .search-input:focus{border-color:var(--al)}
  .search-input::placeholder{color:var(--muted)}
  table{width:100%;border-collapse:collapse}
  th{background:var(--surface2);padding:10px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--muted);border-bottom:1px solid var(--border)}
  td{padding:11px 14px;border-bottom:1px solid var(--border);font-size:13px;color:var(--text);vertical-align:middle}
  tr:last-child td{border-bottom:none}
  tr:hover td{background:var(--surface2)}
  .dec-badge{display:inline-flex;align-items:center;gap:4px;font-size:11px;font-weight:600;padding:2px 9px;border-radius:20px;white-space:nowrap;border:1px solid}
  .dec-aprovada{background:rgba(16,185,129,.12);color:#34d399;border-color:rgba(16,185,129,.45)}
  .dec-aguardando{background:rgba(245,158,11,.12);color:#fbbf24;border-color:rgba(245,158,11,.45)}
  .dec-reprovada{background:rgba(239,68,68,.12);color:#f87171;border-color:rgba(239,68,68,.45)}
  .dec-ressalva{background:rgba(59,130,246,.12);color:#60a5fa;border-color:rgba(59,130,246,.45)}
  .dec-misto{background:rgba(16,185,129,.08);color:#a7f3d0;border-color:rgba(16,185,129,.35)}
  html[data-tema^="claro"] .dec-aprovada{color:#047857}
  html[data-tema^="claro"] .dec-aguardando{color:#b45309}
  html[data-tema^="claro"] .dec-reprovada{color:#b91c1c}
  html[data-tema^="claro"] .dec-ressalva{color:#1d4ed8}
  html[data-tema^="claro"] .dec-misto{color:#065f46}
  .va-dec{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:6px;font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted)}
  .va-dec span:not(.dec-badge){margin-left:2px}
  .va-dec .dec-badge{text-transform:none;letter-spacing:0}
  .status-badge{display:inline-flex;align-items:center;gap:5px;font-size:11px;font-weight:600;padding:3px 9px;border-radius:20px}
  .status-parcial{background:rgba(245,158,11,.15);color:#fbbf24;border:1px solid rgba(245,158,11,.3)}
  .status-aguardando{background:rgba(148,163,184,.1);color:var(--subtle);border:1px solid var(--border)}
  .pct-bar-cell{min-width:120px}
  .pct-inline{display:flex;align-items:center;gap:8px}
  .pct-track{flex:1;height:5px;background:var(--surface2);border-radius:99px;overflow:hidden}
  .pct-fill{height:100%;border-radius:99px}
  .pct-num{font-family:'DM Mono',monospace;font-size:11px;color:var(--subtle);min-width:32px;text-align:right}
  .count-pill{background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-family:'DM Mono',monospace;font-size:12px;color:var(--subtle)}
  .setor-tag{font-size:11px;padding:2px 8px;border-radius:4px;font-weight:500}
  .setor-corp{background:rgba(37,99,235,.15);color:#93c5fd}
  .setor-tec{background:rgba(124,58,237,.15);color:#c4b5fd}
  .empresa-toggle{display:flex;gap:8px;margin-bottom:20px}
  .toggle-btn{background:var(--surface);border:1px solid var(--border);border-radius:8px;padding:8px 18px;font-size:13px;font-weight:500;color:var(--muted);cursor:pointer;font-family:'Inter',sans-serif;transition:all .15s}
  .toggle-btn:hover{color:var(--text)}
  .toggle-btn.al.active{background:var(--al-glow);color:#60a5fa;border-color:#2563eb55}
  .toggle-btn.se.active{background:var(--se-glow);color:#a78bfa;border-color:#7c3aed55}
  .toggle-btn.re.active{background:var(--re-glow);color:#fdba74;border-color:#f9731655}
  .alert{border-radius:10px;padding:14px 18px;margin-bottom:20px;font-size:13px;display:flex;align-items:flex-start;gap:10px;border:1px solid}
  .alert-red{background:rgba(239,68,68,.08);border-color:rgba(239,68,68,.3);color:#fca5a5}
  .alert-icon{font-size:16px;margin-top:1px;flex-shrink:0}
  footer{text-align:center;padding:24px 40px;font-size:12px;color:var(--muted);border-top:1px solid var(--border);margin-top:48px}
</style>
<style>
  /* GRUPO_NOME_PH ocupa so o slot 'se'; os slots herdados da Novatec ficam ocultos */
  #filtro-todas,#filtro-se,#filtro-al,#filtro-rec,#bloco-al,#bloco-rec,
  [id$="-toggle-se"],[id$="-toggle-al"],[id$="-toggle-rec"]{display:none!important}
  #empresa-grid{grid-template-columns:1fr!important}
  #resumo-grid>:nth-child(n+2),#anl-resumo-grid>:nth-child(n+2),
  #anl-col-al,#anl-col-rec,#anl-col-total{display:none!important}
  #resumo-grid,#anl-resumo-grid,#anl-cols-grid{grid-template-columns:minmax(0,1fr)!important}
  .se-corpo{display:grid;grid-template-columns:minmax(300px,1fr) 2fr}
  .se-esq{display:flex;flex-direction:column;justify-content:space-between}
  /* grafico da empresa ocupa a coluna (antes ficava pequeno no canto) */
  .se-esq .donut-wrap{flex:1;justify-content:center;gap:34px;padding:28px 24px}
  .se-esq .donut-canvas{width:190px;height:190px}
  .se-esq .donut-legend{gap:16px}
  .se-esq .legend-item{font-size:15px}
  .se-esq .legend-dot{width:11px;height:11px}
  .se-esq .legend-item{white-space:nowrap}
  @media(max-width:1400px){.se-esq .donut-canvas{width:150px;height:150px}.se-esq .donut-wrap{gap:22px;padding:24px 18px}.se-esq .legend-item{font-size:14px}}
  @media(max-width:1100px){.se-esq .donut-canvas{width:140px;height:140px}.se-esq .donut-wrap{gap:20px}}
  @media(max-width:600px){.se-esq .donut-canvas{width:130px;height:130px}.se-esq .legend-item{font-size:13px}.se-esq .donut-wrap{gap:18px;padding:20px 16px}}
  .se-setores{border-left:1px solid var(--border);padding:20px 24px 22px}
  .setor-titulo{font-size:13px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--muted);margin-bottom:14px}
  .prox-atual b{color:var(--white);font-variant-numeric:tabular-nums;font-weight:700}
  .prox-lbl-c{display:none}
  @media(max-width:600px){.prox-lbl{display:none}.prox-lbl-c{display:inline}}
  .logo-texto{height:52px;display:flex;align-items:center;padding:0 16px;border-radius:8px;background:var(--surface2);
    border:1px solid var(--border);color:var(--white);font-weight:800;font-size:18px;white-space:nowrap}
  @media(max-width:600px){.logo-texto{height:38px;font-size:14px;padding:0 10px}}
  .setor-titulo span{font-weight:400;text-transform:none;letter-spacing:0;margin-left:6px;font-size:12px}
  .setor-lista{display:grid;grid-template-columns:1fr 1fr;gap:12px 28px}
  .setor-top{display:flex;justify-content:space-between;gap:10px;font-size:13px;margin-bottom:5px}
  .setor-nome{color:var(--white);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .setor-num{color:var(--muted);white-space:nowrap}
  .setor-num strong{color:var(--white)}
  .setor-track{height:6px;background:var(--surface2);border-radius:99px;overflow:hidden}
  .setor-fill{height:100%;border-radius:99px}
  .setor-vazio{color:var(--muted);font-size:13px}
  /* ── Paletas (botao de paleta no cabecalho) ── */
  html[data-tema="marinho"]{--bg:#0a1324;--surface:#0f1d36;--surface2:#16284a;--border:#24385e;--text:#dbe5f5;--muted:#8fa3c4;--subtle:#c3d0e6;--white:#f5f8ff}
  html[data-tema="grafite"]{--bg:#121212;--surface:#1b1b1d;--surface2:#242427;--border:#34343a;--text:#e4e4e7;--muted:#a1a1aa;--subtle:#d4d4d8;--white:#fafafa}
  html[data-tema="millena"]{--bg:#0a1836;--surface:#0f2350;--surface2:#163066;--border:#26437e;--text:#e3eaf8;--muted:#9fb1d6;--subtle:#cdd8ee;--white:#ffffff;--al:#f5c400;--se:#f5c400;--se-glow:rgba(245,196,0,.16)}
  html[data-tema="claro"]{--bg:#f3f5f9;--surface:#ffffff;--surface2:#eef2f7;--border:#d9e0ea;--text:#1f2937;--muted:#64748b;--subtle:#334155;--white:#0f172a}
  html[data-tema="claro-millena"]{--bg:#fff8dc;--surface:#ffffff;--surface2:#fff1b8;--border:#efd98a;--text:#1f2a44;--muted:#5b6b8c;--subtle:#24345c;--white:#0b2a6b;--al:#1d4fb8;--se:#1d4fb8;--se-glow:rgba(29,79,184,.12)}
  html[data-tema^="claro"] #filtro-data-de,html[data-tema^="claro"] #filtro-data-ate,
  html[data-tema^="claro"] #filtro-turno,html[data-tema^="claro"] #vg-search{background:var(--surface2)!important;border-color:var(--border)!important;color:var(--text)!important;color-scheme:light!important}
  html[data-tema^="claro"] [style*="rgba(255,255,255,.06)"]{background:var(--surface2)!important}
  html[data-tema^="claro"] [style*="rgba(255,255,255,.08)"],html[data-tema^="claro"] [style*="rgba(255,255,255,.14)"]{border-color:var(--border)!important}
  html[data-tema^="claro"] .card,html[data-tema^="claro"] .table-wrap,html[data-tema^="claro"] .empresa-block{box-shadow:0 1px 3px rgba(15,23,42,.06)}
  html[data-tema^="claro"] .btn-ver{color:#5b21b6}
  html[data-tema^="claro"] .alert-red{color:#b91c1c;background:rgba(239,68,68,.07)}
  html[data-tema^="claro"] .badge-se{color:var(--se);border-color:var(--se)}
  html[data-tema^="claro"] #filtro-turno option{background:#fff;color:#1f2937}
  html[data-tema^="claro"] #modal-atividades{background:rgba(15,23,42,.45)}
  body,header,.nav,.card,.table-wrap,.empresa-block,th,td{transition:background-color .25s,border-color .25s,color .25s}
  .tema-wrap{position:relative;display:flex;align-items:center;gap:10px}
  .tema-btn{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:8px;height:34px;padding:0 12px;
    display:flex;align-items:center;gap:6px;cursor:pointer;font-family:inherit;font-size:12px;font-weight:600}
  .tema-btn:hover{border-color:var(--muted)}
  .tema-pop{position:absolute;right:0;top:42px;background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:12px;
    width:250px;box-shadow:0 12px 40px rgba(0,0,0,.35);z-index:200;display:none}
  .tema-pop.aberto{display:block}
  .tema-pop h4{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:0 0 8px;font-weight:600}
  .tema-op{display:flex;align-items:center;gap:10px;width:100%;background:none;border:1px solid transparent;border-radius:8px;padding:7px 8px;
    cursor:pointer;color:var(--text);font-family:inherit;font-size:13px;text-align:left}
  .tema-op:hover{background:var(--surface2)}
  .tema-op.ativo{border-color:var(--al);background:var(--surface2)}
  .tema-amostra{display:flex;border-radius:6px;overflow:hidden;border:1px solid var(--border);flex-shrink:0}
  .tema-amostra i{display:block;width:12px;height:22px}
  @media(max-width:600px){.tema-btn span{display:none}}
  /* ── Celular e tablet (Android e iOS) ── */
  html{-webkit-text-size-adjust:100%;text-size-adjust:100%}
  body{overflow-x:hidden}
  button,.nav-btn,.setor-row,.tema-op{-webkit-tap-highlight-color:transparent;touch-action:manipulation}
  .table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  @media(hover:none){.card:hover{transform:none;box-shadow:none}tr:hover td{background:transparent}}
  @media(max-width:900px){
    header{padding:12px 16px;padding-top:max(12px,env(safe-area-inset-top));padding-left:max(16px,env(safe-area-inset-left));padding-right:max(16px,env(safe-area-inset-right))}
    .nav{padding:0 8px;overflow-x:auto;overflow-y:hidden;flex-wrap:nowrap;scrollbar-width:none;-webkit-overflow-scrolling:touch;position:sticky;top:0;z-index:90}
    .nav::-webkit-scrollbar{display:none}
    .nav-btn{flex:0 0 auto;white-space:nowrap;padding:12px 14px;font-size:14px}
    header{position:relative}
    main{padding:18px 14px calc(28px + env(safe-area-inset-bottom));padding-left:max(14px,env(safe-area-inset-left));padding-right:max(14px,env(safe-area-inset-right))}
    .cards{grid-template-columns:repeat(2,minmax(0,1fr))!important;gap:10px}
    .cards > #card-prazo{grid-column:1/-1}
    .card{padding:10px 12px}
    .card svg{width:34px!important;height:34px!important}
    .card-value{font-size:22px}
    [style*="grid-template-columns:repeat(3"],[style*="grid-template-columns:repeat(4"]{grid-template-columns:repeat(2,minmax(0,1fr))!important}
    [style*="grid-template-columns:1fr 1fr"]{grid-template-columns:minmax(0,1fr)!important}
    .table-header{flex-wrap:wrap;gap:8px}
    .table-header input,.search-input,#vg-search{max-width:none!important;width:100%;flex:1 1 100%}
    table{min-width:620px}
    th,td{padding:9px 10px}
    .donut-wrap{padding:18px 16px;gap:16px}
    .se-setores{padding:16px}
    .setor-lista{grid-template-columns:1fr}
    /* iOS amplia a tela quando o campo tem fonte menor que 16px */
    input,select,textarea{font-size:16px!important}
    .btn-ver{padding:8px 12px;font-size:12px}
    .filtros-linha{width:100%}
    .filtros-linha > div{width:100%}
    .filtro-periodo{flex-wrap:wrap}
    .filtro-periodo input{flex:1 1 130px;min-width:0}
    #filtro-turno{flex:1}
    #modal-atividades{padding:0;align-items:flex-start}
    .va-caixa{max-width:none;border-radius:0;min-height:100vh;min-height:100dvh;border:0}
    .va-cab{border-radius:0;padding:12px 14px;padding-top:max(12px,env(safe-area-inset-top));flex-wrap:wrap}
    .va-corpo{padding:12px 14px calc(20px + env(safe-area-inset-bottom))}
    .va-titulo{font-size:16px}
    .tema-pop{position:fixed;right:12px;left:auto;top:calc(56px + env(safe-area-inset-top));width:min(260px,calc(100vw - 24px))}
  }
  @media(max-width:520px){
    .logo-millena{height:34px}
    .logo-text{font-size:14px!important}
    .logo-sub{font-size:11px}
    .card-label{font-size:10px}
    .card-value{font-size:20px}
    .va-acoes{width:100%;justify-content:space-between}
  }


  .btn-ver{background:rgba(124,58,237,.15);border:1px solid rgba(124,58,237,.5);color:#c4b5fd;border-radius:6px;
    padding:4px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;font-family:inherit}
  .btn-ver:hover{background:rgba(124,58,237,.3);color:#fff}
  #modal-atividades{position:fixed;inset:0;background:rgba(0,0,0,.65);z-index:1000;align-items:flex-start;justify-content:center;padding:40px 16px;overflow:auto}
  .va-caixa{background:var(--surface);border:1px solid var(--border);border-radius:14px;width:100%;max-width:820px;box-shadow:0 20px 60px rgba(0,0,0,.5)}
  .va-cab{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;padding:18px 22px;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--surface);border-radius:14px 14px 0 0}
  .va-titulo{font-size:18px;font-weight:700;color:var(--white)}
  .va-subt{font-size:12px;color:var(--muted);margin-top:3px}
  .va-fechar{background:none;border:1px solid var(--border);color:var(--text);border-radius:8px;width:32px;height:32px;cursor:pointer;font-size:14px;flex-shrink:0}
  .va-fechar:hover{background:rgba(255,255,255,.08)}
  .va-acoes{display:flex;gap:8px;align-items:center;flex-shrink:0}
  .va-exportar{background:#10b981;border:none;color:#fff;border-radius:8px;height:32px;padding:0 12px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;white-space:nowrap}
  .va-exportar:hover{background:#059669}
  .va-exportar:disabled{opacity:.6;cursor:wait}
  .va-corpo{padding:14px 22px 22px;display:flex;flex-direction:column;gap:10px}
  .va-item{border:1px solid var(--border);border-radius:10px;padding:12px 14px;background:var(--surface2,var(--surface))}
  .va-topo{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;color:var(--white);font-size:14px}
  .va-meta{font-size:11px;color:var(--muted);margin:4px 0 6px}
  .va-det{font-size:13px;color:var(--text);white-space:pre-line;line-height:1.45}
  .va-campo{font-size:12px;color:var(--text);margin-top:6px}
  .va-campo span{display:block;font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:1px}
  .va-vazio{color:var(--muted);font-size:13px;padding:10px 0}
  .setor-titulo{display:flex;align-items:center;flex-wrap:wrap;gap:6px}
  .setor-row{cursor:pointer;border-radius:8px;padding:6px 8px;margin:-6px -8px;transition:background .15s,opacity .15s}
  .setor-row:hover{background:rgba(255,255,255,.05)}
  .setor-row:focus-visible{outline:2px solid #7c3aed;outline-offset:1px}
  .setor-row.ativo{background:rgba(124,58,237,.18);box-shadow:inset 0 0 0 1px rgba(124,58,237,.6)}
  .setor-row.apagado{opacity:.45}
  .setor-chip{margin-left:auto;background:rgba(124,58,237,.18);border:1px solid rgba(124,58,237,.6);color:var(--white);
    border-radius:99px;padding:3px 10px;font-size:12px;font-family:inherit;text-transform:none;letter-spacing:0;cursor:pointer}
  @media(max-width:1100px){.setor-lista{grid-template-columns:1fr}}
  @media(max-width:800px){.se-corpo{grid-template-columns:1fr}.se-setores{border-left:0;border-top:1px solid var(--border)}}

  /* ── Saudacao e relogio no cabecalho ── */
  .saudacao{display:flex;align-items:center;gap:12px;background:var(--surface2);border:1px solid var(--border);border-radius:12px;padding:6px 14px 6px 6px}
  .sd-avatar{width:40px;height:40px;border-radius:50%;border:2px solid var(--border);background:linear-gradient(135deg,var(--se),var(--al));
    color:#fff;font-weight:700;font-size:14px;cursor:pointer;flex-shrink:0;font-family:inherit}
  .sd-texto{display:flex;flex-direction:column;line-height:1.2;min-width:0}
  #sd-ola{font-size:12px;color:var(--muted)}
  .sd-nome{background:none;border:0;padding:0;color:var(--white);font-weight:700;font-size:15px;cursor:pointer;text-align:left;font-family:inherit;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:180px}
  .sd-nome.vazio{font-size:12px;font-weight:600;color:var(--al);text-decoration:underline dotted}
  .sd-recarregar{background:none;border:0;color:var(--muted);cursor:pointer;width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center}
  .sd-recarregar:hover{color:var(--white);background:rgba(127,127,127,.15)}
  .sd-recarregar.girando svg{animation:sdGira .6s linear infinite}
  @keyframes sdGira{to{transform:rotate(360deg)}}
  .sd-sep{width:1px;align-self:stretch;background:var(--border);margin:4px 0}
  .sd-relogio{display:flex;flex-direction:column;align-items:flex-end;line-height:1.15}
  #sd-data{font-size:12px;color:var(--muted);white-space:nowrap}
  #sd-hora{font-size:22px;font-weight:700;color:var(--white);font-variant-numeric:tabular-nums}
  #filtro-data-info{display:none!important}
  .card-detalhe,#det-novas-se{display:none!important}  /* uma empresa so: o detalhe repetia o numero sem filtro */  /* aviso amarelo removido: o botao Filtros (n) indica o filtro */
  #card-novas{cursor:pointer}
  #card-novas:hover{border-color:#f59e0b}
  #card-novas.novas-ativo{border-color:#f59e0b;box-shadow:0 0 0 2px rgba(245,158,11,.45)}
  #card-novas:focus-visible{outline:2px solid #f59e0b;outline-offset:2px}
  /* ── Filtros escondidos ── */
  .filtros-area{display:flex;flex-direction:column;align-items:flex-end;gap:8px;margin-left:auto}
  .filtros-btn{display:flex;align-items:center;gap:7px;background:var(--surface2);border:1px solid var(--border);color:var(--text);
    border-radius:8px;padding:7px 12px;font-size:13px;font-weight:600;cursor:pointer;font-family:inherit}
  .filtros-btn:hover,.filtros-btn.ativo{border-color:var(--al);color:var(--white)}
  .filtros-qtd{background:var(--al);color:#fff;border-radius:99px;font-size:11px;min-width:18px;height:18px;padding:0 5px;display:none;align-items:center;justify-content:center}
  .filtros-qtd:not(:empty){display:inline-flex}
  #filtros-painel:not(.aberto){display:none!important}
  .vg-ordem-box{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:8px;padding:3px 4px 3px 10px}
  .vg-ordem-box span{font-size:11px;color:var(--muted);white-space:nowrap}
  .vg-ordem-box select{background:var(--surface2);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:12px;padding:4px 6px;font-family:inherit;max-width:190px}
  .vg-ordem-box button{font-size:11px;padding:5px 9px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;font-weight:600;white-space:nowrap;font-family:inherit}
  tr.vg-grupo td{background:rgba(124,58,237,.14)!important;color:var(--white);font-weight:700;font-size:12px;letter-spacing:.04em;text-transform:uppercase;padding:8px 12px!important;text-align:left!important}
  tr.vg-grupo td small{font-weight:500;color:var(--muted);text-transform:none;letter-spacing:0;margin-left:6px;font-size:11px}
  @media(max-width:640px){.vg-ordem-box{width:100%;box-sizing:border-box;min-width:0}.vg-ordem-box select{flex:1;width:0;min-width:0;max-width:none;font-size:14px;padding:7px}.vg-ordem-box button{padding:8px 10px}
    .table-wrap tr.vg-grupo{padding:0;border-bottom:0}.table-wrap tr.vg-grupo td{padding:10px 12px!important}}
  .setor-picker{position:relative}
  .sp-btn{display:flex;align-items:center;gap:8px;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:6px 10px;
    color:var(--text);font-family:inherit;cursor:pointer;width:100%;min-height:34px}
  .sp-label{flex:1;text-align:left;font-size:12px;font-weight:600;color:var(--white);max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .sp-seta{color:var(--muted);font-size:11px}
  .sp-menu{display:none;position:absolute;right:0;top:calc(100% + 6px);z-index:60;min-width:260px;max-width:92vw;background:var(--surface);
    border:1px solid var(--border);border-radius:10px;box-shadow:0 12px 30px rgba(0,0,0,.45);padding:8px}
  .setor-picker.aberto .sp-menu{display:block}
  .sp-acoes{display:flex;gap:6px;margin-bottom:6px}
  .sp-acoes button{flex:1;font-size:12px;padding:7px 8px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text);cursor:pointer;font-family:inherit;font-weight:600}
  .sp-lista{max-height:300px;overflow:auto}
  .sp-item{display:flex;align-items:center;gap:10px;padding:9px 8px;border-radius:6px;font-size:13px;color:var(--text);cursor:pointer}
  .sp-item:hover{background:var(--surface2)}
  .sp-item input{width:18px;height:18px;accent-color:#7c3aed;flex:none}
  .sp-vazio{color:var(--muted);font-size:12px;padding:8px}
  @media(max-width:900px){.setor-picker{width:100%}.sp-label{max-width:none}.sp-menu{left:0;right:0;min-width:0}}
  @media(max-width:900px){
    .saudacao{padding:4px 8px 4px 4px;gap:8px}
    .sd-avatar{width:34px;height:34px;font-size:12px}
    .sd-texto,.sd-sep,#sd-data{display:none}
    #sd-hora{font-size:16px}
    .filtros-area{width:100%;align-items:stretch}
    .filtros-btn{justify-content:center}
  }
  /* ── Celular: ajustes finais (precisam vir por ultimo) ── */
  @media(max-width:900px){
    header{flex-wrap:nowrap!important;gap:10px}
    header .logo{min-width:0;flex:1}
    header .logo > div{min-width:0}
    .tema-wrap{flex-shrink:0}
    #modal-atividades{padding:0!important;align-items:flex-start!important}
    .va-caixa{height:auto!important}
    .va-caixa{max-width:none!important;border-radius:0!important;min-height:100vh;min-height:100dvh;border:0!important}
    .va-cab{border-radius:0!important}
  }
  .va-topo .status-badge{white-space:nowrap;flex-shrink:0}
  @media(max-width:640px){
    /* tabelas viram cartoes: cada linha vira um bloco com o nome do campo ao lado */
    .table-wrap table{min-width:0!important;width:100%}
    .table-wrap thead{display:none}
    .table-wrap tbody,.table-wrap tr{display:block;width:100%}
    .table-wrap tr{border-bottom:1px solid var(--border);padding:8px 12px}
    .table-wrap td{display:flex;justify-content:space-between;align-items:center;gap:12px;border:0!important;padding:4px 0!important;text-align:right!important;white-space:normal!important;max-width:none!important}
    .table-wrap td::before{content:attr(data-rotulo);font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);text-align:left;flex-shrink:0}
    .table-wrap td[data-rotulo=""]::before,.table-wrap td:not([data-rotulo])::before{display:none}
    .table-wrap td:first-child{font-size:15px;justify-content:flex-start;text-align:left!important;padding-top:2px!important}
    .table-wrap td:first-child::before{display:none}
    .table-wrap td:empty{display:none}
    .table-wrap td .btn-ver{width:100%;padding:10px;font-size:13px}
    .table-wrap td:has(.btn-ver)::before{display:none}
    tr:hover td{background:transparent!important}
    /* cabecalho em duas linhas: logo + titulo; saudacao + tema */
    header{flex-wrap:wrap!important;row-gap:10px}
    header .logo{flex:1 1 100%}
    .tema-wrap{flex:1 1 100%;gap:8px}
    .saudacao{flex:1;min-width:0}
    .sd-texto{display:flex;flex:1}
    .sd-nome{max-width:100%}
    .sd-texto{min-width:0;overflow:hidden}
    .tema-wrap{min-width:0}
    .sd-sep,#sd-data{display:flex}
    #sd-data{font-size:11px}
  }
  @media(max-width:380px){ #sd-data,.sd-sep{display:none!important} .saudacao{gap:6px} }
</style>
</head>
<body>
<script>try{var _t=localStorage.getItem('painel_tema');if(_t)document.documentElement.setAttribute('data-tema',_t);}catch(e){}</script>
<header>
  <div class="logo">
    LOGO_HTML_PH
    <div>
      <div class="logo-text">Status do Mapeamento</div>
      <div class="logo-sub">Atualizado em HORA_REF<span class="prox-atual" id="prox-atual" title="Próxima atualização automática dos dados"></span></div>
    </div>
  </div>
  <div class="tema-wrap">
    <div class="saudacao">
      <button type="button" class="sd-avatar" id="sd-avatar" title="Informar seu nome">?</button>
      <div class="sd-texto"><span id="sd-ola">Olá,</span>
        <button type="button" class="sd-nome" id="sd-nome" title="Informar seu nome">Toque para se identificar</button></div>
      <button type="button" class="sd-recarregar" id="sd-recarregar" title="Buscar dados novos" aria-label="Buscar dados novos">
        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg></button>
      <div class="sd-sep"></div>
      <div class="sd-relogio"><span id="sd-data">DATE_HOJE</span><strong id="sd-hora">DATE_HORA</strong></div>
    </div>
    <button type="button" class="tema-btn" id="tema-btn" aria-haspopup="true" aria-expanded="false" title="Trocar cores">🎨 <span>Tema</span></button>
    <div class="tema-pop" id="tema-pop" role="menu">
      <h4>Paleta de cores</h4>
      <div id="tema-lista"></div>
    </div>
  </div>
</header>

<nav class="nav">
  <button class="nav-btn active" onclick="showTab('visao-geral',this)">✨ Visão Geral</button>
  <button class="nav-btn" onclick="showTab('por-cargo',this)">Por Cargo</button>
  <button class="nav-btn" onclick="showTab('nao-enviaram',this)">Não Enviaram</button>
  <button class="nav-btn" onclick="showTab('enviaram',this)">Enviaram</button>
  <button class="nav-btn" onclick="showTab('atividades',this)">Atividades</button>
  <button class="nav-btn" onclick="showTab('analise-atividades',this)">Análise das Atividades</button>
  <button class="nav-btn" id="nav-aprovacoes" onclick="showTab('aprovacoes',this)" style="display:APROV_NAV_DISPLAY">Aprovações</button>
</nav>

<main>

<!-- VISÃO GERAL -->
<div id="tab-visao-geral" class="tab-panel active">
  <div id="filtro-data-info" style="display:none;font-size:12px;color:#f59e0b;margin-bottom:14px;padding:8px 14px;background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.25);border-radius:8px"></div>
  ALERT_CRITICOS_HTML
  <div class="empresa-toggle" style="margin-bottom:16px;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
    <div style="display:flex;gap:8px;flex-wrap:wrap">
      <button class="toggle-btn active" id="filtro-todas" onclick="filterEmpresa('todas')">Todas as filiais</button>
      <button class="toggle-btn se" id="filtro-se" onclick="filterEmpresa('se')">GRUPO_NOME_PH</button>
      <button class="toggle-btn al" id="filtro-al" onclick="filterEmpresa('al')">Novatec Alagoas</button>
      <button class="toggle-btn re" id="filtro-rec" onclick="filterEmpresa('rec')">Novatec Recife</button>
    </div>
    <div class="filtros-area">
    <button type="button" class="filtros-btn" id="filtros-btn" aria-expanded="false" aria-controls="filtros-painel">
      <svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 5h18l-7 8v6l-4 2v-8z"/></svg>
      Filtros <span class="filtros-qtd" id="filtros-qtd"></span></button>
    <div class="filtros-linha" id="filtros-painel" style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
      <div class="filtro-periodo" style="display:flex;align-items:center;gap:6px;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:6px 10px">
        <span style="font-size:11px;color:var(--muted);white-space:nowrap">📅 Período</span>
        <input type="date" id="filtro-data-de" onchange="aplicarFiltroData()" style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.14);border-radius:6px;color:var(--text);font-size:11px;padding:3px 6px;font-family:'DM Mono',monospace;color-scheme:dark">
        <span style="font-size:11px;color:var(--muted)">até</span>
        <input type="date" id="filtro-data-ate" onchange="aplicarFiltroData()" style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.14);border-radius:6px;color:var(--text);font-size:11px;padding:3px 6px;font-family:'DM Mono',monospace;color-scheme:dark">
        <button onclick="limparFiltroData()" title="Limpar filtro de data" style="font-size:11px;padding:4px 9px;border-radius:6px;border:none;cursor:pointer;background:rgba(255,255,255,.08);color:var(--subtle);font-weight:600">Limpar</button>
      </div>
      <div style="display:flex;align-items:center;gap:6px;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:6px 10px">
        <span style="font-size:11px;color:var(--muted);white-space:nowrap">🕐 Turno</span>
        <select id="filtro-turno" onchange="aplicarFiltroData()" style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.14);border-radius:6px;color:var(--text);font-size:11px;padding:3px 6px;font-family:Inter,sans-serif">
          <option value="todos">Todos</option>
          <option value="manha">Manhã</option>
          <option value="tarde">Tarde</option>
          <option value="noite">Noite</option>
        </select>
      </div>
      <div id="setor-picker" class="setor-picker">
        <button type="button" id="setor-picker-btn" class="sp-btn" onclick="abrirSetorPicker(event)" aria-expanded="false">
          <span style="font-size:11px;color:var(--muted);white-space:nowrap">🏷️ Setores</span>
          <span id="setor-picker-label" class="sp-label">Todos</span><span class="sp-seta">▾</span></button>
        <div class="sp-menu">
          <div class="sp-acoes"><button type="button" onclick="setorPickerTodos(true)">Marcar todos</button><button type="button" onclick="setorPickerTodos(false)">Limpar</button></div>
          <div id="setor-picker-lista" class="sp-lista"></div>
        </div>
      </div>
    </div>
    </div>
  </div>
  <div class="cards" style="grid-template-columns:repeat(7,1fr)">
    <div class="card al">
      <svg style="position:absolute;right:12px;top:50%;transform:translateY(-50%);opacity:.18;width:52px;height:52px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M12 12c2.7 0 4.8-2.1 4.8-4.8S14.7 2.4 12 2.4 7.2 4.5 7.2 7.2 9.3 12 12 12zm0 2.4c-3.2 0-9.6 1.6-9.6 4.8v2.4h19.2v-2.4c0-3.2-6.4-4.8-9.6-4.8z"/></svg>
      <div class="card-label">Cadastrados</div>
      <div class="card-value" id="val-cadastrados">TOTAL_USERS</div>
      <div class="card-detail card-detalhe">GRUPO_NOME_PH: SE_TOTAL</div>
    </div>
    <div class="card green">
      <svg style="position:absolute;right:12px;top:50%;transform:translateY(-50%);opacity:.18;width:52px;height:52px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 14.5l-4-4 1.41-1.41L10 13.67l6.59-6.59L18 8.5l-8 8z"/></svg>
      <div class="card-label">Concluídos</div>
      <div class="card-value" id="val-concluidos">ENV_TOTAL</div>
      <div class="card-detail card-detalhe">GRUPO_NOME_PH: SE_ENV</div>
    </div>
    <div class="card red">
      <svg style="position:absolute;right:12px;top:50%;transform:translateY(-50%);opacity:.18;width:52px;height:52px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z"/></svg>
      <div class="card-label">Pendentes</div>
      <div class="card-value" id="val-pendentes">NENV_TOTAL</div>
      <div class="card-detail card-detalhe">GRUPO_NOME_PH: SE_NENV</div>
    </div>
    <div class="card se">
      <svg style="position:absolute;right:12px;top:50%;transform:translateY(-50%);opacity:.18;width:52px;height:52px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M5 9.2h3V19H5V9.2zM10.6 5h2.8v14h-2.8V5zm5.6 8H19v6h-2.8v-6z"/></svg>
      <div class="card-label">Atividades</div>
      <div class="card-value" id="val-atividades">ATV_TOTAL</div>
      <div class="card-detail card-detalhe">GRUPO_NOME_PH: SE_ATV</div>
    </div>
    <div class="card teal">
      <svg style="position:absolute;right:12px;top:50%;transform:translateY(-50%);opacity:.18;width:52px;height:52px" viewBox="0 0 24 24" fill="var(--white)" xmlns="http://www.w3.org/2000/svg"><path d="M3.5 18.49l6-6.01 4 4L22 6.92l-1.41-1.41-7.09 7.97-4-4L2 16.99z"/></svg>
      <div class="card-label">Adesão Global</div>
      <div class="card-value" id="val-adesao">PCT_GLOBAL%</div>
    </div>
    NOVOS_ENVIOS_CARD
    <div class="card orange" id="card-prazo">
      <div class="card-label">Temporazador</div>
      <div id="prazo-countdown" style="display:flex;gap:3px;align-items:flex-end;margin:2px 0 6px">
        <div style="text-align:center"><span id="p-dias" style="font-size:22px;font-weight:700;color:var(--white);font-family:'DM Mono',monospace">--</span><div style="font-size:10px;color:var(--white);margin-top:1px;letter-spacing:.04em">DIAS</div></div>
        <span style="font-size:16px;font-weight:700;color:var(--muted);margin-bottom:10px">:</span>
        <div style="text-align:center"><span id="p-horas" style="font-size:22px;font-weight:700;color:var(--white);font-family:'DM Mono',monospace">--</span><div style="font-size:10px;color:var(--muted);margin-top:1px;letter-spacing:.04em">HRS</div></div>
        <span style="font-size:16px;font-weight:700;color:var(--muted);margin-bottom:10px">:</span>
        <div style="text-align:center"><span id="p-min" style="font-size:22px;font-weight:700;color:var(--white);font-family:'DM Mono',monospace">--</span><div style="font-size:10px;color:var(--muted);margin-top:1px;letter-spacing:.04em">MIN</div></div>
        <span style="font-size:16px;font-weight:700;color:var(--muted);margin-bottom:10px">:</span>
        <div style="text-align:center"><span id="p-seg" style="font-size:22px;font-weight:700;color:var(--white);font-family:'DM Mono',monospace">--</span><div style="font-size:10px;color:var(--muted);margin-top:1px;letter-spacing:.04em">SEG</div></div>
      </div>
      <div class="card-detail" id="prazo-txt" style="color:var(--white)">Prazo PRAZO_TXT</div>
    </div>
  </div>

  <div class="empresa-grid" id="empresa-grid">
    <div class="empresa-block" id="bloco-se">
      <div class="empresa-header se">
        <span class="empresa-name">GRUPO_NOME_PH <span id="se-trophy">SE_TROPHY_HTML</span></span>
        <span class="empresa-badge badge-se"><span id="se-pct-badge">SE_PCT</span>% adesão</span>
      </div>
      <div class="se-corpo">
      <div class="se-esq">
      <div class="donut-wrap">
        <canvas class="donut-canvas" id="donut-se" width="420" height="420"></canvas>
        <div class="donut-legend">
          <div class="legend-item"><div class="legend-dot" style="background:#10b981"></div><span>Enviaram — <strong id="se-enviaram-count">SE_ENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div><span>Não enviaram — <strong id="se-nao-enviaram-count">SE_NENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#7c3aed"></div><span>Atividades — <strong id="se-atividades-count">SE_ATV</strong></span></div>
        </div>
      </div>
      <div class="progress-section">
        <div class="progress-label"><span>Taxa de envio</span><span style="color:var(--white);font-weight:700" id="se-pct-label">SE_PCT%</span></div>
        <div class="progress-track"><div class="progress-fill" id="se-progress-fill" style="width:SE_PCT_RAW%;background:SE_PROGRESS_COLOR"></div></div>
      </div>
      </div>
      <div class="se-setores">
        <div id="emp-area" style="display:none;margin-bottom:22px">
        <div class="setor-titulo">Adesão por empresa <span>enviaram / cadastrados · clique para filtrar</span>
          <button id="emp-chip" class="setor-chip" style="display:none" onclick="filtrarEmpresaGrupo('')">Empresa: <b></b> ✕</button></div>
        <div class="setor-lista" id="emp-lista"></div>
        </div>
        <div class="setor-titulo">Adesão por setor <span>enviaram / cadastrados · clique para filtrar</span>
          <button id="setor-chip" class="setor-chip" style="display:none" onclick="filtrarSetor('')">Setor: <b></b> ✕</button></div>
        <div class="setor-lista" id="setor-lista">SETOR_HTML_PLACEHOLDER</div>
      </div>
      </div>
    </div>
    <div class="empresa-block" id="bloco-al">
      <div class="empresa-header al">
        <span class="empresa-name">Novatec Alagoas <span id="al-trophy">AL_TROPHY_HTML</span></span>
        <span class="empresa-badge badge-al"><span id="al-pct-badge">PCT_AL</span>% adesão</span>
      </div>
      <div class="donut-wrap">
        <canvas class="donut-canvas" id="donut-al" width="120" height="120"></canvas>
        <div class="donut-legend">
          <div class="legend-item"><div class="legend-dot" style="background:#10b981"></div><span>Enviaram — <strong id="al-enviaram-count">AL_ENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div><span>Não enviaram — <strong id="al-nao-enviaram-count">AL_NENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#2563eb"></div><span>Atividades — <strong id="al-atividades-count">AL_ATV</strong></span></div>
        </div>
      </div>
      <div class="progress-section">
        <div class="progress-label"><span>Taxa de envio</span><span style="color:var(--white);font-weight:700" id="al-pct-label">PCT_AL%</span></div>
        <div class="progress-track"><div class="progress-fill" id="al-progress-fill" style="width:PCT_AL_RAW%;background:AL_PROGRESS_COLOR"></div></div>
      </div>
    </div>
    <div class="empresa-block" id="bloco-rec">
      <div class="empresa-header re">
        <span class="empresa-name">Novatec Recife <span id="rec-trophy">RE_TROPHY_HTML</span></span>
        <span class="empresa-badge badge-re"><span id="rec-pct-badge">RE_PCT</span>% adesão</span>
      </div>
      <div class="donut-wrap">
        <canvas class="donut-canvas" id="donut-rec" width="120" height="120"></canvas>
        <div class="donut-legend">
          <div class="legend-item"><div class="legend-dot" style="background:#10b981"></div><span>Enviaram — <strong id="rec-enviaram-count">RE_ENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#ef4444"></div><span>Não enviaram — <strong id="rec-nao-enviaram-count">RE_NENV</strong></span></div>
          <div class="legend-item"><div class="legend-dot" style="background:#f97316"></div><span>Atividades — <strong id="rec-atividades-count">RE_ATV</strong></span></div>
        </div>
      </div>
      <div class="progress-section">
        <div class="progress-label"><span>Taxa de envio</span><span style="color:var(--white);font-weight:700" id="rec-pct-label">RE_PCT%</span></div>
        <div class="progress-track"><div class="progress-fill" id="rec-progress-fill" style="width:RE_PCT_RAW%;background:REC_PROGRESS_COLOR"></div></div>
      </div>
    </div>
    <div id="bloco-chart" style="display:none;background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:22px;align-self:stretch;flex-direction:column">
      <div style="font-size:13px;font-weight:700;color:var(--white);letter-spacing:.03em;margin-bottom:4px" id="chart-titulo">Envios por Data</div>
      <div style="font-size:11px;color:var(--muted);margin-bottom:16px" id="chart-subtitulo">Funcionários únicos que enviaram atividades por dia</div>
      <div style="flex:1;position:relative;min-height:180px">
        <canvas id="chart-envios-data" style="width:100%;display:block"></canvas>
        <div id="chart-no-data" style="display:none;position:absolute;inset:0;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:8px">
          <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="#475569" stroke-width="1.5"><path d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
          <span style="font-size:12px;color:var(--muted)">Sem dados de data disponíveis no relatório</span>
        </div>
      </div>
    </div>
  </div>

  <div class="table-wrap">
    <div class="table-header" style="flex-wrap:wrap;gap:8px;align-items:center">
      <div style="min-width:180px">
        <div class="table-title">👥 Funcionários — Enviaram e Pendentes</div>
        <span class="table-count" id="count-tb-enviaram">—</span>
      </div>
      <input id="vg-search" type="text" placeholder="🔍 Buscar funcionário..." oninput="filtrarEnviaram(window._vgFiltro)" style="flex:1;min-width:180px;max-width:300px;padding:6px 12px;border-radius:8px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.06);color:var(--text);font-size:12px;outline:none" />
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-left:auto">
        <div style="display:flex;gap:3px;background:rgba(255,255,255,.06);border-radius:8px;padding:3px;border:1px solid rgba(255,255,255,.08)">
          <button id="fg-btn-todos" onclick="filtrarEnviaram('todos')" style="font-size:11px;padding:4px 11px;border-radius:6px;border:none;cursor:pointer;font-weight:600;transition:all .15s">Todos</button>
          <button id="fg-btn-enviaram" onclick="filtrarEnviaram('enviaram')" style="font-size:11px;padding:4px 11px;border-radius:6px;border:none;cursor:pointer;font-weight:400;transition:all .15s">Enviaram</button>
          <button id="fg-btn-pendentes" onclick="filtrarEnviaram('pendentes')" style="font-size:11px;padding:4px 11px;border-radius:6px;border:none;cursor:pointer;font-weight:400;transition:all .15s">Pendentes</button>
        </div>
        <div class="vg-ordem-box">
          <span>↕ Ordenar</span>
          <select id="vg-ordem" onchange="mudarOrdemVG()">
            <option value="setor">Setor</option>
            <option value="nome">Nome</option>
            <option value="cargo">Cargo</option>
            <option value="atividades">Atividades</option>
            <option value="gestor">Gestor (aguardando aprovação)</option>
            <option value="consultor">Consultor (aguardando aprovação)</option>
            <option value="status">Status (pendentes primeiro)</option>
            <option value="empresa">Empresa</option>
          </select>
          <button type="button" id="vg-dir" onclick="inverterOrdemVG()" title="Inverter a ordem">A→Z</button>
        </div>
        <button onclick="exportarVisaoGeral(event)" style="font-size:11px;padding:5px 12px;border-radius:6px;border:1px solid rgba(16,185,129,.3);cursor:pointer;background:rgba(16,185,129,.12);color:#34d399;font-weight:600;white-space:nowrap">⬇ Exportar Excel</button>
      </div>
    </div>
    <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Filial</th><th>Atividades</th><th>Gestor</th><th>Consultor</th><th></th></tr></thead>
    <tbody id="tb-enviaram"></tbody></table>
  </div>
</div>

<!-- POR CARGO -->
<div id="tab-por-cargo" class="tab-panel">
  <div class="section-title">Análise por Cargo</div>
  <div class="section-sub">Adesão e volume de atividades por função</div>
  <div class="empresa-toggle">
    <button class="toggle-btn se active" id="cargo-toggle-se" onclick="cargoToggle('se')">GRUPO_NOME_PH</button>
    <button class="toggle-btn al" id="cargo-toggle-al" onclick="cargoToggle('al')">Novatec Alagoas</button>
    <button class="toggle-btn re" id="cargo-toggle-rec" onclick="cargoToggle('rec')">Novatec Recife</button>
  </div>
  <div id="cargo-se">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">GRUPO_NOME_PH — por cargo</div>
        <span class="table-count">N_CARGOS_SE cargos</span>
      </div>
      <table><thead><tr><th>Cargo</th><th>Cadastrados</th><th>Enviaram</th><th>Atividades</th><th>Adesão</th></tr></thead>
      <tbody id="tb-cargo-se"></tbody></table>
    </div>
  </div>
  <div id="cargo-al" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Alagoas — por cargo</div>
        <span class="table-count">N_CARGOS_AL cargos</span>
      </div>
      <table><thead><tr><th>Cargo</th><th>Cadastrados</th><th>Enviaram</th><th>Atividades</th><th>Adesão</th></tr></thead>
      <tbody id="tb-cargo-al"></tbody></table>
    </div>
  </div>
  <div id="cargo-rec" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Recife — por cargo</div>
        <span class="table-count">N_CARGOS_REC cargos</span>
      </div>
      <table><thead><tr><th>Cargo</th><th>Cadastrados</th><th>Enviaram</th><th>Atividades</th><th>Adesão</th></tr></thead>
      <tbody id="tb-cargo-rec"></tbody></table>
    </div>
  </div>
</div>

<!-- NÃO ENVIARAM -->
<div id="tab-nao-enviaram" class="tab-panel">
  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:6px">
    <div class="section-title" style="margin-bottom:0">Cadastrados sem Envio</div>
    <button onclick="exportarPendentes()" style="display:inline-flex;align-items:center;gap:8px;background:#10b981;border:none;border-radius:8px;padding:9px 18px;font-size:13px;font-weight:600;color:#fff;cursor:pointer;font-family:Inter,sans-serif;transition:opacity .15s" onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Exportar Excel
    </button>
  </div>
  <div class="section-sub">Funcionários registrados no sistema que não submeteram nenhuma atividade</div>
  <div class="empresa-toggle">
    <button class="toggle-btn se active" id="nao-toggle-se" onclick="naoToggle('se')">GRUPO_NOME_PH — SE_NENV pessoas</button>
    <button class="toggle-btn al" id="nao-toggle-al" onclick="naoToggle('al')">Alagoas — AL_NENV pessoas</button>
    <button class="toggle-btn re" id="nao-toggle-rec" onclick="naoToggle('rec')">Recife — RE_NENV pessoas</button>
  </div>
  <div id="nao-se">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">GRUPO_NOME_PH — sem envio</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('nao-se-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="width:90px">Status</th></tr></thead>
      <tbody id="nao-se-body"></tbody></table>
    </div>
  </div>
  <div id="nao-al" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Alagoas — sem envio</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('nao-al-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="width:90px">Status</th></tr></thead>
      <tbody id="nao-al-body"></tbody></table>
    </div>
  </div>
  <div id="nao-rec" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Recife — sem envio</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('nao-rec-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="width:90px">Status</th></tr></thead>
      <tbody id="nao-rec-body"></tbody></table>
    </div>
  </div>
</div>

<!-- ENVIARAM -->
<div id="tab-enviaram" class="tab-panel">
  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:6px">
    <div class="section-title" style="margin-bottom:0">Cadastrados com Envio</div>
    <button onclick="exportarEnviaram()" style="display:inline-flex;align-items:center;gap:8px;background:#2563eb;border:none;border-radius:8px;padding:9px 18px;font-size:13px;font-weight:600;color:#fff;cursor:pointer;font-family:Inter,sans-serif;transition:opacity .15s" onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Exportar Excel
    </button>
  </div>
  <div class="section-sub">Funcionários que submeteram atividades no sistema</div>
  <div class="empresa-toggle">
    <button class="toggle-btn se active" id="env-toggle-se" onclick="envToggle('se')">GRUPO_NOME_PH — SE_ENV pessoas</button>
    <button class="toggle-btn al" id="env-toggle-al" onclick="envToggle('al')">Alagoas — AL_ENV pessoas</button>
    <button class="toggle-btn re" id="env-toggle-rec" onclick="envToggle('rec')">Recife — RE_ENV pessoas</button>
  </div>
  <div id="env-se">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">GRUPO_NOME_PH — enviaram</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('env-se-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="text-align:center">Atividades</th></tr></thead>
      <tbody id="env-se-body"></tbody></table>
    </div>
  </div>
  <div id="env-al" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Alagoas — enviaram</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('env-al-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="text-align:center">Atividades</th></tr></thead>
      <tbody id="env-al-body"></tbody></table>
    </div>
  </div>
  <div id="env-rec" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Recife — enviaram</div>
        <input class="search-input" placeholder="Filtrar por nome ou cargo…" oninput="filterTable('env-rec-body',this.value)">
      </div>
      <table><thead><tr><th>Nome</th><th>Cargo</th><th>Setor</th><th>Perfil</th><th style="text-align:center">Atividades</th></tr></thead>
      <tbody id="env-rec-body"></tbody></table>
    </div>
  </div>
</div>

<!-- ATIVIDADES -->
<div id="tab-atividades" class="tab-panel">
  <div class="section-title">Atividades Enviadas</div>
  <div class="section-sub">Listagem completa das atividades submetidas — ATV_TOTAL no total</div>

  <!-- Resumo executivo -->
  <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:24px" id="resumo-grid">
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · GRUPO_NOME_PH</div>
      <div id="resumo-se-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · Novatec Alagoas</div>
      <div id="resumo-al-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · Novatec Recife</div>
      <div id="resumo-rec-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
  </div>

  <!-- Versões & Parecer -->
  <div style="margin-bottom:24px">
    <div class="section-title" style="font-size:18px;margin-bottom:4px">Versões &amp; Parecer</div>
    <div class="section-sub" style="margin-bottom:12px">Edições e pareceres do consultor registrados no relatório</div>
    <div class="empresa-toggle" style="margin-bottom:12px">
      <button class="toggle-btn se active" id="ver-toggle-se" onclick="verToggle('se')">GRUPO_NOME_PH</button>
      <button class="toggle-btn al" id="ver-toggle-al" onclick="verToggle('al')">Alagoas</button>
      <button class="toggle-btn re" id="ver-toggle-rec" onclick="verToggle('rec')">Recife</button>
    </div>
    <div id="ver-se">
      <div class="table-wrap">
        <table><thead><tr><th>Funcionário</th><th>Atividade</th><th>Tipo</th><th>Autor</th><th>Data/hora</th><th>Conteúdo</th></tr></thead>
        <tbody id="ver-se-body"></tbody></table>
      </div>
    </div>
    <div id="ver-al" style="display:none">
      <div class="table-wrap">
        <table><thead><tr><th>Funcionário</th><th>Atividade</th><th>Tipo</th><th>Autor</th><th>Data/hora</th><th>Conteúdo</th></tr></thead>
        <tbody id="ver-al-body"></tbody></table>
      </div>
    </div>
    <div id="ver-rec" style="display:none">
      <div class="table-wrap">
        <table><thead><tr><th>Funcionário</th><th>Atividade</th><th>Tipo</th><th>Autor</th><th>Data/hora</th><th>Conteúdo</th></tr></thead>
        <tbody id="ver-rec-body"></tbody></table>
      </div>
    </div>
  </div>

  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:20px">
    <div class="empresa-toggle" style="margin-bottom:0">
      <button class="toggle-btn se active" id="atv-toggle-se" onclick="atvToggle('se')">GRUPO_NOME_PH — SE_ATV atividades</button>
      <button class="toggle-btn al" id="atv-toggle-al" onclick="atvToggle('al')">Alagoas — AL_ATV atividades</button>
      <button class="toggle-btn re" id="atv-toggle-rec" onclick="atvToggle('rec')">Recife — RE_ATV atividades</button>
    </div>
    <button onclick="exportarAtividades(event)" style="display:inline-flex;align-items:center;gap:8px;background:#7c3aed;border:none;border-radius:8px;padding:9px 18px;font-size:13px;font-weight:600;color:#fff;cursor:pointer;font-family:Inter,sans-serif;transition:opacity .15s" onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Exportar Excel
    </button>
  </div>
  <div id="atv-se">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">GRUPO_NOME_PH</div>
        <input class="search-input" placeholder="Filtrar atividades…" oninput="filterTable('atv-se-body',this.value)">
      </div>
      <table style="table-layout:fixed;width:100%"><thead><tr><th style="width:140px">Funcionário</th><th style="width:70px">Horário</th><th style="width:120px">Cargo</th><th style="width:160px">Atividade</th><th>Detalhamento</th><th style="width:110px">Periodicidade</th><th style="width:70px">Tempo</th><th style="width:100px">Status</th></tr></thead>
      <tbody id="atv-se-body"></tbody></table>
    </div>
  </div>
  <div id="atv-al" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Alagoas</div>
        <input class="search-input" placeholder="Filtrar atividades…" oninput="filterTable('atv-al-body',this.value)">
      </div>
      <table style="table-layout:fixed;width:100%"><thead><tr><th style="width:140px">Funcionário</th><th style="width:70px">Horário</th><th style="width:120px">Cargo</th><th style="width:160px">Atividade</th><th>Detalhamento</th><th style="width:110px">Periodicidade</th><th style="width:70px">Tempo</th><th style="width:100px">Status</th></tr></thead>
      <tbody id="atv-al-body"></tbody></table>
    </div>
  </div>
  <div id="atv-rec" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Recife</div>
        <input class="search-input" placeholder="Filtrar atividades…" oninput="filterTable('atv-rec-body',this.value)">
      </div>
      <table style="table-layout:fixed;width:100%"><thead><tr><th style="width:140px">Funcionário</th><th style="width:70px">Horário</th><th style="width:120px">Cargo</th><th style="width:160px">Atividade</th><th>Detalhamento</th><th style="width:110px">Periodicidade</th><th style="width:70px">Tempo</th><th style="width:100px">Status</th></tr></thead>
      <tbody id="atv-rec-body"></tbody></table>
    </div>
  </div>
</div>

<!-- ANÁLISE DAS ATIVIDADES -->
<div id="tab-analise-atividades" class="tab-panel">
  <div class="section-title">Análise das Atividades</div>
  <div class="section-sub">Totais e composição das atividades registradas</div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px" id="anl-resumo-grid">
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · GRUPO_NOME_PH</div>
      <div id="anl-resumo-se-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · Novatec Alagoas</div>
      <div id="anl-resumo-al-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · Novatec Recife</div>
      <div id="anl-resumo-rec-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px">
      <div style="font-size:12px;font-weight:700;text-transform:uppercase;color:var(--muted);letter-spacing:.05em;margin-bottom:12px">Resumo · Todas as Empresas</div>
      <div id="anl-resumo-total-cards" style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px"></div>
    </div>
  </div>

  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:4px">
    <div class="section-title" style="margin-bottom:0;font-size:18px">Atividades — totais por coluna</div>
    <button onclick="exportarAnalise()" style="display:inline-flex;align-items:center;gap:8px;background:#14b8a6;border:none;border-radius:8px;padding:9px 18px;font-size:13px;font-weight:600;color:#fff;cursor:pointer;font-family:Inter,sans-serif;transition:opacity .15s" onmouseover="this.style.opacity='.85'" onmouseout="this.style.opacity='1'">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Exportar Excel
    </button>
  </div>
  <div class="section-sub" style="margin-bottom:16px">Periodicidade, forma de execução, status, campos preenchidos e criação por mês — exclui Horário fixo, Decisão gestor, Decisão consultor, Edições do consultor e Parecer do consultor</div>

  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px" id="anl-cols-grid">
    <div id="anl-col-se"></div>
    <div id="anl-col-al"></div>
    <div id="anl-col-rec"></div>
    <div id="anl-col-total"></div>
  </div>
</div>

<!-- APROVAÇÕES -->
<div id="tab-aprovacoes" class="tab-panel">
  <div class="section-title">Aprovações</div>
  <div class="section-sub">Cruzamento entre atividades enviadas (Relatório) e status de aprovação</div>

  <div class="cards">
    <div class="card al">
      <div class="card-label">Total Aprovações</div>
      <div class="card-value" id="apr-val-total">APR_TOTAL</div>
      <div class="card-detail apr-det">GRUPO_NOME_PH: APR_SE_TOTAL</div>
    </div>
    <div class="card green">
      <div class="card-label">Aprovados</div>
      <div class="card-value" id="apr-val-aprov">APR_APROV</div>
      <div class="card-detail apr-det">GRUPO_NOME_PH: APR_SE_APROV</div>
    </div>
    <div class="card orange">
      <div class="card-label">Pendentes</div>
      <div class="card-value" id="apr-val-pend">APR_PEND</div>
      <div class="card-detail apr-det">GRUPO_NOME_PH: APR_SE_PEND</div>
    </div>
    <div class="card red">
      <div class="card-label">Reprovados</div>
      <div class="card-value" id="apr-val-reprov">APR_REPROV</div>
      <div class="card-detail apr-det">GRUPO_NOME_PH: APR_SE_REPROV</div>
    </div>
  </div>

  <div class="empresa-toggle">
    <button class="toggle-btn se active" id="apr-toggle-se" onclick="aprToggle('se')">GRUPO_NOME_PH — APR_SE_TOTAL registros</button>
    <button class="toggle-btn al" id="apr-toggle-al" onclick="aprToggle('al')">Alagoas — APR_AL_TOTAL registros</button>
    <button class="toggle-btn re" id="apr-toggle-rec" onclick="aprToggle('rec')">Recife — APR_REC_TOTAL registros</button>
  </div>

  <div id="apr-se">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">GRUPO_NOME_PH — Aprovações</div>
        <input class="search-input" placeholder="Filtrar por nome ou atividade…" oninput="filterTable('apr-se-body',this.value)">
      </div>
      <table><thead><tr><th>Funcionário</th><th>Cargo</th><th>Atividade</th><th>Status</th><th>Aprovador</th><th>Data</th><th>No Relatório</th></tr></thead>
      <tbody id="apr-se-body"></tbody></table>
    </div>
  </div>
  <div id="apr-al" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Alagoas — Aprovações</div>
        <input class="search-input" placeholder="Filtrar por nome ou atividade…" oninput="filterTable('apr-al-body',this.value)">
      </div>
      <table><thead><tr><th>Funcionário</th><th>Cargo</th><th>Atividade</th><th>Status</th><th>Aprovador</th><th>Data</th><th>No Relatório</th></tr></thead>
      <tbody id="apr-al-body"></tbody></table>
    </div>
  </div>
  <div id="apr-rec" style="display:none">
    <div class="table-wrap">
      <div class="table-header">
        <div class="table-title">Novatec Recife — Aprovações</div>
        <input class="search-input" placeholder="Filtrar por nome ou atividade…" oninput="filterTable('apr-rec-body',this.value)">
      </div>
      <table><thead><tr><th>Funcionário</th><th>Cargo</th><th>Atividade</th><th>Status</th><th>Aprovador</th><th>Data</th><th>No Relatório</th></tr></thead>
      <tbody id="apr-rec-body"></tbody></table>
    </div>
  </div>
</div>

</main>
<footer>GRUPO_NOME_PH — Relatório de Mapeamento de Atividades · Gerado em DATE_HOJE</footer>

<script>
const DATA = DATA_JSON_PLACEHOLDER;

// ── Filtro de data (recalcula tudo no navegador a partir de DATA.*.atividades) ──
function _normJS(s){
  if(!s) return '';
  return s.toString().trim().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'');
}
function _parseDataISO(raw){
  if(!raw) return null;
  let v=String(raw).split(',')[0].split('T')[0].split(' ')[0].trim();
  if(!v) return null;
  const parts=v.replace(/-/g,'/').split('/');
  if(parts.length!==3) return null;
  let d,m,y;
  if(parts[0].length===4){ y=parts[0]; m=parts[1]; d=parts[2]; } else { d=parts[0]; m=parts[1]; y=parts[2]; }
  d=d.padStart(2,'0'); m=m.padStart(2,'0');
  if(y.length!==4||!/^\d{4}$/.test(y)) return null;
  return `${y}-${m}-${d}`;
}
function _dentroIntervalo(raw,de,ate){
  if(!de && !ate) return true;
  const iso=_parseDataISO(raw);
  if(!iso) return false;
  if(de && iso<de) return false;
  if(ate && iso>ate) return false;
  return true;
}
function _extrairHorario(raw){
  if(!raw) return null;
  const s=String(raw).trim();
  let resto=null;
  if(s.includes(',')) resto=s.split(',')[1];
  else if(s.includes('T')) resto=s.split('T')[1];
  else if(s.includes(' ')) resto=s.split(' ')[1];
  if(!resto) return null;
  const m=resto.trim().match(/^(\d{1,2}):(\d{2})/);
  if(!m) return null;
  const h=parseInt(m[1],10), min=parseInt(m[2],10);
  if(h>23||min>59) return null;
  return {h, m:min, texto:String(h).padStart(2,'0')+':'+String(min).padStart(2,'0')};
}
function _classificarTurno(raw){
  const info=_extrairHorario(raw);
  if(!info) return null;
  if(info.h<12) return 'manha';
  if(info.h<18) return 'tarde';
  return 'noite';
}
function _passaFiltroAtividade(a,de,ate,turno){
  if(!_dentroIntervalo(a.criada_em,de,ate)) return false;
  if(turno && turno!=='todos' && _classificarTurno(a.criada_em)!==turno) return false;
  return true;
}
function _computeEnviosPorData(atividades){
  const porData={};
  atividades.forEach(a=>{
    const iso=_parseDataISO(a.criada_em);
    if(!iso) return;
    const [y,m,d]=iso.split('-');
    const label=d+'/'+m+'/'+y;
    if(!porData[label]) porData[label]=new Set();
    porData[label].add(_normJS(a.funcionario));
  });
  const resultado={};
  Object.keys(porData).sort((a,b)=>{
    const [da,ma,ya]=a.split('/'), [db,mb,yb]=b.split('/');
    return (ya+ma+da).localeCompare(yb+mb+db);
  }).forEach(k=>{ resultado[k]=porData[k].size; });
  return resultado;
}
function _registroEmpresa(raw){
  const map=new Map();
  (raw.enviaram_list||[]).forEach(p=>map.set(_normJS(p.nome),{nome:p.nome,cargo:p.cargo,setor:p.setor,perfil:p.perfil,empresa:p.empresa,status_usr:'Ativo'}));
  (raw.nao_enviaram_list||[]).forEach(p=>{const k=_normJS(p.nome); if(!map.has(k)) map.set(k,{nome:p.nome,cargo:p.cargo,setor:p.setor,perfil:p.perfil,empresa:p.empresa,status_usr:p.status_usr});});
  return map;
}
window._registros={};
window._setorFiltro=[];   // varios setores: lista vazia = todos
function _temSetor(f){return Array.isArray(f)?f.length>0:!!f;}
function _setorBate(f,st){return Array.isArray(f)?f.includes(st):f===st;}
function _setorDe(p){return ((p&&p.setor)||'').toString().trim()||'(sem setor)';}
window._empFiltro='';
function _empDe(p){return ((p&&p.empresa)||'').toString().trim()||'(sem empresa)';}
function _pessoaPassa(info,setorF,empF){
  if(setorF===undefined) setorF=window._setorFiltro;
  if(empF===undefined) empF=window._empFiltro;
  return !!info && (!_temSetor(setorF)||_setorBate(setorF,_setorDe(info))) && (!empF||_empDe(info)===empF);
}
function _calcularEmpresaFiltrada(empKey,de,ate,turno,setorF){
  const raw=DATA[empKey];
  if(setorF===undefined) setorF=window._setorFiltro;
  const empF=window._empFiltro;
  const temSetor=_temSetor(setorF);
  if(!de && !ate && (!turno||turno==='todos') && !temSetor && !empF) return raw;
  if(!window._registros[empKey]) window._registros[empKey]=_registroEmpresa(raw);
  let registro=window._registros[empKey];
  if(temSetor||empF) registro=new Map([...registro].filter(([n,info])=>_pessoaPassa(info,setorF,empF)));
  const atividadesFiltradas=(raw.atividades||[]).filter(a=>_passaFiltroAtividade(a,de,ate,turno)
    && (!(temSetor||empF) || registro.has(_normJS(a.funcionario))));
  const contagem={};
  atividadesFiltradas.forEach(a=>{const n=_normJS(a.funcionario); contagem[n]=(contagem[n]||0)+1;});
  const enviaram_list=[], nao_enviaram_list=[];
  registro.forEach((info,n)=>{
    if(contagem[n]) enviaram_list.push({nome:info.nome,cargo:info.cargo,setor:info.setor,perfil:info.perfil,empresa:info.empresa,atividades:contagem[n]});
    else nao_enviaram_list.push({nome:info.nome,cargo:info.cargo,setor:info.setor,perfil:info.perfil,empresa:info.empresa,status_usr:info.status_usr||'Ativo'});
  });
  enviaram_list.sort((a,b)=>b.atividades-a.atividades);
  nao_enviaram_list.sort((a,b)=>a.nome.localeCompare(b.nome));
  const cargoMap={};
  registro.forEach(info=>{const c=info.cargo||''; if(!c) return; if(!cargoMap[c]) cargoMap[c]={total:0,enviaram:0,atividades:0}; cargoMap[c].total++;});
  enviaram_list.forEach(p=>{const c=p.cargo||''; if(c && cargoMap[c]){cargoMap[c].enviaram++; cargoMap[c].atividades+=p.atividades;}});
  const cargo_data=Object.entries(cargoMap).map(([cargo,d])=>({cargo,total:d.total,enviaram:d.enviaram,atividades:d.atividades,pct:Math.round(d.enviaram/Math.max(d.total,1)*100)})).sort((a,b)=>b.atividades-a.atividades);
  const envios_por_data=_computeEnviosPorData(atividadesFiltradas);
  return {...raw, total:(temSetor||empF)?registro.size:raw.total, enviaram:enviaram_list.length, nao_enviaram:nao_enviaram_list.length, total_atividades:atividadesFiltradas.length, enviaram_list, nao_enviaram_list, cargo_data, atividades:atividadesFiltradas, envios_por_data};
}
function _valOuVazioJS(v){v=(v||'').toString().trim(); return v?v:'(vazio)';}
function _contarCampoJS(atividades,campo){
  const c={};
  atividades.forEach(a=>{const k=_valOuVazioJS(a[campo]); c[k]=(c[k]||0)+1;});
  return Object.fromEntries(Object.entries(c).sort((x,y)=>y[1]-x[1]));
}
function _analiseFromAtividades(atividades){
  const pontos_criticos_count=atividades.filter(a=>(a.pontos_criticos||'').toString().trim()).length;
  const sugestoes_count=atividades.filter(a=>(a.sugestoes||'').toString().trim()).length;
  const porMes={};
  atividades.forEach(a=>{
    const iso=_parseDataISO(a.criada_em);
    let mes='(vazio)';
    if(iso){const [y,m]=iso.split('-'); mes=`${m}/${y}`;}
    porMes[mes]=(porMes[mes]||0)+1;
  });
  const chaveMes=x=>x==='(vazio)'?'9999-99':x.split('/').reverse().join('-');
  const por_mes=Object.fromEntries(Object.entries(porMes).sort((a,b)=>chaveMes(a[0]).localeCompare(chaveMes(b[0]))));
  return {
    total:atividades.length,
    periodicidade:_contarCampoJS(atividades,'periodicidade'),
    como_realiza:_contarCampoJS(atividades,'como_realiza'),
    status:_contarCampoJS(atividades,'status'),
    pontos_criticos_count, sugestoes_count, por_mes,
  };
}
window._empDataAtual={alagoas:DATA.alagoas, sergipe:DATA.sergipe, recife:DATA.recife};
window._analiseAtual={
  alagoas: DATA.analise && DATA.analise.alagoas,
  sergipe: DATA.analise && DATA.analise.sergipe,
  recife:  DATA.analise && DATA.analise.recife,
  total:   DATA.analise && DATA.analise.total,
};
function EMP(key){return window._empDataAtual[key];}
function ANALISE(key){return window._analiseAtual[key];}
function _recalcularFiltro(de,ate,turno){
  window._empDataAtual={
    alagoas:_calcularEmpresaFiltrada('alagoas',de,ate,turno),
    sergipe:_calcularEmpresaFiltrada('sergipe',de,ate,turno),
    recife:_calcularEmpresaFiltrada('recife',de,ate,turno),
  };
  const ativo=!!(de||ate||(turno&&turno!=='todos')||_temSetor(window._setorFiltro)||window._empFiltro);
  const atvTotal=[...window._empDataAtual.alagoas.atividades,...window._empDataAtual.sergipe.atividades,...window._empDataAtual.recife.atividades];
  window._analiseAtual={
    alagoas: ativo?_analiseFromAtividades(window._empDataAtual.alagoas.atividades):(DATA.analise&&DATA.analise.alagoas),
    sergipe: ativo?_analiseFromAtividades(window._empDataAtual.sergipe.atividades):(DATA.analise&&DATA.analise.sergipe),
    recife:  ativo?_analiseFromAtividades(window._empDataAtual.recife.atividades):(DATA.analise&&DATA.analise.recife),
    total:   ativo?_analiseFromAtividades(atvTotal):(DATA.analise&&DATA.analise.total),
  };
}
function _empresaAtiva(){
  for(const e of ['al','se','rec']){
    const btn=document.getElementById('filtro-'+e);
    if(btn && btn.classList.contains('active')) return e;
  }
  return 'todas';
}
function _salvarFiltroLocal(de,ate,turno){
  try{ localStorage.setItem('novatec_filtro_data', JSON.stringify({de,ate,turno})); }catch(e){}
}
function _carregarFiltroLocal(){
  try{ return JSON.parse(localStorage.getItem('novatec_filtro_data')||'null'); }catch(e){ return null; }
}
const _TURNO_LABEL={manha:'Manhã (00h–12h)', tarde:'Tarde (12h–18h)', noite:'Noite (18h–24h)'};
function _hojeISO(){ const d=new Date(); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); }
function _filtroEhHoje(){ const h=_hojeISO(); return document.getElementById('filtro-data-de').value===h && document.getElementById('filtro-data-ate').value===h; }
function _atualizarNovas(){
  const card=document.getElementById('card-novas'); if(!card) return;
  const h=_hojeISO(), filtra=!!(_temSetor(window._setorFiltro)||window._empFiltro);
  const reg=_registroEmpresa(DATA.sergipe);
  const n=(DATA.sergipe.atividades||[]).filter(a=>_parseDataISO(a.criada_em)===h
     && (!filtra || _pessoaPassa(reg.get(_normJS(a.funcionario))))).length;
  const txt=n?('+'+n):'0';
  card.dataset.total=txt; card.dataset.se=txt;
  const v=document.getElementById('val-novas'); if(v) v.textContent=txt;
  const d=document.getElementById('det-novas-se'); if(d) d.textContent=n?('GRUPO_NOME_PH: '+txt):'nenhuma hoje';
  card.classList.toggle('pink',n>0);
  card.classList.toggle('novas-ativo',_filtroEhHoje());
  card.setAttribute('aria-pressed',_filtroEhHoje());
}
function filtrarNovasHoje(){
  const de=document.getElementById('filtro-data-de'), ate=document.getElementById('filtro-data-ate');
  if(_filtroEhHoje()){ de.value=''; ate.value=''; } else { de.value=_hojeISO(); ate.value=_hojeISO(); }
  aplicarFiltroData();
  if(_filtroEhHoje()){ const t=document.getElementById('tb-enviaram'); if(t) t.closest('.table-wrap').scrollIntoView({behavior:'smooth',block:'start'}); }
}
document.addEventListener('click',e=>{ if(e.target.closest&&e.target.closest('#card-novas')) filtrarNovasHoje(); });
document.addEventListener('keydown',e=>{ if((e.key==='Enter'||e.key===' ')&&e.target.id==='card-novas'){ e.preventDefault(); filtrarNovasHoje(); } });
function _atualizarInfoFiltro(){
  _atualizarNovas();
  const de=document.getElementById('filtro-data-de').value;
  const ate=document.getElementById('filtro-data-ate').value;
  const turno=document.getElementById('filtro-turno').value;
  const box=document.getElementById('filtro-data-info');
  const setorF=_temSetor(window._setorFiltro)?window._setorFiltro.join(', '):'';
  const qtd=(de||ate?1:0)+(turno&&turno!=='todos'?1:0)+(setorF?1:0);
  _renderSetorPicker();
  const q=document.getElementById('filtros-qtd'); if(q){ q.textContent=qtd||''; q.style.display=qtd?'':'none'; }
  if(!de && !ate && (!turno||turno==='todos') && !setorF && !window._empFiltro){ box.style.display='none'; return; }
  const fmt=iso=>iso?iso.split('-').reverse().join('/'):'';
  const partes=[];
  if(de && ate) partes.push('período '+fmt(de)+' até '+fmt(ate));
  else if(de) partes.push('a partir de '+fmt(de));
  else if(ate) partes.push('até '+fmt(ate));
  if(turno && turno!=='todos') partes.push('turno '+_TURNO_LABEL[turno]);
  if(setorF) partes.push('setor '+setorF);
  let texto='📅 Filtrando por '+partes.join(' · ');
  texto+=' — cartões, tabelas, gráficos e análise consideram apenas esse filtro.'
    +(setorF?' Aprovações também são filtradas pelo setor; o Resumo do relatório não.':' (Aprovações e Resumo não são afetados.)');
  box.textContent=texto;
  box.style.display='block';
}
function _pctColorJS(p){return p>=50?'var(--green)':'var(--red)';}
function atualizarBlocosEmpresa(){
  [{key:'alagoas',pre:'al'},{key:'sergipe',pre:'se'},{key:'recife',pre:'rec'}].forEach(({key,pre})=>{
    const d=EMP(key);
    const pct=d.enviaram/Math.max(d.total,1)*100;
    const pctFmt=pct.toFixed(1).replace('.',',');
    document.getElementById(pre+'-pct-badge').textContent=pctFmt;
    document.getElementById(pre+'-enviaram-count').textContent=d.enviaram;
    document.getElementById(pre+'-nao-enviaram-count').textContent=d.nao_enviaram;
    document.getElementById(pre+'-atividades-count').textContent=d.total_atividades;
    document.getElementById(pre+'-pct-label').textContent=pctFmt+'%';
    document.getElementById(pre+'-progress-fill').style.width=pct+'%';
    document.getElementById(pre+'-progress-fill').style.background=_pctColorJS(pct);
    document.getElementById(pre+'-trophy').innerHTML=pct>=100?' <span style="font-size:20px;vertical-align:middle;filter:drop-shadow(0 0 8px gold)">\U0001F3C6</span>':'';
  });
}
function aplicarFiltroData(){
  const de=document.getElementById('filtro-data-de').value;
  const ate=document.getElementById('filtro-data-ate').value;
  const turno=document.getElementById('filtro-turno').value;
  if(de && ate && de>ate){ alert('A data inicial não pode ser depois da data final.'); return; }
  _recalcularFiltro(de,ate,turno);
  _salvarFiltroLocal(de,ate,turno);
  populate();
  atualizarBlocosEmpresa();
  atualizarDonuts();
  filterEmpresa(_empresaAtiva());
  _atualizarInfoFiltro();
}
function limparFiltroData(){
  document.getElementById('filtro-data-de').value='';
  document.getElementById('filtro-data-ate').value='';
  document.getElementById('filtro-turno').value='todos';
  window._setorFiltro=[]; window._empFiltro='';
  _recalcularFiltro('','','todos');
  _salvarFiltroLocal('','','todos');
  populate();
  atualizarBlocosEmpresa();
  atualizarDonuts();
  filterEmpresa(_empresaAtiva());
  _atualizarInfoFiltro();
}

function renderSetores(){
  const box=document.getElementById('setor-lista'); if(!box) return;
  const de=document.getElementById('filtro-data-de').value;
  const ate=document.getElementById('filtro-data-ate').value;
  const turno=document.getElementById('filtro-turno').value;
  const d=_calcularEmpresaFiltrada('sergipe',de,ate,turno,'');
  const reg=_registroEmpresa(DATA.sergipe), env={};
  (d.enviaram_list||[]).forEach(p=>env[_normJS(p.nome)]=p.atividades||0);
  const m={};
  renderEmpresas(reg,de,ate,turno);
  reg.forEach((info,n)=>{if(!_pessoaPassa(info,'')) return; const st=_setorDe(info); const x=m[st]||(m[st]={setor:st,total:0,enviaram:0,atividades:0});
    x.total++; if(n in env){x.enviaram++; x.atividades+=env[n];}});
  const lista=Object.values(m).map(x=>({...x,pct:Math.round(x.enviaram/Math.max(x.total,1)*100)}))
    .sort((a,b)=>(b.enviaram/b.total-a.enviaram/a.total)||(b.total-a.total)||a.setor.localeCompare(b.setor));
  const sel=window._setorFiltro, temSel=_temSetor(sel);
  window._setoresDisponiveis=lista.map(x=>x.setor).sort((a,b)=>a.localeCompare(b,'pt-BR'));
  box.innerHTML=lista.map(x=>{const cor=x.pct>=70?'#10b981':(x.pct>=30?'#f59e0b':'#ef4444');
    const cls='setor-row'+(temSel?(_setorBate(sel,x.setor)?' ativo':' apagado'):'');
    const corBarra='#8b5cf6';   // setores: roxo
    return `<div class="${cls}" data-setor="${esc(x.setor)}" role="button" tabindex="0" title="Filtrar a página por ${esc(x.setor)}"><div class="setor-top"><span class="setor-nome">${esc(x.setor)}</span><span class="setor-num"><strong>${x.enviaram}</strong>/${x.total} · ${x.atividades} ativ. · <strong style="color:${cor}">${x.pct}%</strong></span></div><div class="setor-track"><div class="setor-fill" style="width:${Math.max(x.pct,2)}%;background:${corBarra}"></div></div></div>`;}).join('')
    || '<div class="setor-vazio">Sem setores no cadastro.</div>';
  const chip=document.getElementById('setor-chip');
  if(chip){chip.style.display=temSel?'':'none'; chip.querySelector('b').textContent=sel.length===1?sel[0]:(sel.length+' setores');}
  _renderSetorPicker();
}
const _PALETA_BARRAS=['#3b82f6','#10b981','#f59e0b','#a855f7','#ec4899','#06b6d4','#f97316','#84cc16','#6366f1','#14b8a6','#eab308','#f43f5e','#0ea5e9','#22c55e','#d946ef','#fb923c'];
function _coresPorNome(nomes){
  const m={}; [...new Set(nomes)].sort((a,b)=>a.localeCompare(b,'pt-BR')).forEach((n,i)=>m[n]=_PALETA_BARRAS[i%_PALETA_BARRAS.length]);
  return m;
}
function renderEmpresas(reg,de,ate,turno){
  const caixa=document.getElementById('emp-area'), box=document.getElementById('emp-lista'); if(!box) return;
  // contagem por empresa ignora o proprio filtro de empresa e o de setor
  const salvo=window._empFiltro; window._empFiltro='';
  const d=_calcularEmpresaFiltrada('sergipe',de,ate,turno,''); window._empFiltro=salvo;
  const env={}; (d.enviaram_list||[]).forEach(p=>env[_normJS(p.nome)]=p.atividades||0);
  const m={};
  reg.forEach((info,n)=>{const e=_empDe(info); const x=m[e]||(m[e]={emp:e,total:0,enviaram:0,atividades:0});
    x.total++; if(n in env){x.enviaram++; x.atividades+=env[n];}});
  const lista=Object.values(m).map(x=>({...x,pct:Math.round(x.enviaram/Math.max(x.total,1)*100)}))
    .sort((a,b)=>(b.enviaram/b.total-a.enviaram/a.total)||(b.total-a.total)||a.emp.localeCompare(b.emp));
  if(caixa) caixa.style.display=lista.length>1?'':'none';
  const sel=window._empFiltro;
  box.innerHTML=lista.map(x=>{const cor=x.pct>=70?'#10b981':(x.pct>=30?'#f59e0b':'#ef4444');
    const cls='setor-row emp-row'+(sel?(x.emp===sel?' ativo':' apagado'):'');
    const corBarra='#0ea5e9';   // empresas: azul
    return `<div class="${cls}" data-emp="${esc(x.emp)}" role="button" tabindex="0" title="Filtrar a página por ${esc(x.emp)}"><div class="setor-top"><span class="setor-nome">${esc(x.emp)}</span><span class="setor-num"><strong>${x.enviaram}</strong>/${x.total} · ${x.atividades} ativ. · <strong style="color:${cor}">${x.pct}%</strong></span></div><div class="setor-track"><div class="setor-fill" style="width:${Math.max(x.pct,2)}%;background:${corBarra}"></div></div></div>`;}).join('');
  const chip=document.getElementById('emp-chip');
  if(chip){chip.style.display=sel?'':'none'; chip.querySelector('b').textContent=sel;}
}
function filtrarEmpresaGrupo(e){
  window._empFiltro=(window._empFiltro===e)?'':(e||'');
  window._setorFiltro=[];   // os setores mudam conforme a empresa
  aplicarFiltroData();
}
function filtrarSetor(st){
  // clicar num setor liga/desliga ele; vazio limpa todos
  if(!st){ window._setorFiltro=[]; }
  else { const a=[...window._setorFiltro]; const i=a.indexOf(st); if(i>=0) a.splice(i,1); else a.push(st); window._setorFiltro=a; }
  aplicarFiltroData();
}
function definirSetores(lista){ window._setorFiltro=[...(lista||[])]; aplicarFiltroData(); }
// ── Seletor de varios setores (dentro de Filtros) ──
function _renderSetorPicker(){
  const lb=document.getElementById('setor-picker-label'), box=document.getElementById('setor-picker-lista');
  if(!lb||!box) return;
  const sel=window._setorFiltro||[], todos=window._setoresDisponiveis||[];
  lb.textContent=!sel.length?'Todos':(sel.length===1?sel[0]:(sel.length+' selecionados'));
  const sig=todos.join('|')+'#'+sel.join('|');
  if(box.dataset.sig===sig) return;
  box.dataset.sig=sig;
  box.innerHTML=todos.map(st=>`<label class="sp-item"><input type="checkbox" value="${esc(st)}"${sel.includes(st)?' checked':''}> <span>${esc(st)}</span></label>`).join('')
    || '<div class="sp-vazio">Sem setores.</div>';
}
function abrirSetorPicker(ev){
  if(ev) ev.stopPropagation();
  const p=document.getElementById('setor-picker'); if(!p) return;
  p.classList.toggle('aberto');
  document.getElementById('setor-picker-btn').setAttribute('aria-expanded',p.classList.contains('aberto'));
}
function _setorPickerMudou(){
  const box=document.getElementById('setor-picker-lista');
  definirSetores([...box.querySelectorAll('input:checked')].map(i=>i.value));
}
function setorPickerTodos(marcar){ definirSetores(marcar?(window._setoresDisponiveis||[]):[]); }
document.addEventListener('change',e=>{ if(e.target.closest&&e.target.closest('#setor-picker-lista')) _setorPickerMudou(); });
document.addEventListener('click',e=>{ const p=document.getElementById('setor-picker');
  if(p&&p.classList.contains('aberto')&&!(e.target.closest&&e.target.closest('#setor-picker'))){ p.classList.remove('aberto'); document.getElementById('setor-picker-btn').setAttribute('aria-expanded','false'); } });
function _descricaoFiltros(){
  const de=document.getElementById('filtro-data-de').value, ate=document.getElementById('filtro-data-ate').value;
  const turno=document.getElementById('filtro-turno').value, fmt=iso=>iso?iso.split('-').reverse().join('/'):'';
  return [['Período',de||ate?((de?fmt(de):'início')+' até '+(ate?fmt(ate):'hoje')):'Todo o período'],
    ['Turno',turno&&turno!=='todos'?_TURNO_LABEL[turno]:'Todos'],
    ['Empresa',window._empFiltro||'Todas'],
    ['Setores',_temSetor(window._setorFiltro)?window._setorFiltro.join(', '):'Todos'],
    ['Exportado em',new Date().toLocaleString('pt-BR')]];
}
function _sufixoSetores(){
  const s=window._setorFiltro||[];
  if(!s.length) return '';
  const t=s.length===1?s[0]:(s.length+'_setores');
  return '_'+t.normalize('NFD').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_+|_+$/g,'').slice(0,40);
}
// Planilha .xlsx simples: aba de dados + aba "Filtros" dizendo o que foi aplicado
function _exportarTabela(cols,linhas,base,btn,aba){
  const txt=btn?btn.innerHTML:''; if(btn){ btn.disabled=true; btn.textContent='Gerando…'; }
  return _carregarExcelJS().then(EJ=>{
    const wb=new EJ.Workbook();
    const ws=wb.addWorksheet(aba||'Dados',{views:[{state:'frozen',ySplit:1}]});
    ws.columns=cols.map(c=>({header:c[0],width:c[1]}));
    linhas.forEach(l=>ws.addRow(l));
    ws.eachRow({includeEmpty:true},(row,n)=>{
      row.eachCell({includeEmpty:true},c=>{
        if(n===1){ c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9222A'}};
          c.font={name:'Calibri',size:10,bold:true,color:{argb:'FFFFFFFF'}}; c.alignment={vertical:'middle',wrapText:true}; return; }
        c.font={name:'Calibri',size:10,color:{argb:'FF1F2937'}}; c.alignment={vertical:'middle',wrapText:true};
        c.border={bottom:{style:'thin',color:{argb:'FFE5E7EB'}}};
        if(n%2===1) c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF5F6F8'}};
      });
    });
    ws.getRow(1).height=24;
    ws.autoFilter={from:{row:1,column:1},to:{row:1,column:cols.length}};
    const wf=wb.addWorksheet('Filtros');
    wf.columns=[{header:'Filtro',width:16},{header:'Valor',width:70}];
    _descricaoFiltros().forEach(l=>wf.addRow(l));
    wf.addRow(['Linhas',linhas.length]);
    wf.getRow(1).font={bold:true};
    return wb.xlsx.writeBuffer();
  }).then(buf=>{
    _baixarArquivo(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),base+'.xlsx');
  }).catch(()=>{
    const csv=[cols.map(c=>c[0]),...linhas].map(r=>r.map(c=>'"'+String(c===null||c===undefined?'':c).split('"').join('""')+'"').join(';')).join(String.fromCharCode(13,10));
    _baixarArquivo(new Blob([String.fromCharCode(65279)+csv],{type:'text/csv;charset=utf-8'}),base+'.csv');
  }).finally(()=>{ if(btn){ btn.disabled=false; btn.innerHTML=txt; } });
}
function _filtrarAprovacoes(){
  const st=_temSetor(window._setorFiltro)||window._empFiltro;
  const reg=_registroEmpresa(DATA.sergipe);
  const raw=(DATA.aprovacoes&&DATA.aprovacoes.sergipe&&DATA.aprovacoes.sergipe.itens)||[];
  const itens=st?raw.filter(a=>_pessoaPassa(reg.get(_normJS(a.funcionario)))):raw;
  const cont=k=>itens.filter(a=>(a.status||'').toLowerCase().startsWith(k)).length;
  const put=(id,v)=>{const el=document.getElementById(id); if(el) el.textContent=v;};
  put('apr-val-total',itens.length); put('apr-val-aprov',cont('aprov'));
  put('apr-val-pend',cont('pend')); put('apr-val-reprov',cont('reprov'));
  document.querySelectorAll('.apr-det').forEach(el=>el.style.display=st?'none':'');
  return itens;
}
document.addEventListener('click',e=>{
  const r=e.target.closest&&e.target.closest('.setor-row'); if(!r) return;
  if(r.dataset.emp!==undefined) filtrarEmpresaGrupo(r.dataset.emp); else filtrarSetor(r.dataset.setor);
});
document.addEventListener('keydown',e=>{
  const r=e.target.closest&&e.target.closest('.setor-row');
  if(r&&(e.key==='Enter'||e.key===' ')){e.preventDefault();
    if(r.dataset.emp!==undefined) filtrarEmpresaGrupo(r.dataset.emp); else filtrarSetor(r.dataset.setor);}
});
function setor(s){return s?`<span class="setor-tag setor-corp">${esc(s)}</span>`:'<span style="color:var(--muted)">—</span>';}
function esc(s){const d=document.createElement('div');d.textContent=s||'';return d.innerHTML;}
window._chartBars=[];
function _initChartClick(){
  const canvas=document.getElementById('chart-envios-data');
  if(!canvas || canvas._clickBound) return;
  canvas._clickBound=true;
  canvas.style.cursor='pointer';
  canvas.title='Clique em uma barra para filtrar por aquele dia';
  canvas.addEventListener('click', e=>{
    const rect=canvas.getBoundingClientRect();
    const scaleX=canvas.width/rect.width;
    const x=(e.clientX-rect.left)*scaleX;
    const bar=window._chartBars.find(b=>x>=b.x0 && x<b.x1);
    if(!bar || !bar.iso) return;
    const deAtual=document.getElementById('filtro-data-de').value;
    const ateAtual=document.getElementById('filtro-data-ate').value;
    if(deAtual===bar.iso && ateAtual===bar.iso){
      limparFiltroData();
    } else {
      document.getElementById('filtro-data-de').value=bar.iso;
      document.getElementById('filtro-data-ate').value=bar.iso;
      aplicarFiltroData();
    }
  });
}
function drawBarChart(labels, values, color){
  const canvas=document.getElementById('chart-envios-data');
  const noData=document.getElementById('chart-no-data');
  if(!canvas) return;
  _initChartClick();
  window._chartBars=[];
  if(!labels.length){canvas.style.display='none';noData.style.display='flex';return;}
  noData.style.display='none'; canvas.style.display='block';
  const W=canvas.parentElement.offsetWidth||520, H=240;
  canvas.width=W; canvas.height=H;
  const ctx=canvas.getContext('2d');
  ctx.clearRect(0,0,W,H);
  const pad={top:28,right:16,bottom:52,left:38};
  const cW=W-pad.left-pad.right, cH=H-pad.top-pad.bottom;
  const maxV=Math.max(...values,1);
  const n=labels.length;
  const step=cW/n;
  const barW=Math.max(14, Math.min(52, step*0.6));
  // abrevia: "25/06/2026" → "25/06"
  const shortLabels=labels.map(l=>{const p=l.split('/');return p.length>=2?p[0]+'/'+p[1]:l;});
  const filtroDe=(document.getElementById('filtro-data-de')||{}).value||'';
  const filtroAte=(document.getElementById('filtro-data-ate')||{}).value||'';
  // linhas de grade horizontais
  ctx.strokeStyle=_corTema('--border'); ctx.lineWidth=1;
  for(let i=0;i<=4;i++){
    const y=pad.top+cH*(1-i/4);
    ctx.beginPath(); ctx.moveTo(pad.left,y); ctx.lineTo(W-pad.right,y); ctx.stroke();
    ctx.fillStyle=_corTema('--muted'); ctx.font='12px Inter,sans-serif'; ctx.textAlign='right';
    ctx.fillText(Math.round(maxV/4*i), pad.left-6, y+4);
  }
  // barras + labels
  values.forEach((v,i)=>{
    const cx=pad.left+step*i+step/2;
    const x=cx-barW/2;
    const bH=Math.max(4,(v/maxV)*cH);
    const y=pad.top+cH-bH;
    const isoDia=_parseDataISO(labels[i]);
    const selecionado=isoDia && filtroDe===isoDia && filtroAte===isoDia;
    window._chartBars.push({label:labels[i], iso:isoDia, x0:pad.left+step*i, x1:pad.left+step*(i+1)});
    // barra com gradiente
    const grad=ctx.createLinearGradient(0,y,0,y+bH);
    grad.addColorStop(0,color); grad.addColorStop(1,color+'44');
    ctx.fillStyle=grad;
    ctx.beginPath();
    if(ctx.roundRect) ctx.roundRect(x,y,barW,bH,[4,4,0,0]);
    else ctx.rect(x,y,barW,bH);
    ctx.fill();
    if(selecionado){
      ctx.lineWidth=2; ctx.strokeStyle='#fbbf24';
      ctx.beginPath();
      if(ctx.roundRect) ctx.roundRect(x-1,y-1,barW+2,bH+2,[4,4,0,0]);
      else ctx.rect(x-1,y-1,barW+2,bH+2);
      ctx.stroke();
    }
    // valor no topo da barra
    ctx.fillStyle=selecionado?'#fbbf24':'#f1f5f9'; ctx.font='bold 13px Inter,sans-serif'; ctx.textAlign='center';
    ctx.fillText(v, cx, y-7);
    // data horizontal embaixo
    ctx.fillStyle=selecionado?'#fbbf24':'#94a3b8'; ctx.font='bold 12px Inter,sans-serif'; ctx.textAlign='center';
    ctx.fillText(shortLabels[i], cx, H-12);
    // linha separadora embaixo da barra
    ctx.strokeStyle=color+'33'; ctx.lineWidth=1;
    ctx.beginPath(); ctx.moveTo(cx,pad.top+cH); ctx.lineTo(cx,pad.top+cH+6); ctx.stroke();
  });
}
function pct_bar(pct){const color=pct===100?'#10b981':pct>0?'#f59e0b':'#ef4444';return `<div class="pct-inline"><div class="pct-track"><div class="pct-fill" style="width:${pct}%;background:${color}"></div></div><span class="pct-num">${pct}%</span></div>`;}
function status_badge(s){return s.includes('consultor')?'<span class="status-badge status-parcial">Ag. consultor</span>':'<span class="status-badge status-aguardando">Ag. gestor</span>';}

function showTab(id,btn){
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('tab-'+id).classList.add('active');
  btn.classList.add('active');
  try{ btn.scrollIntoView({block:'nearest',inline:'center'}); sessionStorage.setItem('painel_aba',id); }catch(e){}
}
function filterEmpresa(emp){
  ['todas','al','se','rec'].forEach(e=>{
    const btn=document.getElementById('filtro-'+e);
    if(btn) btn.className='toggle-btn'+(e==='al'?' al':e==='se'?' se':e==='rec'?' re':'')+(e===emp?' active':'');
  });
  const d={al:EMP('alagoas'),se:EMP('sergipe'),rec:EMP('recife')};
  let total,env,nenv,atv,pct;
  if(emp==='todas'){
    total=d.al.total+d.se.total+d.rec.total;
    env  =d.al.enviaram+d.se.enviaram+d.rec.enviaram;
    nenv =d.al.nao_enviaram+d.se.nao_enviaram+d.rec.nao_enviaram;
    atv  =d.al.total_atividades+d.se.total_atividades+d.rec.total_atividades;
    pct  =(env/Math.max(total,1)*100).toFixed(1).replace('.',',')+'%';
  } else {
    const r=d[emp];
    total=r.total; env=r.enviaram; nenv=r.nao_enviaram; atv=r.total_atividades;
    pct=(env/Math.max(total,1)*100).toFixed(1).replace('.',',')+'%';
  }
  document.getElementById('val-cadastrados').textContent=total;
  document.getElementById('val-concluidos').textContent=env;
  document.getElementById('val-pendentes').textContent=nenv;
  document.getElementById('val-atividades').textContent=atv;
  document.getElementById('val-adesao').textContent=pct;
  // Detalhes por filial nos cartões: ocultar quando filtro único
  document.querySelectorAll('.card-detalhe').forEach(el=>{el.style.display=emp==='todas'?'':'none';});
  // Blocos de empresa
  ['al','se','rec'].forEach(e=>{
    const el=document.getElementById('bloco-'+e);
    if(el) el.style.display=(emp==='todas'||emp===e)?'':'none';
  });
  // Tabela de enviaram — re-renderiza via filtrarEnviaram para contar corretamente
  window._vgEmpresa=emp;
  filtrarEnviaram(window._vgFiltro);
  // Alerta crítico dinâmico
  const alertBox=document.getElementById('alerta-criticos');
  if(alertBox){
    document.querySelectorAll('.alerta-linha').forEach(el=>{
      el.style.display=(emp==='todas'||el.dataset.filial===emp)?'':'none';
    });
    const algumVisivel=[...document.querySelectorAll('.alerta-linha')].some(el=>el.style.display!=='none');
    alertBox.style.display=algumVisivel?'':'none';
  }
  // Cartão Novas Atividades dinâmico
  const cardNovas=document.getElementById('card-novas');
  const valNovas=document.getElementById('val-novas');
  if(cardNovas && valNovas){
    if(emp==='todas'){
      valNovas.textContent=cardNovas.dataset.total;
      ['se','al','rec'].forEach(e=>{const el=document.getElementById('det-novas-'+e);if(el)el.style.display='';});
    } else {
      valNovas.textContent=cardNovas.dataset[emp];
      ['se','al','rec'].forEach(e=>{const el=document.getElementById('det-novas-'+e);if(el)el.style.display=e===emp?'':'none';});
    }
  }
  // Gráfico de envios por data
  const grid=document.getElementById('empresa-grid');
  const chartBloco=document.getElementById('bloco-chart');
  const colorMap={al:'#2563eb',se:'#7c3aed',rec:'#f97316'};
  const empKey={al:'alagoas',se:'sergipe',rec:'recife'};
  const nomeFilial={al:'Novatec Alagoas',se:'GRUPO_NOME_PH',rec:'Novatec Recife'};
  if(emp==='todas'){
    if(grid) grid.style.gridTemplateColumns='repeat(3,1fr)';
    if(chartBloco) chartBloco.style.display='none';
  } else {
    if(grid) grid.style.gridTemplateColumns='1fr 2fr';
    if(chartBloco){ chartBloco.style.display='flex'; chartBloco.style.flexDirection='column'; }
    const empData=EMP(empKey[emp]);
    const datas=empData.envios_por_data||{};
    const labels=Object.keys(datas);
    const values=Object.values(datas);
    document.getElementById('chart-titulo').textContent='Envios por Data — '+nomeFilial[emp];
    requestAnimationFrame(()=>requestAnimationFrame(()=>drawBarChart(labels,values,colorMap[emp])));
  }
}
function cargoToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('cargo-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('cargo-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function naoToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('nao-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('nao-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function envToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('env-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('env-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function exportarEnviaram(){
  const toggles={al:'env-toggle-al',se:'env-toggle-se',rec:'env-toggle-rec'};
  let empAtiva='todas';
  for(const [e,id] of Object.entries(toggles)){
    const btn=document.getElementById(id);
    if(btn && btn.classList.contains('active')){empAtiva=e;break;}
  }
  const fontes={
    al:{filial:'Novatec Alagoas',lista:EMP('alagoas').enviaram_list},
    se:{filial:'GRUPO_NOME_PH',lista:EMP('sergipe').enviaram_list},
    rec:{filial:'Novatec Recife', lista:EMP('recife').enviaram_list},
  };
  const selecionadas=empAtiva==='todas'?Object.values(fontes):[fontes[empAtiva]];
  const linhas=[];
  selecionadas.forEach(({filial,lista})=>{
    lista.forEach(p=>{linhas.push([p.empresa||filial,p.nome,p.cargo,p.setor,p.perfil,p.atividades]);});
  });
  _exportarTabela([['Empresa',22],['Nome',30],['Cargo',24],['Setor',22],['Perfil',14],['Atividades',11]],linhas,
    'Enviaram'+_sufixoSetores()+'_'+_hojeArq(),null,'Enviaram');
  return;
  const csv=linhas.map(r=>r.map(c=>'"'+String(c||'').replace(/"/g,'""')+'"').join(';')).join('\\r\\n');
  const hoje=new Date().toLocaleDateString('pt-BR').replace(/\//g,'');
  const sufixo=empAtiva==='todas'?'Todas':fontes[empAtiva].filial.replace(/\\s+/g,'_');
  const blob=new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='Enviaram_'+sufixo+'_'+hoje+'.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
// ── Exportar atividades no formato do modelo do mapeia ──
const _XL_COLS=[['#',5],['Empresa',16],['Setor',18],['Funcionário',20],['Cargo',20],['Nome',30],['Detalhamento',55],
  ['Periodicidade',16],['Como realiza',16],['Tempo',9],['Horário fixo',12],['Pontos críticos',34],['Sugestões de melhoria',34],
  ['Status',14],['Decisão gestor',14],['Decisão consultor',14],['Edições do consultor',12],['Parecer do consultor',40],
  ['Criada em',20],['Atualizada em',20]];
function _xlData(v){
  if(v===null||v===undefined||v==='') return '';
  const s=String(v);
  if(s.length>=19&&s[4]==='-'&&s[10]==='T'){ const d=new Date(s); if(!isNaN(d)) return d.toLocaleString('pt-BR',{timeZone:'America/Sao_Paulo'}); }
  return s;
}
function _xlSN(v){ if(v===true) return 'Sim'; if(v===false) return 'Não'; return v===null||v===undefined?'':v; }
function _xlLinha(a,i,extra){
  extra=extra||{};
  const st=String(a.status||'');
  const dg=a.decisao_gestor||(st.toLowerCase().indexOf('aguardando gestor')===0?'Aguardando':'');
  const ed=(a.edicoes===''||a.edicoes===null||a.edicoes===undefined)?0:(isNaN(Number(a.edicoes))?a.edicoes:Number(a.edicoes));
  return [i+1,a.empresa||'GRUPO_NOME_PH',a.setor||extra.setor||'',a.funcionario||extra.nome||'',a.cargo||extra.cargo||'',
    a.nome||'',a.detalhamento||'',a.periodicidade||'',a.como_realiza||'',a.tempo||'',_xlSN(a.horario_fixo),
    a.pontos_criticos||'',a.sugestoes||'',st,dg,a.decisao_consultor||'—',ed,a.parecer||'',
    _xlData(a.criada_em),_xlData(a.atualizada_em||a.criada_em)];
}
function _carregarExcelJS(){
  if(window.ExcelJS) return Promise.resolve(window.ExcelJS);
  return new Promise((ok,erro)=>{
    const sc=document.createElement('script');
    sc.src='https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
    sc.onload=()=>ok(window.ExcelJS); sc.onerror=erro; document.head.appendChild(sc);
  });
}
function _baixarArquivo(blob,nome){
  const u=URL.createObjectURL(blob), l=document.createElement('a'); l.href=u; l.download=nome;
  document.body.appendChild(l); l.click(); l.remove(); setTimeout(()=>URL.revokeObjectURL(u),3000);
}
function _exportarModelo(linhas,base,btn){
  const txt=btn?btn.innerHTML:''; if(btn){ btn.disabled=true; btn.textContent='Gerando…'; }
  return _carregarExcelJS().then(EJ=>{
    const wb=new EJ.Workbook();
    const ws=wb.addWorksheet('Atividades',{views:[{state:'frozen',ySplit:1}]});
    ws.columns=_XL_COLS.map(c=>({header:c[0],width:c[1]}));
    linhas.forEach(l=>ws.addRow(l));
    const fonte={name:'Calibri',size:10,color:{argb:'FF1F2937'}};
    const linhaBorda={bottom:{style:'thin',color:{argb:'FFE5E7EB'}}};
    ws.eachRow({includeEmpty:true},(row,n)=>{
      row.eachCell({includeEmpty:true},(c,col)=>{
        if(n===1){
          c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFD9222A'}};
          c.font={name:'Calibri',size:10,bold:true,color:{argb:'FFFFFFFF'}};
          c.alignment={vertical:'middle',horizontal:'left',wrapText:true};
          return;
        }
        c.font=fonte; c.border=linhaBorda;
        c.alignment={vertical:'middle',wrapText:true,horizontal:(col===1||col===10||col===17)?'right':'left'};
        if(n%2===1) c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF5F6F8'}};
        if(col===15&&String(c.value||'').toLowerCase().indexOf('aguard')===0)
          c.font={name:'Calibri',size:10,bold:true,color:{argb:'FFD97706'}};
      });
    });
    ws.getRow(1).height=26;
    ws.autoFilter={from:{row:1,column:1},to:{row:1,column:_XL_COLS.length}};
    return wb.xlsx.writeBuffer();
  }).then(buf=>{
    _baixarArquivo(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),base+'.xlsx');
  }).catch(()=>{
    const cab=_XL_COLS.map(c=>c[0]);
    const csv=[cab,...linhas].map(r=>r.map(c=>'"'+String(c===null||c===undefined?'':c).split('"').join('""')+'"').join(';')).join(String.fromCharCode(13,10));
    _baixarArquivo(new Blob([String.fromCharCode(65279)+csv],{type:'text/csv;charset=utf-8'}),base+'.csv');
  }).finally(()=>{ if(btn){ btn.disabled=false; btn.innerHTML=txt; } });
}
function _hojeArq(){ return new Date().toLocaleDateString('pt-BR').split('/').join(''); }
function exportarAtividades(ev){
  const lista=EMP('sergipe').atividades||[];
  const btn=ev&&ev.currentTarget?ev.currentTarget:null;
  _exportarModelo(lista.map((a,i)=>_xlLinha(a,i)),'Atividades_GRUPO_ARQ_PH_'+_hojeArq(),btn);
}
function exportarAtividadesPessoa(){
  const at=window._vaAtual; if(!at||!at.lista.length) return;
  const pessoa=[...(EMP('sergipe').enviaram_list||[]),...(EMP('sergipe').nao_enviaram_list||[])].find(p=>_normJS(p.nome)===_normJS(at.nome))||{};
  const extra={nome:at.nome,cargo:pessoa.cargo,setor:pessoa.setor};
  _exportarModelo(at.lista.map((a,i)=>_xlLinha(a,i,extra)),'Atividades_'+at.nome.trim().split(' ').join('_')+'_'+_hojeArq(),document.getElementById('va-exportar'));
}
function exportarAnalise(){
  const fontes=[
    {filial:'GRUPO_NOME_PH', d:ANALISE('sergipe')},
  ];
  const linhas=[['Filial','Categoria','Item','Quantidade','Percentual']];
  fontes.forEach(({filial,d})=>{
    if(!d) return;
    const total=d.total||0;
    const pct=q=>total?Math.round(q/total*100)+'%':'';
    const addBreak=(cat,obj)=>{
      Object.entries(obj||{}).forEach(([item,qtd])=>{linhas.push([filial,cat,item,qtd,pct(qtd)]);});
    };
    addBreak('Periodicidade', d.periodicidade);
    addBreak('Como realiza', d.como_realiza);
    addBreak('Status', d.status);
    linhas.push([filial,'Pontos Críticos','Preenchidos',d.pontos_criticos_count,pct(d.pontos_criticos_count)]);
    linhas.push([filial,'Sugestões','Preenchidas',d.sugestoes_count,pct(d.sugestoes_count)]);
    addBreak('Criadas por mês', d.por_mes);
  });
  const csv=linhas.map(r=>r.map(c=>'"'+String(c||'').replace(/"/g,'""')+'"').join(';')).join('\\r\\n');
  const hoje=new Date().toLocaleDateString('pt-BR').replace(/\//g,'');
  const blob=new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='Analise_Atividades_'+hoje+'.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
function atvToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('atv-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('atv-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function verToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('ver-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('ver-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function resumoCard(label,val,cor){
  return `<div style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:8px 10px">
    <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px">${label}</div>
    <div style="font-size:20px;font-weight:700;color:${cor||'var(--white)'}">${val}</div>
  </div>`;
}
function renderResumo(prefix){
  if(!DATA.resumo) return;
  prefix = prefix || 'resumo';
  const r={al:DATA.resumo.alagoas, se:DATA.resumo.sergipe, rec:DATA.resumo.recife, total:DATA.resumo.total};
  ['se','al','rec','total'].forEach(e=>{
    const d=r[e];
    const el=document.getElementById(`${prefix}-${e}-cards`);
    if(!el) return;
    el.innerHTML=
      resumoCard('Aprovadas',d.aprovadas,'#10b981')+
      resumoCard('Ressalvas',d.ressalvas,'#f59e0b')+
      resumoCard('Ag. Gestor',d.ag_gestor,'#f59e0b')+
      resumoCard('Ag. Consultor',d.ag_consultor,'#60a5fa')+
      resumoCard('Editadas',d.editadas,'var(--white)')+
      resumoCard('Com Parecer',d.com_parecer,'var(--white)')+
      `<div style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:8px 10px;grid-column:span 3">
        <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px">Tempo Total Mapeado</div>
        <div style="font-size:20px;font-weight:700;color:#a78bfa">${d.tempo_total}h</div>
      </div>`;
  });
}
function breakdownBlock(title, obj, total, cor){
  cor = cor || '#334155';
  const tarja = `<div style="font-size:11px;font-weight:700;text-transform:uppercase;color:#fff;letter-spacing:.05em;background:${cor}33;border-left:3px solid ${cor};padding:5px 10px;border-radius:5px;margin-bottom:8px">${title}</div>`;
  const entries=Object.entries(obj||{});
  if(!entries.length) return `<div style="margin-bottom:14px">${tarja}<div style="font-size:12px;color:var(--muted)">Sem registros</div></div>`;
  const rows=entries.map(([k,v])=>{
    const pct=total?Math.round(v/total*100):0;
    return `<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:12px;padding:4px 0;border-bottom:1px solid var(--border)">
      <span style="color:var(--subtle)">${esc(k)}</span>
      <span style="display:flex;align-items:center;gap:6px"><span class="count-pill">${v}</span><span style="color:var(--muted);font-size:11px;min-width:32px;text-align:right">${pct}%</span></span>
    </div>`;
  }).join('');
  return `<div style="margin-bottom:14px">
    ${tarja}
    ${rows}
  </div>`;
}
function renderAnalise(){
  if(!DATA.analise) return;
  const a={al:ANALISE('alagoas'), se:ANALISE('sergipe'), rec:ANALISE('recife'), total:ANALISE('total')};
  const nomes={al:'Novatec Alagoas', se:'GRUPO_NOME_PH', rec:'Novatec Recife', total:'Todas as Empresas'};
  const corBarra={al:'#2563eb', se:'#7c3aed', rec:'#f97316', total:'#14b8a6'};
  const corTexto={al:'#60a5fa', se:'#a78bfa', rec:'#fdba74', total:'#2dd4bf'};
  ['se','al','rec','total'].forEach(e=>{
    const d=a[e];
    const el=document.getElementById(`anl-col-${e}`);
    if(!el||!d) return;
    const totalAtv=d.total||0;
    const mesesRows=Object.entries(d.por_mes||{}).map(([mes,q])=>
      `<div style="display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:12px;padding:4px 0;border-bottom:1px solid var(--border)">
        <span style="color:var(--subtle)">${mes}</span><span class="count-pill">${q}</span>
      </div>`).join('') || '<div style="font-size:12px;color:var(--muted)">Sem registros</div>';
    el.innerHTML=`
      <div style="background:var(--surface);border:1px solid var(--border);border-top:3px solid ${corBarra[e]};border-radius:12px;padding:16px;height:100%">
        <div style="font-size:13px;font-weight:700;color:${corTexto[e]};margin-bottom:10px">${nomes[e]} <span class="count-pill" style="margin-left:6px">${totalAtv} atividades</span></div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px">
          <div style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:8px 10px">
            <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px">Pontos Críticos</div>
            <div style="font-size:18px;font-weight:700;color:var(--white)">${d.pontos_criticos_count} <span style="font-size:11px;color:var(--muted);font-weight:500">/ ${totalAtv}</span></div>
          </div>
          <div style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:8px 10px">
            <div style="font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:3px">Sugestões</div>
            <div style="font-size:18px;font-weight:700;color:var(--white)">${d.sugestoes_count} <span style="font-size:11px;color:var(--muted);font-weight:500">/ ${totalAtv}</span></div>
          </div>
        </div>
        ${breakdownBlock('Periodicidade', d.periodicidade, totalAtv, corBarra[e])}
        ${breakdownBlock('Como realiza', d.como_realiza, totalAtv, corBarra[e])}
        ${breakdownBlock('Status', d.status, totalAtv, corBarra[e])}
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;color:#fff;letter-spacing:.05em;background:${corBarra[e]}33;border-left:3px solid ${corBarra[e]};padding:5px 10px;border-radius:5px;margin-bottom:8px">Criadas por mês</div>
        ${mesesRows}
      </div>`;
  });
}
function renderVersoes(){
  if(!DATA.versoes) return;
  const _vmap={al:'alagoas',se:'sergipe',rec:'recife'};
  ['al','se','rec'].forEach(e=>{
    const tb=document.getElementById(`ver-${e}-body`);
    const itens=DATA.versoes[_vmap[e]];
    if(!itens||itens.length===0){
      tb.innerHTML='<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:20px">Nenhum registro de versão ou parecer</td></tr>';
      return;
    }
    tb.innerHTML=itens.map(v=>`<tr>
      <td><strong style="font-size:12px">${v.funcionario}</strong></td>
      <td style="font-size:12px">${v.atividade}</td>
      <td><span class="status-badge status-aguardando" style="font-size:10px">${v.tipo}</span></td>
      <td style="font-size:12px;color:var(--subtle)">${v.autor}</td>
      <td style="font-size:11px;color:var(--muted);font-family:'DM Mono',monospace">${v.data}</td>
      <td style="font-size:12px;color:var(--text);max-width:300px">${v.conteudo}</td>
    </tr>`).join('');
  });
}
function aprToggle(emp){
  ['al','se','rec'].forEach(e=>{document.getElementById('apr-'+e).style.display=emp===e?'':'none';});
  ['al','se','rec'].forEach(e=>{document.getElementById('apr-toggle-'+e).className='toggle-btn '+e+(emp===e?' active':'');});
}
function apr_status_badge(s){
  if(s==='Aprovado') return '<span class="status-badge" style="background:rgba(16,185,129,.15);color:#34d399;border:1px solid rgba(16,185,129,.3)">Aprovado</span>';
  if(s==='Reprovado') return '<span class="status-badge" style="background:rgba(239,68,68,.15);color:#f87171;border:1px solid rgba(239,68,68,.3)">Reprovado</span>';
  return '<span class="status-badge status-aguardando">Pendente</span>';
}
function exportarPendentes(){
  // Detecta qual filial está ativa
  const toggles={al:'nao-toggle-al',se:'nao-toggle-se',rec:'nao-toggle-rec'};
  let empAtiva='todas';
  for(const [e,id] of Object.entries(toggles)){
    const btn=document.getElementById(id);
    if(btn && btn.classList.contains('active')){empAtiva=e;break;}
  }
  const fontes={
    al:{filial:'Novatec Alagoas', lista:EMP('alagoas').nao_enviaram_list},
    se:{filial:'GRUPO_NOME_PH', lista:EMP('sergipe').nao_enviaram_list},
    rec:{filial:'Novatec Recife',  lista:EMP('recife').nao_enviaram_list},
  };
  const selecionadas = empAtiva==='todas' ? Object.values(fontes) : [fontes[empAtiva]];
  const linhas=[];
  selecionadas.forEach(({filial,lista})=>{
    lista.forEach(p=>{linhas.push([p.empresa||filial, p.nome, p.cargo, p.setor, p.perfil, p.status_usr]);});
  });
  _exportarTabela([['Empresa',22],['Nome',30],['Cargo',24],['Setor',22],['Perfil',14],['Status',10]],linhas,
    'Pendentes'+_sufixoSetores()+'_'+_hojeArq(),null,'Pendentes');
  return;
  const csv=linhas.map(r=>r.map(c=>'"'+String(c||'').replace(/"/g,'""')+'"').join(';')).join('\\r\\n');
  const hoje=new Date().toLocaleDateString('pt-BR').replace(/\//g,'');
  const sufixo=empAtiva==='todas'?'Todas':fontes[empAtiva].filial.replace(/\s+/g,'_');
  const blob=new Blob(['﻿'+csv],{type:'text/csv;charset=utf-8'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url; a.download='Pendentes_'+sufixo+'_'+hoje+'.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}
function filterTable(tbodyId,q){
  document.getElementById(tbodyId).querySelectorAll('tr').forEach(r=>{
    r.style.display=r.textContent.toLowerCase().includes(q.toLowerCase())?'':'none';
  });
}

// dentro do portal (iframe) quem busca a versao nova e o portal; sozinho, recarrega a pagina
function _recarregarPagina(){ try{ if(window.parent!==window && window.parent.recarregarPainel){ window.parent.recarregarPainel(); return; } }catch(e){} location.reload(); }
// ── cabecalho: saudacao, relogio e recarregar ──
const _MESES=['jan.','fev.','mar.','abr.','mai.','jun.','jul.','ago.','set.','out.','nov.','dez.'];
const _DIAS=['dom.','seg.','ter.','qua.','qui.','sex.','sáb.'];
function _nomeSalvo(){ try{ return localStorage.getItem('painel_nome')||''; }catch(e){ return ''; } }
function _iniciais(n){ const p=n.trim().split(' ').filter(Boolean); return p.length?(p[0][0]+(p.length>1?p[p.length-1][0]:'')).toUpperCase():'?'; }
function _atualizarSaudacao(){
  const agora=new Date(), h=agora.getHours();
  const ola=h<12?'Bom dia,':h<18?'Boa tarde,':'Boa noite,';
  const nome=_nomeSalvo();
  const el=id=>document.getElementById(id);
  if(!el('sd-ola')) return;
  el('sd-ola').textContent=ola;
  el('sd-nome').textContent=nome||'Toque para se identificar';
  el('sd-nome').classList.toggle('vazio',!nome);
  el('sd-avatar').textContent=nome?_iniciais(nome):'?';
  el('sd-data').textContent=_DIAS[agora.getDay()]+', '+agora.getDate()+' de '+_MESES[agora.getMonth()];
  el('sd-hora').textContent=String(h).padStart(2,'0')+':'+String(agora.getMinutes()).padStart(2,'0');
  const vg=el('vg-hoje'); if(vg) vg.textContent=agora.toLocaleDateString('pt-BR');
}
function _pedirNome(){
  const atual=_nomeSalvo();
  const n=window.prompt('Como você quer ser chamado neste painel?', atual);
  if(n===null) return;
  try{ n.trim()?localStorage.setItem('painel_nome',n.trim().slice(0,40)):localStorage.removeItem('painel_nome'); }catch(e){}
  _atualizarSaudacao();
}
document.addEventListener('click',e=>{
  const t=e.target.closest&&e.target.closest('#sd-nome,#sd-avatar,#sd-recarregar,#filtros-btn'); if(!t) return;
  if(t.id==='sd-recarregar'){ t.classList.add('girando'); setTimeout(_recarregarPagina,350); return; }
  if(t.id==='filtros-btn'){ const p=document.getElementById('filtros-painel'); const ab=!p.classList.contains('aberto');
    p.classList.toggle('aberto',ab); t.setAttribute('aria-expanded',ab); t.classList.toggle('ativo',ab); return; }
  _pedirNome();
});
_atualizarSaudacao(); setInterval(_atualizarSaudacao, 15*1000);
// abas fixas logo abaixo do cabecalho (que tambem e fixo no computador)
function _ajustarAbas(){ const h=document.querySelector('header'); if(!h) return;
  const fixo=getComputedStyle(h).position==='sticky';
  document.documentElement.style.setProperty('--hdr-h',(fixo?h.offsetHeight:0)+'px'); }
window.addEventListener('resize',_ajustarAbas); window.addEventListener('load',_ajustarAbas); setTimeout(_ajustarAbas,300);
// ── cronometro da proxima atualizacao (o robo roda a cada INTERVALO min, no inicio da hora) ──
const _INTERVALO_MIN=INTERVALO_MIN_PH, _ATRASO_MIN=2;   // ~2 min para buscar, gerar e gravar
// no portal, o admin escolhe o intervalo e pode pausar (aba Empresas)
function _confAtualizacao(){
  try{ if(window.parent!==window && typeof window.parent.atualizacaoDoPainel==='function'){
    const c=window.parent.atualizacaoDoPainel(); if(c) return c; } }catch(e){}
  return {ativa:true, intervalo:_INTERVALO_MIN};
}
function _proximaAtualizacao(){
  // janelas contadas a partir da meia-noite (igual ao robo), mais o tempo de processamento
  const iv=Math.max(10,_confAtualizacao().intervalo||60);
  const agora=new Date(), b=new Date(agora); b.setHours(0,_ATRASO_MIN,0,0);
  while(b<=agora) b.setMinutes(b.getMinutes()+iv);
  return b;
}
let _buscaAgendada=false;
function _tickProxima(){
  const el=document.getElementById('prox-atual'); if(!el) return;
  if(!_confAtualizacao().ativa){ el.innerHTML='<span class="prox-sep"> · </span><b>atualização automática pausada</b>'; return; }
  const ms=_proximaAtualizacao()-new Date();
  const h=Math.floor(ms/3600000), m=Math.floor(ms%3600000/60000), sg=Math.floor(ms%60000/1000);
  const t=(h?h+':'+String(m).padStart(2,'0'):String(m).padStart(2,'0'))+':'+String(sg).padStart(2,'0');
  el.innerHTML='<span class="prox-sep"> · </span><span class="prox-lbl">próxima atualização em </span><span class="prox-lbl-c">próx. </span><b>'+t+'</b>';
  // chegou a hora: da 1 min de folga e busca a versao nova (no portal so baixa se mudou)
  if(ms<1500 && !_buscaAgendada){ _buscaAgendada=true; setTimeout(()=>{ _buscaAgendada=false; _recarregarPagina(); },60*1000); }
}
_tickProxima(); setInterval(_tickProxima,1000);
function _corTema(v){return getComputedStyle(document.documentElement).getPropertyValue(v).trim()||'#888';}
const _TEMAS=[
  {id:'',nome:'Escuro',cores:['#0d1117','#161b22','#7c3aed']},
  {id:'marinho',nome:'Azul-marinho',cores:['#0a1324','#0f1d36','#2563eb']},
  {id:'grafite',nome:'Grafite',cores:['#121212','#242427','#a1a1aa']},
  {id:'millena',nome:'GRUPO_NOME_PH',cores:['#0a1836','#163066','#f5c400']},
  {id:'claro',nome:'Claro',cores:['#f3f5f9','#ffffff','#2563eb']},
  {id:'claro-millena',nome:'Claro GRUPO_NOME_PH',cores:['#fff8dc','#ffffff','#1d4fb8']},
];
function _temaAtual(){return document.documentElement.getAttribute('data-tema')||'';}
function aplicarTema(id){
  if(id) document.documentElement.setAttribute('data-tema',id); else document.documentElement.removeAttribute('data-tema');
  try{ id?localStorage.setItem('painel_tema',id):localStorage.removeItem('painel_tema'); }catch(e){}
  _montarTemas();
  try{ const m=document.querySelector('meta[name="theme-color"]'); if(m) m.setAttribute('content',_corTema('--surface')); }catch(e){}
  try{ atualizarDonuts(); filterEmpresa(_empresaAtiva()); }catch(e){}
}
function _montarTemas(){
  const box=document.getElementById('tema-lista'); if(!box) return;
  const at=_temaAtual();
  box.innerHTML=_TEMAS.map(t=>`<button type="button" role="menuitemradio" aria-checked="${t.id===at}" class="tema-op${t.id===at?' ativo':''}" data-tema-id="${t.id}">
    <span class="tema-amostra">${t.cores.map(c=>`<i style="background:${c}"></i>`).join('')}</span>${t.nome}</button>`).join('');
}
document.addEventListener('click',e=>{
  const pop=document.getElementById('tema-pop'), btn=document.getElementById('tema-btn'); if(!pop) return;
  const op=e.target.closest&&e.target.closest('.tema-op');
  if(op){aplicarTema(op.dataset.temaId); return;}
  if(e.target.closest&&e.target.closest('#tema-btn')){const ab=!pop.classList.contains('aberto'); pop.classList.toggle('aberto',ab); btn.setAttribute('aria-expanded',ab); if(ab)_montarTemas(); return;}
  if(!(e.target.closest&&e.target.closest('#tema-pop'))){pop.classList.remove('aberto'); btn.setAttribute('aria-expanded','false');}
});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){const pop=document.getElementById('tema-pop'); if(pop) pop.classList.remove('aberto');}});
window._donutUltimo={};
function drawDonut(id,sent,total,colorSent){
  const c=document.getElementById(id);if(!c)return;
  // tudo proporcional ao tamanho do canvas (o da empresa e grande e nitido)
  const W=c.width,ctx=c.getContext('2d'),cx=W/2,cy=c.height/2,r=W*0.40,lw=W*0.10,TAU=Math.PI*2,pct=sent/Math.max(total,1);
  // anima so a partir do valor anterior (antes recomecava do zero a cada 10 s)
  const de=(id in window._donutUltimo)?window._donutUltimo[id]:0; window._donutUltimo[id]=pct;
  const duration=de===pct?0:1200,start=performance.now();
  function ease(t){return t<.5?2*t*t:1-Math.pow(-2*t+2,2)/2;}
  function frame(now){
    const t=duration?Math.min((now-start)/duration,1):1,current=de+(pct-de)*ease(t);
    ctx.clearRect(0,0,c.width,c.height);
    ctx.beginPath();ctx.arc(cx,cy,r,0,TAU);ctx.lineWidth=lw;ctx.strokeStyle=_corTema('--surface2');ctx.stroke();
    if(current>0){ctx.beginPath();ctx.arc(cx,cy,r,-Math.PI/2,-Math.PI/2+TAU*current);ctx.lineWidth=lw;ctx.strokeStyle=colorSent;ctx.lineCap='round';ctx.stroke();}
    ctx.fillStyle=_corTema('--white');ctx.font='bold '+Math.round(W*0.19)+'px Inter,sans-serif';ctx.textAlign='center';ctx.textBaseline='middle';
    ctx.fillText(Math.round(current*100)+'%',cx,cy);
    if(t<1)requestAnimationFrame(frame);
  }
  requestAnimationFrame(frame);
}
const _DONUTS=[
  {id:'donut-al', key:'alagoas'},
  {id:'donut-se', key:'sergipe'},
  {id:'donut-rec',key:'recife'},
];
function atualizarDonuts(){
  _DONUTS.forEach(d=>drawDonut(d.id,EMP(d.key).enviaram,EMP(d.key).total,'#10b981'));
}
function startDonutLoop(){
  atualizarDonuts();
  setInterval(atualizarDonuts,10000);
}
function setHtml(id, rows){document.getElementById(id).innerHTML=rows.join('');}
function _dataHora(raw){
  const iso=_parseDataISO(raw), h=_extrairHorario(raw);
  if(!iso) return h?h.texto:'';
  const [y,m,d]=iso.split('-');
  return d+'/'+m+'/'+y+(h?' às '+h.texto:'');
}
function _rotularTabelas(){
  document.querySelectorAll('.table-wrap table').forEach(t=>{
    const ths=[...t.querySelectorAll('thead th')].map(th=>th.textContent.trim());
    if(!ths.length) return;
    t.querySelectorAll('tbody tr').forEach(tr=>{ [...tr.children].forEach((td,i)=>{ if(!td.hasAttribute('data-rotulo')) td.setAttribute('data-rotulo', ths[i]||''); }); });
  });
}
(function(){ let fila=null; new MutationObserver(()=>{ clearTimeout(fila); fila=setTimeout(_rotularTabelas,50); })
  .observe(document.documentElement,{childList:true,subtree:true}); })();
function verAtividades(nome){
  const n=_normJS(nome);
  const lista=(EMP('sergipe').atividades||[]).filter(a=>_normJS(a.funcionario)===n);
  window._vaAtual={nome,lista};
  const pessoa=[...(EMP('sergipe').enviaram_list||[]),...(EMP('sergipe').nao_enviaram_list||[])].find(p=>_normJS(p.nome)===n)||{};
  const min=lista.reduce((t,a)=>{const m=String(a.tempo||'').match(/^(\d+):(\d+)/);return t+(m?(+m[1])*60+(+m[2]):0);},0);
  const fmtT=m=>`${Math.floor(m/60)}h${String(m%60).padStart(2,'0')}`;
  document.getElementById('va-nome').textContent=nome;
  document.getElementById('va-sub').textContent=[pessoa.cargo,pessoa.setor,`${lista.length} atividade${lista.length===1?'':'s'}`,min?`${fmtT(min)} mapeadas`:''].filter(Boolean).join(' · ');
  const campo=(r,v)=>v&&String(v).trim()?`<div class="va-campo"><span>${r}</span>${esc(v)}</div>`:'';
  document.getElementById('va-lista').innerHTML=lista.length?lista.map((a,i)=>{return `<div class="va-item">
      <div class="va-topo"><strong>${i+1}. ${esc(a.nome)}</strong>${status_badge(a.status||'')}</div>
      <div class="va-dec"><span>Gestor</span>${_badgeDec(_tipoDecisao(a.decisao_gestor))}<span>Consultor</span>${_badgeDec(_tipoDecisao(a.decisao_consultor))}</div>
      <div class="va-meta">${[a.periodicidade,a.como_realiza,a.tempo?('⏱ '+a.tempo):'',_dataHora(a.criada_em)?('enviada em '+_dataHora(a.criada_em)):''].filter(Boolean).map(esc).join(' · ')}</div>
      ${a.detalhamento?`<div class="va-det">${esc(a.detalhamento)}</div>`:''}
      ${campo('Pontos críticos',a.pontos_criticos)}${campo('Sugestões de melhoria',a.sugestoes)}
    </div>`;}).join(''):'<div class="va-vazio">Nenhuma atividade no filtro atual.</div>';
  const m=document.getElementById('modal-atividades'); m.style.display='flex'; document.body.style.overflow='hidden';
  m.querySelector('.va-fechar').focus();
}
function _carregarXLSX(){
  if(window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((ok,erro)=>{
    const sc=document.createElement('script');
    sc.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    sc.onload=()=>ok(window.XLSX); sc.onerror=erro; document.head.appendChild(sc);
  });
}
function fecharAtividades(){document.getElementById('modal-atividades').style.display='none';document.body.style.overflow='';}
document.addEventListener('click',e=>{
  const b=e.target.closest&&e.target.closest('.btn-ver'); if(b){verAtividades(b.dataset.func);return;}
  if(e.target.id==='modal-atividades') fecharAtividades();
});
document.addEventListener('keydown',e=>{if(e.key==='Escape') fecharAtividades();});
// ── Decisao do gestor e do consultor (como no mapeia) ──
function _tipoDecisao(v){
  const t=String(v||'').trim().toLowerCase();
  if(t.startsWith('aprov')) return 'aprovada';
  if(t.startsWith('reprov')||t.startsWith('recus')) return 'reprovada';
  if(t.startsWith('ressalv')) return 'ressalva';
  return 'aguardando';   // vazio, "—" ou "Aguardando"
}
const _DEC_ROT={aprovada:'Aprovada',aguardando:'Aguardando',reprovada:'Reprovada',ressalva:'Ressalva'};
const _DEC_ICO={aprovada:'✓',aguardando:'◌',reprovada:'✕',ressalva:'!'};
function _badgeDec(tipo,texto,titulo){
  return `<span class="dec-badge dec-${tipo}"${titulo?` title="${esc(titulo)}"`:''}>${_DEC_ICO[tipo]||''} ${esc(texto||_DEC_ROT[tipo])}</span>`;
}
function _decisoesPessoa(nome){
  const n=_normJS(nome), r={gestor:{},consultor:{},total:0};
  (EMP('sergipe').atividades||[]).forEach(a=>{
    if(_normJS(a.funcionario)!==n) return;
    r.total++;
    const g=_tipoDecisao(a.decisao_gestor), c=_tipoDecisao(a.decisao_consultor);
    r.gestor[g]=(r.gestor[g]||0)+1; r.consultor[c]=(r.consultor[c]||0)+1;
  });
  return r;
}
function _resumoDec(cont,total){
  if(!total) return {html:'<span style="color:var(--muted)">—</span>',txt:''};
  const tipos=Object.keys(cont);
  const det=tipos.map(t=>cont[t]+' '+_DEC_ROT[t].toLowerCase()).join(' · ');
  if(tipos.length===1){ const t=tipos[0]; return {html:_badgeDec(t,_DEC_ROT[t]+(total>1?' ('+total+')':''),det),txt:_DEC_ROT[t]}; }
  const ap=cont.aprovada||0;
  const t=cont.reprovada?'reprovada':(ap?'misto':'aguardando');
  return {html:_badgeDec(t,ap+'/'+total+' aprovadas',det),txt:det};
}
function _vgRow(p){
  if(p.pendente) return `<tr style="background:rgba(245,158,11,.07)"><td><span style="color:#f59e0b;font-weight:700">⚠ ${p.nome}</span></td><td style="font-size:12px;color:#f59e0b">${p.cargo}</td><td>${setor(p.setor)}</td><td><span style="font-size:11px;font-weight:600;color:${p.cor}">${esc(p.empresa||p.filial)}</span></td><td><span style="font-size:11px;color:#f59e0b;font-weight:600">Pendente</span></td><td></td><td></td><td></td></tr>`;
  const d=_decisoesPessoa(p.nome);
  return `<tr><td><strong>${p.nome}</strong></td><td style="font-size:12px;color:var(--subtle)">${p.cargo}</td><td>${setor(p.setor)}</td><td><span style="font-size:11px;font-weight:600;color:${p.cor}">${esc(p.empresa||p.filial)}</span></td><td><span class="count-pill">${p.atividades}</span></td><td>${_resumoDec(d.gestor,d.total).html}</td><td>${_resumoDec(d.consultor,d.total).html}</td><td style="text-align:right"><button type="button" class="btn-ver" data-func="${esc(p.nome)}">Ver atividades</button></td></tr>`;
}
window._vgFiltro='todos';
window._vgOrdem='setor'; window._vgDesc=false;
try{ const o=JSON.parse(localStorage.getItem('painel_vg_ordem')||'null'); if(o&&o.campo){ window._vgOrdem=o.campo; window._vgDesc=!!o.desc; } }catch(e){}
const _VG_NUM={atividades:1,gestor:1,consultor:1};   // numericos: comecam do maior
function _vgAguardando(p,quem){ if(p.pendente) return -1; const d=_decisoesPessoa(p.nome); return d[quem].aguardando||0; }
function _vgChave(p,campo){
  if(campo==='nome') return p.nome||'';
  if(campo==='cargo') return (p.cargo||'').trim()||'(sem cargo)';
  if(campo==='setor') return _setorDe(p);
  if(campo==='empresa') return p.empresa||p.filial||'';
  if(campo==='status') return p.pendente?0:1;
  if(campo==='atividades') return p.pendente?0:(p.atividades||0);
  if(campo==='gestor'||campo==='consultor') return _vgAguardando(p,campo);
  return '';
}
function _vgOrdenar(rows){
  const campo=window._vgOrdem, num=!!_VG_NUM[campo]||campo==='status';
  // numericos: padrao do maior para o menor; texto: A-Z. O botao inverte.
  const inv=(_VG_NUM[campo]?!window._vgDesc:window._vgDesc)?-1:1;
  const cmp=(a,b)=>{const x=_vgChave(a,campo), y=_vgChave(b,campo);
    const r=num?(x-y):String(x).localeCompare(String(y),'pt-BR',{sensitivity:'base'});
    return r*inv || (a.nome||'').localeCompare(b.nome||'','pt-BR');};
  return [...rows].sort(cmp);
}
function _vgComGrupos(rows){
  const campo=window._vgOrdem;
  if(!['setor','cargo','empresa'].includes(campo)) return rows.map(_vgRow);
  const html=[]; let atual=null;
  const cont={}; rows.forEach(p=>{const k=_vgChave(p,campo); cont[k]=(cont[k]||0)+1;});
  const env={}; rows.forEach(p=>{const k=_vgChave(p,campo); if(!p.pendente) env[k]=(env[k]||0)+1;});
  rows.forEach(p=>{const k=_vgChave(p,campo);
    if(k!==atual){ atual=k; html.push(`<tr class="vg-grupo"><td colspan="8">${esc(k)}<small>${cont[k]} pessoa${cont[k]>1?'s':''} · ${env[k]||0} enviaram</small></td></tr>`); }
    html.push(_vgRow(p));});
  return html;
}
function _vgAtualizarDir(){
  const b=document.getElementById('vg-dir'), sel=document.getElementById('vg-ordem'); if(!b||!sel) return;
  sel.value=window._vgOrdem;
  const num=!!_VG_NUM[window._vgOrdem]||window._vgOrdem==='status';
  b.textContent=num?(window._vgDesc?'↑ Menor':'↓ Maior'):(window._vgDesc?'Z→A':'A→Z');
  if(window._vgOrdem==='status') b.textContent=window._vgDesc?'Enviaram 1º':'Pendentes 1º';
}
function mudarOrdemVG(){
  window._vgOrdem=document.getElementById('vg-ordem').value; window._vgDesc=false;
  try{ localStorage.setItem('painel_vg_ordem',JSON.stringify({campo:window._vgOrdem,desc:window._vgDesc})); }catch(e){}
  filtrarEnviaram(window._vgFiltro);
}
function inverterOrdemVG(){
  window._vgDesc=!window._vgDesc;
  try{ localStorage.setItem('painel_vg_ordem',JSON.stringify({campo:window._vgOrdem,desc:window._vgDesc})); }catch(e){}
  filtrarEnviaram(window._vgFiltro);
}
window._vgEmpresa='todas';
function filtrarEnviaram(tipo){
  window._vgFiltro=tipo;
  const busca=(document.getElementById('vg-search').value||'').toLowerCase().trim();
  const emp=window._vgEmpresa||'todas';
  const filialMap={al:'Alagoas',se:'GRUPO_NOME_PH',rec:'Recife'};
  const filialFiltro=emp!=='todas'?filialMap[emp]:null;
  let envBase=filialFiltro?window._vgEnv.filter(p=>p.filial===filialFiltro):window._vgEnv;
  let naoBase=filialFiltro?window._vgNao.filter(p=>p.filial===filialFiltro):window._vgNao;
  let rows=tipo==='enviaram'?envBase:tipo==='pendentes'?naoBase:[...envBase,...naoBase];
  if(busca) rows=rows.filter(p=>p.nome.toLowerCase().includes(busca));
  rows=_vgOrdenar(rows);
  window._vgRows=rows;   // a exportacao sai na mesma ordem e com a mesma busca
  _vgAtualizarDir();
  setHtml('tb-enviaram',_vgComGrupos(rows));
  const total=envBase.length+naoBase.length;
  let label;
  if(busca) label=rows.length+' resultado'+(rows.length!==1?'s':'')+' de '+total+' — "'+busca+'"';
  else if(tipo==='todos') label=total+' pessoas · '+envBase.length+' enviaram · '+naoBase.length+' pendentes';
  else if(tipo==='enviaram') label=rows.length+' de '+total+' enviaram';
  else label=rows.length+' de '+total+' pendentes';
  document.getElementById('count-tb-enviaram').textContent=label;
  const cores={todos:'var(--se)',enviaram:'#10b981',pendentes:'#f59e0b'};
  ['todos','enviaram','pendentes'].forEach(t=>{
    const b=document.getElementById('fg-btn-'+t);
    const ativo=t===tipo;
    b.style.background=ativo?cores[t]:'transparent';
    b.style.color=ativo?'#fff':'var(--subtle)';
    b.style.fontWeight=ativo?'600':'400';
  });
}
function exportarVisaoGeral(ev){
  const rows=window._vgRows||[...window._vgEnv,...window._vgNao];
  const cols=[['Nome',30],['Cargo',24],['Setor',22],['Empresa',22],['Status',11],['Atividades',11],['Gestor',26],['Consultor',26]];
  const linhas=rows.map(p=>{
    const d=p.pendente?{gestor:{},consultor:{},total:0}:_decisoesPessoa(p.nome);
    return [p.nome,p.cargo,p.setor||'',p.empresa||p.filial,p.pendente?'Pendente':'Enviou',p.pendente?0:p.atividades,_resumoDec(d.gestor,d.total).txt,_resumoDec(d.consultor,d.total).txt];});
  const btn=ev&&ev.currentTarget?ev.currentTarget:null;
  _exportarTabela(cols,linhas,'Visao_Geral'+_sufixoSetores()+'_'+_hojeArq(),btn,'Visão Geral');
}
function populate(){
  // Enviaram + Não enviaram — armazena globalmente para filtro
  window._vgEnv=[...EMP('alagoas').enviaram_list.map(p=>({...p,filial:'Alagoas',cor:'#2563eb',pendente:false})),
   ...EMP('sergipe').enviaram_list.map(p=>({...p,filial:'GRUPO_NOME_PH',cor:'#7c3aed',pendente:false})),
   ...EMP('recife').enviaram_list.map(p=>({...p,filial:'Recife',cor:'#f97316',pendente:false}))]
    .sort((a,b)=>b.atividades-a.atividades);
  window._vgNao=[...EMP('alagoas').nao_enviaram_list.map(p=>({...p,filial:'Alagoas',cor:'#2563eb',pendente:true,atividades:0})),
   ...EMP('sergipe').nao_enviaram_list.map(p=>({...p,filial:'GRUPO_NOME_PH',cor:'#7c3aed',pendente:true,atividades:0})),
   ...EMP('recife').nao_enviaram_list.map(p=>({...p,filial:'Recife',cor:'#f97316',pendente:true,atividades:0}))]
    .sort((a,b)=>a.nome.localeCompare(b.nome));
  filtrarEnviaram('todos');
  // Cargo
  setHtml('tb-cargo-al',EMP('alagoas').cargo_data.map(c=>`<tr><td><strong style="font-size:12px">${c.cargo}</strong></td><td style="text-align:center">${c.total}</td><td style="text-align:center">${c.enviaram}</td><td style="text-align:center"><span class="count-pill">${c.atividades}</span></td><td class="pct-bar-cell">${pct_bar(c.pct)}</td></tr>`));
  setHtml('tb-cargo-se',EMP('sergipe').cargo_data.map(c=>`<tr><td><strong style="font-size:12px">${c.cargo}</strong></td><td style="text-align:center">${c.total}</td><td style="text-align:center">${c.enviaram}</td><td style="text-align:center"><span class="count-pill">${c.atividades}</span></td><td class="pct-bar-cell">${pct_bar(c.pct)}</td></tr>`));
  setHtml('tb-cargo-rec',EMP('recife').cargo_data.map(c=>`<tr><td><strong style="font-size:12px">${c.cargo}</strong></td><td style="text-align:center">${c.total}</td><td style="text-align:center">${c.enviaram}</td><td style="text-align:center"><span class="count-pill">${c.atividades}</span></td><td class="pct-bar-cell">${pct_bar(c.pct)}</td></tr>`));
  // Nao enviaram
  const statusBadge=s=>s==='Inativo'?'<span class="status-badge" style="background:rgba(239,68,68,.15);color:#f87171;border:1px solid rgba(239,68,68,.3)">Inativo</span>':'<span class="status-badge" style="background:rgba(16,185,129,.15);color:#34d399;border:1px solid rgba(16,185,129,.3)">Ativo</span>';
  const naoRow=p=>`<tr><td>${p.nome}</td><td style="font-size:12px;color:var(--subtle)">${p.cargo}</td><td>${setor(p.setor)}</td><td style="font-size:11px;color:var(--muted)">${p.perfil}</td><td>${statusBadge(p.status_usr)}</td></tr>`;
  setHtml('nao-al-body',EMP('alagoas').nao_enviaram_list.map(naoRow));
  setHtml('nao-se-body',EMP('sergipe').nao_enviaram_list.map(naoRow));
  setHtml('nao-rec-body',EMP('recife').nao_enviaram_list.map(naoRow));
  // Enviaram detalhado
  setHtml('env-al-body',EMP('alagoas').enviaram_list.map(p=>`<tr><td><strong>${p.nome}</strong></td><td style="font-size:12px;color:var(--subtle)">${p.cargo}</td><td>${setor(p.setor)}</td><td style="font-size:11px;color:var(--muted)">${p.perfil}</td><td style="text-align:center"><span class="count-pill">${p.atividades}</span></td></tr>`));
  setHtml('env-se-body',EMP('sergipe').enviaram_list.map(p=>`<tr><td><strong>${p.nome}</strong></td><td style="font-size:12px;color:var(--subtle)">${p.cargo}</td><td>${setor(p.setor)}</td><td style="font-size:11px;color:var(--muted)">${p.perfil}</td><td style="text-align:center"><span class="count-pill">${p.atividades}</span></td></tr>`));
  setHtml('env-rec-body',EMP('recife').enviaram_list.map(p=>`<tr><td><strong>${p.nome}</strong></td><td style="font-size:12px;color:var(--subtle)">${p.cargo}</td><td>${setor(p.setor)}</td><td style="font-size:11px;color:var(--muted)">${p.perfil}</td><td style="text-align:center"><span class="count-pill">${p.atividades}</span></td></tr>`));
  renderSetores();
  renderResumo('resumo');
  renderResumo('anl-resumo');
  renderAnalise();
  renderVersoes();
  // Atividades
  const atvRow=a=>{const _h=_extrairHorario(a.criada_em); return `<tr><td style="font-size:12px"><strong>${esc(a.funcionario)}</strong></td><td style="font-size:11px;color:var(--muted);font-family:'DM Mono',monospace">${_h?_h.texto:'—'}</td><td style="font-size:11px;color:var(--subtle)">${esc(a.cargo)}</td><td>${esc(a.nome)}</td><td style="font-size:11px;color:var(--text);white-space:pre-line;overflow:hidden;display:-webkit-box;-webkit-line-clamp:4;-webkit-box-orient:vertical;cursor:pointer" onclick="this.style.webkitLineClamp=this.style.webkitLineClamp==='unset'?'4':'unset'">${esc(a.detalhamento)}</td><td style="font-size:11px;color:var(--subtle)">${esc(a.periodicidade)}</td><td><span class="count-pill">${esc(a.tempo)}</span></td><td>${status_badge(a.status)}</td></tr>`;};
  setHtml('atv-al-body',EMP('alagoas').atividades.map(atvRow));
  setHtml('atv-se-body',EMP('sergipe').atividades.map(atvRow));
  setHtml('atv-rec-body',EMP('recife').atividades.map(atvRow));
  // Aprovações
  const aprRow=a=>{const rel=a.tem_relatorio?'<span style="color:#10b981;font-size:11px">Sim</span>':'<span style="color:#ef4444;font-size:11px">Não</span>';return `<tr><td><strong style="font-size:12px">${a.funcionario}</strong></td><td style="font-size:11px;color:var(--subtle)">${a.cargo}</td><td>${a.atividade}</td><td>${apr_status_badge(a.status)}</td><td style="font-size:11px;color:var(--muted)">${a.aprovador}</td><td style="font-family:'DM Mono',monospace;font-size:11px;color:var(--muted)">${a.data}</td><td>${rel}</td></tr>`;};
  if(DATA.aprovacoes){
    if(DATA.aprovacoes.alagoas) setHtml('apr-al-body',DATA.aprovacoes.alagoas.itens.map(aprRow));
    if(DATA.aprovacoes.sergipe) setHtml('apr-se-body',_filtrarAprovacoes().map(aprRow));
    if(DATA.aprovacoes.recife)  setHtml('apr-rec-body',DATA.aprovacoes.recife.itens.map(aprRow));
  }
}
// prazo definido pelo administrador no portal (aba Empresas); fora do portal vale o do robo
function _prazoISO(){
  try{ if(window.parent!==window && typeof window.parent.prazoDoPainel==='function'){
    const p=window.parent.prazoDoPainel(); if(p!==null && p!==undefined) return p ? p+'T23:59:59' : '';
  } }catch(e){}
  return 'PRAZO_ISO';
}
function calcularPrazo(){
  const iso = _prazoISO(), prazo = new Date(iso);
  const txt = document.getElementById('prazo-txt');
  if(txt) txt.textContent = 'Prazo ' + (isNaN(prazo) ? 'a definir' : iso.slice(0,10).split('-').reverse().join('/'));
  if(isNaN(prazo)){ ['p-dias','p-horas','p-min','p-seg'].forEach(id=>{const el=document.getElementById(id); if(el){el.textContent='--'; el.style.color='';}}); return; }
  const agora = new Date();
  const diff  = prazo - agora;
  const card  = document.getElementById('card-prazo');
  const pad   = n => String(Math.floor(n)).padStart(2,'0');
  const setColor = c => {
    ['p-dias','p-horas','p-min','p-seg'].forEach(id=>{
      document.getElementById(id).style.color = c;
    });
  };
  if(diff <= 0){
    ['p-dias','p-horas','p-min','p-seg'].forEach(id=>{document.getElementById(id).textContent='00';});
    setColor('#f59e0b');
    return;
  }
  const dias  = Math.floor(diff / 86400000);
  const horas = Math.floor((diff % 86400000) / 3600000);
  const min   = Math.floor((diff % 3600000)  / 60000);
  const seg   = Math.floor((diff % 60000)    / 1000);
  document.getElementById('p-dias').textContent  = pad(dias);
  document.getElementById('p-horas').textContent = pad(horas);
  document.getElementById('p-min').textContent   = pad(min);
  document.getElementById('p-seg').textContent   = pad(seg);
  setColor('#f59e0b');
}
document.addEventListener('DOMContentLoaded',()=>{
  const filtroSalvo=_carregarFiltroLocal();
  if(filtroSalvo && (filtroSalvo.de||filtroSalvo.ate||(filtroSalvo.turno&&filtroSalvo.turno!=='todos'))){
    document.getElementById('filtro-data-de').value=filtroSalvo.de||'';
    document.getElementById('filtro-data-ate').value=filtroSalvo.ate||'';
    document.getElementById('filtro-turno').value=filtroSalvo.turno||'todos';
    _recalcularFiltro(filtroSalvo.de||'', filtroSalvo.ate||'', filtroSalvo.turno||'todos');
  }
  populate();
  atualizarBlocosEmpresa();
  startDonutLoop();
  filterEmpresa(_empresaAtiva());
  _atualizarInfoFiltro();
  calcularPrazo();
  setInterval(calcularPrazo, 1000);
  try{ const aba=sessionStorage.getItem('painel_aba'); const b=aba&&document.querySelector(`.nav-btn[onclick*="'${aba}'"]`); if(b) showTab(aba,b); }catch(e){}
  const _recarregar=()=>{
    const ocupado=document.getElementById('modal-atividades').style.display==='flex'
      || (document.getElementById('tema-pop')||{classList:{contains:()=>false}}).classList.contains('aberto')
      || ['INPUT','SELECT','TEXTAREA'].includes((document.activeElement||{}).tagName);
    if(ocupado) setTimeout(_recarregar, 30*1000); else _recarregarPagina();
  };
  setTimeout(_recarregar, 2 * 60 * 1000);
});
</script>
<div id="modal-atividades" role="dialog" aria-modal="true" aria-labelledby="va-nome" style="display:none">
  <div class="va-caixa">
    <div class="va-cab"><div><div id="va-nome" class="va-titulo"></div><div id="va-sub" class="va-subt"></div></div>
      <div class="va-acoes"><button type="button" id="va-exportar" class="va-exportar" onclick="exportarAtividadesPessoa()">⬇ Exportar Excel</button>
      <button type="button" class="va-fechar" onclick="fecharAtividades()" aria-label="Fechar">✕</button></div></div>
    <div id="va-lista" class="va-corpo"></div>
  </div>
</div>
</body>
</html>"""

# ── Alerta dinâmico ───────────────────────────────────────────────────────────
_threshold_critico = 30
_crits = []
for _en, _key, _ed, _ep in [(GRUPO_NOME, 'se', se, pct_se)]:
    if _ep < _threshold_critico:
        _crits.append(
            f'<div class="alerta-linha" data-filial="{_key}">'
            f'<strong>{_en} com adesão crítica:</strong> apenas {_ed["enviaram"]} de {_ed["total"]} '
            f'funcionários cadastrados realizaram envio ({fmt(_ep)}%). Requer ação imediata de acompanhamento.'
            f'</div>'
        )
_alert_html = (
    '<div class="alert alert-red" id="alerta-criticos" style="margin-bottom:20px">'
    '<span class="alert-icon">⚠️</span>'
    '<div>' + ''.join(_crits) + '</div></div>'
) if _crits else '<div id="alerta-criticos" style="display:none"></div>'

def _pct_color(p): return 'var(--green)' if p >= 50 else 'var(--red)'

# ── Substituições ──────────────────────────────────────────────────────────────
replacements = {
    'PRAZO_TXT':     PRAZO_TXT,
    'PRAZO_ISO':     PRAZO_ISO,
    'DATE_HOJE':     HOJE,
    'DATE_HORA':     HORA,
    'HORA_REF':      _hora_ref,
    'DATA_JSON_PLACEHOLDER': DATA_JSON,
    'TOTAL_USERS':   str(total_users),
    'AL_TOTAL':      str(al['total']),
    'SE_TOTAL':      str(se['total']),
    'RE_TOTAL':      str(rec['total']),
    # NENV_TOTAL deve vir ANTES de ENV_TOTAL (evitar substituição parcial)
    'NENV_TOTAL':    str(total_nenv),
    'AL_NENV':       str(al['nao_enviaram']),
    'SE_NENV':       str(se['nao_enviaram']),
    'RE_NENV':       str(rec['nao_enviaram']),
    'ENV_TOTAL':     str(total_env),
    'AL_ENV':        str(al['enviaram']),
    'SE_ENV':        str(se['enviaram']),
    'RE_ENV':        str(rec['enviaram']),
    'ATV_TOTAL':     str(total_atv),
    'AL_ATV':        str(al['total_atividades']),
    'SE_ATV':        str(se['total_atividades']),
    'RE_ATV':        str(rec['total_atividades']),
    'PCT_GLOBAL':    fmt(pct_global),
    'PCT_AL':        fmt(pct_al),
    'SE_PCT':        fmt(pct_se),
    'RE_PCT':        fmt(pct_rec),
    'PCT_AL_RAW':    str(pct_al),
    'SE_PCT_RAW':    str(pct_se),
    'RE_PCT_RAW':    str(pct_rec),
    'AL_PROGRESS_COLOR':  _pct_color(pct_al),
    'SE_PROGRESS_COLOR':  _pct_color(pct_se),
    'REC_PROGRESS_COLOR': _pct_color(pct_rec),
    'N_CARGOS_AL':   str(n_cargos_al),
    'N_CARGOS_SE':   str(n_cargos_se),
    'N_CARGOS_REC':  str(n_cargos_rec),
    # Aprovações — específicos antes dos genéricos para evitar colisão
    'APR_REC_TOTAL':  str(apr_rec['total']),
    'APR_REC_APROV':  str(apr_rec['aprovados']),
    'APR_REC_PEND':   str(apr_rec['pendentes']),
    'APR_REC_REPROV': str(apr_rec['reprovados']),
    'APR_AL_TOTAL':   str(apr_al['total']),
    'APR_SE_TOTAL':   str(apr_se['total']),
    'APR_AL_APROV':   str(apr_al['aprovados']),
    'APR_SE_APROV':   str(apr_se['aprovados']),
    'APR_AL_PEND':    str(apr_al['pendentes']),
    'APR_SE_PEND':    str(apr_se['pendentes']),
    'APR_AL_REPROV':  str(apr_al['reprovados']),
    'APR_SE_REPROV':  str(apr_se['reprovados']),
    'APROV_NAV_DISPLAY': 'inline-flex' if TEM_APROVACOES else 'none',
    'APR_TOTAL':     str(apr_total),
    'APR_APROV':     str(apr_aprov),
    'APR_PEND':      str(apr_pend),
    'APR_REPROV':    str(apr_reprov),
    'NOVOS_ENVIOS_CARD':  novos_card,
    'ALERT_CRITICOS_HTML': _alert_html,
    'AL_TROPHY_HTML': trophy_al,
    'SE_TROPHY_HTML': trophy_se,
    'RE_TROPHY_HTML': trophy_rec,
}

import html as _html_mod
_NOME_ESC = _html_mod.escape(GRUPO_NOME)
_LOGO_FUNDO = _CFG.get('logo_fundo') or ''   # logo branca/transparente: fundo proprio (legivel nos temas claros)
_LOGO_EST = f' style="background:{_html_mod.escape(_LOGO_FUNDO)};padding:7px 12px;box-sizing:border-box"' if _LOGO_FUNDO else ''
LOGO_HTML = (f'<img class="logo-millena" src="{LOGO_SRC}" alt="{_NOME_ESC}"{_LOGO_EST}>' if LOGO_SRC
             else f'<div class="logo-texto">{_NOME_ESC}</div>')
html = HTML.replace('LOGO_HTML_PH', LOGO_HTML)
for k, v in sorted({**replacements, 'SETOR_HTML_PLACEHOLDER': SETOR_HTML, 'LOGO_SRC_PLACEHOLDER': LOGO_SRC}.items(), key=lambda kv: -len(kv[0])):
    html = html.replace(k, v)
# nome do grupo por ultimo (o nome nao pode ser confundido com outro marcador)
html = html.replace('GRUPO_ARQ_PH', ARQ).replace('GRUPO_NOME_PH', _NOME_ESC)
html = html.replace('INTERVALO_MIN_PH', str(int(_CFG.get('intervalo_min') or 60)))

# ── Salvar ────────────────────────────────────────────────────────────────────
output = BASE / f'Dashboard_{ARQ}_{DATA_ARQUIVO}.html'
output.write_text(html, encoding='utf-8')
print(f'Dashboard gerado: {output.name}')

# ── ZIP para envio por e-mail (evita bloqueio de antivírus) ──────────────────
zip_output = BASE / f'Dashboard_{ARQ}_{DATA_ARQUIVO}.zip'
with zipfile.ZipFile(zip_output, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.write(output, output.name)
print(f'ZIP gerado:       {zip_output.name}  (use este para enviar por e-mail)')

# ── ZIP com nome fixo (link permanente no OneDrive — sempre a versão mais recente) ──
zip_latest = BASE / f'Dashboard_{ARQ}_LATEST.zip'
with zipfile.ZipFile(zip_latest, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.write(output, output.name)
print(f'ZIP permanente:   {zip_latest.name}  (compartilhe o link OneDrive deste)')

print(f'{GRUPO_NOME:<13}: {se["total"]} usuarios | {se["enviaram"]} enviaram | {se["nao_enviaram"]} nao enviaram | {se["total_atividades"]} atividades')

# ── Abrir no navegador apenas na primeira execução do dia ────────────────────
# Usa nome fixo para que a aba já aberta seja reutilizada pelo browser.
# O HTML tem setTimeout que recarrega a cada 2 min — sem nova aba a cada ciclo.
html_atual = BASE / f'Dashboard_{ARQ}_ATUAL.html'
html_atual.write_text(html, encoding='utf-8')

_browser_flag = BASE / '.browser_opened.flag'
_abrir_browser = True
if _browser_flag.exists():
    try:
        _flag_data = json.loads(_browser_flag.read_text(encoding='utf-8'))
        if _flag_data.get('data') == HOJE:
            _abrir_browser = False
    except Exception:
        pass

if _abrir_browser:
    webbrowser.open(html_atual.as_uri())
    _browser_flag.write_text(json.dumps({'data': HOJE}), encoding='utf-8')
    print(f'Navegador aberto: {html_atual.name}')
else:
    print(f'Dashboard atualizado — aba existente recarregara automaticamente em ate 2 min.')

# ── Exportar Atividades por Empresa (.xlsx) ───────────────────────────────────
_PASTA_ATV = BASE.parent / "ATIVIDADES ENVIADAS"
_PASTA_ATV.mkdir(parents=True, exist_ok=True)

_HEADERS_ATV = ['Funcionário', 'Cargo', 'Setor', 'Atividade', 'Detalhamento', 'Periodicidade', 'Tempo', 'Status']

def _exportar_xlsx(dados_empresa, nome_empresa):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Atividades'
    ws.append(_HEADERS_ATV)
    for a in dados_empresa['atividades']:
        ws.append([
            a.get('funcionario', ''),
            a.get('cargo', ''),
            a.get('setor', ''),
            a.get('nome', ''),
            a.get('detalhamento', ''),
            a.get('periodicidade', ''),
            a.get('tempo', ''),
            a.get('status', ''),
        ])
    # Auto-ajuste de largura das colunas
    for col in ws.columns:
        max_len = max((len(str(cell.value or '')) for cell in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
    destino = _PASTA_ATV / f'Atividades_{nome_empresa}_{DATA_ARQUIVO}.xlsx'
    try:
        wb.save(str(destino))
        print(f'  Exportado: {destino.name}')
    except PermissionError:
        print(f'  [aviso] {destino.name} esta aberto — feche o Excel e rode novamente para atualizar.')

print('\n--- ATIVIDADES XLSX ---')
_exportar_xlsx(se,  ARQ)

# Mantém apenas o arquivo mais recente por empresa
for _prefixo in [f'Atividades_{ARQ}_']:
    _arqs = sorted(_PASTA_ATV.glob(f'{_prefixo}*.xlsx'), key=lambda f: f.stat().st_mtime, reverse=True)
    for _old in _arqs[1:]:
        _old.unlink()
        print(f'  Removido obsoleto: {_old.name}')
