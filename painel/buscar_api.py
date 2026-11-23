"""
buscar_api.py — roda no GitHub Actions.

Busca as atividades de TODAS as empresas na API do mapeia (uma consulta so: a API aceita
1 consulta de atividades a cada 5 min), separa por grupo (painel/grupos.json) e monta, numa
pasta por grupo, os arquivos que o gerar_dashboard.py le (Relatorio, Aprovacoes e Usuarios).

Uso: python buscar_api.py <pasta_de_trabalho>   ->  <pasta>/<slug>/... e <pasta>/grupos.txt
A chave vem da variavel de ambiente MAPEIA_API_KEY (segredo do repositorio).

Paineis restritos (tabela "filtros" do portal): se existir <pasta>/agora.json (montado pelo
workflow), alem dos grupos completos monta uma pasta <slug>__<usuario> por usuario restrito,
so com as pessoas que ele pode ver (setor, cargo ou nome). Tudo vai para <pasta>/paineis.txt:
"<pasta> <caminho no Storage> <usuario|-> <empresa> <base_em|->", uma linha por painel.
Tambem grava <pasta>/opcoes.json (setores, cargos e nomes de cada grupo, para o Admin escolher).
"""
import json, os, sys, shutil, gzip, unicodedata
import urllib.request, urllib.parse, urllib.error
from datetime import datetime
from pathlib import Path
import openpyxl

API   = "https://ruwfvtgacnimpebelxfd.supabase.co/functions/v1/api-relatorio"
AQUI  = Path(__file__).parent
RAIZ  = AQUI.parent
GRUPOS = json.loads((AQUI / "grupos.json").read_text(encoding="utf-8"))["grupos"]
HOJE  = datetime.now().strftime("%d%m%Y")

COLS_REL = ['#', 'Empresa', 'Setor', 'Funcionário', 'Cargo', 'Nome', 'Detalhamento',
            'Periodicidade', 'Como realiza', 'Tempo', 'Horário fixo', 'Pontos críticos',
            'Sugestões de melhoria', 'Status', 'Decisão gestor', 'Decisão consultor',
            'Edições do consultor', 'Parecer do consultor', 'Criada em', 'Atualizada em']


def _get(params, tentativa=1):
    chave = os.environ.get("MAPEIA_API_KEY", "").strip()
    if not chave:
        sys.exit("ERRO: segredo MAPEIA_API_KEY nao configurado no repositorio.")
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {chave}",
                                               "Accept-Encoding": "gzip"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            corpo = r.read()
            if r.headers.get("Content-Encoding") == "gzip":
                corpo = gzip.decompress(corpo)
    except urllib.error.HTTPError as e:
        corpo = e.read()
        if e.headers.get("Content-Encoding") == "gzip" or corpo[:2] == b"\x1f\x8b":
            corpo = gzip.decompress(corpo)
        msg = corpo[:300].decode("utf-8", "replace")
        if e.code == 429 and tentativa == 1:
            # a API aceita 1 consulta a cada 5 min: espera o tempo pedido e tenta de novo
            import re, time
            m = re.search(r"(\d+)\s*min", msg)
            espera = min(int(m.group(1)) if m else 5, 5) * 60 + 15
            print(f"API pediu para esperar ({msg[:120]}); tentando de novo em {espera} s")
            time.sleep(espera)
            return _get(params, 2)
        print(f"::error::API respondeu {e.code} ({params.get('recurso', 'atividades')}): {msg}")
        sys.exit(1)
    return json.loads(corpo.decode("utf-8"))


def _lista(resp):
    """A resposta pode ser uma lista ou um objeto com a lista dentro."""
    if isinstance(resp, list):
        return resp
    for v in resp.values():
        if isinstance(v, list):
            return v
    return []


def _chave(s):
    """'Funcionário', 'funcionario', 'decisao_gestor' -> 'funcionario', 'decisaogestor'."""
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower()
    return "".join(ch for ch in s if ch.isalnum())


# nomes alternativos que a API pode usar para cada coluna do Excel
_SINONIMOS = {
    "Funcionário": ["funcionario", "funcionarionome", "colaborador", "usuario", "autor"],
    "Nome": ["nome", "titulo", "atividade", "nomeatividade"],
    "Criada em": ["criadaem", "criadoem", "createdat", "datacriacao"],
    "Atualizada em": ["atualizadaem", "atualizadoem", "updatedat"],
    "Status": ["status", "statusgeral", "situacao"],
    "Sugestões de melhoria": ["sugestoesdemelhoria", "sugestoesmelhoria", "sugestoes", "sugestao"],
    "Pontos críticos": ["pontoscriticos", "pontocritico"],
    "Empresa": ["empresa", "empresanome", "filial"],
    "Horário fixo": ["horariofixo", "temhorariofixo", "horario"],
    "Decisão gestor": ["decisaogestor", "decisaodogestor", "statusgestor"],
    "Decisão consultor": ["decisaoconsultor", "decisaodoconsultor", "statusconsultor"],
    "Edições do consultor": ["edicoesdoconsultor", "edicoesconsultor", "edicoes", "qtdedicoes"],
    "Parecer do consultor": ["parecerdoconsultor", "parecerconsultor", "parecer"],
}


def _padronizar(a):
    """Devolve a atividade com as chaves exatamente como as colunas do Excel antigo."""
    idx = {_chave(k): v for k, v in a.items()}
    out = {}
    for c in COLS_REL[1:]:
        for k in [_chave(c)] + _SINONIMOS.get(c, []):
            if k in idx and idx[k] is not None:
                v = idx[k]
                if isinstance(v, dict):  # ex.: {"nome": "..."}
                    v = v.get("nome") or v.get("name") or ""
                out[c] = v
                break
        else:
            out[c] = ""
    return out


def _horas(t):
    try:
        h, m = str(t).split(":")[:2]
        return int(h) + int(m) / 60
    except Exception:
        return 0


def _salvar(caminho, abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nome, linhas in abas:
        ws = wb.create_sheet(nome)
        for l in linhas:
            ws.append(l)
    # o gerador le a aba ativa: como no export do mapeia, ela e a de "Atividades"
    nomes = [n for n, _ in abas]
    wb.active = nomes.index("Atividades") if "Atividades" in nomes else 0
    wb.save(caminho)


def _do_grupo(g, empresa):
    e = _chave(empresa)
    return bool(e) and any(_chave(t) in e for t in g["empresas"])


def _montar_grupo(trab, g, atv, usuarios, pasta=None):
    """Grava, em <trab>/<pasta ou slug>/, o que o gerador precisa para o painel do grupo."""
    raiz = trab / (pasta or g["slug"])
    dest = raiz / "ResumoGerencialMapeamentoMillena"
    dest.mkdir(parents=True, exist_ok=True)
    arq = g["arquivo"]
    shutil.copy(AQUI / "gerar_dashboard.py", dest / "gerar_dashboard.py")
    if (AQUI / g["logo"]).exists():
        shutil.copy(AQUI / g["logo"], raiz / g["logo"])
    (raiz / "config.json").write_text(json.dumps(
        {"nome": g["nome"], "arquivo": arq, "prazo": g.get("prazo", ""), "logo": g["logo"],
         "logo_fundo": g.get("logo_fundo", "")},
        ensure_ascii=False), encoding="utf-8")

    linhas = [COLS_REL]
    for i, a in enumerate(atv, 1):
        linhas.append([i] + [a.get(c, "") if a.get(c) is not None else "" for c in COLS_REL[1:]])
    st = [str(a.get("Status", "")) for a in atv]
    resumo = [
        ["Relatório de Atividades", None],
        [f"Gerado em {datetime.now():%d/%m/%Y, %H:%M:%S} (API)", None],
        [None, None], ["Filtros aplicados", None], ["Empresa", g["nome"]], [None, None],
        ["Resumo executivo", None],
        ["Total de atividades", len(atv)],
        ["Funcionários distintos", len({a.get("Funcionário") for a in atv})],
        ["Aprovadas", sum(s == "Aprovada" for s in st)],
        ["Ressalvas", sum(s == "Ressalva" for s in st)],
        ["Aguardando gestor", sum(s == "Aguardando gestor" for s in st)],
        ["Aguardando consultor", sum(s == "Aguardando consultor" for s in st)],
        ["Tempo total mapeado (horas)", round(sum(_horas(a.get("Tempo")) for a in atv), 1)],
        ["Atividades editadas pelo consultor",
         sum(str(a.get("Edições do consultor", "0")).strip() not in ("", "0") for a in atv)],
        ["Atividades com parecer do consultor",
         sum(bool(str(a.get("Parecer do consultor", "") or "").strip()) for a in atv)],
    ]
    versoes = [["Atividade", "Empresa", "Setor", "Funcionário", "Tipo", "Autor", "Data/hora", "Conteúdo"],
               ["(Nenhuma edição ou parecer)"] + [None] * 7]
    _salvar(dest / f"Relatorio_{arq}_{HOJE}.xlsx",
            [("Resumo", resumo), ("Atividades", linhas), ("Versões & Parecer", versoes)])

    cab = ['#', 'Funcionário', 'Cargo', 'Nome', 'Status geral', 'Decisão gestor',
           'Decisão consultor', 'Criada em', 'Atualizada em']
    apr = [cab] + [[i, a.get("Funcionário", ""), a.get("Cargo", ""), a.get("Nome", ""),
                    a.get("Status", ""), a.get("Decisão gestor", ""), a.get("Decisão consultor", ""),
                    a.get("Criada em", ""), a.get("Atualizada em", "")] for i, a in enumerate(atv, 1)]
    _salvar(dest / f"Aprovacoes_{arq}_{HOJE}.xlsx", [("Atividades", apr)])

    cab_u = ['Nome', 'Cargo', 'Empresa', 'Setor', 'Perfil', 'Status']
    _salvar(dest / f"Usuarios_{arq}_{HOJE}.xlsx",
            [("Usuarios", [cab_u] + [[u.get(c.lower(), "") for c in cab_u] for u in usuarios])])


def _filtrar(atv, usuarios, f):
    """So as atividades e cadastrados que o filtro libera: quem estiver em QUALQUER item marcado
    (um dos setores, um dos cargos ou um dos nomes). Filtro vazio = ninguem."""
    S = {_chave(x) for x in f.get("setores") or []} - {""}
    C = {_chave(x) for x in f.get("cargos") or []} - {""}
    N = {_chave(x) for x in f.get("funcionarios") or []} - {""}
    ok = lambda setor, cargo, nome: _chave(setor) in S or _chave(cargo) in C or _chave(nome) in N
    return ([a for a in atv if ok(a.get("Setor", ""), a.get("Cargo", ""), a.get("Funcionário", ""))],
            [u for u in usuarios if ok(u.get("setor", ""), u.get("cargo", ""), u.get("nome", ""))])


def _opcoes(atv, usuarios):
    """Setores, cargos e funcionarios do grupo (para o Admin montar os filtros)."""
    pessoas = {}
    for u in usuarios:
        n = str(u.get("nome", "") or "").strip()
        if n:
            pessoas.setdefault(_chave(n), {"nome": n, "setor": u.get("setor", "") or "", "cargo": u.get("cargo", "") or ""})
    for a in atv:
        n = str(a.get("Funcionário", "") or "").strip()
        if n:
            pessoas.setdefault(_chave(n), {"nome": n, "setor": a.get("Setor", "") or "", "cargo": a.get("Cargo", "") or ""})
    lista = lambda vals: sorted({str(v).strip() for v in vals if str(v or "").strip()}, key=_chave)
    return {"setores": lista([a.get("Setor") for a in atv] + [u.get("setor") for u in usuarios]),
            "cargos": lista([a.get("Cargo") for a in atv] + [u.get("cargo") for u in usuarios]),
            "funcionarios": sorted(pessoas.values(), key=lambda p: _chave(p["nome"]))}


def main():
    trab = Path(sys.argv[1]).resolve()
    trab.mkdir(parents=True, exist_ok=True)

    # empresas que a chave enxerga (diagnostico: a consulta de filtros nao conta no limite de 5 min)
    try:
        filtros = _get({"recurso": "filtros"})
        emps_f = filtros.get("empresas") if isinstance(filtros, dict) else _lista(filtros)
        nomes_f = [str(e.get("nome") or e.get("name") or e) if isinstance(e, dict) else str(e) for e in (emps_f or [])]
        print(f"::notice::A chave da API enxerga {len(nomes_f)} empresa(s): {', '.join(nomes_f)[:600]}")
    except SystemExit:
        print("::warning::nao consegui ler os filtros da API")

    # atividades de todas as empresas que a chave enxerga (uma consulta so)
    brutas = _lista(_get({"recurso": "atividades", "formato": "json"}))
    from collections import Counter
    print("::notice::Atividades por empresa na API: " + ", ".join(
        f"{k or '(vazio)'}={v}" for k, v in Counter(_padronizar(a).get("Empresa", "") for a in brutas).most_common()))
    if brutas:
        print("::notice::Campos da API: " + ", ".join(sorted(brutas[0].keys())))
    atv = [_padronizar(a) for a in brutas]

    # cadastrados (lista diaria do robo de usuarios, todas as empresas)
    usr = json.loads((RAIZ / "dados" / "usuarios.json").read_text(encoding="utf-8"))
    usuarios = usr["usuarios"]
    print(f"Usuarios: {len(usuarios)} (lista de {usr.get('atualizado_em', '?')})")

    # quais paineis gerar: agora.json (do workflow) ou, sem ele, todos os grupos e nenhum restrito
    arq_agora = trab / "agora.json"
    if arq_agora.exists():
        cfg = json.loads(arq_agora.read_text(encoding="utf-8"))
        agora, filtros = cfg.get("grupos") or [], cfg.get("filtros") or []
    else:
        agora, filtros = [g["slug"] for g in GRUPOS], []

    gerados, usadas, paineis, opcoes = [], set(), [], {}
    for g in GRUPOS:
        a_g = [a for a in atv if _do_grupo(g, a.get("Empresa", ""))]
        u_g = [u for u in usuarios if _do_grupo(g, u.get("empresa", ""))]
        emps = sorted({a.get("Empresa") for a in a_g} | {u.get("empresa") for u in u_g})
        usadas.update(emps)
        opcoes[g["slug"]] = _opcoes(a_g, u_g)
        if g["slug"] in agora:
            if not a_g and not u_g:
                print(f"::warning::Grupo {g['nome']}: nenhuma empresa encontrada para {g['empresas']}")
            else:
                _montar_grupo(trab, g, a_g, u_g)
                gerados.append(g["slug"])
                paineis.append(f"{g['slug']} {g['slug']}/index.html - {g['slug']} -")
                print(f"::notice::{g['nome']}: {len(a_g)} atividades, {len(u_g)} cadastrados "
                      f"({len(emps)} empresas: {', '.join(emps)})")
        # paineis restritos deste grupo (sem nomes no log: o log do GitHub e publico)
        meus = [f for f in filtros if f.get("empresa") == g["slug"]]
        for i, f in enumerate(meus, 1):
            a_f, u_f = _filtrar(a_g, u_g, f)
            pasta = f"{g['slug']}__{f['usuario']}"
            _montar_grupo(trab, g, a_f, u_f, pasta)
            paineis.append(f"{pasta} {g['slug']}/u/{f['usuario']}.html {f['usuario']} {g['slug']} {f['alterado_em']}")
            print(f"::notice::{g['nome']}: painel restrito {i}/{len(meus)} com {len(a_f)} atividades e {len(u_f)} cadastrados")
    sobra = sorted({a.get("Empresa") for a in atv} - usadas - {""})
    if sobra:
        print(f"::notice::Empresas sem grupo (sem painel): {', '.join(sobra)}")
    (trab / "grupos.txt").write_text(" ".join(gerados), encoding="utf-8")
    (trab / "paineis.txt").write_text("\n".join(paineis) + ("\n" if paineis else ""), encoding="utf-8")
    (trab / "opcoes.json").write_text(json.dumps(
        {"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"), "grupos": opcoes}, ensure_ascii=False),
        encoding="utf-8")


if __name__ == "__main__":
    main()
