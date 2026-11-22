"""
baixar_usuarios.py — roda no GitHub Actions uma vez por dia.

A API do mapeia ainda nao traz o cadastro de colaboradores. Ate trazer, este script entra no
mapeia com MAPEIA_EMAIL / MAPEIA_SENHA (segredos do repositorio), exporta a lista de Usuarios
de cada empresa dos grupos (painel/grupos.json) e grava so nome, cargo, empresa, setor, perfil
e status em dados/usuarios.json (sem e-mail nem matricula). O workflow guarda o arquivo na
area privada do Supabase; ele nao fica no repositorio.
"""
import json, os, re, sys, time, tempfile
from datetime import datetime
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright

AQUI = Path(__file__).parent
SAIDA = AQUI.parent / "dados" / "usuarios.json"
GRUPOS = json.loads((AQUI / "grupos.json").read_text(encoding="utf-8"))["grupos"]
EMAIL, SENHA = os.environ.get("MAPEIA_EMAIL", ""), os.environ.get("MAPEIA_SENHA", "")


def _chave(t):
    import unicodedata
    t = unicodedata.normalize("NFD", str(t)).encode("ascii", "ignore").decode().lower()
    return "".join(ch for ch in t if ch.isalnum())


def _de_algum_grupo(empresa):
    e = _chave(empresa)
    return bool(e) and any(_chave(t) in e for g in GRUPOS for t in g["empresas"])


def _combo_empresa(page):
    page.evaluate("""() => {
      const lab = [...document.querySelectorAll('label,span,div,p')]
        .find(e => e.childElementCount === 0 && e.textContent.trim() === 'Empresa');
      for (let c = lab && lab.parentElement, i = 0; c && i < 3; c = c.parentElement, i++) {
        const ctl = c.querySelector('button[role=combobox], button');
        if (ctl) { ctl.setAttribute('data-filtro', 'empresa'); return; }
      }
    }""")
    return page.locator('[data-filtro=empresa]').first


def listar_empresas(page):
    combo = _combo_empresa(page)
    combo.click()
    page.locator('[role="option"]:visible').first.wait_for(timeout=10_000)
    # o mapeia tem empresas com o MESMO nome (ex.: duas "Guanabara"): guarda (nome, ordem)
    nomes, vistos = [], {}
    for t in page.locator('[role="option"]:visible').all_inner_texts():
        t = t.strip()
        if t:
            nomes.append((t, vistos.get(t, 0)))
            vistos[t] = vistos.get(t, 0) + 1
    page.keyboard.press("Escape")
    time.sleep(0.5)
    return nomes


def selecionar_empresa(page, nome, ordem=0):
    combo = _combo_empresa(page)
    if ordem or combo.inner_text().strip() != nome:
        combo.click()
        page.locator('[role="option"]:visible').filter(has_text=re.compile("^" + re.escape(nome) + "$")).nth(ordem).click()
        time.sleep(1)
    if combo.inner_text().strip() != nome:
        raise RuntimeError("filtro de empresa nao aplicado")


def exportar(page, nome, ordem=0):
    """Exporta os usuarios da empresa e devolve as linhas (dicionarios)."""
    for tent in range(3):
        try:
            page.goto("https://mapeia.sinerggia.com.br/admin/usuarios")
            page.get_by_text("Empresa", exact=True).first.wait_for(timeout=15_000)
            selecionar_empresa(page, nome, ordem)
            time.sleep(2)
            botao = page.locator('button:has-text("Exportar Excel")').first
            if botao.count() == 0 or botao.is_disabled():
                print(f"::notice::{nome}: nenhum usuario cadastrado (botao Exportar indisponivel)")
                return []
            with page.expect_download(timeout=30_000) as dl:
                botao.click()
            arq = Path(tempfile.gettempdir()) / "usuarios.xlsx"
            dl.value.save_as(str(arq))
            break
        except Exception as e:
            print(f"::warning::exportar usuarios de {nome}, tentativa {tent + 1}: {str(e).splitlines()[0][:200]}")
            if tent == 2:
                raise
    ws = openpyxl.load_workbook(arq).active
    cab = [str(c.value or "") for c in ws[1]]
    linhas = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(cab, ["" if v is None else str(v).strip() for v in row]))
        r.setdefault("Empresa", nome)
        if r.get("Nome") and _chave(r.get("Empresa") or nome) == _chave(nome):
            linhas.append({k.lower(): r.get(k, "") or (nome if k == "Empresa" else "")
                           for k in ("Nome", "Cargo", "Empresa", "Setor", "Perfil", "Status")})
    return linhas


def main():
    if not (EMAIL and SENHA):
        print("::warning::Sem MAPEIA_EMAIL/MAPEIA_SENHA: lista de usuarios mantida como esta.")
        return
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(accept_downloads=True)
        page = ctx.new_page()
        respostas_login = []

        def _resp(r):
            # so registra falhas do login (codigo + mensagem de erro); sucesso traz tokens e nao e lido
            if "/auth/v1/token" in r.url and r.status >= 400:
                try:
                    j = r.json()
                    msg = j.get("error_description") or j.get("msg") or j.get("message") or j.get("error_code") or ""
                except Exception:
                    msg = ""
                respostas_login.append(f"{r.status} {msg}"[:160])
        page.on("response", _resp)

        entrou = False
        for tent in range(3):
            page.goto("https://mapeia.sinerggia.com.br/login", wait_until="networkidle")
            campo_email = page.locator('input[type="email"], input[name="email"], input[placeholder*="mail" i]').first
            campo_email.wait_for(state="visible", timeout=20_000)
            time.sleep(2)  # deixa a pagina terminar de montar antes de digitar
            campo_email.fill(EMAIL.strip())
            time.sleep(0.5)
            page.locator('input[type="password"]').first.fill(SENHA.strip())
            time.sleep(0.5)
            page.locator('button[type="submit"], button:has-text("Entrar"), button:has-text("Acessar")').first.click()
            try:
                page.wait_for_url(lambda u: "login" not in u.lower(), timeout=30_000)
                entrou = True
                break
            except Exception:
                # mostra so o aviso da tela (sem e-mail nem senha) para saber o motivo
                avisos = page.locator('[role="alert"], [data-sonner-toast], .text-destructive, .text-red-500')
                textos = [t.strip() for t in avisos.all_inner_texts() if t.strip()]
                print(f"::warning::login tentativa {tent + 1}: nao entrou. Aviso na tela: {textos or 'nenhum'}; "
                      f"url={page.url.split('?')[0]}; resposta do login={respostas_login or 'nenhuma requisicao de login'}; "
                      f"tamanho do e-mail={len(EMAIL.strip())}, da senha={len(SENHA.strip())}")
                respostas_login.clear()
        if not entrou:
            print("::error::login no mapeia falhou (confira MAPEIA_EMAIL e MAPEIA_SENHA nos segredos)")
            sys.exit(1)
        page.goto("https://mapeia.sinerggia.com.br/admin/usuarios")
        page.get_by_text("Empresa", exact=True).first.wait_for(timeout=15_000)
        empresas = [(n, k) for n, k in listar_empresas(page) if _de_algum_grupo(n)]
        print(f"::notice::Empresas dos grupos no mapeia: {', '.join(n + (f' ({k + 1})' if k else '') for n, k in empresas) or 'nenhuma'}")
        usuarios, falhas = [], []
        for nome, ordem in empresas:
            try:
                usuarios += exportar(page, nome, ordem)
            except Exception:
                falhas.append(nome)
        b.close()
    if falhas:
        print(f"::error::nao foi possivel exportar os usuarios de: {', '.join(falhas)}")
        sys.exit(1)
    if not usuarios:
        print("::error::nenhum usuario exportado")
        sys.exit(1)

    usuarios.sort(key=lambda u: u["nome"])
    antigo = json.loads(SAIDA.read_text(encoding="utf-8"))["usuarios"] if SAIDA.exists() else None
    if usuarios == antigo:
        print(f"::notice::Usuarios sem mudanca ({len(usuarios)}).")
        return
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(json.dumps({"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                 "usuarios": usuarios}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"::notice::Usuarios atualizados: {len(usuarios)}")


if __name__ == "__main__":
    main()
