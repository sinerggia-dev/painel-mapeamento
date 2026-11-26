# Testes da Gestão de Contratos

Tudo aqui usa **dados inventados**. Nenhuma planilha, logo ou dado real de cliente entra nesta pasta
(o repositório é público).

| Arquivo | O que é |
|---|---|
| `mock.js` | Supabase e SharePoint simulados (empresas, filiais, contratos, documentos, naturezas, pessoas de exemplo). |
| `build.py` | Monta o protótipo: `contratos/index.html` + `mock.js` + grupos e categorias tirados do SQL. É o mesmo arquivo publicado como protótipo no claude.ai. |
| `rodar.sh` | Roda todos os testes de tela (Playwright/Chromium) no computador, Android e iPhone simulados. |
| `teste*.js`, `mob.js` | Os testes: painel, lista, contrato, cadastro, importação, naturezas, modos de visão, personalização, pessoas, reajuste, versões, tipo do contrato e celular. |
| `sql.sh`, `sql/` | Testes do banco num Postgres 16 local, com um "stub" do Supabase (login simulado por `test.uid`). |

## Como rodar

```bash
tests/contratos/rodar.sh /tmp/ctr-teste
sudo tests/contratos/sql.sh
```

- Precisa de Node com Playwright e Chromium, Python 3 e (para o banco) Postgres 16.
- A biblioteca de Excel (`xlsx@0.18.5`) é baixada pelo `npm pack` na primeira vez.
- Opcional, só na sua máquina: `<pasta>/logo.png` (logo da empresa no protótipo) e `<pasta>/up/nat.xlsx`
  (planilha real de naturezas, para testar a importação de verdade). Sem eles, essas partes são puladas.

## Depois de mudar o sistema

1. Rode os dois scripts. Teste que falhou porque o comportamento mudou de propósito: ajuste o teste.
2. Para atualizar o protótipo no claude.ai, publique `<pasta>/gestao-contratos.html` (gerado pelo `build.py`).
