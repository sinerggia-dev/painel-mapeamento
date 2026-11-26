# Painel de Mapeamento — Sinergia

Portal com login que mostra, para cada cliente da Sinergia, o andamento do **mapeamento de
atividades** feito no sistema **mapeia** (https://mapeia.sinerggia.com.br). Hoje há 3 painéis
(um por grupo): **Millena**, **Novatec** (8 unidades) e **Metropolitana**.

- Portal: https://sinerggia-dev.github.io/painel-mapeamento/
- Dono: Natanael Silva (Sinerggia). Fale sempre em **português**, de forma simples e prática.
  Ele não é programador: quando precisar que ele faça algo em outro site (GitHub, Supabase,
  cron-job.org, mapeia), dê **um passo de cada vez**, com o link exato e o nome do botão
  como aparece na tela. Ele costuma responder com prints.

## Como funciona (visão geral)

```
cron-job.org (a cada 10 min) ──► GitHub Actions "Atualizar painel"
                                   1. pergunta ao Supabase quais grupos estão na hora
                                   2. API do mapeia (1 consulta para TODAS as empresas)
                                   3. gera 1 HTML por grupo (painel/gerar_dashboard.py)
                                   4. grava em Supabase Storage (bucket PRIVADO "paineis/<slug>/index.html")
                                   5. em push: publica o portal (portal/index.html) no GitHub Pages

cron-job.org (1x por dia, 00:06) ──► GitHub Actions "Atualizar lista de usuarios"
                                   Playwright entra no mapeia, exporta os cadastrados de cada empresa
                                   e grava em Supabase Storage "paineis/_privado/usuarios.json"

Portal (GitHub Pages, público) ──► login Supabase Auth ──► baixa o HTML do grupo do Storage
                                   (regras RLS: só quem tem acesso à empresa) e mostra num iframe
```

**Nenhum dado de funcionário fica no repositório** (ele é público). Dados e painéis ficam só
no Supabase. Não commitar `dados/`, xlsx, html gerados, nem nada com nomes de pessoas.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `portal/index.html` | O site inteiro do portal: login, seletor de empresa, iframe do painel, Área Admin (usuários, acessos, empresas, prazo, intervalo, senha, WhatsApp/e-mail). Vira a raiz do Pages. |
| `painel/gerar_dashboard.py` | Gera o HTML de um painel (template gigante dentro do Python). Lê os xlsx da pasta de trabalho e o `config.json` do grupo. |
| `painel/buscar_api.py` | Busca a API do mapeia e monta `trabalho/<slug>/` (xlsx + config.json + gerador + logo) para cada grupo. |
| `painel/baixar_usuarios.py` | Robô Playwright que exporta os cadastrados do mapeia. |
| `painel/grupos.json` | **Lista de grupos**: slug, nome, prefixo de arquivo, empresas (pedaços do nome no mapeia, sem acento), prazo padrão, logo, fundo da logo. Grupo novo = nova linha aqui (+ logo em `painel/`). |
| `painel/logo_*.png` | Logos dos grupos. `logo_fundo` no grupos.json põe um fundo atrás da logo (logo branca/transparente). |
| `painel/pwa/` | Ícones (alfa da Sinergia) e manifesto do portal. |
| `.github/workflows/atualizar-painel.yml` | Robô dos painéis (ver fluxo acima). |
| `.github/workflows/atualizar-usuarios.yml` | Robô da lista de cadastrados. |
| `supabase/2026-11-23_filtros.sql` | Filtro do que cada usuário vê dentro do painel (tabelas `filtros` e `paineis_gerados`, função `pode_ler_painel`, regra do Storage). |
| `supabase/estrutura.sql` | Estrutura do banco (tabelas, regras, funções). Já foi rodada; serve de referência. Mudança no banco = SQL novo que o **usuário** roda no SQL Editor do Supabase. |

## Detalhes importantes

- **Gerador**: o HTML é uma string Python normal (não f-string). Marcadores em MAIÚSCULAS são
  trocados no fim, do maior para o menor; `GRUPO_NOME_PH`, `GRUPO_ARQ_PH` e `INTERVALO_MIN_PH`
  são trocados por último. O painel herdou da Novatec 3 "slots" (al/se/rec): **o grupo sempre
  ocupa o slot `se` / `DATA.sergipe`**; al e rec ficam vazios e escondidos.
- **Filtros do painel**: período/turno, setor (`window._setorFiltro`) e empresa do grupo
  (`window._empFiltro`) passam por `_calcularEmpresaFiltrada` e `_pessoaPassa`.
- **Painel dentro do portal**: o painel chama `parent.recarregarPainel()`,
  `parent.prazoDoPainel()` e `parent.atualizacaoDoPainel()` (prazo e intervalo vêm da tabela
  `empresas`, editados pelo admin). Fora do portal usa os valores gravados pelo robô.
- **Portal só recarrega quando há versão nova** no Storage e troca o iframe sem piscar.
- **Filtro por usuário (o que ele vê dentro do painel)**: no Admin, botão 👁 ao lado de cada
  empresa marcada. Sem linha em `filtros` = vê o painel completo (`<slug>/index.html`). Com
  linha = vê só `<slug>/u/<id do usuário>.html`, que o robô gera só com quem estiver em
  QUALQUER setor/cargo/nome marcado (`_filtrar` em `buscar_api.py`); o completo fica bloqueado
  pela regra `pode_ler_painel`. Filtro mudado = painel antigo bloqueado até o robô gerar o novo
  (`filtros.alterado_em` > `paineis_gerados.base_em`); o robô gera os pendentes em até 10 min.
  As opções (setores, cargos, nomes) ficam em `paineis/_privado/opcoes.json` (só admin lê).
  O log do Actions é público: não imprimir nomes nem filtros.
- **API do mapeia**: `https://ruwfvtgacnimpebelxfd.supabase.co/functions/v1/api-relatorio`,
  Bearer `MAPEIA_API_KEY`, **1 consulta de atividades a cada 5 min** (429 → o robô espera e
  repete). Não tem cadastro de usuários (por isso o robô Playwright).
- **Supabase do portal**: projeto `dcdbsiuveisxoypcblzd` (São Paulo). Tabelas `empresas`
  (slug, nome, ativo, prazo, intervalo_min, atualizacao_ativa), `perfis`, `acessos`;
  funções `eh_admin()`, `pode_ver(slug)`, `admin_definir_senha()`. A chave publishable no
  `portal/index.html` é pública (pode ficar). A chave secreta **nunca** vai para o código.
- **Segredos do GitHub** (Settings → Secrets → Actions): `MAPEIA_API_KEY`, `MAPEIA_EMAIL`,
  `MAPEIA_SENHA`, `SUPABASE_SERVICE_KEY`. Nunca pedir esses valores no chat; o usuário cola
  direto no GitHub.
- **cron-job.org** (conta do usuário): job "Paineis Sinergia" a cada 10 min →
  `.../actions/workflows/atualizar-painel.yml/dispatches`; job "Usuarios Sinergia" diário →
  `atualizar-usuarios.yml`. Usam um token fine-grained "cron-painel" (Actions read/write, só
  este repositório).
- Erro "No artifacts named github-pages" no deploy é passageiro (o workflow já tenta de novo).

## Como alterar e testar

1. Altere o arquivo (para o gerador, prefira trocas exatas e pequenas).
2. Teste localmente antes de publicar, **com dados inventados** (nunca dados reais no repo):
   monte uma pasta com `Relatorio_<Arq>_<data>.xlsx` (aba "Atividades", colunas como em
   `COLS_REL` de `buscar_api.py`), `Usuarios_<Arq>_<data>.xlsx` e `config.json`, rode
   `python gerar_dashboard.py` com `BROWSER=true` e abra o HTML (Playwright) em 1500 px e
   390 px; confira que não há erro de JavaScript.
3. Para o portal: `node -e` com `new Function(...)` em cada `<script>` para checar a sintaxe.
4. `git push` na branch `main`: o robô gera os painéis e republica o portal sozinho
   (~2 min). Acompanhe em https://github.com/sinerggia-dev/painel-mapeamento/actions.
5. Depois avise o usuário para apertar **Ctrl+F5** no portal (ou tocar no ↻ do painel).

## Gestão de Contratos (`contratos/`)

Sistema separado do mapeamento, publicado em `.../painel-mapeamento/contratos/`. Usa o **mesmo
Supabase e o mesmo login** do portal (tabela `perfis`; admin do portal = admin dos contratos).
Os **arquivos ficam no SharePoint**, nunca no Supabase nem no GitHub. **Ainda não instalado.**

```
navegador (contratos/index.html) ──► Supabase (tabelas ctr_*, regras RLS)
        │  envio em partes direto para o link de upload do SharePoint
        └──► Edge Function "ctr-arquivos" (guarda o segredo da Microsoft, confere o nível
             do usuário no banco, cria a pasta, abre o envio, registra, gera link de ver/baixar)
                    └──► SharePoint: site GestaoContratos (sinerggiacombr), biblioteca Documentos,
                         Contratos/<Empresa>/<Categoria>/<CTR-AAAA-NNNNN - Fornecedor>/
```

| Arquivo | Para que serve |
|---|---|
| `contratos/index.html` | Telas: painel, lista, contrato (dados, documentos obrigatórios, arquivos, histórico), novo, importar (planilha modelo), Admin (acessos, categorias, unidades, identidade da empresa, SharePoint). |
| `supabase/2026-11-23_contratos.sql` | Tabelas `ctr_*`, lista padrão de grupos e categorias, função `ctr_nivel`, gatilhos (número CTR, categoria válida, histórico automático, histórico e documentos imutáveis). |
| `supabase/functions/ctr-arquivos/index.ts` | Edge Function (ações iniciar, registrar, ver, baixar, testar). |

- **Organização igual para todo cliente (não depende do sistema dele)**: GRUPO (`ctr_grupos`, fixo,
  só muda por SQL) > CATEGORIA (`ctr_categorias`: `empresa` vazia = lista padrão da Sinergia, só muda
  por SQL; com `empresa` = criada pelo admin) > contrato. Cada categoria traz `aviso_dias`, `reajuste`
  e `documentos` obrigatórios (além do contrato assinado) e uma `dica`. Os tipos de documento estão em
  `TIPOS_DOC` (telas) e `CATEGORIAS`/`COM_VALIDADE` (Edge Function): **manter os dois iguais**.
- **Modos de visão (admin)**, no menu lateral, seletor **Ver como**:
  - **Sinergia (administração)** (`MODO_ADM = true`): sempre com a identidade da Sinergia, selo de admin e menu
    "Administração · só Sinergia". O seletor **Empresa em gestão** escolhe a empresa ou "Todas as empresas"
    (`EMP = '*'`: painel consolidado com quadro por empresa; cadastro, importação e configuração pedem uma empresa).
  - **Uma empresa** (`MODO_ADM = false`): a tela exatamente como a empresa vê (identidade dela, sem nada de
    administração, acesso completo). O seletor **Usuário da empresa** mostra só o que uma pessoa vê (filtra no
    navegador com a mesma regra de `ctr_nivel`); esse modo é só visualização: `db()` bloqueia gravações e `fn()`
    bloqueia envio de arquivo. Toda leitura/gravação passa por `db(tabela)`, não por `sb.from`.
  - `ehAdmin()` = admin real **e** em modo administração **e** sem usuário escolhido; `ehAdminReal()` = papel admin.
- **Identidade da empresa** (`ctr_config`): nome, cor, logo (imagem reduzida, guardada na própria tabela)
  e como a empresa chama as unidades (`termo_unidade`/`termo_unidades`, ex.: Filial/Filiais). Os usuários
  da empresa sempre veem a identidade dela; sem linha = Sinergia. Login e visão consolidada = Sinergia.
  Use `tU()`, `tUs()`, `tu()`, `tus()` em vez de escrever "Unidade" nas telas.
  Tela **Personalização** (admin): também cor do menu (`menu_estilo`: escuro, claro, marca), tema padrão
  (`tema`: auto, claro, escuro) e o **link de entrada da empresa** `.../contratos/?e=<slug>`, que mostra nome,
  logo e cor na tela de login (função pública `ctr_marca_publica`, só esses três dados). A logo vira o ícone da
  aba. **Meu perfil** (`#perfil`, clicando no próprio nome): tema só para a pessoa (localStorage `ctr-tema`,
  vale mais que o da empresa) e troca de senha.
- **Naturezas do ERP** (`ctr_naturezas`; a Millena Móveis usa o ERP **Xadrex**): o admin importa o cadastro de
  naturezas (Admin → Naturezas do ERP; "#" no começo = desativada, opção ligada por padrão) e revisa a
  sugestão automática (`sugerirCategoria`): quais têm contrato e a qual categoria cada uma se liga.
  Com `ctr_config.usa_natureza`, a empresa cadastra pela natureza (form, filtros, planilha modelo com a coluna
  "Natureza", painel por classe, pasta do SharePoint "702 - ALUGUEL DE IMOVEIS"); a categoria vem da
  natureza (gatilho no banco) e continua dando as regras e a visão consolidada. Mudou a categoria de uma
  natureza = os contratos dela acompanham. A planilha de naturezas do cliente **nunca** vai para o repositório.
- **Anexar documento**: tipos agrupados (O contrato / Documentos com validade / Outros). "Contrato atualizado
  (nova versão)", aditivo, renovação e distrato são **mudanças** (`ehMudanca`): pedem "o que mudou em relação ao
  anterior" (obrigatório), vigência e, se houver, novo valor/vencimento; viram uma linha em `ctr_alteracoes` ligada
  ao arquivo. A lista marca "versão vigente" e "versão anterior · guardada". Nenhum arquivo pode ser apagado.
  **O SharePoint só aparece para o admin da Sinergia** (`ehAdmin()`): nunca para a empresa, nem com acesso completo.
- **Tipo do contrato** (coluna e filtro na lista, selo na tela do contrato, coluna no Excel): `tipoC()` decide pela
  linha do tempo e pelos arquivos: Distrato (arquivo de distrato ou alteração "Distrato: ...") > Renovação
  (`tipo_contrato = 'renovacao'` no cadastro, alteração ou arquivo de renovação) > Aditivado (aditivo) > Novo.
  O filtro "Distrato" mostra também os encerrados. Planilha modelo: coluna opcional "Tipo (Novo ou Renovação)" no fim.
- **Documentos com validade** (certidões, AVCB, alvará, seguro...): `ctr_documentos.validade`; vencido
  ou faltando = "Documentos pendentes"; alerta com 30 dias (`AVISAR_DOC`).
- **Códigos do sistema do cliente** (`fornecedor_codigo`, `natureza_cliente`) são opcionais no
  contrato: servem para a futura conferência dos pagamentos. CPF/CNPJ (`fornecedor_doc`) é validado,
  inclusive o CNPJ com letras (desde jul/2026).
- **Importação**: planilha modelo gerada no navegador (abas Contratos, Exemplo, Instruções,
  Categorias, Unidades). Confere tudo antes; com qualquer erro nada é gravado. Admin e nível 3
  importam; unidade nova só o admin cria. **Nunca commitar planilhas de clientes.**
- **Permissão**: `ctr_acessos` por usuário e empresa: nível 1 ver, 2 ver e baixar, 3 editar, enviar
  e importar; `unidades` e `grupos` vazios = todos; `contratos` com itens = vê **só** esses (e não cadastra
  novos: `podeCriar()`).
- **Pessoas** (Admin → Acessos): busca por nome, e-mail, cargo, função ou telefone. Cadastro com cargo, função e
  WhatsApp (`ctr_contatos`), nível e o que vê. A senha é gerada (`gerarSenha`, formato xxxx-xxxx) e a tela
  "Enviar acesso" manda por WhatsApp (wa.me) ou e-mail (mailto) com o link `?e=<empresa>`. `trocar_senha` obriga
  a criar a senha no 1º acesso (depois `ctr_senha_trocada()`). "Gerar nova senha e enviar" usa
  `admin_definir_senha` (do portal) e liga `trocar_senha` de novo.
- **Busca rápida** no topo de todas as telas (contrato, fornecedor, CNPJ, número, natureza/categoria, unidade). Contrato confidencial: só admin, quem criou e
  `ctr_liberados`. A Edge Function usa a mesma regra (`ctr_nivel_contrato`).
- **Microsoft**: app "Sistema de Gestão de Contratos" no Entra, permissão de aplicativo
  `Sites.Selected` com acesso `write` só ao site GestaoContratos (concedido via Graph Explorer,
  POST `/sites/{id}/permissions`). Segredos da Edge Function: `MS_TENANT_ID`, `MS_CLIENT_ID`,
  `MS_CLIENT_SECRET`, `SP_SITE_ID` (e opcional `SP_PASTA_RAIZ`). O segredo do Entra **vence**:
  anotar a data e renovar antes.
- **Testes** (`tests/contratos/`, só dados inventados; ver o README de lá): `tests/contratos/rodar.sh /tmp/ctr-teste`
  roda todos os testes de tela (computador, Android, iPhone) e gera o protótipo em `/tmp/ctr-teste/gestao-contratos.html`
  (é o que vai para o claude.ai); `sudo tests/contratos/sql.sh` testa o banco (Postgres 16 local com stub do Supabase).
  **Rode os dois depois de qualquer mudança.** Logo e planilha real de cliente só na pasta de trabalho, nunca no repositório.
- **Reajuste e linha do tempo**: regra no contrato (`reajuste` = índice, `mes_reajuste`, `reajuste_teto`,
  `reajuste_negativo` mantém/reduz, `reajuste_percentual` para "Percentual fixo", `reajuste_substituto`).
  `proximoReajuste()` acha a data-base devida (12 meses depois do início, depois todo ano no mês do reajuste, no
  dia do início). `calcularReajuste()` usa os 12 meses de índice ANTES do mês da data-base (Edge Function, ação
  `indice`: Banco Central SGS 433 IPCA, 189 IGP-M, 188 INPC), aplica regra do negativo e teto e mostra o % com
  4 casas (o mesmo usado na conta). Convenção coletiva/Outro: % digitado. Registrar grava em `ctr_alteracoes`
  (imutável), e o gatilho atualiza valor/vencimento do contrato e o histórico. "Registrar aditivo": valor, prazo,
  renovação, com o documento ligado. Painel tem o cartão "Reajustes"; lista tem o filtro "Reajuste devido".
  **A consulta ao Banco Central não pôde ser testada daqui (rede bloqueada): conferir na instalação.**
- Ainda não feito: e-mail de aviso de vencimento, backup diário do banco para o SharePoint,
  conferência contrato × pagamento.
