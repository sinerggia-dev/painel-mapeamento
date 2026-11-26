#!/bin/bash
# Testes de tela da Gestão de Contratos: monta o protótipo (Supabase simulado, dados inventados)
# e roda todos os testes no Chromium (computador, Android e iPhone simulados).
#   tests/contratos/rodar.sh [pasta de trabalho]      (padrão: /tmp/ctr-teste)
# Opcional: <pasta>/logo.png (logo da empresa) e <pasta>/up/nat.xlsx (planilha real de naturezas).
# Nada disso vai para o repositório.
set -e
AQUI=$(cd "$(dirname "$0")" && pwd)
W=${1:-/tmp/ctr-teste}
mkdir -p "$W"
# biblioteca de Excel usada pelas telas (o CDN pode estar bloqueado na rede do teste)
if [ ! -f "$W/xlsxpkg/package/dist/xlsx.full.min.js" ]; then
  (cd "$W" && npm pack xlsx@0.18.5 >/dev/null 2>&1 && mkdir -p xlsxpkg && tar xzf xlsx-0.18.5.tgz -C xlsxpkg)
fi
# sem a logo real, usa uma logo de exemplo (quadradinho inventado), para os testes de logo rodarem
[ -f "$W/logo.png" ] || echo "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAYAAADED76LAAAAHElEQVQoU2NkYPj/n4EIwDiqEF8YjIYC/jAYDQUAbBMIAf3Yq5kAAAAASUVORK5CYII=" | base64 -d > "$W/logo.png"
python3 "$AQUI/build.py" "$W" >/dev/null
export NODE_PATH=$(npm root -g)
falhas=0
for t in teste teste_perso teste_novo teste_pessoas teste_reaj teste_versao teste_tipo teste_fornecedor; do
  echo "== $t"
  saida=$(timeout 400 node "$AQUI/$t.js" "$W" 2>&1) || true
  echo "$saida" | grep -v "^OK" | grep -v "^   " | tail -4
  echo "$saida" | grep -q "ERROS/FALHAS: nenhum" || falhas=$((falhas+1))
done
echo "== celular (larguras e zoom do iPhone)"
timeout 200 node "$AQUI/mob.js" "$W" 2>&1 | grep -A4 PROBLEMAS
echo
[ $falhas -eq 0 ] && echo "TUDO CERTO" || { echo "$falhas teste(s) com falha"; exit 1; }
