#!/bin/bash
# Testes do banco da Gestão de Contratos num Postgres 16 local, com um "stub" do Supabase
# (auth.uid() lido de test.uid, papéis anon/authenticated/service_role) e dados inventados.
#   sudo tests/contratos/sql.sh
# Instala o SQL duas vezes (tem que poder rodar de novo) e roda os testes de permissão (sql/t.sql).
set -e
AQUI=$(cd "$(dirname "$0")" && pwd); RAIZ=$(cd "$AQUI/../.." && pwd)
D=/var/lib/postgresql/ctrtest; PGB=$(ls -d /usr/lib/postgresql/*/bin | tail -1)
mkdir -p $D && chown postgres $D
if ! psql -h $D -p 5499 -U postgres -tAc "select 1" >/dev/null 2>&1; then
  su postgres -c "$PGB/pg_ctl -D $D/data stop -m immediate" >/dev/null 2>&1 || true
  [ -d $D/data ] && mv $D/data $D/data.old.$$
  su postgres -c "$PGB/initdb -D $D/data -U postgres -A trust" >/dev/null
  su postgres -c "$PGB/pg_ctl -D $D/data -o '-p 5499 -k $D' -l $D/log start" >/dev/null; sleep 2
fi
P="psql -h $D -p 5499 -U postgres -v ON_ERROR_STOP=1 -q"
$P -c "drop database if exists t" -c "create database t"
$P -d t -f "$AQUI/sql/stub.sql"
$P -d t -f "$AQUI/sql/base.sql"
$P -d t -f "$AQUI/sql/dados.sql"
$P -d t -f "$RAIZ/supabase/2026-11-23_contratos.sql" 2>&1 | grep -v NOTICE || true
$P -d t -f "$RAIZ/supabase/2026-11-23_contratos.sql" 2>&1 | grep -v NOTICE || true
echo "SQL instalado duas vezes sem erro."
echo "== testes de permissão (as linhas 'esperado ERRO' devem vir seguidas de ERROR)"
psql -h $D -p 5499 -U postgres -d t -q -f "$AQUI/sql/t.sql" 2>&1 | grep -E "esperado|ERROR|---" | cut -c1-140
echo "== partes novas: identidade, naturezas, linha do tempo, contatos (banco limpo)"
$P -c "drop database if exists t" -c "create database t"
$P -d t -f "$AQUI/sql/stub.sql"; $P -d t -f "$AQUI/sql/base.sql"; $P -d t -f "$AQUI/sql/dados.sql"
$P -d t -f "$RAIZ/supabase/2026-11-23_contratos.sql" 2>&1 | grep -v NOTICE || true
psql -h $D -p 5499 -U postgres -d t -q -f "$AQUI/sql/t2.sql" 2>&1 | grep -v "^$" | cut -c1-140
