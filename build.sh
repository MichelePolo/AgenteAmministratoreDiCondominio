#!/usr/bin/env bash
# Sincronizza il tutorial nel plugin, valida il manifest, compila gli script, esegue i test e produce
# amministratore-condominio.plugin (zip del plugin) da allegare a una release.
set -euo pipefail
cd "$(dirname "$0")"
P=plugins/amministratore-condominio
# Interprete con openpyxl: .venv del repo se presente (python3 -m venv .venv && .venv/bin/pip install openpyxl)
PY=python3
[ -x .venv/bin/python ] && PY=.venv/bin/python
cp TUTORIAL.md "$P/TUTORIAL.md"
find "$P" -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
"$PY" -m py_compile "$P"/skills/*/scripts/*.py
"$PY" -m unittest discover -s tests
find "$P" tests -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
if command -v claude >/dev/null 2>&1; then
  claude plugin validate "$P/.claude-plugin/plugin.json"
fi
rm -f amministratore-condominio.plugin
(cd "$P" && zip -qr ../../amministratore-condominio.plugin . -x '*.DS_Store' -x '*__pycache__*')
echo "creato amministratore-condominio.plugin ($(du -h amministratore-condominio.plugin | cut -f1))"
