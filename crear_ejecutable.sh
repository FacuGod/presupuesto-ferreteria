#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install -r requirements_presupuestos_ferreteria.txt
.venv/bin/python -m pip install pyinstaller

.venv/bin/python -m PyInstaller \
  --noconfirm \
  --clean \
  --onefile \
  --windowed \
  --name "PresupuestosFerreteria" \
  presupuestos_ferreteria_web.py

chmod +x "dist/PresupuestosFerreteria"

echo "Ejecutable creado en: $(pwd)/dist/PresupuestosFerreteria"
