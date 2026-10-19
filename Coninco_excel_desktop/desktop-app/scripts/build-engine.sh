#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/../.." && pwd)"
VENV_DIR="$ROOT_DIR/python-engine/.venv"
OUTPUT_DIR="$ROOT_DIR/desktop-app/src-tauri/engine"

"$VENV_DIR/bin/pip" install -r "$ROOT_DIR/python-engine/requirements.txt"
rm -f "$OUTPUT_DIR/excel-auditor-engine" "$OUTPUT_DIR/excel-auditor-engine.exe"
"$VENV_DIR/bin/pyinstaller" \
  --noconfirm \
  --clean \
  --onefile \
  --name excel-auditor-engine \
  --paths "$ROOT_DIR/python-engine/src" \
  --collect-all pandas \
  --collect-all openpyxl \
  --collect-all xlrd \
  --distpath "$OUTPUT_DIR" \
  --workpath "$ROOT_DIR/target/pyinstaller-work" \
  --specpath "$ROOT_DIR/target" \
  "$ROOT_DIR/python-engine/engine_entry.py"

chmod +x "$OUTPUT_DIR/excel-auditor-engine"
echo "Built $OUTPUT_DIR/excel-auditor-engine"
