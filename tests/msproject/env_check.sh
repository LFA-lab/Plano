#!/usr/bin/env bash
set -euo pipefail
echo "== MS Project Env Check (Linux/Wine) =="
WINEPREFIX="${WINEPREFIX:-$HOME/.wine}"
echo "WINEPREFIX=$WINEPREFIX"

CANDIDATES=(
  "$WINEPREFIX/drive_c/Program Files/Microsoft Office/root/Office16/WINPROJ.EXE"
  "$WINEPREFIX/drive_c/Program Files (x86)/Microsoft Office/root/Office16/WINPROJ.EXE"
)
found=""

for p in "${CANDIDATES[@]}"; do
  if [[ -f "$p" ]]; then echo "Found WINPROJ: $p"; found="$p"; fi
done

if [[ -z "$found" ]]; then
  echo "Not found. Install MS Project into your Wine prefix."
  exit 1
fi

echo "Launching version check (window may appear)..."
wine "$found" /safe || true
echo "Done."
