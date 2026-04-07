#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/../../.." && pwd)"
cd "$ROOT"

SUMMARY_INPUT="data/excel/启源A06/启源A06_双平台口碑摘要.xlsx"
AUTOHOME_INPUT="data/excel/启源A06/ZJ口碑_启源A06_2026-03-31.xlsx"
DCD_INPUT="data/excel/启源A06/DCD口碑_启源A06_2026-03-31.xlsx"
OUT_BASE="data/output/koubei-wordcloud-smoke"

mkdir -p "$OUT_BASE"
rm -rf "$OUT_BASE/summary" "$OUT_BASE/raw"

echo "[1/2] summary input smoke test"
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input "$SUMMARY_INPUT" \
  --output-dir "$OUT_BASE/summary" \
  --model-name 启源A06 \
  --mode compact \
  --json

echo "[2/2] raw fallback smoke test"
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --autohome-input "$AUTOHOME_INPUT" \
  --dcd-input "$DCD_INPUT" \
  --output-dir "$OUT_BASE/raw" \
  --model-name 启源A06 \
  --mode compact \
  --json

echo "Smoke test passed. Outputs under: $OUT_BASE"
