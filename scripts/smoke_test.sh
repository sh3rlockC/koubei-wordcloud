#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TMP_ROOT="$(mktemp -d "${TMPDIR:-/tmp}/koubei-wordcloud-smoke.XXXXXX")"
trap 'rm -rf "$TMP_ROOT"' EXIT

SUMMARY_INPUT="$TMP_ROOT/fixtures/测试车_双平台口碑摘要.xlsx"
AUTOHOME_INPUT="$TMP_ROOT/fixtures/ZJ口碑_测试车.xlsx"
DCD_INPUT="$TMP_ROOT/fixtures/DCD口碑_测试车.xlsx"
OUT_BASE="$TMP_ROOT/output"

mkdir -p "$TMP_ROOT/fixtures" "$OUT_BASE"

python3 - "$SUMMARY_INPUT" "$AUTOHOME_INPUT" "$DCD_INPUT" <<'PY'
from pathlib import Path
import sys

import pandas as pd

summary_input = Path(sys.argv[1])
autohome_input = Path(sys.argv[2])
dcd_input = Path(sys.argv[3])

summary_sheets = {
    "汽车之家_满意摘要": pd.DataFrame([
        {"方向": "空间宽敞", "提及次数": 6, "代表性原句1": "后排空间宽敞，座椅舒服"},
        {"方向": "续航扎实", "提及次数": 4, "代表性原句1": "续航表现扎实，通勤省心"},
    ]),
    "汽车之家_不满意摘要": pd.DataFrame([
        {"方向": "胎噪明显", "提及次数": 5, "代表性原句1": "高速胎噪明显"},
        {"方向": "悬架偏硬", "提及次数": 3, "代表性原句1": "过减速带悬架偏硬"},
    ]),
    "懂车帝_正向摘要": pd.DataFrame([
        {"方向": "车机顺手", "提及次数": 5, "代表性原句1": "车机操作顺手，响应快"},
        {"方向": "外观好看", "提及次数": 4, "代表性原句1": "外观看着好看"},
    ]),
    "懂车帝_负向摘要": pd.DataFrame([
        {"方向": "充电偏慢", "提及次数": 4, "代表性原句1": "快充速度偏慢"},
        {"方向": "异味明显", "提及次数": 2, "代表性原句1": "新车异味明显"},
    ]),
}

with pd.ExcelWriter(summary_input, engine="openpyxl") as writer:
    for sheet_name, df in summary_sheets.items():
        df.to_excel(writer, index=False, sheet_name=sheet_name)

pd.DataFrame([
    {"最满意": "空间宽敞，续航扎实", "最不满意": "胎噪明显，悬架偏硬"},
    {"最满意": "座椅舒服，外观好看", "最不满意": "充电偏慢，异味明显"},
]).to_excel(autohome_input, index=False, sheet_name="口碑")

pd.DataFrame([
    {"优点": "车机顺手，外观好看", "缺点": "充电偏慢，胎噪明显"},
    {"优点": "空间宽敞，座椅舒服", "缺点": "悬架偏硬，异味明显"},
]).to_excel(dcd_input, index=False, sheet_name="口碑")
PY

echo "[1/2] summary input smoke test"
python3 "$SCRIPT_DIR/generate_wordcloud.py" \
  --input "$SUMMARY_INPUT" \
  --output-dir "$OUT_BASE/summary" \
  --model-name 测试车 \
  --mode compact \
  --json

echo "[2/2] raw fallback smoke test"
python3 "$SCRIPT_DIR/generate_wordcloud.py" \
  --autohome-input "$AUTOHOME_INPUT" \
  --dcd-input "$DCD_INPUT" \
  --output-dir "$OUT_BASE/raw" \
  --model-name 测试车 \
  --mode compact \
  --json

python3 - "$OUT_BASE" <<'PY'
from pathlib import Path
import sys

import pandas as pd

out_base = Path(sys.argv[1])
expected_sheets = ["summary", "positive_terms", "negative_terms", "overall_terms", "platform_breakdown"]

for case in ["summary", "raw"]:
    output_dir = out_base / case
    images = sorted(path.name for path in output_dir.glob("*.png"))
    expected_images = ["测试车_优点词云.png", "测试车_总体词云.png", "测试车_槽点词云.png"]
    if images != expected_images:
        raise SystemExit(f"{case} 词云图片不符合预期: {images}")

    excel_path = output_dir / "测试车_词云词项清单.xlsx"
    sheets = pd.ExcelFile(excel_path).sheet_names
    if sheets != expected_sheets:
        raise SystemExit(f"{case} Excel sheet 不符合预期: {sheets}")
PY

echo "Smoke test passed. Outputs under: $OUT_BASE"
