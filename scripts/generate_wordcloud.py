#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from wordcloud_utils import (
    aggregate_terms,
    build_compact_groups,
    build_expanded_groups,
    detect_font,
    ensure_non_empty_groups,
    export_term_excel,
    extract_terms_from_summary_excel,
    load_autohome_raw_terms,
    load_dcd_raw_terms,
    load_stopwords,
    load_synonym_map,
    render_wordcloud,
)

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_STOPWORDS = BASE_DIR / "references" / "stopwords_zh.txt"
DEFAULT_SYNONYM_MAP = BASE_DIR / "references" / "synonym_map.example.yaml"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从口碑摘要 Excel 生成词云 PNG 和词项清单 Excel")
    parser.add_argument("--input", help="koubei-keyword-summary 输出的摘要 Excel")
    parser.add_argument("--autohome-input", help="汽车之家原始口碑 Excel")
    parser.add_argument("--dcd-input", help="懂车帝原始口碑 Excel")
    parser.add_argument("--output-dir", required=True, help="输出目录")
    parser.add_argument("--model-name", required=True, help="车型名，用于输出文件名和标题")
    parser.add_argument("--mode", choices=["compact", "expanded"], default="compact", help="compact 输出 2 张总图，expanded 输出分平台图")
    parser.add_argument("--top-n", type=int, default=80, help="每张词云最多展示多少个词")
    parser.add_argument("--min-weight", type=float, default=1.0, help="最低词权重")
    parser.add_argument("--stopwords", default=str(DEFAULT_STOPWORDS), help="停用词 txt 路径")
    parser.add_argument("--synonym-map", default=str(DEFAULT_SYNONYM_MAP), help="同义词映射 yaml 路径")
    parser.add_argument("--font-path", help="中文字体路径")
    parser.add_argument("--json", action="store_true", help="输出 JSON 结果")
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    if not args.input and not args.autohome_input and not args.dcd_input:
        raise SystemExit("请提供 --input，或至少提供 --autohome-input / --dcd-input 之一")
    if args.input and (args.autohome_input or args.dcd_input):
        raise SystemExit("--input 与 --autohome-input/--dcd-input 不能混用")
    if args.top_n <= 0:
        raise SystemExit("--top-n 必须大于 0")
    if args.min_weight < 0:
        raise SystemExit("--min-weight 不能小于 0")


def load_terms(args: argparse.Namespace, stopwords: set[str], synonym_map: dict[str, str]):
    if args.input:
        return extract_terms_from_summary_excel(args.input, stopwords, synonym_map), "summary", args.input

    terms = []
    input_parts = []
    if args.autohome_input:
        terms.extend(load_autohome_raw_terms(args.autohome_input, stopwords, synonym_map))
        input_parts.append(args.autohome_input)
    if args.dcd_input:
        terms.extend(load_dcd_raw_terms(args.dcd_input, stopwords, synonym_map))
        input_parts.append(args.dcd_input)
    if not terms:
        raise ValueError("原始 Excel 中未提取到可用词项")
    return terms, "raw_fallback", ", ".join(input_parts)


def main() -> int:
    args = parse_args()
    validate_args(args)

    stopwords = load_stopwords(args.stopwords)
    synonym_map = load_synonym_map(args.synonym_map)
    font_path = detect_font(args.font_path)

    raw_terms, input_type, input_path = load_terms(args, stopwords, synonym_map)
    merged_terms = aggregate_terms(raw_terms)

    if args.mode == "compact":
        groups = build_compact_groups(merged_terms, args.top_n, args.min_weight, args.model_name)
    else:
        groups = build_expanded_groups(merged_terms, args.top_n, args.min_weight, args.model_name)
    ensure_non_empty_groups(groups)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    image_paths = []
    for group in groups:
        image_path = output_dir / group["filename"]
        render_wordcloud(image_path, group["title"], group["frequencies"], font_path, group["direction"])
        image_paths.append(str(image_path))

    excel_path = output_dir / f"{args.model_name}_词云词项清单.xlsx"
    export_term_excel(
        excel_path,
        merged_terms,
        groups,
        meta={
            "model_name": args.model_name,
            "mode": args.mode,
            "input_type": input_type,
            "input_path": input_path,
            "top_n": args.top_n,
            "min_weight": args.min_weight,
            "font_path": font_path,
        },
    )

    result = {
        "model_name": args.model_name,
        "mode": args.mode,
        "input_type": input_type,
        "image_paths": image_paths,
        "excel_path": str(excel_path),
        "font_path": font_path,
        "group_count": len(groups),
        "term_count": len(merged_terms),
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"生成完成: {len(image_paths)} 张词云, 清单 {excel_path}")
        for image_path in image_paths:
            print(image_path)
        print(excel_path)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
