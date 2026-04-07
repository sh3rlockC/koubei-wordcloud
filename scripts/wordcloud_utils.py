#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import yaml
from wordcloud import WordCloud

NOISE_TERMS = {
    "感觉", "觉得", "还是", "这个", "那个", "整体", "方面", "表现", "使用", "体验",
    "目前", "现在", "有点", "一些", "这种", "确实", "真的", "就是", "其实", "如果",
    "没有", "可以", "以及", "还有", "的话", "一下", "比较", "非常", "用户", "评价",
    "集中", "正面", "负面", "反馈", "提及", "摘要", "口碑", "汽车之家", "懂车帝",
    "车型", "这车", "这个车", "辆车", "风云", "启源", "plus", "PLUS", "A06", "x3",
}

SUMMARY_SHEET_RULES = [
    {"sheet": "汽车之家_满意摘要", "direction": "positive", "platform": "autohome"},
    {"sheet": "汽车之家_不满意摘要", "direction": "negative", "platform": "autohome"},
    {"sheet": "懂车帝_正向摘要", "direction": "positive", "platform": "dcd"},
    {"sheet": "懂车帝_负向摘要", "direction": "negative", "platform": "dcd"},
]

POSITIVE_HINTS = ["满意", "优点", "喜欢", "不错", "好评", "香", "给力", "舒服", "宽敞", "好看", "快", "省油", "顺手", "扎实", "稳定", "灵敏"]
NEGATIVE_HINTS = ["不满意", "缺点", "槽点", "不好", "拉跨", "一般", "差", "噪音", "异味", "不足", "没有", "问题", "偏贵", "偏硬", "卡顿", "费油", "虚标"]

AUTOHOME_POSITIVE_COLUMNS = ["最满意"]
AUTOHOME_NEGATIVE_COLUMNS = ["最不满意"]
DCD_DIRECTION_CANDIDATE_COLUMNS = [
    "最满意", "最不满意", "标题", "标签", "评价标签", "评价摘要", "评价内容", "评价全文", "内容",
]
DCD_POSITIVE_COLUMNS = ["最满意", "优点", "正向", "正向标签"]
DCD_NEGATIVE_COLUMNS = ["最不满意", "缺点", "负向", "负向标签"]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\u3000", " ").strip()
    return re.sub(r"\s+", " ", text)


def to_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        return float(value)
    text = clean_text(value)
    if not text:
        return default
    match = re.search(r"-?\d+(?:\.\d+)?", text.replace(",", ""))
    return float(match.group()) if match else default


def load_stopwords(path: str | Path | None = None) -> set[str]:
    words = set(NOISE_TERMS)
    if path:
        p = Path(path)
        if p.exists():
            for line in p.read_text(encoding="utf-8").splitlines():
                word = clean_text(line)
                if word and not word.startswith("#"):
                    words.add(word)
    return words


def load_synonym_map(path: str | Path | None = None) -> dict[str, str]:
    if not path:
        return {}
    p = Path(path)
    if not p.exists():
        return {}
    data = yaml.safe_load(p.read_text(encoding="utf-8")) or {}
    result = {}
    if isinstance(data, dict):
        for raw, norm in data.items():
            raw_text = clean_text(raw)
            norm_text = clean_text(norm)
            if raw_text and norm_text:
                result[raw_text] = norm_text
    return result


def normalize_term(term: Any, stopwords: set[str], synonym_map: dict[str, str]) -> str:
    text = clean_text(term)
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"^[\-•·\d\.:：、\s]+", "", text)
    text = re.sub(r"[，,。；;！!？?]+$", "", text)
    text = re.sub(r"^(用户对|主要是|主要在|主要短板集中在|核心卖点集中在|核心槽点集中在)", "", text)
    text = re.sub(r"(评价集中偏正面|评价集中偏负面|反馈偏正向|反馈偏负向|的负面反馈较集中|的正面反馈较集中)$", "", text)
    text = clean_text(text)
    if not text:
        return ""
    lower_map = {k.lower(): v for k, v in synonym_map.items()}
    if text.lower() in lower_map:
        text = lower_map[text.lower()]
    elif text in synonym_map:
        text = synonym_map[text]
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return ""
    if len(text) <= 1:
        return ""
    if text.lower() in {w.lower() for w in stopwords}:
        return ""
    return text


def infer_weight(row: pd.Series, preferred_columns: Iterable[str]) -> float:
    for col in preferred_columns:
        if col in row.index:
            value = to_float(row[col], 0.0)
            if value > 0:
                return value
    return 1.0


def extract_terms_from_summary_excel(input_path: str | Path, stopwords: set[str], synonym_map: dict[str, str]) -> list[dict[str, Any]]:
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"输入文件不存在: {path}")

    workbook = pd.ExcelFile(path)
    terms: list[dict[str, Any]] = []

    for rule in SUMMARY_SHEET_RULES:
        if rule["sheet"] not in workbook.sheet_names:
            continue
        df = pd.read_excel(path, sheet_name=rule["sheet"])
        if df.empty:
            continue
        if "方向" not in df.columns:
            raise ValueError(f"sheet {rule['sheet']} 缺少‘方向’列")

        for _, row in df.iterrows():
            base_weight = infer_weight(row, ["提及次数", "主方向口径", "条目口径", "语句归并口径"])
            direction_term = normalize_term(row.get("方向"), stopwords, synonym_map)
            merged_from: list[str] = []
            if direction_term:
                merged_from.append(clean_text(row.get("方向")))
                terms.append({
                    "term": direction_term,
                    "weight": base_weight,
                    "direction": rule["direction"],
                    "platform": rule["platform"],
                    "source_type": "summary_direction",
                    "source_sheet": rule["sheet"],
                    "source_column": "方向",
                    "normalized_term": direction_term,
                    "merged_from": merged_from.copy(),
                })

            quote_cols = [c for c in df.columns if str(c).startswith("代表性原句")]
            seen_quotes: set[str] = set()
            for quote_col in quote_cols:
                for frag in split_fragments(row.get(quote_col)):
                    raw_frag = clean_text(frag)
                    normalized = normalize_term(raw_frag, stopwords, synonym_map)
                    if not normalized:
                        continue
                    key = (normalized, quote_col)
                    if str(key) in seen_quotes:
                        continue
                    seen_quotes.add(str(key))
                    terms.append({
                        "term": normalized,
                        "weight": max(1.0, round(base_weight * 0.35, 2)),
                        "direction": rule["direction"],
                        "platform": rule["platform"],
                        "source_type": "quote",
                        "source_sheet": rule["sheet"],
                        "source_column": quote_col,
                        "normalized_term": normalized,
                        "merged_from": [raw_frag],
                    })

    if not terms:
        raise ValueError("未能从摘要 Excel 中识别到可用词项")
    return terms


def split_fragments(text: Any) -> list[str]:
    value = clean_text(text)
    if not value:
        return []
    parts = re.split(r"[，,、；;。\n\r]+", value)
    result = []
    for part in parts:
        item = clean_text(part)
        if item:
            result.append(item)
    return result


def load_autohome_raw_terms(path: str | Path, stopwords: set[str], synonym_map: dict[str, str]) -> list[dict[str, Any]]:
    workbook = pd.ExcelFile(path)
    terms: list[dict[str, Any]] = []
    seen_reasons = set()
    for sheet in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        available_cols = set(df.columns)
        if not (available_cols & set(AUTOHOME_POSITIVE_COLUMNS + AUTOHOME_NEGATIVE_COLUMNS)):
            continue
        for col, direction in [(name, "positive") for name in AUTOHOME_POSITIVE_COLUMNS] + [(name, "negative") for name in AUTOHOME_NEGATIVE_COLUMNS]:
            if col not in df.columns:
                continue
            for value in df[col].fillna(""):
                for frag in split_fragments(value):
                    normalized = normalize_term(frag, stopwords, synonym_map)
                    if not normalized:
                        continue
                    reason_key = (sheet, col, clean_text(frag))
                    if reason_key in seen_reasons:
                        continue
                    seen_reasons.add(reason_key)
                    terms.append({
                        "term": normalized,
                        "weight": 1.0,
                        "direction": direction,
                        "platform": "autohome",
                        "source_type": "raw_sentence",
                        "source_sheet": sheet,
                        "source_column": col,
                        "normalized_term": normalized,
                        "merged_from": [clean_text(frag)],
                    })
    if not terms:
        raise ValueError("汽车之家原始 Excel 未找到可用的‘最满意/最不满意’词项")
    return terms


def load_dcd_raw_terms(path: str | Path, stopwords: set[str], synonym_map: dict[str, str]) -> list[dict[str, Any]]:
    workbook = pd.ExcelFile(path)
    terms: list[dict[str, Any]] = []
    for sheet in workbook.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet)
        if df.empty:
            continue
        available_cols = set(df.columns)
        direction_cols = [col for col in DCD_DIRECTION_CANDIDATE_COLUMNS if col in available_cols]
        explicit_positive = [col for col in DCD_POSITIVE_COLUMNS if col in available_cols]
        explicit_negative = [col for col in DCD_NEGATIVE_COLUMNS if col in available_cols]
        if not direction_cols and not explicit_positive and not explicit_negative:
            continue

        for col in explicit_positive:
            terms.extend(extract_raw_terms_from_column(df, sheet, col, "positive", "dcd", stopwords, synonym_map))
        for col in explicit_negative:
            terms.extend(extract_raw_terms_from_column(df, sheet, col, "negative", "dcd", stopwords, synonym_map))

        for _, row in df.iterrows():
            row_direction = infer_direction_from_row(row)
            if not row_direction:
                continue
            for col in direction_cols:
                value = row.get(col)
                for frag in split_fragments(value):
                    normalized = normalize_term(frag, stopwords, synonym_map)
                    if not normalized:
                        continue
                    terms.append({
                        "term": normalized,
                        "weight": 1.0,
                        "direction": row_direction,
                        "platform": "dcd",
                        "source_type": "raw_sentence",
                        "source_sheet": sheet,
                        "source_column": col,
                        "normalized_term": normalized,
                        "merged_from": [clean_text(frag)],
                    })
    if not terms:
        raise ValueError("懂车帝原始 Excel 未识别到可用词项，请确认存在评价/标签/最满意/最不满意等列")
    return dedupe_raw_terms(terms)


def extract_raw_terms_from_column(
    df: pd.DataFrame,
    sheet: str,
    column: str,
    direction: str,
    platform: str,
    stopwords: set[str],
    synonym_map: dict[str, str],
) -> list[dict[str, Any]]:
    terms: list[dict[str, Any]] = []
    for value in df[column].fillna(""):
        for frag in split_fragments(value):
            normalized = normalize_term(frag, stopwords, synonym_map)
            if not normalized:
                continue
            terms.append({
                "term": normalized,
                "weight": 1.0,
                "direction": direction,
                "platform": platform,
                "source_type": "raw_sentence",
                "source_sheet": sheet,
                "source_column": column,
                "normalized_term": normalized,
                "merged_from": [clean_text(frag)],
            })
    return terms


def infer_direction_from_row(row: pd.Series) -> str | None:
    positive_text = " ".join(clean_text(row.get(col)) for col in DCD_POSITIVE_COLUMNS if col in row.index)
    negative_text = " ".join(clean_text(row.get(col)) for col in DCD_NEGATIVE_COLUMNS if col in row.index)
    if positive_text and not negative_text:
        return "positive"
    if negative_text and not positive_text:
        return "negative"

    combined = " ".join(clean_text(row.get(col)) for col in DCD_DIRECTION_CANDIDATE_COLUMNS if col in row.index)
    return infer_direction_from_text(combined)


def infer_direction_from_text(text: str) -> str | None:
    value = clean_text(text)
    if not value:
        return None
    pos_hits = sum(1 for hint in POSITIVE_HINTS if hint in value)
    neg_hits = sum(1 for hint in NEGATIVE_HINTS if hint in value)
    if neg_hits > pos_hits:
        return "negative"
    if pos_hits > neg_hits:
        return "positive"
    return None


def dedupe_raw_terms(terms: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen = set()
    result = []
    for item in terms:
        key = (
            item.get("platform"),
            item.get("direction"),
            item.get("normalized_term"),
            item.get("source_sheet"),
            item.get("source_column"),
            tuple(item.get("merged_from", [])),
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    return result


def aggregate_terms(terms: list[dict[str, Any]]) -> list[dict[str, Any]]:
    bucket: dict[tuple[str, str, str], dict[str, Any]] = {}
    for item in terms:
        key = (item["normalized_term"], item["direction"], item["platform"])
        if key not in bucket:
            bucket[key] = {
                "term": item["normalized_term"],
                "weight": 0.0,
                "direction": item["direction"],
                "platform": item["platform"],
                "source_types": set(),
                "source_sheets": set(),
                "source_columns": set(),
                "merged_from": [],
            }
        current = bucket[key]
        current["weight"] += float(item.get("weight", 1.0))
        current["source_types"].add(item.get("source_type", "unknown"))
        current["source_sheets"].add(item.get("source_sheet", ""))
        current["source_columns"].add(item.get("source_column", ""))
        for source in item.get("merged_from", []):
            source_text = clean_text(source)
            if source_text and source_text not in current["merged_from"]:
                current["merged_from"].append(source_text)

    rows = []
    for value in bucket.values():
        rows.append({
            "term": value["term"],
            "weight": round(value["weight"], 2),
            "direction": value["direction"],
            "platform": value["platform"],
            "source_type": ", ".join(sorted(value["source_types"])),
            "source_sheet": ", ".join(sorted(filter(None, value["source_sheets"]))),
            "source_column": ", ".join(sorted(filter(None, value["source_columns"]))),
            "normalized_term": value["term"],
            "merged_from": " | ".join(value["merged_from"][:10]),
        })
    rows.sort(key=lambda x: (x["direction"], x["platform"], -x["weight"], x["term"]))
    return rows


def build_compact_groups(rows: list[dict[str, Any]], top_n: int, min_weight: float, model_name: str) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for row in rows:
        grouped[row["direction"]][row["term"]] += float(row["weight"])
    result = []
    meta = {
        "positive": (f"{model_name}_优点词云.png", f"{model_name} 优点词云", "#2E8B57"),
        "negative": (f"{model_name}_槽点词云.png", f"{model_name} 槽点词云", "#C0392B"),
    }
    for direction in ["positive", "negative"]:
        freq = {term: weight for term, weight in grouped.get(direction, {}).items() if weight >= min_weight}
        freq = dict(sorted(freq.items(), key=lambda x: (-x[1], x[0]))[:top_n])
        if not freq:
            continue
        filename, title, color = meta[direction]
        result.append({
            "group_key": direction,
            "direction": direction,
            "platform": "combined",
            "title": title,
            "filename": filename,
            "color": color,
            "frequencies": freq,
        })
    return result


def build_expanded_groups(rows: list[dict[str, Any]], top_n: int, min_weight: float, model_name: str) -> list[dict[str, Any]]:
    platform_name = {"autohome": "汽车之家", "dcd": "懂车帝"}
    direction_name = {"positive": "优点", "negative": "槽点"}
    color_map = {"positive": "#2E8B57", "negative": "#C0392B"}
    grouped: dict[tuple[str, str], dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for row in rows:
        grouped[(row["platform"], row["direction"])][row["term"]] += float(row["weight"])
    result = []
    for (platform, direction), values in sorted(grouped.items()):
        freq = {term: weight for term, weight in values.items() if weight >= min_weight}
        freq = dict(sorted(freq.items(), key=lambda x: (-x[1], x[0]))[:top_n])
        if not freq:
            continue
        result.append({
            "group_key": f"{platform}_{direction}",
            "direction": direction,
            "platform": platform,
            "title": f"{model_name} {platform_name.get(platform, platform)} {direction_name.get(direction, direction)}词云",
            "filename": f"{model_name}_{platform_name.get(platform, platform)}_{direction_name.get(direction, direction)}词云.png",
            "color": color_map[direction],
            "frequencies": freq,
        })
    return result


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    value = hex_color.strip().lstrip("#")
    if len(value) != 6:
        raise ValueError(f"无效颜色值: {hex_color}")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def _mix_rgb(start_rgb: tuple[int, int, int], end_rgb: tuple[int, int, int], ratio: float) -> tuple[int, int, int]:
    ratio = max(0.0, min(1.0, ratio))
    return tuple(round(start_rgb[i] + (end_rgb[i] - start_rgb[i]) * ratio) for i in range(3))


def _build_frequency_color_func(frequencies: dict[str, float], direction: str):
    if not frequencies:
        return lambda *args, **kwargs: (0, 0, 0, 255)

    values = list(frequencies.values())
    min_weight = min(values)
    max_weight = max(values)
    if direction == "positive":
        start_rgb = _hex_to_rgb("#7FDBFF")
        end_rgb = _hex_to_rgb("#0B8F5A")
    elif direction == "negative":
        start_rgb = _hex_to_rgb("#FFD27D")
        end_rgb = _hex_to_rgb("#D62828")
    else:
        start_rgb = _hex_to_rgb("#D0D8E8")
        end_rgb = _hex_to_rgb("#2F4F6F")

    def color_func(word, *args, **kwargs):
        weight = frequencies.get(word, min_weight)
        if max_weight == min_weight:
            ratio = 1.0
        else:
            ratio = (weight - min_weight) / (max_weight - min_weight)
        rgb = _mix_rgb(start_rgb, end_rgb, ratio)
        alpha = round(255 * (0.42 + 0.5 * ratio))
        return rgb + (alpha,)

    return color_func


def detect_font(font_path: str | Path | None = None) -> str:
    if font_path:
        p = Path(font_path)
        if p.exists():
            return str(p)
        raise FileNotFoundError(f"字体文件不存在: {font_path}")

    search_dirs = [
        Path("/System/Library/Fonts"),
        Path("/Library/Fonts"),
        Path.home() / "Library/Fonts",
        Path("/Windows/Fonts"),
        Path("/mnt/c/Windows/Fonts"),
    ]
    preferred = [
        "Microsoft YaHei UI.ttf",
        "Microsoft YaHei.ttf",
        "msyh.ttc",
        "msyh.ttf",
        "msyhbd.ttf",
        "微软雅黑.ttf",
        "PingFang.ttc",
        "PingFang SC.ttc",
        "STHeiti Light.ttc",
        "STHeiti Medium.ttc",
        "Songti.ttc",
        "Hiragino Sans GB.ttc",
        "Arial Unicode.ttf",
    ]
    candidates = []
    for directory in search_dirs:
        if not directory.exists():
            continue
        for candidate in preferred:
            for path in directory.rglob(candidate):
                if path.exists():
                    candidates.append(path)
    if candidates:
        # 优先命中微软雅黑；如果本机没有，再退回到其他可用中文字体。
        for path in candidates:
            if re.search(r"microsoft yahei|msyh|微软雅黑", path.name, re.IGNORECASE):
                return str(path)
        return str(candidates[0])
    raise FileNotFoundError("未找到可用中文字体，优先需要微软雅黑；请通过 --font-path 显式指定字体文件")


def render_wordcloud(output_path: str | Path, title: str, frequencies: dict[str, float], font_path: str, direction: str) -> None:
    if not frequencies:
        raise ValueError(f"词云无可用词项: {title}")
    color_func = _build_frequency_color_func(frequencies, direction)
    wc = WordCloud(
        width=1400,
        height=900,
        background_color="white",
        font_path=font_path,
        prefer_horizontal=0.9,
        collocations=False,
        max_words=max(len(frequencies), 20),
        margin=8,
        color_func=color_func,
    )
    image = wc.generate_from_frequencies(frequencies).to_image()

    from PIL import Image, ImageDraw, ImageFont

    title_font = ImageFont.truetype(font_path, 42)
    canvas = Image.new("RGB", (1400, 980), "white")
    draw = ImageDraw.Draw(canvas)
    draw.text((50, 22), title, fill="#1F2937", font=title_font)
    canvas.paste(image, (0, 80))
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path)


def export_term_excel(path: str | Path, all_terms: list[dict[str, Any]], groups: list[dict[str, Any]], meta: dict[str, Any]) -> None:
    summary_df = pd.DataFrame([
        {"字段": "model_name", "值": meta.get("model_name", "")},
        {"字段": "mode", "值": meta.get("mode", "")},
        {"字段": "input_type", "值": meta.get("input_type", "")},
        {"字段": "input_path", "值": meta.get("input_path", "")},
        {"字段": "top_n", "值": meta.get("top_n", "")},
        {"字段": "min_weight", "值": meta.get("min_weight", "")},
        {"字段": "font_path", "值": meta.get("font_path", "")},
        {"字段": "group_count", "值": len(groups)},
        {"字段": "term_count", "值": len(all_terms)},
    ])
    terms_df = pd.DataFrame(all_terms)
    positive_df = terms_df[terms_df["direction"] == "positive"].sort_values(["weight", "platform"], ascending=[False, True]) if not terms_df.empty else pd.DataFrame()
    negative_df = terms_df[terms_df["direction"] == "negative"].sort_values(["weight", "platform"], ascending=[False, True]) if not terms_df.empty else pd.DataFrame()
    breakdown_rows = []
    for group in groups:
        for term, weight in group["frequencies"].items():
            breakdown_rows.append({
                "group_key": group["group_key"],
                "title": group["title"],
                "platform": group["platform"],
                "direction": group["direction"],
                "term": term,
                "weight": weight,
                "filename": group["filename"],
            })
    breakdown_df = pd.DataFrame(breakdown_rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        positive_df.to_excel(writer, index=False, sheet_name="positive_terms")
        negative_df.to_excel(writer, index=False, sheet_name="negative_terms")
        if not breakdown_df.empty:
            breakdown_df.to_excel(writer, index=False, sheet_name="platform_breakdown")


def ensure_non_empty_groups(groups: list[dict[str, Any]]) -> None:
    directions = {group["direction"] for group in groups}
    if "positive" not in directions or "negative" not in directions:
        raise ValueError("至少需要识别出一组正向词项和一组负向词项")
