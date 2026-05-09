from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from wordcloud_utils import build_compact_groups, export_term_excel  # noqa: E402


class CompactWordcloudGroupTests(unittest.TestCase):
    def test_compact_groups_include_positive_negative_and_overall_with_direction_colors(self) -> None:
        rows = [
            {"term": "空间宽敞", "weight": 4.0, "direction": "positive", "platform": "autohome"},
            {"term": "续航扎实", "weight": 3.0, "direction": "positive", "platform": "dcd"},
            {"term": "胎噪明显", "weight": 2.5, "direction": "negative", "platform": "autohome"},
            {"term": "车机卡顿", "weight": 1.5, "direction": "negative", "platform": "dcd"},
        ]

        groups = build_compact_groups(rows, top_n=80, min_weight=1.0, model_name="测试车")

        self.assertEqual(["positive", "negative", "overall"], [group["group_key"] for group in groups])
        self.assertEqual("#2E8B57", groups[0]["color"])
        self.assertEqual("#C0392B", groups[1]["color"])
        self.assertEqual("#2563EB", groups[2]["color"])
        self.assertEqual("测试车_总体词云.png", groups[2]["filename"])
        self.assertEqual("测试车 总体词云", groups[2]["title"])
        self.assertEqual(
            {
                "空间宽敞": 4.0,
                "续航扎实": 3.0,
                "胎噪明显": 2.5,
                "车机卡顿": 1.5,
            },
            groups[2]["frequencies"],
        )

    def test_export_term_excel_includes_overall_terms_sheet(self) -> None:
        rows = [
            {
                "term": "空间宽敞",
                "weight": 4.0,
                "direction": "positive",
                "platform": "autohome",
                "source_type": "summary_direction",
                "source_sheet": "汽车之家_满意摘要",
                "source_column": "方向",
                "normalized_term": "空间宽敞",
                "merged_from": "空间宽敞",
            },
            {
                "term": "胎噪明显",
                "weight": 2.5,
                "direction": "negative",
                "platform": "autohome",
                "source_type": "summary_direction",
                "source_sheet": "汽车之家_不满意摘要",
                "source_column": "方向",
                "normalized_term": "胎噪明显",
                "merged_from": "胎噪明显",
            },
        ]
        groups = build_compact_groups(rows, top_n=80, min_weight=1.0, model_name="测试车")

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_path = Path(tmp_dir) / "terms.xlsx"
            export_term_excel(output_path, rows, groups, {"model_name": "测试车", "mode": "compact"})

            workbook = pd.ExcelFile(output_path)
            self.assertIn("overall_terms", workbook.sheet_names)

            overall_df = pd.read_excel(output_path, sheet_name="overall_terms")
            self.assertEqual(["空间宽敞", "胎噪明显"], overall_df["term"].tolist())
            self.assertEqual([4.0, 2.5], overall_df["weight"].tolist())


if __name__ == "__main__":
    unittest.main()
