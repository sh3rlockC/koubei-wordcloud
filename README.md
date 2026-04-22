# koubei-wordcloud

把汽车口碑关键词结果做成能直接拿去汇报的词云图。

支持两种输入路径：
- `koubei-keyword-summary` 输出的摘要 Excel
- 汽车之家 / 懂车帝原始口碑 Excel（raw fallback）

输出产物：
- 词云 PNG
- 词项清单 Excel
- 可打包分发的 `.skill` 文件

## 适用场景

适合这些场景：
- 把口碑摘要继续可视化
- 快速做优点 / 槽点词云
- 做周报、汇报、一页纸素材
- 上游摘要还没产出时，直接从原始口碑走 fallback

## 功能

- 默认 `compact` 模式，输出 2 张总图
- 支持 `expanded` 模式，输出最多 4 张分平台图
- 支持基础停用词过滤
- 支持轻量同义词归并
- 支持 raw fallback，从原始口碑里直接抽词
- 导出 Excel 追溯词项来源

## 目录结构

```text
skills/koubei-wordcloud/
├── README.md
├── SKILL.md
├── references/
│   ├── stopwords_zh.txt
│   └── synonym_map.example.yaml
└── scripts/
    ├── generate_wordcloud.py
    └── wordcloud_utils.py
```

## 环境要求

- Python 3.10+
- `pandas`
- `openpyxl`
- `pyyaml`
- `wordcloud`
- `pillow`

## 快速开始

安装依赖：

```bash
pip install pandas openpyxl pyyaml wordcloud pillow
```

如果本机没有中文字体，需要准备一个 `.ttf` / `.ttc`，运行时通过 `--font-path` 指定。

## 快速验收

直接跑最小自测：

```bash
skills/koubei-wordcloud/scripts/smoke_test.sh
```

默认会验证两条链路：
- 摘要 Excel 输入
- raw fallback 输入

输出目录：
- `data/output/koubei-wordcloud-smoke/summary/`
- `data/output/koubei-wordcloud-smoke/raw/`

## 用法

### 1) 摘要 Excel 输入

sample command：

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input data/excel/启源A06/启源A06_双平台口碑摘要.xlsx \
  --output-dir data/output/koubei-wordcloud-demo-summary/ \
  --model-name 启源A06 \
  --mode compact
```


```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input data/output/koubei-keyword-summary/启源A06_双平台口碑摘要.xlsx \
  --output-dir data/output/koubei-wordcloud/ \
  --model-name 启源A06
```

### 2) raw fallback 输入

sample command：

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --autohome-input data/excel/启源A06/ZJ口碑_启源A06_2026-03-31.xlsx \
  --dcd-input data/excel/启源A06/DCD口碑_启源A06_2026-03-31.xlsx \
  --output-dir data/output/koubei-wordcloud-demo-raw/ \
  --model-name 启源A06 \
  --mode compact
```


```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --autohome-input data/raw/autohome.xlsx \
  --dcd-input data/raw/dcd.xlsx \
  --output-dir data/output/koubei-wordcloud-raw/ \
  --model-name 启源A06
```

### 3) expanded 分平台模式

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input data/output/koubei-keyword-summary/启源A06_双平台口碑摘要.xlsx \
  --output-dir data/output/koubei-wordcloud-expanded/ \
  --model-name 启源A06 \
  --mode expanded
```

## 参数

- `--input`：摘要 Excel
- `--autohome-input`：汽车之家原始 Excel
- `--dcd-input`：懂车帝原始 Excel
- `--output-dir`：输出目录，必填
- `--model-name`：车型名，必填
- `--mode`：`compact` 或 `expanded`
- `--top-n`：每张图最多展示多少词，默认 `80`
- `--min-weight`：最低词权重，默认 `1.0`
- `--stopwords`：停用词文件路径
- `--synonym-map`：同义词映射文件路径
- `--font-path`：中文字体路径
- `--json`：输出 JSON 结果

## 输出说明

### compact

输出 2 张图：
- `<车型名>_优点词云.png`
- `<车型名>_槽点词云.png`

### expanded

最多输出 4 张图：
- `<车型名>_汽车之家_优点词云.png`
- `<车型名>_汽车之家_槽点词云.png`
- `<车型名>_懂车帝_优点词云.png`
- `<车型名>_懂车帝_槽点词云.png`

### Excel 清单

额外输出：
- `<车型名>_词云词项清单.xlsx`

其中会包含：
- summary：本次生成元信息
- positive_terms：正向词项
- negative_terms：负向词项
- platform_breakdown：每张图实际采用的词项清单

## raw fallback 规则

### 汽车之家

优先读取：
- `最满意`
- `最不满意`

按列直接判定正负向，再拆分短语做聚合。

### 懂车帝

优先读取显式列：
- `最满意` / `优点` / `正向` / `正向标签`
- `最不满意` / `缺点` / `负向` / `负向标签`

如果没有显式列，会再从这些列中尝试轻量方向推断：
- `标题`
- `标签`
- `评价标签`
- `评价摘要`
- `评价内容`
- `评价全文`
- `内容`

注意，这一层是启发式规则，目标是稳定可用，不是学术级情感分类。

## 校验与报错

脚本会校验：
- 输入存在
- `--input` 不与 raw 参数混用
- 至少有正向和负向两组词
- 字体可用
- 图片和 Excel 成功写出

输入结构不对时会直接报错，不输出伪结果。

## 现状

已验证：
- `compact` 模式可成功生成 2 张图
- `expanded` 模式可成功生成 4 张图
- `smoke_test.sh` 可同时跑通 summary + raw fallback 两条链路

如果后面要继续做，可以往这几个方向扩：
- 更强的行业词典 / 同义词库
- 更稳的 raw 情感判定
- 更适合汇报的主题皮肤
- 批量车型处理

## Changelog & Releases

- User-visible changes are tracked in [`CHANGELOG.md`](./CHANGELOG.md).
- For a new release, update the `Unreleased` section first, then cut the versioned release.
- GitHub Release notes should match the same user-visible changes, not just raw commit history.

