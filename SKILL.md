---
name: koubei-wordcloud
description: 从 `koubei-keyword-summary` 的摘要 Excel 或原始口碑 Excel 生成口碑词云图片，并导出词项清单 Excel。适用于用户希望把口碑关键词结果继续可视化、输出优点/槽点词云、做汇报素材时使用。默认输出 2 张总图（优点词云 / 槽点词云），也支持 expanded 分平台模式与 raw fallback；默认优先使用微软雅黑字体，词云颜色按词频做深浅渐变。
---

# 口碑词云生成

目标：承接 `koubei-keyword-summary` 的输出结果，生成适合汇报使用的词云 PNG 图片，并同时导出词项清单 Excel 便于追溯。

## 1. 当前 MVP 范围

当前可用范围：
- 输入：
  - `koubei-keyword-summary` 输出的摘要 Excel
  - 或汽车之家 / 懂车帝原始口碑 Excel（raw fallback）
- 模式：
  - 默认 `compact`，输出 2 张总图
  - `expanded`，输出最多 4 张分平台图
- 输出：
  - 词云 PNG
  - 词项清单 Excel

视觉规则：
- 默认优先使用微软雅黑字体；如果本机未安装，可用 `--font-path` 显式指定字体文件
- 正向词云使用蓝绿渐变，负向词云使用橙红渐变
- 同一张词云里，词频越高颜色越实、透明度越高；词频越低颜色越淡、透明度越低

当前仍未覆盖：
- 复杂主题皮肤
- 大规模行业词典
- 更细粒度情感分类模型

## 2. 支持的输入

### 方案 A：摘要 Excel（推荐）

优先输入一个已经由 `koubei-keyword-summary` 导出的摘要 Excel。

默认会尝试读取这些 sheet：
- `汽车之家_满意摘要`
- `汽车之家_不满意摘要`
- `懂车帝_正向摘要`
- `懂车帝_负向摘要`

会优先复用方向词与提及频次，必要时从 `代表性原句*` 列补词。

### 方案 B：raw fallback

如果上游摘要 Excel 还没产出，也可以直接喂原始口碑 Excel：
- 汽车之家：识别 `最满意` / `最不满意`
- 懂车帝：优先识别显式正负向列；若没有，再从 `标题 / 标签 / 评价摘要 / 评价全文` 等列里做轻量方向推断

注意：raw fallback 追求“能稳定出业务可读图”，不是精确情感 NLP。

## 3. 默认输出

默认输出到指定目录：
- `<车型名>_优点词云.png`
- `<车型名>_槽点词云.png`
- `<车型名>_词云词项清单.xlsx`

## 4. 推荐执行方式

### 用摘要 Excel 生成

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input /path/to/启源A06_双平台口碑摘要.xlsx \
  --output-dir /path/to/output \
  --model-name 启源A06
```

### 用 raw fallback 生成

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --autohome-input /path/to/autohome_raw.xlsx \
  --dcd-input /path/to/dcd_raw.xlsx \
  --output-dir /path/to/output \
  --model-name 启源A06
```

### expanded 分平台模式

```bash
python3 skills/koubei-wordcloud/scripts/generate_wordcloud.py \
  --input /path/to/启源A06_双平台口碑摘要.xlsx \
  --output-dir /path/to/output \
  --model-name 启源A06 \
  --mode expanded
```

可选参数：
- `--top-n` 词云最多展示多少个词
- `--min-weight` 最低词权重
- `--stopwords` 自定义停用词文件
- `--synonym-map` 自定义同义词映射
- `--font-path` 指定中文字体；默认优先微软雅黑
- `--json` 输出结构化结果

## 5. 输出原则

- 优先复用上游已经整理好的关键词 / 方向词
- 必要时从代表性原句中补词
- 默认做基础清洗：停用词、纯数字、常见噪声词过滤
- 默认做轻量同义归并
- 不伪装成精确 NLP 模型结果，优先保证业务可读性

## 6. 校验要求

生成前至少检查：
- 输入文件存在
- `--input` 与 raw 输入参数不混用
- 至少识别出一组正向词项和一组负向词项
- 字体文件可用
- PNG 成功生成
- 词项清单 Excel 成功导出

如果输入结构不符合预期，直接报错，不输出伪结果。

## 7. 交付物说明

默认会产出：
- compact：2 张图（优点 / 槽点）
- expanded：最多 4 张图（汽车之家优点/槽点、懂车帝优点/槽点）
- 1 份 `词云词项清单.xlsx`

清单里会记录：
- 词项
- 权重
- 平台
- 正负方向
- 来源 sheet / 列
- 归并前原词片段
