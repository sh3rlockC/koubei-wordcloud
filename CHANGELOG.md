# Changelog

All notable changes to this project will be documented in this file.

The format is loosely based on Keep a Changelog and uses semantic-ish project version markers where available.

## [Unreleased]

### Added
- compact 模式现在会额外生成蓝色的 `<车型名>_总体词云.png`，与绿色优点词云、红色槽点词云一起输出。
- 词项清单 Excel 新增 `overall_terms` sheet，汇总正负向合并后的总体词项。

### Changed
- README 和 skill 说明同步更新 compact 默认输出数量与颜色约定。
- `smoke_test.sh` 改为临时生成最小 Excel fixture，支持在独立仓库 worktree 中直接验收。

## [v0.2.0]

### Notes
- Current released version after `v0.1.0`.
- See Git history and GitHub Releases for detailed changes to koubei-wordcloud.

## Conventions

- Add short, user-visible changes under `Unreleased` as work lands.
- When publishing a release, move the relevant entries into a versioned section.
- Keep GitHub Release notes aligned with the same entries so changelog and releases do not drift.
- Prefer describing behavior changes instead of internal-only commit trivia.
