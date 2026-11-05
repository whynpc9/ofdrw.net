---
title: "受限拉丁文/希伯来文段落查询"
description: "受限拉丁文/希伯来文混排查询与候选写出器的方向、源区间和输出边界。"
sidebar:
  order: 120
  label: "拉丁文/希伯来文混排"
---

`HarfBuzzTextPrelayout.CreateMixedParagraph()` 将固定版本的 HarfBuzz 游程塑形器提供给 `TextMixedParagraphEngine`。独立的 `TextMixedParagraphRequest` 接受一个包含 3–128 个 BMP 源字符的 LTR 段落：ASCII 拉丁文单词、U+05D0–U+05EA 希伯来字母，以及单词之间的单个 ASCII 空格。它要求同时包含两种文字体系、一个由调用方提供且覆盖所有字符的不可变常规静态 TrueType 字面、水平排版，以及有限的正数字号和宽度。单词内文字体系变化、控制符、组合标记、数字、标点、双向格式控制符、增补字符、变体选择符、ZWJ 和其他文字体系都会明确失败。它不代表通用 Unicode 双向排版能力。

查询识别单词和空格的源区间，分别按 LTR 塑形拉丁文单词、按 RTL 塑形希伯来文单词，并使用绝对 UTF-16 源簇，再将得到的视觉字形序列放入 LTR 段落。相邻希伯来文单词及其内部空格组成一个视觉 RTL 短语：该短语内的单词位置反转，而结果的游程列表仍保持逻辑源顺序。查询只在完整单词之后选择断点，并重新塑形每个最终单词和空格。软断点保留其尾部空格在源区间和总步进中；可见适配计算排除该空格。最终结果拥有返回给调用方的实际塑形游程，包括所选字面、引擎版本、文字体系、方向、GID、偏移和步进。它绝不替换为宿主机字体，也不退回近似字符宽度。

查询限制字号、宽度、行数、塑形调用次数和累计字体复制工作量。缺字、过长单词、格式错误的塑形结果、预算耗尽和取消都会失败，且不返回部分结果。固定的 R22 DejaVu Sans 2.37 测试样例来自上游发布归档，其许可证和字节哈希记录在 `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json`。E2E 检查源区间、递减的希伯来文 HarfBuzz 簇、合法软换行、覆盖失败、不支持的标点、预算和取消。

通过 PR #72 合并的查询 API 返回按源文本寻址的塑形游程和查询相对坐标，不创建 OFD、PDF、SVG 或 PNG。草案 PR #71 集成了独立的类型化写出器候选。仅凭查询成功，不能授权通过无关写出器进行序列化。[查询验证记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r22-mixed-query.md)跟踪已经合并的查询配置。

## 混排输出草案集成

`TextMixedParagraphOfdWriter.CreateSinglePage` 是有界的 B/D 输出候选。它为每个最终单词或空格写入独立的逻辑 TextObject：拉丁文和字母互不重复的希伯来文单词携带完整的原始 GID `CGTransform` 映射，标称空格则保留为源文本。希伯来文视觉顺序 GID 在 PDF/SVG 中采用显式轮廓路径，逻辑 Unicode 语义另行携带；不推断导入的多 TextCode 映射。写出器在返回包之前拒绝不支持的游程和页面墨迹。暂定的阅读器安全输出配置为单个不换行，且恰好包含五个游程：拉丁文单词、空格、一个希伯来文单词、空格、拉丁文单词。它使用固定的 R22 DejaVu 字面和 4 mm 字号。每个游程最多八个字符；希伯来文单词包含两到八个字母，每个字母对应一个不同的标称字形。相邻或重复的希伯来文单词、位于段落任一端点的希伯来文单词以及换行，都会在创建任何包之前返回带源区间的 `UnsupportedOutput`。查询 API 仍接受其更宽的有界排版配置。即使是暂定写出器模式，发布前也需要基于准确 HEAD 的新 Preview 和包证据。导入的希伯来文 `CGTransform` 区间同样必须声明每个逻辑字母恰好对应一个原始 GID；字形列表过短或过长都会在 PDF/SVG 输出前被拒绝。合法的非希伯来文组合区间继续遵循各自的规则。独立源码、排版、PDF 提取、单词位置和 Preview 门禁见[验证矩阵](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/r22-mixed-paragraph-validation.md)。这些陈述描述源码约定；实际输出行为、PDF 提取、独立阅读器和视觉验收需要准确的候选证据。阿拉伯文数字/标点、印度文多字形簇、段落 RTL、通用 Unicode 断行和混合字面回退仍是独立的 R22 门禁。旧版 Flow 和 R21 拉丁文/汉字查询保持既有行为。
