---
title: "受限拉丁文/希伯来文段落查询"
description: "受限拉丁文/希伯来文混排查询与候选写出器的方向、源区间和输出边界。"
sidebar:
  order: 120
  label: "拉丁文/希伯来文混排"
---

`HarfBuzzTextPrelayout.CreateMixedParagraph()` 将固定版本的 HarfBuzz 游程塑形器提供给 `TextMixedParagraphEngine`。独立的 `TextMixedParagraphRequest` 接受一个包含 3–128 个 BMP 源字符的 LTR 段落：ASCII 拉丁文单词、U+05D0–U+05EA 希伯来字母，以及单词之间的单个 ASCII 空格。它要求同时包含两种文字体系、一个由调用方提供且覆盖所有字符的不可变常规静态 TrueType 字面、水平排版，以及有限的正数字号和宽度。单词内文字体系变化、控制符、组合标记、数字、标点、双向格式控制符、增补字符、变体选择符、ZWJ 和其他文字体系都会明确失败。它不代表通用 Unicode 双向排版能力。

查询识别单词和空格的源区间，分别按 LTR 塑形拉丁文单词、按 RTL 塑形希伯来文单词，并使用绝对 UTF-16 源簇，再将得到的视觉字形序列放入 LTR 段落。相邻希伯来文单词及其内部空格组成一个视觉 RTL 短语：该短语内的单词位置反转，而结果的游程列表仍保持逻辑源顺序。查询只在完整单词之后选择断点，并重新塑形每个最终单词和空格。软断点保留其尾部空格在源区间和总步进中；可见适配计算排除该空格。最终结果拥有返回给调用方的实际塑形游程，包括所选字面、引擎版本、文字体系、方向、GID、偏移和步进。它绝不替换为宿主机字体，也不退回近似字符宽度。

查询限制字号、宽度、行数、塑形调用次数和累计字体复制工作量。缺字、过长单词、格式错误的塑形结果、预算耗尽和取消都会失败，且不返回部分结果。固定的 Ofdrw Test DejaVu Sans 2.37 测试样例来自上游发布归档，其许可证和字节哈希记录在 `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json`。E2E 检查源区间、递减的希伯来文 HarfBuzz 簇、合法软换行、覆盖失败、不支持的标点、预算和取消。

查询 API 返回按源文本寻址的塑形游程和查询相对坐标，不创建 OFD、PDF、SVG 或 PNG；写出需显式调用下述受限写出器。查询通过本身不表示 PDF 阅读器的复制和查找已经通过。[混排验证记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/text-mixed-paragraph-validation.md)保存了测试配置与历史阅读器结果。

## 受限混排输出

`TextMixedParagraphOfdWriter.CreateSinglePage` 将最终单词和空格写为独立的逻辑 `TextObject`，保留拉丁文和希伯来文原始 GID 映射。当前写出配置只接受一行、4 mm 字号和经过验证的 DejaVu Sans 测试字面（当前字体 SHA-256 见仓库 manifest；旧字节身份仍兼容）：一个拉丁文单词、空格、一到两个相邻希伯来文单词（以单个空格分隔）、空格、一个拉丁文单词。每个游程最多八个字符；希伯来文单词需要每个字母对应一个标称字形。其他字面、字号、换行、更多希伯来文单词或不受支持的源簇会在创建包前失败。查询 API 的准入范围比写出器更宽。

PDF/SVG 输出使用显式轮廓绘制希伯来文字形，并单独保留逻辑文本语义。实际复制、查找和视觉结果只按对应输出与阅读器验收记录判断。阿拉伯文、印度文多字形簇、通用 Unicode 断行和混合字面回退属于其他配置；详见[段落排版引擎](https://whynpc9.github.io/ofdrw.net/reference/text-paragraph-layout/)。
