---
title: "段落排版引擎（R22）"
description: "Unicode 17.0.0 双向、断行与字素簇，字体回退、逐行重塑形、对齐/两端对齐、分页，以及直接消费最终排版结果的 OFD 写出器。"
sidebar:
  order: 130
  label: "段落排版引擎"
---

> 自 [0.2.0](https://whynpc9.github.io/ofdrw.net/releases/020/) 起可用。

`Ofdrw.Net.Layout.Text.Paragraphs` 将带样式的多段落文本排版为最终行和页面，`ParagraphLayoutOfdWriter` 则将该结果原样写为 OFD。这是一个新入口：`FlowDocument`、`PlainTextOfdConverter`、DOCX 转换和早期 R21/R22 查询均保持既有行为，不会隐式迁移。

```csharp
var noto = new TextFontFace("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"));
var wenkai = new TextFontFace("LXGW WenKai", File.ReadAllBytes("LXGWWenKai-Regular.ttf"));
var dejavu = new TextFontFace("DejaVu Sans", File.ReadAllBytes("DejaVuSans.ttf"));
var style = new TextStyle(new[]
{
    new TextFontFamily("Noto Sans", noto),
    new TextFontFamily("LXGW WenKai", wenkai),
    new TextFontFamily("DejaVu Sans", dejavu),
}, fontSizeMillimeters: 4.2);
var request = new ParagraphLayoutRequest("排版 layout שלום مرحبا\n第二段。", style,
    new ParagraphLayoutOptions(widthMillimeters: 90, pageHeightMillimeters: 250,
        alignment: TextAlignment.Justify, lineHeight: TextLineHeight.Multiple(1.2)));
var layout = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(request);
var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 110, 280, 10, 15);
await new OfdPackageWriter().WriteAsync(package, output);
```

## 处理流程

1. **准入检查。** 不成对的代理项、制表符、除行/段分隔符外的 C0/C1 控制符、显式双向格式字符（LRE…PDI、LRM/RLM/ALM）、其他格式控制符（ZWJ、ZWNJ、软连字符、单词连接符、BOM）以及变体选择符，都会在任何塑形之前连同源区间一起被拒绝。样式区间必须完整拼接覆盖文本，且不得拆分字素簇。
2. **段落划分。** 按此处准入的 UAX #9 段落分隔符拆分：LF、CR、CRLF、NEL、U+2029。VT、FF 和 U+2028 是段落内的硬换行。终止符属于行的源区间（`TerminatorLength`），但没有字形。
3. **Unicode 分析。** 使用冻结的 Unicode 17.0.0 数据表：UAX #9 rev 51（P2–P3、X1–X10、W1–W7、N0–N2、I1–I2、L1–L2）、UAX #14 rev 55 默认断行和 UAX #29 rev 47 默认扩展字素簇。不进行定制、词典分词或断字。参见[一致性验证](#unicode-一致性验证)。
4. **项目划分。** 按解析后的双向级别、文字体系（通用/继承字符采用相邻文字体系）、所选字面和样式划分。每个字素簇采用样式回退链中首个满足条件的字体家族：其字面在请求的粗体/斜体声明下覆盖该簇的所有标量。通用文字体系的簇在旧字面仍可覆盖时保持使用该字面。不发现宿主机字体，也不合成粗体/斜体：缺少样式对应字面会以 `UnsupportedFontProfile` 失败，缺少覆盖则以 `MissingGlyph` 失败。
5. **塑形。** 使用 `ITextShaper`（默认 HarfBuzz）对每个项目进行测量塑形。
6. **断行。** 在同时属于字素簇边界的 UAX #14 断行机会上采用首次适配贪心算法。尾部空格悬挂。随后重建所选行：被行边缘或 L1 级别变化切开的游程会**重新塑形**并检查实际宽度；若溢出，则尝试前一个断行机会（最多 32 次，随后 `BudgetExceeded`）。由于估算将跨断点字距调整计入左侧，在信任估算结果之前，还会重建最多三个“估算超出版心不超过半个 em”的后续合法断点。没有断行机会的片段会以 `LayoutOverflow` 失败（默认 `TextOverflowPolicy.Reject`），或在字素簇边界断开（`BreakGraphemes`）；绝不拆分簇内部。候选切点都会重建并重新塑形。重建后的宽度不具单调性（阿拉伯文 lam-alef 连字可能比单独的 lam 更窄），因此二分查找后还会进行有界向上扫描，并在连续四个切点溢出后停止。所选切点是已找到的最大适配切点，不保证全局最大。
7. **重排序。** 每行先应用 L1，再对游程应用 L2。游程内部保留塑形器的视觉字形顺序。
8. **度量与对齐。** 所有游程共享一条基线。正常行高使用各字面在相应字号下的上升部、下降部和行间距；也可使用 `Multiple` 和 `Exact`。`Start`/`End` 随段落方向变化；首行缩进位于起始侧。`Justify` 在词间空格以及汉字、假名、谚文或注音符号相邻字素之间拉伸非末行（闭标点前不拉伸）；额外步进按字形记录。
9. **分页。** 按 `PageHeightMillimeters` 划分页面（或使用一个高度不限的页面）。分页处丢弃段间距；高于页面的行会以 `LayoutOverflow` 失败。

## 最终结果约定

`ParagraphLayoutResult` 不可变且完整；失败或取消时不返回部分结果。坐标为相对于各页面内容框的毫米值，Y 向下，只从设计单位转换一次。

| 层级 | 包含内容 |
| --- | --- |
| 结果 | 请求（不可变样式）、页面、`UnicodeVersion` = 17.0.0、塑形器 `EngineVersion`、塑形调用次数 |
| 页面 | 索引、行、已用高度 |
| 行 | 绝对 UTF-16 源区间（含尾部空格和终止符）、段落/页面索引、基础级别、硬换行标记、顶部、基线、高度、上升部、下降部、对齐后的内容左侧位置/宽度、墨迹边界、按视觉顺序排列的游程 |
| 游程 | 逻辑源文本和起点、精确的 `TextShapedRun`（字面字节身份、文字体系、语言、方向、特性）、解析后的双向级别、`TextResolvedStyle`、原点、基线、步进、墨迹 |
| 字形 | 原始 GID、绝对簇起点/长度（多对多）、原点 X/Y、含两端对齐的步进、两端对齐增量、墨迹 |

`TextResolvedStyle` 快照保存画刷（DeviceRGB 填充/描边颜色及 alpha、描边宽度）、字号、语言、请求与字面各自的粗体/斜体状态、字体家族及其回退索引，以及 `Synthetic = None`。写出器不再查询请求样式，不重新选择字体，也不重新塑形。

## OFD 输出

`ParagraphLayoutOfdWriter.CreateDocument(result, pageWidth, pageHeight, left, top)` 为每个排版页面创建一个 OFD 页面，并为每个游程创建一个 `TextObject`：

- `TextCode` 保存游程未经修改的逻辑文本；
- 每个源簇对应一个按逻辑顺序排列的 `CGTransform`（`CodePosition`/`CodeCount` 以 BMP 字符计数），列出原始 GID；
- 字形槽位遵循该顺序，每个槽位都有显式原点（完整的 `DeltaX`/`DeltaY`），因此 RTL 游程使用负增量，而不是重排文本；
- 每个不同字面对应一个字体资源，以保留 GID 的子集嵌入（或通过 `embedFullFonts: true` 完整嵌入）；遵守 OS/2 嵌入限制；
- 写入填充颜色和 alpha；粗体/斜体由游程字面提供。

以下情况会在创建包之前以 `UnsupportedOutput` 拒绝：描边文字或无填充、增补（非 BMP）字符、文本框或墨迹超出页面，以及源文本/GID 身份丢失。

模型将此表示为 `OfdPositionedGlyphMapping.ClusterMap`。写出器在 `https://ofdrw.net/layout-hints` 命名空间中，用写出器自有的提示属性 `ClusterMapV1="true"` 标记此类 TextObject（其他阅读器会忽略它）。读取时，历史有界 CGTransform 配置仍优先；只有被这些配置拒绝的已标记 TextObject，才会按完整簇映射接纳：一个 TextCode、每个代码单元恰好属于一个递增 CGTransform、每个字形槽位都有原点。未标记的外来 CGTransform 继续保持历史失败关闭规则。PDF 和 SVG 根据接纳映射中每个字形的原始轮廓进行绘制，并将逻辑文本保留为不可见 PDF 文本或 SVG `title`/`aria-label`。

## Unicode 一致性验证

数据表由 `scripts/unicode/generate_unicode_tables.py` 根据 UCD 17.0.0 文件生成，这些文件的 SHA-256 固定在 `scripts/unicode/ucd-17.0.0-manifest.json` 中（Unicode License v3）。完整官方测试在 `UnicodeConformanceTests` 中运行：

| 文件 | 用例数 | 失败数 |
| --- | ---: | ---: |
| BidiTest.txt | 770,241 | 0 |
| BidiCharacterTest.txt | 91,707 | 0 |
| LineBreakTest.txt | 19,338 | 0 |
| GraphemeBreakTest.txt | 766 | 0 |

`UnicodeTextAnalysis` 公开相同算法（字素边界、断行机会、双向级别和视觉顺序）。双向 API 不应用 L3/L4；镜像由塑形器处理。

## 限制与不承诺的范围

- 预算：`MaximumTextLength`（262,144 个 UTF-16 代码单元）、`MaximumLines`（10,000）、`MaximumShapeCalls`（100,000）、每个塑形游程 4,096 个 UTF-16 代码单元、坐标范围 ±1,000,000 mm。
- 已验证的输出文字体系以验证报告为准：使用所列字体的拉丁文、汉字、希伯来文、阿拉伯文和泰米尔文。其他文字体系经过相同流程排版，但不作通过承诺。
- 此配置不包括：制表符、ZWJ/ZWNJ/变体选择符、垂直文本、kashida 两端对齐、断字、语言定制、合成样式、描边文本输出、非 BMP 输出。
- 0.2.0 候选的三个固定段落输出页面已通过 macOS Preview 检查。PDF 选择几何、Copy 和 Find 仍未验证。详见[发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)。
