---
title: "段落排版引擎"
description: "Unicode 17.0.0 双向、断行与字素簇，字体回退、逐行重塑形、对齐/两端对齐、分页，以及直接消费最终排版结果的 OFD 写出器。"
sidebar:
  order: 130
  label: "段落排版引擎"
---

> 自 [0.2.0](https://whynpc9.github.io/ofdrw.net/releases/020/) 起可用。制表位、合成粗体/斜体、增补字符输出、装饰线与高亮、字距、ZWJ/ZWNJ/变体选择符、静态 CFF 字面以及 Flow/纯文本/DOCX 的显式开关自 [0.3.0](/releases/030/) 起可用（Issue #109 F08 与引擎开关）。

`Ofdrw.Net.Layout.Text.Paragraphs` 将带样式的多段落文本排版为最终行和页面，`ParagraphLayoutOfdWriter` 则将该结果原样写为 OFD。`FlowDocument`、`PlainTextOfdConverter` 和 DOCX Native 默认仍使用既有兼容排版；只有显式设置 [`FlowDocumentOptions.TextEngine` 或 `DocxConversionOptions.NativeParagraphLayoutEngine`](https://whynpc9.github.io/ofdrw.net/guides/final-text-layout/) 时才使用本引擎。既有预排版和混排查询保持各自的准入限制。

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

1. **准入检查。** 不成对的代理项、除制表符和行/段分隔符外的 C0/C1 控制符、显式双向格式字符（LRE…PDI、LRM/RLM/ALM）以及 ZWJ/ZWNJ 之外的格式控制符（软连字符、单词连接符、BOM 等），都会在任何塑形之前连同源区间一起被拒绝。ZWJ、ZWNJ 和变体选择符（U+FE00–FE0F、U+E0100–E01EF）随所在字素簇一起塑形，字面 cmap 无需覆盖它们。样式区间必须完整拼接覆盖文本，且不得拆分字素簇。
2. **段落划分。** 按此处准入的 UAX #9 段落分隔符拆分：LF、CR、CRLF、NEL、U+2029。VT、FF 和 U+2028 是段落内的硬换行。终止符属于行的源区间（`TerminatorLength`），但没有字形。
3. **Unicode 分析。** 使用冻结的 Unicode 17.0.0 数据表：UAX #9 rev 51（P2–P3、X1–X10、W1–W7、N0–N2、I1–I2、L1–L2）、UAX #14 rev 55 默认断行和 UAX #29 rev 47 默认扩展字素簇。不进行定制、词典分词或断字。参见[一致性验证](#unicode-一致性验证)。
4. **项目划分。** 按解析后的双向级别、文字体系（通用/继承字符采用相邻文字体系）、所选字面和样式划分。每个字素簇采用样式回退链中首个满足条件的字体家族：其字面在请求的粗体/斜体声明下覆盖该簇的所有标量。通用文字体系的簇在旧字面仍可覆盖时保持使用该字面。不发现宿主机字体。默认（`TextSyntheticStylePolicy.Reject`）不合成粗体/斜体：缺少样式对应字面会以 `UnsupportedFontProfile` 失败；缺少覆盖则以 `MissingGlyph` 失败。`options.WithSyntheticStyles(TextSyntheticStylePolicy.Synthesize)` 时选择最接近的已声明字面（只补不减），用填充色轮廓描边合成粗体（汉字 1% em，其余 2.5% em，与既有 Native OFD 一致），用以基线为轴的 0.2 斜切合成斜体；步进不变，墨迹边界计入描边和斜切。合成粗体不能与请求的描边或无填充画刷组合。每个制表符自成一项，不送入塑形器。
5. **塑形。** 使用 `ITextShaper`（默认 HarfBuzz）对每个项目进行测量塑形。
6. **断行。** 在同时属于字素簇边界的 UAX #14 断行机会上采用首次适配贪心算法。尾部空格悬挂。随后重建所选行：被行边缘或 L1 级别变化切开的游程会**重新塑形**并检查实际宽度；若溢出，则尝试前一个断行机会（最多 32 次，随后 `BudgetExceeded`）。由于估算将跨断点字距调整计入左侧，在信任估算结果之前，还会重建最多三个“估算超出版心不超过半个 em”的后续合法断点。没有断行机会的片段会以 `LayoutOverflow` 失败（默认 `TextOverflowPolicy.Reject`），或在字素簇边界断开（`BreakGraphemes`）；绝不拆分簇内部。候选切点都会重建并重新塑形。重建后的宽度不具单调性（阿拉伯文 lam-alef 连字可能比单独的 lam 更窄），因此二分查找后还会进行有界向上扫描，并在连续四个切点溢出后停止。所选切点是已找到的最大适配切点，不保证全局最大。
7. **重排序。** 每行先应用 L1，再对游程应用 L2。游程内部保留塑形器的视觉字形顺序。
8. **制表位与字距。** 制表符（UAX #9 L1 后位于段落级别）按逻辑顺序从行起始边推进到下一个制表位：显式 `TextTabStop`（`Start`/`Center`/`End`/`Decimal`，最多 64 个，从版心起始边量起），其后为 `DefaultTabOriginMillimeters + k × DefaultTabIntervalMillimeters`（默认 0 与 12.7 mm）的隐式起始对齐制表位；超出版心的制表位截到版心边缘。`options.WithMinimumTabAdvance(em)`（0–4，默认 0，即 Word 规则）要求起始对齐和隐式制表位与制表符起点至少相距 em × 制表符字号，不足时改用下一个制表位（同 CSS `tab-size` 不足半个字符时跳到下一个制表位的规则）；居中、右对齐和小数制表位不受影响。制表符写为字面空格字形（GID 来自 cmap U+0020），其步进为到制表位的距离；逻辑文本保留 U+0009。`TextStyle.WithLetterSpacing` 在同一样式、同一行内后接内容的每个字素簇之后增加固定步进。
9. **度量与对齐。** 所有游程共享一条基线。正常行高使用各字面在相应字号下的上升部、下降部和行间距；也可使用 `Multiple`、`AtLeast`（不小于给定值，多余空间位于文字下方）和 `Exact`。`Start`/`End` 随段落方向变化；首行缩进位于起始侧，可为负值（悬挂缩进，范围 (−宽度, 宽度)）。`Justify` 在词间空格以及汉字、假名、谚文或注音符号相邻字素之间拉伸非末行（闭标点前不拉伸，行内最后一个制表符之前不拉伸）；额外步进按字形记录。
10. **装饰。** `TextStyle.WithDecorations`（下划线、删除线，可指定颜色，默认填充色）和 `WithHighlight` 生成填充矩形：高亮跨上升部到下降部，下划线和删除线使用字面 `post` / `OS/2` 中的位置和粗细（缺失或不合理时使用常规默认值）。几何取自样式回退链第一个字体家族的字面，使同一样式跨字体回退时保持水平。行尾悬挂空格不装饰。
11. **分页。** 按 `PageHeightMillimeters` 划分页面（或使用一个高度不限的页面）。分页处丢弃段间距；高于页面的行会以 `LayoutOverflow` 失败。

## 最终结果约定

`ParagraphLayoutResult` 不可变且完整；失败或取消时不返回部分结果。坐标为相对于各页面内容框的毫米值，Y 向下，只从设计单位转换一次。

| 层级 | 包含内容 |
| --- | --- |
| 结果 | 请求（不可变样式）、页面、`UnicodeVersion` = 17.0.0、塑形器 `EngineVersion`、塑形调用次数 |
| 页面 | 索引、行、已用高度 |
| 行 | 绝对 UTF-16 源区间（含尾部空格和终止符）、段落/页面索引、基础级别、硬换行标记、顶部、基线、高度、上升部、下降部、对齐后的内容左侧位置/宽度、墨迹边界、按视觉顺序排列的游程 |
| 游程 | 逻辑源文本和起点、精确的 `TextShapedRun`（字面字节身份、文字体系、语言、方向、特性）、解析后的双向级别、`TextResolvedStyle`、原点、基线、步进、墨迹 |
| 字形 | 原始 GID、绝对簇起点/长度（多对多）、原点 X/Y、含两端对齐的步进、两端对齐增量、墨迹 |

`TextResolvedStyle` 快照保存画刷（DeviceRGB 填充/描边颜色及 alpha、描边宽度；合成粗体时含填充色描边）、字号、语言、请求与字面各自的粗体/斜体状态、字体家族及其回退索引、`Synthetic`（`None`/`Bold`/`Oblique` 标志）、`SyntheticBoldStrokeMillimeters`、`ObliqueSkew`、装饰线与颜色、高亮色和字距。游程另带 `Decorations`（高亮、下划线、删除线矩形，按绘制顺序）和 `IsTab`。写出器不再查询请求样式，不重新选择字体，也不重新塑形。

## OFD 输出

`ParagraphLayoutOfdWriter.CreateDocument(result, pageWidth, pageHeight, left, top)` 为每个排版页面创建一个 OFD 页面，并为每个游程创建一个 `TextObject`：

- `TextCode` 保存游程未经修改的逻辑文本；
- 每个源簇对应一个按逻辑顺序排列的 `CGTransform`（`CodePosition`/`CodeCount` 按 Unicode 标量即 XML 字符计数，增补字符计为 1；BMP 文本与 0.2.0 相同），列出原始 GID；
- 字形槽位遵循该顺序，每个槽位都有显式原点（完整的 `DeltaX`/`DeltaY`），因此 RTL 游程使用负增量，而不是重排文本；
- 每个不同字面对应一个字体资源，以保留 GID 的子集嵌入（或通过 `embedFullFonts: true` 完整嵌入）；遵守 OS/2 嵌入限制。TrueType 字面沿用 glyf 子集；静态 CFF1 字面（含 CID-keyed）保留所有 GID 槽位，未使用的 CharString 原位改为 `endchar`、未调用的子程序改为 `return` 并补零，不改动任何 CFF 偏移，由 OFD 的 deflate 压缩去除填充；
- 写入填充颜色和 alpha；描边（含合成粗体）写为 `Stroke`/`LineWidth`/`StrokeColor`，不写 `Weight`；合成斜体写为以游程基线为轴的剪切 `CTM`，不写 `Italic`；
- 高亮在该行所有 TextObject 之前、下划线和删除线在其之后，写为填充 PathObject。同一行内种类、颜色和纵向范围相同且首尾相接的矩形（例如字体回退分成拉丁文和汉字游程的同一高亮）合并为一个路径，避免抗锯齿阅读器在两块之间画出接缝（Issue #115）；结果中的 `Decorations` 仍按游程给出。

以下情况会在创建包之前以 `UnsupportedOutput` 拒绝：解析样式与字面不一致、文本框或墨迹超出页面，以及源文本/GID 身份丢失。

模型将此表示为 `OfdPositionedGlyphMapping.ClusterMap`。写出器在 `https://ofdrw.net/layout-hints` 命名空间中，用写出器自有的提示属性 `ClusterMapV1="true"` 标记此类 TextObject（其他阅读器会忽略它）。读取时，历史有界 CGTransform 配置仍优先；只有被这些配置拒绝的已标记 TextObject，才会按完整簇映射接纳：一个 TextCode（可含制表符、ZWJ/ZWNJ 和成对代理项，其他控制/格式字符仍拒绝）、每个标量恰好属于一个递增 CGTransform、每个字形槽位都有原点。簇内的隐藏连接符/选择符字形可以无轮廓，只要同簇另有带墨迹的字形。未标记的外来 CGTransform 继续保持历史失败关闭规则。PDF 和 SVG 根据接纳映射中每个字形的原始轮廓进行绘制（CFF 字面由托管 Type 2 CharString 解释器读取三次曲线），并将逻辑文本保留为不可见 PDF 文本或 SVG `title`/`aria-label`。

PDF 导出只嵌入 TrueType 程序（FontFile2）：静态 CFF1 字面在导出时转换为保留 GID、cmap 和水平度量的 glyf 版本（每段三次曲线以不超过 1 设计单位的误差拆为二次曲线），仅用于 PDF；OFD 保留原 CFF 字节。增补字符的语义文本沿用既有私用区别名绑定。PDFsharp 文本层无法保留 ZWJ/ZWNJ 和变体选择符，因此含这些字符的 TextObject 导出 PDF 时失败关闭（OFD 与 SVG 支持）。最终排版写出的汉字合成粗体在 PDF 中沿用 Issue #107 的仅水平加粗，避免细密笔画闭合。

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
- 字面：静态 TrueType（glyf）和 OpenType/CFF1（`TextFontFace.OutlineFormat`），含 CID-keyed CFF（如 Source Han Sans、Noto Sans CJK OTF）。可变字体（fvar/gvar）、CFF2、彩色字体（COLR/SVG/CBDT/sbix）、seac 风格的 CFF 重音 endchar 仍拒绝。
- 此配置不包括：垂直文本、kashida 两端对齐、断字、语言定制、制表符前导符（leader）、制表位上的双向分段重排之外的复杂制表语义、显式双向格式字符、彩色 emoji 字体。ZWJ emoji 序列只在提供单色字形的字面中按塑形结果输出，未单独验收。
- 0.2.0 候选的三个固定段落输出页面已通过 macOS Preview 检查。PDF 选择几何、Copy 和 Find 仍未验证。详见[发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)。
