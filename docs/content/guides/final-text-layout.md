---
title: "启用最终文字排版引擎"
description: "为 FlowDocument、纯文本转换和 DOCX Native OFD 显式启用段落排版引擎：真实字体塑形、制表位、合成粗斜体、装饰线、CFF/OTF 字体与增补汉字。"
sidebar:
  order: 35
  label: "最终文字排版（可选）"
---

> **main，尚未发布。** 这是可选开关（Issue #109）。默认行为不变：`FlowDocument`、`PlainTextOfdConverter` 和 DOCX Native 继续使用兼容排版和名称引用字体。默认值是否切换将在后续版本根据比对结果另行决定。

[段落排版引擎](https://whynpc9.github.io/ofdrw.net/reference/text-paragraph-layout/)用实际字体文件做 HarfBuzz 塑形，按 Unicode 17.0.0 双向、断行和字素簇排版，并把每个字形的原始 GID 写入 OFD（簇映射 `CGTransform`，字体以保留 GID 的子集嵌入）。开启后，已有的分页、表格、页眉页脚和回调逻辑保持不变，只是文本行由引擎产生。

## FlowDocument 与纯文本

需要额外引用 `Ofdrw.Net.Text.HarfBuzz`。字体由调用方显式提供，不搜索系统字体：

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Text.HarfBuzz;

var engine = new FlowTextEngine(HarfBuzzTextPrelayout.CreateParagraphLayout(), new[]
{
    new TextFontFamily("Noto Sans", new TextFontFace("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"))),
    // CFF/OTF（含 CID-keyed 的思源黑体、Noto Sans CJK OTF）可直接使用
    new TextFontFamily("Source Han Sans SC", new TextFontFace("Source Han Sans SC", File.ReadAllBytes("SourceHanSansSC-Regular.otf"))),
});

var document = new FlowDocument();
document.Options.TextEngine = engine;
document.Blocks.Add(new Paragraph("中文 English\t制表位 与 𠂇 增补汉字"));
var package = document.Render();

var options = new PlainTextOfdOptions();
options.Layout.TextEngine = engine;          // PlainTextOfdConverter 复制该设置
var plain = new PlainTextOfdConverter(options).Render(text);
```

- `Span.FontFamily`（或 `DefaultFontFamily`）按名称（不区分大小写）匹配到的字体家族排在回退链首位，其余家族按给定顺序作为回退；未匹配时使用全部家族的原顺序。
- `Bold`/`Italic`/`FontWeight ≥ 600` 优先选择声明了对应样式的字面；缺失时默认合成（`TextSyntheticStylePolicy.Synthesize`），可在构造函数中改为 `Reject`。
- `Underline` 使用字面的下划线位置和粗细，`UnderlineOffsetMillimeters`/`UnderlineWidthMillimeters` 只影响兼容排版。`LetterSpacingMillimeters` 按字素簇生效。
- 换行、末尾换行和换页（`\f`）沿用兼容排版的语义；制表符按默认 12.7 mm 制表位推进（兼容排版按四个空格测量）。
- 含内嵌图片的段落仍使用兼容排版。引擎错误以 `InvalidOperationException` 抛出，`InnerException` 为带源区间的 `TextLayoutException`。

## DOCX Native

```csharp
var options = new DocxConversionOptions { OfdMode = DocxToOfdMode.Native };
options.FontDirectories.Add("/path/to/fonts");
options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
await new DocxToOfdConverter(options).ConvertAsync(input, output);
```

- 字体来自 `FontDirectories`（未配置时为平台 CJK 目录）。请求的字体家族在前，`FontFallbackFamilies` 与常见中文字体替代顺序中实际存在的家族随后；所有使用的字面都会嵌入。名称引用的 SimSun/SimHei 等阅读器本地字体不再保持未嵌入，改为嵌入实际用于排版的字面。
- 读取 `w:tabs`（左/居中/右/小数，`clear` 合并样式链）和 `settings.xml` 的 `w:defaultTabStop`；制表位按 Word 从文本栏起始处量起，悬挂缩进的位置本身也是首行的制表位。
- 读取 `w:strike`/`w:dstrike`（一条删除线）、`w:highlight` 和字符底纹 `w:shd`（高亮矩形）。
- `NativeEmbeddedBoldBehavior = Throw` 时不合成粗体（缺少粗体字面即失败）；默认 `TextStroke` 合成，写为填充色描边而不是 `Weight`。
- 含内嵌图片的段落使用兼容排版，并报告一次 `DOCX_NATIVE_TEXT_ENGINE_FALLBACK`。不受支持的字面报告 `DOCX_FONT_TEXT_ENGINE_UNSUPPORTED` 并被跳过。
- OFD 包带自定义标签 `docx-native-text-engine = ParagraphLayout`。

## 与兼容排版的差异

同一 DOCX、同一字体下，[比对记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/r22-engine-switch-comparison.md)中的样例页数和逻辑文本完全一致，差异来自真实字体度量：行高取字面上升部/下降部（兼容排版为 1.3 em），拉丁文按实际字面宽度换行，制表符对齐到制表位。默认切换前请用自己的文档比对两种输出。

## 限制

- 不支持：彩色字体、可变字体、CFF2、竖排、制表符前导符、断字、显式双向格式字符。
- 含 ZWJ/ZWNJ/变体选择符的文字可写入 OFD 和 SVG，但导出 PDF 时失败关闭（PDFsharp 文本层无法保留它们）。
- 本开关尚无 macOS Preview 验收记录；PDF Copy/Find 未在真实阅读器验证。
