---
title: "流式排版"
description: "使用 FlowDocument、Paragraph 和 Span 生成多页 OFD，配置局部样式、图片、Canvas、回调与分页预算。"
sidebar:
  order: 30
  label: "流式排版"
---

`Ofdrw.Net.Layout` 的 `FlowDocument` 用段落和其他块生成多页 OFD，调用方不需要手动计算每个 `TextObject` 的页面坐标。它支持文本、局部样式、表格、固定尺寸图片和 Canvas 块、折行及分页。

> **可选最终文字排版（自 0.3.0 起）：** 设置 `FlowDocumentOptions.TextEngine` 后，段落和表格单元格文本改用段落排版引擎，以真实字体塑形并嵌入字体子集（支持制表位、合成粗斜体、CFF/OTF 与增补汉字）。默认行为不变，见[启用最终文字排版引擎](https://whynpc9.github.io/ofdrw.net/guides/final-text-layout/)。

所有页面尺寸、边距、缩进、字间距和字号都以**毫米**表示。需要固定列网格、合并单元格和表格分页时，继续阅读[流式表格](https://whynpc9.github.io/ofdrw.net/guides/public-tables/)；需要显式页面坐标时使用 `OfdDocumentBuilder`。

## 生成多页文档

安装 `Ofdrw.Net.Layout` 和 `Ofdrw.Net.Packaging`：

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var document = new FlowDocument();
var paragraph = new Paragraph { Alignment = ParagraphAlignment.Center };
paragraph.Spans.Add(new Span("中文 English "));
paragraph.Spans.Add(new Span("重点")
{
    Bold = true,
    Color = new OfdColor(192, 0, 0)
});
document.Blocks.Add(paragraph);

for (var i = 0; i < 100; i++)
    document.Blocks.Add(new Paragraph($"第 {i + 1} 段内容。"));

var package = document.Render();
await using var output = File.Create("report.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

默认页面为 A4、四边距 25.4 mm、10.5 pt 等值字号（3.704 mm）、`SimSun` 字体名。两端对齐和浮动盒尚未实现。

## Span 与局部样式

每个 Span 的粗体、斜体和颜色相互独立，不会传播到相邻 Span。

- `FontWeight` 可取 100–900，步长 100；与 `Bold=true` 同时设置时只接受 700。
- `LetterSpacingMillimeters` 在同一 Span 的相邻字素之间增加非负间距，自动折行末尾不增加尾间距。
- `Underline` 可配置偏移和线宽，沿最终文字游程绘制原生 OFD 路径。
- 段落和表格单元格共用这些样式规则。
- 相邻 Span 会合并为逻辑文本，`CRLF` / `CR` 归一为 `LF`；显式 `\n` 换行，`\f` 换页。一个字素跨 Span 时采用起始字符的样式。

这些样式仍使用下述 Flow 字宽估算，不代表真实字体预排版或复杂文字 shaping。

## 图片与 Canvas 块

图片块和 Canvas 块占据明确宽高，是不可拆分的分页单元。宽度超过正文区或高度超过新页正文高度会失败，不会自动缩小。

```csharp
using Ofdrw.Net.Layout.Graphics;

var pngBytes = File.ReadAllBytes("chart.png");
document.Blocks.Add(new ImageBlock(pngBytes, 60, 40)
{
    Alignment = ParagraphAlignment.Center,
    SpaceAfterMillimeters = 4
});

document.Blocks.Add(new CanvasBlock(100, 28, (graphics, box) =>
{
    graphics.FillRectangle(
        new OfdBrush(new OfdColor(225, 240, 225)),
        0, 0, box.WidthMillimeters, box.HeightMillimeters);
})
{
    Alignment = ParagraphAlignment.Center
});
```

图片接受[原生绘图](https://whynpc9.github.io/ofdrw.net/guides/graphics/)入口支持的静态 PNG 或基线 JPEG，构造时复制输入字节。Canvas 回调在排版完成后运行，接收以最终块左上角为原点、初始裁剪到块矩形的 `OfdGraphics`，可绘制路径、文字或图片。

## 页面与元素回调

```csharp
document.OnPage = (graphics, page) =>
{
    // 使用整页毫米坐标绘制页背景。
};

document.OnElement = (graphics, placement) =>
{
    // placement.Block 是最终块片段，placement.Boundary 是页内毫米坐标。
};
```

- `OnPage` 按页序各调用一次，绘制对象位于正文下方。
- `OnElement` 按文档顺序对每个最终片段调用一次，绘制对象紧随该片段。
- 段落按行、表格按行、固定块与命名区域按整体提供片段。
- `FlowBoundary.PageIndex` 从零开始，X/Y/Width/Height 是物理页上的毫米坐标。

测量失败时不运行回调；回调抛错或取消会使 `Render` 失败且不返回包。回调运行时不要修改 `Blocks` 或其内容，调用方仍需正确使用绘图上下文的变换和裁剪。

## 字宽、折行与空格

Flow 使用确定性的兼容度量，而不是运行机器的实际字体测量：

- Basic Latin 与 Latin-1 使用固定的 Arial 兼容比例字宽表。
- CJK（含 Hangul 和增补汉字）使用 1 em。
- 其他脚本暂按固定 0.6 em 近似。
- `FontFamily` 决定 OFD 声明字体；默认只声明名称，不嵌入字体字节。实际字体宽度差异较大时，需要单独检查视觉结果。
- 写出的 OFD `DeltaX` 与分页测量一致。字体子集和嵌入复用见[字体参考](https://whynpc9.github.io/ofdrw.net/reference/font-subset-and-reuse/)。

过长英文单词按 Unicode 文本元素拆分。Basic Latin/Latin-1 字母和数字支持按字素基字符预测折行，包含 `café` 和分解形式的 `café`。预组合与分解形式可组成受支持 Latin-1 字母时使用组合字母字宽，OFD 仍保留原始字符序列。

普通分隔空格在自动折行末尾保留字符、推进量归零，不影响居中或右对齐；显式换行和末行保留空格推进量。

NBSP（U+00A0）、窄 NBSP（U+202F）和数字空格（U+2007）按 [Unicode GL 空格](https://www.unicode.org/reports/tr14/#GL)处理，相连文字作为不可拆分单元：

- 普通 NBSP 等于普通空格宽。
- 窄 NBSP 固定为 0.2 em。
- 数字空格等于对应粗斜体样式的 `0` 字宽。
- 不可拆分单元超出可用行宽时明确失败，不删除缩进或额外生成空首行。

这不是完整 Unicode 行断算法。DOCX Native 保留其自己的超宽内容溢出策略和字体度量入口。

## 空行与分页

- 段末 `\n` 的空行在后续内容出现时计入高度，不单独创建文末空白页。
- 连续段末换行保留全部空行高度；显式空行使用换行 Span/run 的字号和段落最小行高，并逐行分页。
- `PageBreakBefore` 在已有正文或延期显式空行时开启新页并清除尾间距；首块直接设置不会产生前置空页。
- DOCX 的非末节结束时，用所属节页面尺寸提交延期空行；末节到文末仍不生成空白尾页。
- Native 正文 `PAGE` 在首个正文行页号确定后重新测量；同段后续自动分页或正文后的显式换页沿用该上下文。页眉页脚的逐页计数遵循独立契约。

PDF 导出绘制单个受支持 Latin-1 字素时使用对应组合字形，复制文本为规范等价形式，避免分解字符的基字符点和重音叠绘。

## 渲染、预算与并发

`Render()` 每次生成新包，不修改之前返回的包。调用方可以复用描述对象，但不要在其他线程同时修改 `Blocks`、`Spans` 或 `Options`。

超高行、无可用宽度、页数、字符数和文本元素数量超限会明确失败，不会静默裁切。预算由 `MaxPageCount`、`MaxCharacters`、`MaxTextElements` 控制，取消通过 `CancellationToken` 传入。

## 运行仓库样例

```sh
dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release -- artifacts/flow-layout
```

样例生成公开 Flow、显式 Native DOCX 和默认 DOCX 的 OFD/PDF。视觉检查应打开本次 OFD 导出的 PDF；自动测试和 PNG 渲染不代替目标阅读器验收。实际检查范围见[验收记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/flow-layout-2026-09-29.md)。

只运行图片/Canvas/Span 样例：

```sh
dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release -- \
  artifacts/flow-fixed \
  e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx \
  --flow-fixed-only
```
