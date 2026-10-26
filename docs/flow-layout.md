# 公开流式布局 API

`Ofdrw.Net.Layout` 的 `FlowDocument` 让调用方用段落和 Span 生成多页 OFD。页面尺寸、边距、缩进和字号都以毫米计；调用方不用设置 TextObject 坐标。支持文本段落、局部样式、表格、固定大小的图片和 Canvas 块、折行与分页。公开表格见 [Table/Row/Cell 教程](public-tables.md)；两端对齐和浮动盒仍未实现。

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var document = new FlowDocument();
var paragraph = new Paragraph { Alignment = ParagraphAlignment.Center };
paragraph.Spans.Add(new Span("中文 English "));
paragraph.Spans.Add(new Span("重点") { Bold = true, Color = new OfdColor(192, 0, 0) });
document.Blocks.Add(paragraph);
for (var i = 0; i < 100; i++) document.Blocks.Add(new Paragraph($"第 {i + 1} 段内容。"));

var package = document.Render();
await using var output = File.Create("report.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

图片块和 Canvas 块占据明确的宽高，作为不可拆分页单元；宽度超过正文区或高度超过新页正文高度会失败，不会自动缩小。图片接受 R09 绘图入口支持的静态 PNG 或基线 JPEG，构造时复制输入字节。Canvas 回调在排版完成后调用，接收以最终块左上角为原点、初始裁剪到块矩形的 `OfdGraphics`。回调可以使用同一绘图入口画路径、文字或图片。示例：

```csharp
using Ofdrw.Net.Layout.Graphics;

var pngBytes = File.ReadAllBytes("chart.png");
document.Blocks.Add(new ImageBlock(pngBytes, 60, 40)
{ Alignment = ParagraphAlignment.Center, SpaceAfterMillimeters = 4 });
document.Blocks.Add(new CanvasBlock(100, 28, (graphics, box) =>
{
    graphics.FillRectangle(new OfdBrush(new OfdColor(225, 240, 225)),
        0, 0, box.WidthMillimeters, box.HeightMillimeters);
}) { Alignment = ParagraphAlignment.Center });
document.OnPage = (graphics, page) => { /* 使用整页毫米坐标画页背景 */ };
document.OnElement = (graphics, placement) =>
{ /* placement.Block 和 placement.Boundary 是最终块片段与页坐标 */ };
```

`OnPage` 按页序各调用一次，绘制对象位于正文下方；`OnElement` 按文档顺序对每个最终片段调用一次，绘制对象紧随该片段。段落按行、表格按行、固定块与命名区域按整体提供片段。`FlowBoundary.PageIndex` 从零开始，X/Y/Width/Height 是物理页上的毫米坐标。测量失败时不运行回调；回调抛错或取消则 `Render` 失败，不返回包。回调运行时不要修改 `Blocks` 或其内容；Canvas 绘图上下文的公开变换和裁剪操作仍由调用方负责正确使用。

`Span.FontWeight` 可选 100–900、步长 100；与 `Bold=true` 同时设置时只接受 700。`LetterSpacingMillimeters` 在同一 Span 的相邻字素之间增加非负间距，自动折行末尾不增加尾间距。`Underline` 可选偏移与线宽，沿最终文字游程绘制原生路径。段落与表格单元格共用这些样式规则。上述间距仍沿用本页下述的传统字宽估算，不代表真实字体预排版或复杂文字 shaping。

默认 A4、四边 25.4 mm、10.5 pt 等值字号（3.704 mm）、`SimSun` 字体名。每个 Span 的粗斜体是独立的，不会传播到相邻 Span。相邻 Span 合并为逻辑文本后，`CRLF`/`CR` 归一为 `LF`；显式 `\n` 换行、`\f` 换页，过长英文单词按 Unicode 文本元素拆分。段末 `\n` 的空行在后续内容出现时计入高度，不单独创建文末空白页。一个字素跨 Span 时采用起始字符的样式。首版用固定的 Arial 兼容比例字宽表测量 Basic Latin 与 Latin-1，CJK（含 Hangul 和增补汉字）使用 1 em；其他脚本暂按固定 0.6 em 近似。字宽不依赖运行机器安装的字体，写出的 OFD `DeltaX` 与分页测量一致；`FontFamily` 指定 OFD 声明字体，但若目标字体实际字形宽度差异很大，视觉仍需单独检查。默认字体仅声明名称，不嵌入字节；跨机器展示需确保目标阅读器有合适 CJK 字体，字体子集和嵌入复用留给 05 票。

普通分隔空格在自动折行末尾保留字符、推进量归零，不影响居中或右对齐；显式换行及末行保留空格推进量。NBSP（U+00A0）、窄 NBSP（U+202F）与数字空格（U+2007）为 [Unicode GL 空格](https://www.unicode.org/reports/tr14/#GL)，连接的文字作为不可拆分单元折行，首版不实现完整 Unicode 行断规则；单元超出可用行宽（首行扣除首行缩进）时公开 API 明确失败，不额外生成空首行或删除缩进，DOCX Native 保留其既有超宽内容溢出策略。

空格字宽规则：普通 NBSP 等于普通空格，窄 NBSP 固定为 0.2 em，数字空格等于对应粗斜体样式的 `0` 字宽。DOCX Native 使用相同规则，但数字字宽继续由其现有字体测量器提供。

单词预测折行覆盖 Basic Latin/Latin-1 字母与数字，并按字素的基字符识别，包含 `café`、`café` 等预组合/分解形式；超出完整行宽的单词仍按字素拆分。非断行空格带组合标记时也保留连接属性。连续段末换行保留全部空行高度，在后续内容出现时使用，不单独生成文末空白页。

显式空行使用换行 Span/run 的字号和段落最小行高，并逐行分页。同段中间或前置空行使用结束该行的换行字号；正文后的延期尾部换行按每个控制后的间距保存，纯换行段的关闭空行重复最后一个换行字号。`PageBreakBefore` 在已有正文或延期显式空行时开启新页并清除这些尾间距；首块直接设置它不产生前置空页。DOCX 的非末节结束时，用所属节的页尺寸提交延期显式空行；末节到达文末仍不生成空白尾页。Native 正文 `PAGE` 在延期空行、段间距及前置 LF/FF 确定段落首个正文行的页号后重新测量；同段后续自动分页或正文后的显式换页仍沿用首个正文行页号上下文，未声明完整正文动态计数支持。页眉页脚的逐页计数契约不变。预组合与分解形式能组成受支持的 Latin-1 字母时，用组合字母字宽测量（如 `í` 与 `í`），OFD 保留原始字符序列。PDF 导出绘制单个受支持的 Latin-1 字素时使用对应组合字形，复制出的 PDF 文本为规范等价形式，避免分解 `í` 的原始 `i` 点与重音叠绘。

`Render()` 每次生成新包，不修改之前返回的包；调用方可复用描述对象，但不要在其他线程同时修改其 `Blocks`、`Spans` 或 `Options`。超高行、无可用宽度、页数、字符数和文本元素数量超限会明确失败，不会静默裁切。预算由 `MaxPageCount`、`MaxCharacters`、`MaxTextElements` 控制，取消通过 `CancellationToken` 传入。`OfdDocumentBuilder` 的按页坐标 API 保持可用。

运行仓库内的无隐私样例：

```bash
dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release -- artifacts/flow-layout
```

样例同时生成公开 Flow、显式 Native DOCX 和默认 DOCX 的 OFD/PDF。视觉验收必须打开本次 OFD 导出的 PDF；测试和 `pdftoppm` 页面图不代替 macOS Preview。实际检查记录见[2026-09-29 验收](validation/flow-layout-2026-09-29.md)。

只运行图片/Canvas/Span 样例：`dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release -- artifacts/flow-fixed e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --flow-fixed-only`。样例生成 `flow-fixed.ofd`、由该 OFD 导出的 `flow-fixed.pdf` 和边界/哈希记录。
