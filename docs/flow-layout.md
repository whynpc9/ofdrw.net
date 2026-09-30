# 公开流式布局 API

`Ofdrw.Net.Layout` 的 `FlowDocument` 让调用方用段落和 Span 生成多页 OFD。页面尺寸、边距、缩进和字号都以毫米计；调用方不用设置 TextObject 坐标。首版支持文本段落、左/中/右对齐、局部粗体、斜体、颜色、折行、显式换页和自动分页。公开表格、Canvas、图片块与两端对齐仍属于后续票。

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
