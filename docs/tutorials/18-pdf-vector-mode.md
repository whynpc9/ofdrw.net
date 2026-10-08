# 可选 PDF → native OFD 矢量模式

安装独立可选包 `Ofdrw.Net.Converter.Pdf.Vector` 后显式使用下面的 API。默认 `PdfToOfdConverter` 和 `ofdrw pdf-to-ofd` 命令继续使用页面图与透明文字双层；默认 CLI 未接入此可选包，不支持 vector 参数。

```csharp
using Ofdrw.Net.Converter.Pdf.Vector;

using var input = File.OpenRead("source.pdf");
using var output = File.Create("vector.ofd");
var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
foreach (var page in result.Pages)
    Console.WriteLine($"source={page.SourcePageIndex + 1}, native={page.IsNative}: {page.Diagnostic}");
```

默认策略 `Fail` 在任何页超出受限范围时抛出 `NotSupportedException`，不写目的流。若希望保留扫描或复杂页，可显式配置整页双层回退，并读取每页报告：

```csharp
var converter = new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
{
    UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage
});
```

原生页具有 PathObject 与可见 TextObject，原文和连续空格保留在带字距的游程内；没有整页图或轮廓字替代原文。回退页丢弃全部已暂存矢量，只有既有页面图与可选透明文字，不会把可见文字再次叠在整页图上。有效纯文字页可以原生输出；扫描、没有可表达 native 内容和真正空白页会明确回退或失败。回退页文字沿用既有 PDF 双层的单词语义，未承诺精确原文/连续空格或 OCR。

首版受限于零原点未旋转页框、CropBox=MediaBox、UserUnit=1、无过滤器的页内容/字体/ToUnicode 流、8-bit RGB/Gray 固色路径、butt/miter=10 描边和横排 fill text。支持内容 affine CTM、文字 affine 矩阵、显式字距、cubic 和 fill rules。clip、图像/Form、ExtGState、透明/混合、dash、shading、标记内容、注释、复杂页框和未知操作均有诊断。压缩流也明确回退/失败；不能用小压缩载荷宣称解码内存有界。

字体只接纳嵌入固定 TrueType 的 Type0 Identity-H/CIDFontType2/identity CIDToGID，并逐字符验证原 Unicode 与相同字节的源 CID/目标 glyph。不兼容的 PDF 子集、竖排、连字、多字符映射、组合 shaping、非 BMP 或仅空白游程回退/失败。字体保持原载荷和真实 face flags；没有共享 resolver/subset/cache 替代，05 的广泛字体服务仍未完成。完整 CJK 字体可让矢量 OFD 明显大于双层 OFD；实际大小见本票证据。

通过 `PdfVectorToOfdOptions` 配置输入/页数、内容字节、操作、事件、路径命令、文字和字体限额；回退像素/图片预算沿用 `Compatibility`。预算与取消始终失败，不以回退绕过。每个选定页面使用新的 PdfPig 文档/资源上下文，避免失败状态污染后一页；多页和重复选择会增加解析成本。PdfPig 打开文档、解析字典和 native font loader 内部并不提供硬堆内存/时间沙箱，本票不作该承诺；只限制所声明的输入、流及输出工作量。

转换写出前的验证失败不改目的流；最终 ZIP 写出阶段的 I/O/取消仍可留下部分流，文件调用方应自行使用临时文件与原子替换。默认 CLI 原子输出契约保持原样。

设计与证据：[实证探针](../pdf-vector-probe-design.md)、[实现契约](../pdf-vector-design-contract.md)。同一 PDF 比较样例位于 `e2e/Ofdrw.Net.Pdf.Vector.E2E`；运行 `scripts/run-pdf-vector-package-e2e.sh` 消费本次默认 11 包加两个可选包。实际视觉结论以已检查的 OFD → PDF → Preview 页面为准，不推断任意 PDF 保真。
