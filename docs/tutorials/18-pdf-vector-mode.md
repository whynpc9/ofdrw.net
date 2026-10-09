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

显式回退遇到可证明的单幅 full-page raw RGB/8 或 literal Indexed/DeviceRGB/8 图像时，会保留原RGB或逐像素展开原literal调色板的采样、分辨率和源 PDF 插值提示，不先做 DPI 放大；报告 `PDFV_ORIGINAL_IMAGE_PAGE`，仍为 raster 内容，`IsNative=false`。其它图像/effects 页报告 `PDFV_RASTER_PAGE` 并沿用双层。图像仅在严格页框/矩阵/操作序列及无裁剪、透明、mask、颜色重解释、滤镜时命中；复杂图像不能根据“看上去像扫描页”推断安全。插值 hint 是本库 PDF/SVG 导出提示，第三方 OFD 阅读器可能忽略；各 renderer 的插值算法、DeviceRGB 显示色彩不作跨设备完全一致保证。

首版受限于零原点未旋转页框、CropBox=MediaBox、UserUnit=1、无过滤器的页内容/字体/ToUnicode 流、8-bit RGB/Gray 固色路径、butt/miter=10 描边和横排 fill text。支持内容 affine CTM、文字 affine 矩阵、显式字距、cubic 和 fill rules。clip、图像/Form、ExtGState、透明/混合、dash、shading、标记内容、注释、复杂页框和未知操作均有诊断。压缩流也明确回退/失败；不能用小压缩载荷宣称解码内存有界。

字体只接纳嵌入固定 TrueType 的 Type0 Identity-H/CIDFontType2/identity CIDToGID，并逐字符验证原 Unicode 与相同字节的源 CID/目标 glyph。不兼容的 PDF 子集、竖排、连字、多字符映射、组合 shaping、非 BMP 或仅空白游程回退/失败。字体保持原载荷和真实 face flags；没有共享 resolver/subset/cache 替代，05 的广泛字体服务仍未完成。完整 CJK 字体可让矢量 OFD 明显大于双层 OFD；实际大小见本票证据。

通过 `PdfVectorToOfdOptions` 配置输入/页数、内容字节、操作、事件、路径命令、文字和字体限额；回退像素/图片预算沿用 `Compatibility`。预算与取消始终失败，不以回退绕过。每个选定页面使用新的 PdfPig 文档/资源上下文，避免失败状态污染后一页；多页和重复选择会增加解析成本。PdfPig 打开文档、解析字典和 native font loader 内部并不提供硬堆内存/时间沙箱，本票不作该承诺；只限制所声明的输入、流及输出工作量。

转换写出前的验证失败不改目的流；最终 ZIP 写出阶段的 I/O/取消仍可留下部分流，文件调用方应自行使用临时文件与原子替换。默认 CLI 原子输出契约保持原样。

设计与证据：[实证探针](../pdf-vector-probe-design.md)、[实现契约](../pdf-vector-design-contract.md)。同一 PDF 比较样例位于 `e2e/Ofdrw.Net.Pdf.Vector.E2E`；运行 `scripts/run-pdf-vector-package-e2e.sh` 消费本次默认 11 包加两个可选包。实际视觉结论以已检查的 OFD → PDF → Preview 页面为准，不推断任意 PDF 保真。

本次已检查的包、源码、失败历史和 Preview 页面记录见 [验收证据](../evidence/pdf-vectors/README.md)。

仅 open move 或支持的 butt move/close 描边按 PDF no-op 消耗路径、不创建事件；若页中仍有支持内容则保持 native，整页没有事件时沿用 `NO_NATIVE_CONTENT` 策略。闭合 singleton 填充可能产生设备像素，因此明确 `DEGENERATE_POINT_FILL` 整页回退/失败，包括与其它段共存的情况。显式 line/cubic 即便退化仍保留；奇异或 float/mm 转换后不可逆矩阵明确 `SINGULAR_SERIALIZED_MATRIX`，不由通用异常捕获掩盖。

原生路径额外限制 producer 将已解析 double 转为 float 时的新增误差：每个页坐标分量 ≤0.0001 PDF pt，非零控制多边形向量的相对误差及线性 CTM 扭曲 ≤0.00001，固定 miter10 的笔宽偏移误差 ≤0.0001 pt。包括 cubic/v/y 控制点、闭合边、signed re 及 double 角点加法；真实重合保留，n 丢弃和不绘制的 trailing move 不产生新的精度回退。超范围明确 `PATH_FLOAT_PRECISION` 按整页policy处理，不能报告已丢失图形的 native。使用有限大小的精确二进制算术比较，复杂度随命令数线性；不修改19/04，也不承诺恢复 PdfPig 解析或累计 double CTM 运算中已丢失的十进制数字。上述是 incoming producer 误差界限，不保证最终边界量化、拓扑、miter分支或像素完全一致。

合法 Encoding CMap / CIDToGIDMap 流是字体范围外的 `FONT_PROFILE`/`CID_MAPPING`，进入明确回退/失败；严格literal Indexed/DeviceRGB数组可保留原采样；其它数组颜色空间跳过原样优化，交给既有整页渲染。用于这些判别的间接引用有深度/环/取消检查，严格读取错误不被吞掉；错误 primitive 类型、缺失引用和损坏仍是输入失败，不以回退掩盖。

R5原Indexed样例已由主/Low真实包及八页Preview复验，原R4失败保留；此有限样例的颜色检查通过，最终票的其它review与验收仍未闭合。严格应用选择Fail；显式RasterizePage结果不等于颜色保真验收。

文本新增 `TEXT_FLOAT_PRECISION`：可见矩阵/字号/累计字距转float产生过大新增误差时整页按policy处理。保留普通小数、CJK、shear、连续空格及可精确表达的大数对照；原文与字体资源不改。此为回调double到producer输出的界限，后续OFD decimal/PDF renderer舍入不属于该界限，不宣称逐像素或任意PDF保真。Indexed省略Decode的8bit值按规范夹取到0..hival，2/255对hival1均读取最后调色板项；错误palette长度仍失败。
