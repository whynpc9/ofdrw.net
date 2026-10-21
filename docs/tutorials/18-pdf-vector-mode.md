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

原生页以 PathObject、可见 TextObject 与受限的 ImageObject 保留绘制顺序；原文和连续空格保留在带字距的游程内。受限图像保留源像素，不生成整页图或轮廓字替代原文。回退页丢弃全部已暂存原生对象，只有既有页面图与可选透明文字，不会把可见文字再次叠在整页图上。有效纯文字页可以原生输出；扫描、没有可表达 native 内容和真正空白页会明确回退或失败。回退页文字沿用既有 PDF 双层的单词语义，未承诺精确原文/连续空格或 OCR。

显式回退遇到可证明的单幅 full-page raw RGB/8 或 literal Indexed/DeviceRGB/8 图像时，会保留原RGB或逐像素展开原literal调色板的采样、分辨率和源 PDF 插值提示，不先做 DPI 放大；报告 `PDFV_ORIGINAL_IMAGE_PAGE`，仍为 raster 内容，`IsNative=false`。其它图像/effects 页报告 `PDFV_RASTER_PAGE` 并沿用双层。图像仅在严格页框/矩阵/操作序列及无裁剪、透明、mask、颜色重解释、滤镜时命中；复杂图像不能根据“看上去像扫描页”推断安全。插值 hint 是本库 PDF/SVG 导出提示，第三方 OFD 阅读器可能忽略；各 renderer 的插值算法、DeviceRGB 显示色彩不作跨设备完全一致保证。

首版受限于零原点未旋转页框、CropBox=MediaBox、UserUnit=1、无过滤器的页内容/字体/ToUnicode 流、8-bit RGB/Gray 固色路径、butt/miter=10 描边和横排 fill text。支持内容 affine CTM、文字 affine 矩阵、显式字距、cubic 和 fill rules。clip、图像/Form、ExtGState、透明/混合、dash、shading、标记内容、注释、复杂页框和未知操作均有诊断。压缩流也明确回退/失败；不能用小压缩载荷宣称解码内存有界。

字体只接纳嵌入固定 TrueType 的 Type0 Identity-H/CIDFontType2/identity CIDToGID，并逐字符验证原 Unicode 与相同字节的源 CID/目标 glyph。不兼容的 PDF 子集、竖排、连字、多字符映射、组合 shaping、非 BMP 或仅空白游程回退/失败。字体保持原载荷和真实 face flags；没有共享 resolver/subset/cache 替代，05 的广泛字体服务仍未完成。完整 CJK 字体可让矢量 OFD 明显大于双层 OFD；实际大小见本票证据。

字体在真正含非空字节的 Tj/TJ 前按当前 PdfPig 文字状态准入。空 Tj、空 TJ 或仅 Tf 选择而未绘制字形不读取/嵌入字体 program，也不验证未消费字体的内部 body/profile；与未选资源相同。有效 Font 字典的类型、声明数量以及页面内容/操作/深度仍检查。真正消费时，字体缺失引用、环、错误类型等输入结构失败不以回退掩盖；已知范围外字体仍按 page policy。含数值但只有空字符串的 TJ 保留每项位移，先认证水平 Identity-H 方向元数据，不能直接当 no-op，也不因此加载 FontFile2。

`MaxFontBytes` 限制每个原字体流；新增 `MaxFontReadBytesPerPage`（默认128,000,000）限制每个选定页面尝试加载的原字体 program 字节，在复制/加载前与剩余额度比较。成功绑定的同名字体重复使用不重复收费；不同 resource ID 的加载各自收费，失败/拒绝不退款。ToUnicode 内容和结构限额独立保持。该限额不代表全部解析内存、SHA 次数、CPU 或时间上限。`MaxTotalFontBytes` 只限制成功文本事件实际引用并输出的字体资源总字节，跨页与重复 selection 继续累计；相同 bytes 的不同资源/页不新增去重。未用或失败绑定及资源帧在页面结束清理，PdfPig 自身解析缓存仅随该页文档释放，不声称 Unload 会清除缓存。

通过 `PdfVectorToOfdOptions` 配置输入/页数、内容字节、操作、事件、路径命令、文字和字体限额；回退像素/图片预算沿用 `Compatibility`。预算与取消始终失败，不以回退绕过。每个选定页面使用新的 PdfPig 文档/资源上下文，避免失败状态污染后一页；多页和重复选择会增加解析成本。PdfPig 打开文档、解析字典和 native font loader 内部并不提供硬堆内存/时间沙箱，本票不作该承诺；只限制所声明的输入、流及输出工作量。

真实或可能绘制的路径在判断范围外 profile/精度前检查剩余事件容量；组合填充/描边 `B/B*/b/b*` 一次须容纳两个事件。非空字形回调先检查已消耗的字符额度及一个文字事件容量，非空游程写出前再次检查容量。额度不足为 `InvalidDataException`，两种 page policy 均失败；容量足够时保留原范围外诊断。空路径、支持的 move-only 描边、空 Tj/TJ 或仅 Tf 不新增事件；路径构造/close、实际字体加载尝试和取消仍走各自硬限制。不预扫描未来字符或把所有 PDF 操作符当作事件收费。

转换写出前的验证失败不改目的流；最终 ZIP 写出阶段的 I/O/取消仍可留下部分流，文件调用方应自行使用临时文件与原子替换。默认 CLI 原子输出契约保持原样。

设计与证据：[实证探针](../pdf-vector-probe-design.md)、[实现契约](../pdf-vector-design-contract.md)。同一 PDF 比较样例位于 `e2e/Ofdrw.Net.Pdf.Vector.E2E`；运行 `scripts/run-pdf-vector-package-e2e.sh` 消费已验证的默认包集合及两个可选包。实际视觉结论以已检查的 OFD → PDF → Preview 页面为准，不推断任意 PDF 保真。

本次已检查的包、源码、失败历史和 Preview 页面记录见 [验收证据](../evidence/pdf-vectors/README.md)。

仅 open move 或支持的 butt move/close 描边按 PDF no-op 消耗路径、不创建事件；若页中仍有支持内容则保持 native，整页没有事件时沿用 `NO_NATIVE_CONTENT` 策略。闭合 singleton 填充可能产生设备像素，因此明确 `DEGENERATE_POINT_FILL` 整页回退/失败，包括与其它段共存的情况。显式 line/cubic 即便退化仍保留；奇异或 float/mm 转换后不可逆矩阵明确 `SINGULAR_SERIALIZED_MATRIX`，不由通用异常捕获掩盖。

原生路径额外限制 producer 将已解析 double 转为 float 时的新增误差：每个页坐标分量 ≤0.0001 PDF pt，非零控制多边形向量的相对误差及线性 CTM 扭曲 ≤0.00001，固定 miter10 的笔宽偏移误差 ≤0.0001 pt。包括 cubic/v/y 控制点、闭合边、signed re 及 double 角点加法；真实重合保留，n 丢弃和不绘制的 trailing move 不产生新的精度回退。超范围明确 `PATH_FLOAT_PRECISION` 按整页policy处理，不能报告已丢失图形的 native。使用有限大小的精确二进制算术比较，复杂度随命令数线性；不修改19/04，也不承诺恢复 PdfPig 解析或累计 double CTM 运算中已丢失的十进制数字。上述是 incoming producer 误差界限，不保证最终边界量化、拓扑、miter分支或像素完全一致。

合法 Encoding CMap / CIDToGIDMap 流是字体范围外的 `FONT_PROFILE`/`CID_MAPPING`，进入明确回退/失败；严格literal Indexed/DeviceRGB数组可保留原采样；其它数组颜色空间跳过原样优化，交给既有整页渲染。用于这些判别的间接引用有深度/环/取消检查，严格读取错误不被吞掉；错误 primitive 类型、缺失引用和损坏仍是输入失败，不以回退掩盖。

R5原Indexed样例已由主/Low真实包及八页Preview复验，原R4失败保留；此有限样例的颜色检查通过，最终票的其它review与验收仍未闭合。严格应用选择Fail；显式RasterizePage结果不等于颜色保真验收。

文本新增 `TEXT_FLOAT_PRECISION`：可见矩阵/字号/累计字距转float产生过大新增误差时整页按policy处理。保留普通小数、CJK、shear、连续空格及可精确表达的大数对照；原文与字体资源不改。此为回调double到producer输出的界限，该float界限之后另检局部笔宽/字号及实际写出CTM的总误差；其它OFD字段及PDF renderer舍入不属于这些界限，不宣称逐像素或任意PDF保真。Indexed省略Decode的8bit值按规范夹取到0..hival，2/255对hival1均读取最后调色板项；错误palette长度仍失败。

正值的描边宽度和字号还须在既有 OFD 三位写出格式下保持正值。检查实际传给 Skia 事件的 float local scalar；pt→mm 和 PDF CTM 仍放在变换中，不重复缩放标量。不可表示值按 `STROKE_WRITER_PRECISION` / `TEXT_WRITER_PRECISION` 整页 Fail/显式回退，保留原 float 精度诊断、事件/字符预算及取消硬失败；不是夹成最小宽度或提高 writer 精度。

原字形 Unicode 还必须能由 OFD 的 XML 1.0 表示。U+FFFE/U+FFFF 即使同时被 ToUnicode 与嵌入 cmap 正确映射，也会在现有字符、事件、取消及精度检查后以 `TEXT_XML_CHARACTER` 判为 native 范围外。合法 BMP 边界 D7FF/E000/FFFD 及 FDD0 保持原文；已有控制字符和代理对限制不变。

对于这类 XML 无法表示的原文，默认双层回退的 Invisible 隐藏文本层也无法序列化，因此该组合仍不支持，实际回退写入会失败；不能把 native profile 诊断当成默认回退成功。调用者可以明确选择 `UnsupportedPagePolicy = RasterizePage` 并设置 `Compatibility.TextLayerMode = PdfTextLayerMode.None`，获得完整整页图像，正常相邻页仍为 native。该选择放弃回退页的原文搜索层；转换器不会自动关闭隐藏层、替换原文或删除字形。


局部笔宽和字号写成正值仍须满足已有页空间精度界。转换器在旧 float 检查后，用实际 `WriterValue(SKfloat)` 的 double 值与实际 Emit 的 mm CTM，比较从原解析值到写出值的总误差。描边沿用 miter10 的保守偏移绝对界；文字同时检查字形/单位 em 四角、相对线性映射、连续基线及两种累计字距路径，不把字号正值当作准确性证明。超过界限仍按 `STROKE_WRITER_PRECISION` / `TEXT_WRITER_PRECISION` 整页 Fail/显式回退，不缩放补偿、不改原文或字距。文字新证书只记录失败，到游程已有硬门和 XML 门后才拒绝，保留后续字形字符额度检查。

成功处理原生内容后，若相同计算所得的物理页宽/高在 OFD 三位毫米格式下不能保持正值，转换器以 `PAGE_WRITER_PRECISION` 的窄 NotSupportedException 分类原子拒绝。两种 page policy 都失败：栅格化仍需要同一可写物理页框，不能解决该格式限制。不会夹到最小纸张或自动缩放。这个分类仅用于成功到达 native Append 前的分支；空页、未知操作、原图回退或先前字体/profile拒绝仍保留原诊断与实际回退限制，不承诺所有极小页面都有相同诊断。

## R06 扩展 profile

在上述首版契约之上，页框现接纳有限正尺寸的 MediaBox、位于其中的 CropBox 与 0/90/180/270 度旋转，输出页按裁切后显示方向计算尺寸。越界或退化 CropBox、非直角旋转仍整页拒绝。路径 `W`/`W*` 裁切按 PDF 的后续 `n` 或绘制操作生效，随 `q/Q` 恢复；裁切几何以本次输出对象坐标写出。只接纳非退化、可精确表示的路径裁切；复杂裁切、透明组、ExtGState、marked content 仍拒绝。

描边接纳 butt/round/square cap、miter/round/bevel join，以及有界的正值 dash 和 phase。实际样式写入 OFD PathObject；不能由现有 PDF/OFD 导出共同表达的组合仍按 `STROKE_PROFILE` 整页处理。圆头或方头的零长度路径可能绘制一个点，故不能作为空操作静默丢弃。

混合路径/文字页可接纳未过滤、无 mask/Decode/透明扩展的原始 DeviceRGB/8 Image XObject；转换保留图像像素、绘制顺序、变换及裁切，输出原生 OFD ImageObject，并按 PNG 字节计入累计图像预算。纯图像页继续走既有 `PDFV_ORIGINAL_IMAGE_PAGE` 或整页回退路径，报告 `IsNative=false`。Form XObject、索引/ICC 或带滤镜的混合图像仍拒绝或显式整页回退，不把先前绘制的原生对象留在回退页。

固定 TrueType 的 Type0 Identity-H/CIDFontType2 现在还接纳有界未过滤的 CIDToGIDMap stream。每个实际字符仍必须由 ToUnicode 给出单个原 Unicode 标量，且同一嵌入字体的 Unicode cmap glyph 与 stream 中的源 CID 映射完全相同。另接纳带显式 WinAnsiEncoding、完整 bounded Widths、nonsymbolic FontDescriptor、ToUnicode 与同一嵌入 TrueType 的简单字体；每个源 code 还必须按 WinAnsi 映射回同一个 Unicode，缺字或映射冲突拒绝。任意 Encoding CMap、Differences、MacRoman、符号字体与字体替换仍不在原生 profile。受限扩展只证明被准入样例的字体字节和字形身份，不推断任意 PDF 字体等价。
