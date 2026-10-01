# OFD 与 PNG/JPEG 图片转换

Issue 02 在 `Ofdrw.Net.Converter.Pdf` 包新增两个公开转换器，转换集合包也可使用。此能力随本票源码交付；既有 `0.1.0-preview.8` 发布包不包含本票增量。

## 按页导出

```csharp
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;

using var input = File.OpenRead("input.ofd");
using var output = File.Create("page-2.png");
await new OfdToImageConverter(new OfdToImageOptions
{
    PixelsPerMillimeter = 4, // 101.6 DPI
    Format = OfdImageFormat.Png // 默认 PNG；也可 Jpeg
}).ConvertAsync(input, output, pageIndex: 1, cancellationToken: cancellationToken);
```

API 页号零基，CLI 页号一基。每次调用导出一页，默认首页；CLI `--pages` 仅接受一个正整数。输出尺寸是每轴 `ceil(页面毫米 × ppm)`，默认 ppm=`144 / 25.4`。格式由 `Format` / `--format` 决定，与文件后缀无关。JPEG 默认质量 90，可设 1–100。PNG/JPEG 都合成为 RGB 白底；透明背景不保留。

链路为 **OFD → 单页 PDF → PDFium → RGB PNG/JPEG**。直接用 pixels-per-point 渲染，不经过已有 PDF 导入栅格器的 72–300 DPI 截断。PDFium 整数像素舍入引起的最多一个像素差异，再缩放到约定的 ceil 网格。极低 ppm 时在预算内增加采样以保证每轴至少一个像素，再缩放到目标尺寸；极细长页面因此可能触发额外采样预算。字体、模板、图像、路径及签章外观沿用 OFD→PDF 预览能力及其保真边界，不承诺任意复杂 OFD 的完整保真。需要 Docnet 支持的 PDFium 本地运行库；加载失败会明确报错。

```bash
ofdrw ofd-to-image input.ofd page-2.png --pages 2 --ppm 4
ofdrw ofd-to-image input.ofd page-2.jpg --pages 2 --ppm 4 --format jpeg --jpeg-quality 95
```

## 一图一页导入

```csharp
using Ofdrw.Net.Core.Models;

using var first = File.OpenRead("first.png");
using var second = File.OpenRead("second.jpg");
using var output = File.Create("images.ofd");
await new ImageToOfdConverter(new ImageToOfdOptions
{
    PixelsPerMillimeter = 4,
    PageSize = new OfdPageSize { WidthMillimeters = 210, HeightMillimeters = 297 }
}).ConvertAsync(new Stream[] { first, second }, output, cancellationToken);
```

单张图片可用 `ConvertAsync(Stream, Stream, CancellationToken)`。输入由文件内容识别，只接受 PNG/JPEG；后缀不能把 GIF 或其他载荷变成合法输入。逐图先检查像素数，再完整解码；多帧载荷拒绝。每图一页，按传入顺序生成；输入流由调用者持有，转换器不关闭它们。

自然尺寸是原始编码像素宽/高除以 ppm。未设置页尺寸时，每页使用自然尺寸；设置后使用 `min(1, 页宽/自然宽, 页高/自然高)` 缩小，等比居中，不放大。忽略内嵌 DPI 和 EXIF 方向，使用原始编码像素轴；需要旋转的照片请先规范化像素。原始 PNG/JPEG 字节保持，不再有损编码；透明 PNG 保持源透明度，白底导出时才合成。重复图片由现有资源写出器去重。

```bash
ofdrw image-to-ofd first.png images.ofd --ppm 4
ofdrw image-to-ofd first.png second.jpg --output images.ofd --ppm 4 --page-width 210 --page-height 297
# --input 可以重复；--page-width / --page-height 必须成对提供
```

## 参数、资源预算与失败

ppm 必须有限且大于零；页宽/高必须有限、大于零且不超过 10000 mm。输出/输入图片每轴不超过 32768 px，像素数也受以下预算和原生 BGRA 长度限制。选项在构造转换器时复制，后续修改选项不改变该实例的契约。

| 范围 | 默认 | API / CLI |
| --- | --- | --- |
| 单张图片/输出页像素 | 40,000,000 | `MaxPixels` 或 `MaxPixelsPerImage` / `--max-pixels` |
| 估计栅格缓冲 | 256 MiB，按 16 bytes/pixel 校验 | `MaxRasterWorkingBytes` / `--max-working-bytes` |
| 导出中间单页 PDF | 128 MiB | `MaxIntermediatePdfBytes` / `--max-pdf-bytes` |
| 导出编码图片 | 64 MiB | `MaxOutputBytes` / `--max-output-bytes` |
| 导入单图编码输入 | 64 MiB | `MaxInputBytesPerImage` / `--max-input-bytes` |
| 导入累计编码输入 | 128 MiB | `MaxTotalInputBytes` / `--max-total-input-bytes` |
| 导入张数 | 1000 | `MaxPageCount` / `--max-pages` |
| 导入 ZIP 条目 | 10000；预检保守上界 `5 + 2 × 张数` | `MaxEntryCount` / `--max-entries` |
| 导入输出 OFD | 256 MiB | `MaxOutputBytes` / `--max-output-bytes` |

所有预算为正整数，导入单图编码输入不得超过 `Int32.MaxValue`。工作缓冲预算与像素预算同时执行；默认 256 MiB / 16 使实际像素上限约 16.7M。它是缓冲估算，并非进程总内存硬限额：编码输入、OFD XML/展开资源、PDF/font/native 状态另有内存开销。导入逐图释放解码缓冲，原始编码输入累计保留供打包；写包时还会产生资源条目副本。

导出的 `PackageLoadOptions` 沿用包预算：压缩输入 512 MiB、10000 条目/页引用、单条目展开 128 MiB、累计展开 512 MiB、压缩比 1000。CLI `--max-input-bytes` 映射压缩 OFD 输入限制；其他包预算可用 API 调整。图片导出只解析选中页的签章外观，嵌套 OFD 继承同一加载预算，超限明确失败。中间 PDF 和编码输出在有界临时文件写入时检查预算，不能通过 seek 越过限制。

转换前和处理阶段的非法页号、格式、超限、取消或生成失败不会向目标流写入半包。**任意调用方 Stream 的最后复制发生 I/O 失败或取消时无法回滚**；不要直接打开需保留的文件并期待 `Stream` API 提供事务。CLI 先暂存，再在同一文件系统原子替换；失败时保留已有文件并清理临时文件，取消返回 130。API 不关闭输出流。解码及原生渲染有同步阶段，取消在阶段前后检查，不能中断正在执行的 PDFium native 调用。

## 可复查样例

[Issue 02 验证记录](../validation/issue02/README.md) 包含实际 OFD、PDF、逐页 PNG/JPEG、无隐私几何样例、manifest 和完整性校验方法。功能测试、自动渲染、PNG 目视与 macOS Preview 分别记录。
