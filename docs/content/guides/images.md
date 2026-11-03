---
title: "图片导入与导出"
description: "将 OFD 页面导出为 PNG/JPEG，或将一张或多张 PNG/JPEG 图片生成一图一页的 OFD。"
sidebar:
  order: 90
  label: "图片导入导出"
---

安装 `Ofdrw.Net.Converter.Pdf`。`OfdToImageConverter` 和 `ImageToOfdConverter` 都位于 `Ofdrw.Net.Converter.Pdf.Converters`，配置类型位于 `Ofdrw.Net.Converter.Pdf`。

## OFD 导出为 PNG/JPEG

```csharp
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;

var options = new OfdToImageOptions
{
    PixelsPerMillimeter = 144d / 25.4d,
    Format = OfdImageFormat.Png
};

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("page-1.png");
await new OfdToImageConverter(options).ConvertAsync(input, output, pageIndex: 0);
```

- 默认输出 PNG、144 DPI；`PixelsPerMillimeter` 表示每毫米像素数，144 DPI 约为 `5.67`。
- API 页索引从 **0** 开始。每次调用导出所选的一个页面，背景合成到白色。
- JPEG 通过 `Format=OfdImageFormat.Jpeg` 和 `JpegQuality` 配置；它是有损格式，验证文字、线框和边缘时优先使用 PNG。
- 输出经 OFD→PDF→PDFium 链路渲染，继承 PDF 出口的字体、图元和原生依赖边界。

CLI：

```sh
ofdrw ofd-to-image input.ofd page-1.png --pages 1 --ppm 5.669291 --format png
ofdrw ofd-to-image input.ofd page-2.jpg --pages 2 --format jpeg --jpeg-quality 90
```

## PNG/JPEG 生成 OFD

下面把两张图片按输入顺序生成两页 A4 OFD：

```csharp
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;

var options = new ImageToOfdOptions
{
    PageSize = new OfdPageSize
    {
        WidthMillimeters = 210,
        HeightMillimeters = 297
    }
};

await using var first = File.OpenRead("page-1.png");
await using var second = File.OpenRead("page-2.jpg");
await using var output = File.Create("images.ofd");
await new ImageToOfdConverter(options)
    .ConvertAsync(new Stream[] { first, second }, output);
```

每张图片独占一页，居中、等比缩小以适应页面，不放大。省略 `PageSize` 时按图片像素和 `PixelsPerMillimeter` 使用自然尺寸；图片内嵌 DPI 和 EXIF 方向不会用于排版。

CLI：

```sh
ofdrw image-to-ofd page-1.png page-2.jpg --output images.ofd \
  --page-width 210 --page-height 297
```

## 输入与输出边界

- 仅接受单帧 PNG/JPEG，不接受 GIF、BMP、TIFF、动画 PNG 或其他格式。
- 图片文字成为位图内容，不会自动获得可搜索的 OFD 语义文本层。
- 图片尺寸、解码像素、工作缓冲、输入/输出字节和页数均有预算；超预算会失败，不会无限分配内存。
- API 在验证和编码阶段保护目标流，但最终复制或 flush 失败仍可能留下部分字节；发布文件时使用同目录暂存与原子替换。CLI 已采用文件暂存机制。

实际样例和已检查页面范围见[图片验收记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/validation/issue02/README.md)。通用转换限制见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。
