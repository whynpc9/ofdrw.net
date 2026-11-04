---
title: "PDF、SVG 与文本转换"
description: "使用 Ofdrw.Net 在 PDF、OFD、SVG 和文本之间转换，理解默认双层、页索引和输出边界。"
sidebar:
  order: 10
  label: "PDF 与 SVG 转换"
---

安装 `Ofdrw.Net.Converter` 元包，或按方向安装 `Ofdrw.Net.Converter.Pdf` / `Ofdrw.Net.Converter.Svg`。以下 API 的页索引均从 **0** 开始；CLI 页号从 **1** 开始。

| 方向 | 入口 | 默认行为 |
| --- | --- | --- |
| PDF → OFD | `PdfToOfdConverter` | 页面图 + 透明文字层 |
| OFD → PDF | `OfdToPdfConverter` | 按页导出 PDF 预览 |
| OFD → SVG | `OfdToSvgConverter` | 导出自包含单页 SVG |
| OFD → 文本 | `OfdTextExtractor` | 根据页面文字对象和坐标重建文本 |

## PDF 转 OFD

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("input.pdf");
await using var output = File.Create("output.ofd");
await new PdfToOfdConverter().ConvertAsync(input, output);
```

默认双层转换用渲染页面保持视觉内容，同时保留可抽取的 PDF 文字作为定位透明文字：

- 扫描 PDF 不会自动获得 OCR 文本。
- PDF 默认抽取顺序不保证多栏、双向文字或复杂页面的逻辑阅读顺序。
- 可选原生 `PathObject` / `TextObject` 输出见[PDF 矢量扩展](https://whynpc9.github.io/ofdrw.net/guides/pdf-vector/)，它是独立源码扩展，不是默认路径。
- 书签和附件迁移默认关闭，需显式启用；见[PDF 导航与附件](https://whynpc9.github.io/ofdrw.net/guides/pdf-import-navigation-attachments/)。

## OFD 转 PDF

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("preview.pdf");

// 省略第三个参数可导出全部页面；这里导出第 3 页、第 1 页、第 3 页。
await new OfdToPdfConverter().ConvertAsync(input, output, new[] { 2, 0, 2 });
```

页选择保留列表顺序和重复页。渲染失败不会用丢失正文的空白页冒充成功。

签章外观无法解析或绘制时，默认仍会写出 PDF 并跳过该外观。需要知道是否丢失签章时，使用 `ConvertWithResultAsync`（main，尚未发布）：

```csharp
var result = await new OfdToPdfConverter().ConvertWithResultAsync(input, output);
Console.WriteLine($"已绘制签章外观：{result.RenderedSignatureAppearances}");
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Code} 页 {diagnostic.PageIndex}: {diagnostic.Message}");
```

## OFD 转 SVG

```csharp
using Ofdrw.Net.Converter.Svg.Converters;

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("page.svg");
await new OfdToSvgConverter().ConvertAsync(input, output, pageIndex: 0);
```

SVG 是自包含的**单页**输出。多页静态预览请使用 [HTML 导出](https://whynpc9.github.io/ofdrw.net/guides/html-preview/)，位图预览请使用[图片导出](https://whynpc9.github.io/ofdrw.net/guides/images/)。

## 读取 OFD 文本

安装 `Ofdrw.Net.Reader` 后：

```csharp
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

await using var input = File.OpenRead("input.ofd");
var package = await new OfdReader().ReadAsync(input);
var pages = new OfdTextExtractor().ExtractPages(package, includeTemplates: true);

foreach (var page in pages)
    Console.WriteLine(page);
```

文本提取按照页面模型和坐标重建行，不保证恢复原始文件的换行、空格或阅读顺序。纯文本生成 OFD 见[纯文本指南](https://whynpc9.github.io/ofdrw.net/guides/plain-text-ofd/)。

## 对应的 CLI 命令

```sh
ofdrw pdf-to-ofd -i input.pdf -o output.ofd
ofdrw ofd-to-pdf -i input.ofd -o preview.pdf --pages 3,1,3
ofdrw ofd-to-svg -i input.ofd -o page.svg --pages 1
ofdrw extract-text input.ofd output.txt --include-templates
```

PDF、SVG、HTML 和图片出口有各自的支持边界，某一出口成功不能证明其他出口也正确。完整的拒绝、回退、取消、流写入和资源限制见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)与[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)。
