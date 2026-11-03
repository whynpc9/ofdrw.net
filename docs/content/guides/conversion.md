---
title: "PDF、SVG 与文本转换"
---

安装转换元包，或按方向安装 Converter.Pdf、Converter.Svg。

## PDF 转 OFD

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("input.pdf");
await using var output = File.Create("output.ofd");
await new PdfToOfdConverter().ConvertAsync(input, output);
```

默认双层转换用渲染页面保持视觉内容，同时保留可抽取的 PDF 文字作为定位透明文字。扫描 PDF 不自动获得 OCR。PDF 的默认抽取顺序不保证逻辑阅读顺序。

可选原生 PathObject / TextObject 输出见[PDF 矢量扩展](https://whynpc9.github.io/ofdrw.net/guides/pdf-vector/)。

## OFD 转 SVG

```csharp
using Ofdrw.Net.Converter.Svg.Converters;

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("page.svg");
await new OfdToSvgConverter().ConvertAsync(input, output, pageIndex: 0);
```

SVG 为自包含单页。PDF、SVG、HTML 和图片有独立的支持边界，不能由某一出口的成功推断其他出口。

纯文本输入使用[PlainTextOfdConverter](https://whynpc9.github.io/ofdrw.net/guides/plain-text-ofd/)，网页预览见[HTML](https://whynpc9.github.io/ofdrw.net/guides/html-preview/)。完整拒绝、回退和资源限制见[能力约定](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)。
