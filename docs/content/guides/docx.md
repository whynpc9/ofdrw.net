---
title: "DOCX 转换与分页"
---

安装 `Ofdrw.Net.Converter.Docx`。默认 `DocxToOfdConverter` 直接写出 Native OFD，适合保留原文和常见段落、表格、内嵌图片及页眉页脚。

```csharp
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx;

var options = new DocxConversionOptions { OfdMode = DocxToOfdMode.Native };
await using var input = File.OpenRead("input.docx");
await using var output = File.Create("output.ofd");
await new DocxToOfdConverter(options).ConvertAsync(input, output);
```

不设置 OfdMode 时也是 Native。字体名称、实际物理字体和回退选择会影响分页，先准备[字体配置](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。

## DualLayer

设置 `OfdMode = DocxToOfdMode.DualLayer` 后，Word / LibreOffice / BuiltIn 渲染视觉页，语义层继续来自原始 OpenXML；不会用 PDF 抽取文本替代 Word 原文。`Engine` 决定渲染器，Auto 在 macOS 优先 Word，然后 LibreOffice，最后 BuiltIn。

Native 与 DualLayer 的依赖和版面能力不同。复杂浮动对象及任意 Word 版式不在通用保真承诺内。

## 页选择与完整性

API 页索引从零开始，CLI 从一开始。页选择以实际转换分页为准。原文保留、附属正文追加和预算失败行为见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。

验证 Native 时应从生成的 OFD 导出 PDF，再查看页面。直接 DOCX→PDF 不能证明 Native OFD 的排版。
