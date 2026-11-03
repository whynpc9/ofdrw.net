---
title: "DOCX 转换与分页"
description: "选择 Native 或 DualLayer 将 DOCX 转为 OFD，读取转换诊断、选择页面并验证原文与实际排版。"
sidebar:
  order: 20
  label: "DOCX 转换"
---

安装 `Ofdrw.Net.Converter.Docx`。`DocxToOfdConverter` 默认使用 **Native**：直接读取 OpenXML 原文并写出 OFD 文字、路径和图片对象，不调用 Office，也不经过 PDF。

## 选择转换模式

| 模式 | 视觉内容 | 机器可读文字 | 适用场景 |
| --- | --- | --- | --- |
| `Native`（默认） | 由内置 Native 排版生成 | 原始 OpenXML | 保留原文、常见段落/表格/图片/页眉页脚 |
| `DualLayer` | Word、LibreOffice 或 BuiltIn 渲染的页面图 | 仍来自原始 OpenXML | 需要由外部渲染器确定视觉页面，同时保留文字层 |

两种模式的依赖和版面能力不同。复杂浮动对象、复杂域和任意 Word 版式不在通用保真承诺内。

## Native：转换并读取诊断

```csharp
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;

var options = new DocxConversionOptions
{
    OfdMode = DocxToOfdMode.Native
};

await using var input = File.OpenRead("input.docx");
await using var output = File.Create("output.ofd");
var result = await new DocxToOfdConverter(options)
    .ConvertWithResultAsync(input, output);

Console.WriteLine($"源分页数：{result.SourcePageCount}");
Console.WriteLine($"原文保留：{result.OriginalTextPreserved}");
Console.WriteLine($"页级映射可靠：{result.PageTextMappingAccurate}");
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

不设置 `OfdMode` 时也是 Native。`ConvertAsync` 适合只需转换的场景；需要判断降级、附属正文或原文保留状态时，使用 `ConvertWithResultAsync` 并检查诊断，不能只看任务是否正常完成。

### 字体与严格失败

字体名称、实际字体文件和回退选择会影响字形、分页和粗体表现。生产部署应显式准备许可明确的 Regular/Bold 等字体，并按[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)设置 `FontDirectories` 与 `FontFallbackFamilies`。

默认 `UnsupportedFeatureBehavior=Placeholder` 会保留部分不支持内容的诊断信息；如果业务不允许降级，可设置 `UnsupportedDocxFeatureBehavior.Throw`。Native 嵌入文字要求真实粗体时，使用 `NativeEmbeddedBoldBehavior=DocxEmbeddedBoldBehavior.Throw`。

## DualLayer：显式使用渲染页

```csharp
var options = new DocxConversionOptions
{
    OfdMode = DocxToOfdMode.DualLayer,
    Engine = DocxConversionEngine.Auto
};
```

`Engine` 决定视觉渲染器。`Auto` 在 macOS 优先 Word，然后 LibreOffice，最后 BuiltIn。PDF 文字只用于定位，输出的语义文字仍来自 OpenXML；无法可靠定位正文时会明确失败，不提交看似成功的部分语义层。

直接 DOCX→PDF 使用 `DocxToPdfConverter`，但它不能替代 Native OFD 的验收结果。

## 选择页面

API 页索引从 **0** 开始，CLI 页号从 **1** 开始：

```csharp
var result = await new DocxToOfdConverter(options)
    .ConvertWithResultAsync(input, output, new[] { 2, 0, 2 });
```

该示例按“第 3 页、第 1 页、第 3 页”的顺序输出。页选择基于转换器实际产生的分页，不应直接沿用 Word UI 中假定的页码。越界、空选择和附属正文归属等详细规则见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。

CLI 对应命令：

```sh
ofdrw docx-to-ofd -i input.docx -o output.ofd --docx-ofd-mode native
ofdrw docx-to-ofd -i input.docx -o selected.ofd --pages 3,1,3
```

## 验证输出

1. 重新读取 OFD，核对原始文本和 `result.Diagnostics`。
2. 将**本次生成的 OFD**导出为 PDF，再逐页查看。
3. 检查中英文、局部样式、表格、页眉页脚、图片、分页、裁切、重叠和空白页。
4. 在业务实际使用的 OFD 阅读器中另行验证；PDF 预览不能证明所有阅读器行为。
