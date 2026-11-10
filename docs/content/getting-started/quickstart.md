---
title: "快速开始"
description: "用一段可运行的 C# 示例完成 DOCX → OFD → PDF 和文本读取，并生成第一份流式排版 OFD。"
sidebar:
  order: 20
---

下面提供两条路径：**转换现有 DOCX**或**用代码生成 OFD**。示例使用 `0.3.1`，C# 代码适用于控制台项目的顶层语句。

## 路径一：DOCX → OFD → PDF

创建项目并安装转换元包：

```sh
dotnet new console -n OfdrwQuickstart
cd OfdrwQuickstart
dotnet add package Ofdrw.Net.Converter --version 0.3.1
```

将 `input.docx` 放到项目目录，再把 `Program.cs` 替换为：

```csharp
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

await using (var input = File.OpenRead("input.docx"))
await using (var output = File.Create("output.ofd"))
{
    await new DocxToOfdConverter().ConvertAsync(input, output);
}

await using (var input = File.OpenRead("output.ofd"))
await using (var output = File.Create("preview.pdf"))
{
    await new OfdToPdfConverter().ConvertAsync(input, output);
}

await using (var input = File.OpenRead("output.ofd"))
{
    var package = await new OfdReader().ReadAsync(input);
    var pages = new OfdTextExtractor().ExtractPages(package, includeTemplates: true);
    for (var index = 0; index < pages.Count; index++)
    {
        Console.WriteLine($"--- 第 {index + 1} 页 ---");
        Console.WriteLine(pages[index]);
    }
}
```

运行：

```sh
dotnet run
```

成功后会得到：

- `output.ofd`：默认 **Native** 模式生成的 OFD，文字直接来自 OpenXML，不经过 PDF。
- `preview.pdf`：由本次生成的 OFD 导出的预览 PDF。
- 终端输出：按页提取的 OFD 文本；它用于核对原文，不证明页面排版完全保真。

处理中文前请按[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)准备实际字体文件。需要诊断、选页或 DualLayer 时继续阅读 [DOCX 转换](https://whynpc9.github.io/ofdrw.net/guides/docx/)。

## 路径二：生成第一份 OFD

安装排版和写包：

```sh
dotnet add package Ofdrw.Net.Layout --version 0.3.1
dotnet add package Ofdrw.Net.Packaging --version 0.3.1
```

使用以下代码生成 `hello.ofd`：

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var document = new FlowDocument();
document.Blocks.Add(new Paragraph("你好，Ofdrw.Net"));

await using var output = File.Create("hello.ofd");
await new OfdPackageWriter().WriteAsync(document.Render(), output);
```

`FlowDocument` 负责折行和分页，`OfdPackageWriter` 负责写出 OFD 包。默认字体只声明 `SimSun` 名称，不嵌入字体字节；跨机器查看前需要确认目标阅读器具有合适的中文字体。

## 检查结果

1. 确认输出文件存在且能重新读取。
2. 提取文本，核对中英文、空格和页数是否符合预期。
3. 在目标阅读器中逐页打开最终 OFD；如通过 PDF 检查 Native DOCX，应使用本次的 **DOCX → Native OFD → PDF** 链路，而不是直接 DOCX→PDF。
4. 检查缺字、断词、局部粗体/斜体、表格、分页、裁切和异常空白页。

文件生成成功、文本可提取和视觉正确是三项不同检查。复杂文档请继续阅读[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)与[已知问题](https://whynpc9.github.io/ofdrw.net/troubleshooting/known-issues/)。
