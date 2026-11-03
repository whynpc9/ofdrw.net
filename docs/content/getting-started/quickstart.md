---
title: "快速开始"
---

先按[安装说明](https://whynpc9.github.io/ofdrw.net/getting-started/installation/)安装转换包。以下示例用于 .NET 控制台程序的顶层语句。

## DOCX 转 OFD

```csharp
using Ofdrw.Net.Converter.Docx.Converters;

await using var input = File.OpenRead("input.docx");
await using var output = File.Create("output.ofd");
await new DocxToOfdConverter().ConvertAsync(input, output);
```

默认走 Native：OpenXML 原文写成可见 OFD 文字对象，不经过 PDF。转换中文前配置合法字体，完整约定见[DOCX 指南](https://whynpc9.github.io/ofdrw.net/guides/docx/)。

## OFD 转 PDF

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("output.ofd");
await using var output = File.Create("preview.pdf");
await new OfdToPdfConverter().ConvertAsync(input, output);
```

在阅读器中打开 preview.pdf 检查实际页面。文件生成成功和文本可提取分别证明不同事项，不能据此断言原 Word 完整保真。

## 生成第一份 OFD

安装 Ofdrw.Net.Layout 和 Ofdrw.Net.Packaging 后运行：

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var document = new FlowDocument();
document.Blocks.Add(new Paragraph("你好，Ofdrw.Net"));
await using var output = File.Create("hello.ofd");
await new OfdPackageWriter().WriteAsync(document.Render(), output);
```

输出一份包含段落的 OFD。Flow 默认以字体名声明文字，跨机显示需要字体部署；报表场景继续阅读[排版](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/)和[表格](https://whynpc9.github.io/ofdrw.net/guides/public-tables/)。
