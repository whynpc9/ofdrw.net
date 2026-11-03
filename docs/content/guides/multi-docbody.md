---
title: "Multi-DocBody 文档选择"
description: "枚举 OFD 中的多个 DocBody，使用零基 API 或一基 CLI 文档号选择读取、导出和受限编辑目标。"
sidebar:
  order: 140
  label: "Multi-DocBody"
---

一个 OFD 包可以在 `OFD.xml` 中声明多个 `DocBody`。Ofdrw.Net 默认选择第一个文档，也允许显式枚举和选择。

## 枚举并选择文档

```csharp
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

var reader = new OfdReader();
await using var input = File.OpenRead("input.ofd");

var documents = await reader.ListDocumentsAsync(input);
Console.WriteLine($"DocBody 数量：{documents.Count}");

// ListDocumentsAsync 会消耗流；可 seek 的输入需要回到开头再读取。
input.Position = 0;
var package = await reader.ReadAsync(
    input,
    new OfdPackageLoadOptions { DocumentIndex = 1 });
```

`ListDocumentsAsync` 按 `OFD.xml` 中的顺序返回描述信息，包括零基索引、规范化 `DocRoot` 路径、可选标题和签名清单路径。它只是清单，不代表文档验证或签名信任结果。缺失、重复或仅大小写不同的 `DocRoot` 路径会被拒绝。

`DocumentIndex` 从 **0** 开始，适用于 Reader 以及 PDF、SVG、HTML 和图片导出。未显式设置时仍为 `0`；页索引始终相对于所选文档。

CLI 的 `--document N` 从 **1** 开始：

```sh
ofdrw ofd-to-pdf -i input.ofd -o second.pdf --document 2
ofdrw extract-text input.ofd second.txt --document 2
ofdrw reorder input.ofd reordered.ofd --document 2 --pages 3,1,2
```

该参数适用于 OFD 读取/导出、`extract-text`、`reorder` 和 `watermark` 等支持的命令。

## 导出与签章外观

PDF/SVG/HTML/图片导出只读取所选文档的 Reader 投影。PDF 和图片的签章外观也只来自所选 `DocBody` 的签名清单，即使另一个文档复用了相同页面 ID。

Reader 可能解析跨文档资源路径并用于渲染，但这不表示该包可以安全编辑和重打包。各格式的渲染限制见[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)；选择文档不会让保留的未知 XML 自动变成可绘制内容。

## 隔离编辑边界

`OfdDocumentEditor` 在内存中编辑选定的 `OfdDocumentPackage`。对于导入的多文档包，`OfdPackageWriter` 只有同时满足以下条件才接受写出：

- 每个 `DocBody` 根位于不同的顶层目录。
- 所有包条目属于这些目录之一或 `OFD.xml`。
- XML 形式的条目没有解析到其他文档目录的引用。
- 包中没有签名声明或 `Signs` 载荷。

Writer 还会比较每个未选文档的条目和 `DocBody` 子树，确保它们没有变化。不支持的情况在写输出流之前抛 `NotSupportedException`；最终文件发布仍应使用暂存和原子提交，以处理最终复制或 I/O 失败。

这个受限模型允许对所选文档重排页面和添加水印，**不允许**多文档 `Split` / `Mix`、跨文档导航/批注/资源迁移、共享根资源或签名包编辑。

`clean-signatures` 是显式的整包操作，验证器也会遍历所有 `DocBody`。Password v2 仍要求单 `DocBody`。重写 `OFD.xml` 后不能声称原签名继续有效；需要编辑签名包时，先显式执行整包清签流程，再对最终结果重新签署。
