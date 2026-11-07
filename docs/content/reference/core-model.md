---
title: "OFD 模型与读写边界"
description: "理解 OfdReader、OfdDocumentPackage、OfdPackageWriter 与 FlowDocument 的职责，并读取 OFD 页面文本。"
sidebar:
  order: 30
  label: "模型与读取"
---

Ofdrw.Net 用强类型模型表示页面、图层、文字游程、路径、图片、字体、模板和附件。常用入口分为三层：

| 层次 | 入口 | 职责 |
| --- | --- | --- |
| 容器读写 | `OfdReader`、`OfdPackageWriter` | 将 OFD ZIP 包读入模型，或将模型写回 OFD |
| 文档模型 | `OfdDocumentPackage`、`OfdPage`、页面元素 | 访问或编辑文档结构、页面和资源 |
| 生成与排版 | `FlowDocument`、`OfdDocumentBuilder` | 用流式块或显式页面坐标创建内容 |

## 读取文档与文本

安装 `Ofdrw.Net.Reader`：

```sh
dotnet add package Ofdrw.Net.Reader --version 0.3.0
```

读取包并按页提取文字：

```csharp
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

await using var input = File.OpenRead("input.ofd");
var package = await new OfdReader().ReadAsync(input);

Console.WriteLine($"页数：{package.Pages.Count}");
var pages = new OfdTextExtractor().ExtractPages(package, includeTemplates: true);
for (var index = 0; index < pages.Count; index++)
{
    Console.WriteLine($"--- 第 {index + 1} 页 ---");
    Console.WriteLine(pages[index]);
}
```

`ExtractPages` 默认不包含模板文字；上面显式传入了 `includeTemplates: true`。文本根据文字对象及坐标重建，不保证恢复原始换行、空格或复杂页面的逻辑阅读顺序。

## 写出修改后的文档

写包需要 `Ofdrw.Net.Packaging`：

```csharp
using Ofdrw.Net.Packaging;

await using var output = File.Create("output.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

生成新文档时，可将 `FlowDocument.Render()` 或 `OfdDocumentBuilder` 产生的包交给同一个 writer。完整生成示例见[快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/)和[流式排版](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/)。

## 包结构与保留策略

OFD 是 ZIP 容器：`OFD.xml` 指向一个或多个 `DocBody`，文档再引用页面与资源。SDK 对未知 XML 和包条目采取保留策略，但以下能力必须分别判断：

- **能保留**：读写往返时不主动丢弃未知内容。
- **能渲染**：PDF/SVG/HTML/图片出口理解并能绘制该内容。
- **能编辑**：公开模型允许安全修改该对象。
- **能迁移**：合并、拆分或混合时能证明并重映射完整引用闭包。

无法证明安全时，相关操作会拒绝或保留并报告诊断；不要把“未知 XML 被保留”理解为“任意内容都能编辑和导出”。

## 单位、页号与文档选择

- 页面坐标和尺寸通常使用**毫米**，Y 轴向下；文字定位还涉及基线。
- API 页索引从 **0** 开始，CLI 页号从 **1** 开始。
- 默认读取第一个 `DocBody`；枚举和选择其他文档见 [Multi-DocBody 指南](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。
- 类型化元数据、权限、版本和动作见[标准对象编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/)。

公开能力以[功能对照](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)和[分出口能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)为准。默认 `DocType` 或局部结构检查不等于医疗、归档或密码规范的完整合规。
