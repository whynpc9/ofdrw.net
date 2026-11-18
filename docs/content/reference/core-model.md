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
dotnet add package Ofdrw.Net.Reader --version 0.5.0
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

## 自定义标签（CustomTags）

`OfdDocumentPackage.CustomTags` 是字符串键值字典，DOCX 转换写入的 `source-text-origin`、`docx-ofd-mode` 等来源信息以及命名区域索引都存放在这里。调用方式不变；以下为包内写出形态。

**自 0.5.0 起**：新生成的文档（Layout、DOCX Native/默认/DualLayer、纯文本等）经中性通道写出，不再使用 EMR 标记：

| 层级 | 位置 | 内容 |
| --- | --- | --- |
| 声明 | `Document.xml` | `<ofd:CustomTags>Tags/CustomTags.xml</ofd:CustomTags>` |
| 索引 | `Tags/CustomTags.xml` | `<ofd:CustomTag TypeID="ofdrw-kv" NameSpace="urn:ofdrw-net:custom-tags:kv:1">`，子元素 `<ofd:FileLoc>CustomTag_KeyValue.xml</ofd:FileLoc>` |
| 载荷 | `Tags/CustomTag_KeyValue.xml` | 根元素 `Tags`，子元素 `Tag`（`Key`、`Value` 属性），均位于 `urn:ofdrw-net:custom-tags:kv:1` 命名空间，与索引声明一致 |

索引不写 `SchemaLoc`：本库没有发布对应的 schema 文件，该结构也未对照 GB/T 33190 核实为符合标准的自定义标引。

0.4.0 及更早版本写出的历史 EMR 通道（`TypeID="EMR"`、`NameSpace="urn:ofdrw-net:custom-tags:emr"`，载荷根为 OFD 命名空间下的 `EMRTags`）继续可读，读到的键值与以前相同。规则如下：

- 标签未改动时原样保存，EMR 声明、索引和载荷不被改写。
- 修改或删除已在 EMR 载荷中的键，在原位置更新；载荷中其他内容保留。
- 新增的键写入中性通道：已有中性载荷时追加到其中，否则新建 `CustomTag_KeyValue.xml` 并在现有索引中追加一条记录。EMR 载荷不会再收到新键。
- 读取时合并两个通道；同一键在多处出现时，按索引中的记录顺序以后者为准。
- 索引记录的 `TypeID`/`NameSpace` 与载荷根元素不匹配时，该载荷按未知内容保留，不读入 `CustomTags`。

拆分保留来源包的标签结构；混合（mix）生成的新文档统一写为中性通道。

## 单位、页号与文档选择

- 页面坐标和尺寸通常使用**毫米**，Y 轴向下；文字定位还涉及基线。
- API 页索引从 **0** 开始，CLI 页号从 **1** 开始。
- 默认读取第一个 `DocBody`；枚举和选择其他文档见 [Multi-DocBody 指南](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。
- 类型化元数据、权限、版本和动作见[标准对象编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/)。

公开能力以[功能对照](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)和[分出口能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)为准。默认 `DocType` 或局部结构检查不等于医疗、归档或密码规范的完整合规。
