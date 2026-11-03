---
title: "强类型标准对象编辑"
description: "通过 OfdReader 读取并编辑 DocInfo、自定义元数据、权限、视图偏好、版本和动作，同时保留未知 XML。"
sidebar:
  order: 130
  label: "标准对象编辑"
---

`OfdReader.ReadAsync` 通过 `OfdDocumentPackage.Options` 暴露所选文档的 `Metadata`、`Permissions`、`ViewingPreferences`、`Versions` 和 `Actions`。`OfdPackageLoadOptions.DocumentIndex` 使用零基索引选择 `DocBody`，默认第一个文档；所有写出都受 [Multi-DocBody 编辑边界](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)约束。

## DocInfo 与自定义元数据

`OfdMetadata.DocUsage` 是可选的 `OfdDocUsage`（`Normal`、`EBook`、`ENewsPaper`、`EMagzine`）。`Cover` 是包内相对图片路径，`CreatorVersion` 是文本，`CustomDatas` 保存有序的 `Name` / 文本值条目。

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

var reader = new OfdReader();
await using var input = File.OpenRead("input.ofd");
var document = await reader.ReadAsync(
    input,
    new OfdPackageLoadOptions { DocumentIndex = 1 });

document.Options.Metadata.DocUsage = OfdDocUsage.EBook;
document.Options.Metadata.CustomDatas ??= new OfdCustomDatas();
document.Options.Metadata.CustomDatas.Add("case", "001");

await using var output = File.Create("output.ofd");
await new OfdPackageWriter().WriteAsync(document, output);
```

新设置或修改的 `Cover` 必须指向包中现有且不是 XML 控制内容的条目；库不会下载或创建封面图片。省略 `DocUsage` 时保留标准默认语义。

`OfdCustomDatas.Items` 和 `OfdVersions.Items` 是有序视图。编辑条目会更新其父对象；`Add`、`RemoveAt` 和 `Clear` 管理已建模条目。没有扩展内容的空 wrapper 在写出时省略，未知属性和子元素保留原位置。

`RemoveAt` 会拒绝含未知源内容的条目；`Clear` 先检查所有条目再执行，因此拒绝时集合保持不变。畸形或有歧义的源节点继续作为保留 XML，不能通过直接安装新 typed 实例静默覆盖。

## 权限、视图偏好与版本

- `OfdPermissions` 表示六个可选 Boolean 声明、`Print`（`Printable`、`Copies`）和 `ValidPeriod`（`StartDate`、`EndDate` XML dateTime 字符串）。可空 Boolean 将“省略”与标准默认值分开保存。
- `OfdViewingPreferences` 表示页面模式、布局、标签显示、隐藏标志，以及互斥的 `ZoomMode` 或数值 `Zoom`。
- `OfdVersions` 表示所选 `DocBody` 的有序 `Version` 描述（`Id`、`Index`、可选 `Current`、`BaseLocation`）。修改或新增的位置必须指向包内现有条目；版本文件本身只保留，不生成。

替换或删除导入的 `Print` / `ValidPeriod` 节点时，如果会丢失未知属性或子内容，操作会失败。应通过现有子对象修改 typed 属性，以保留未知内容。

> 权限只是声明。Reader、Writer 和编辑 API 不把它们执行为访问控制；视图偏好也不会自动控制应用 UI。

## 动作

除已有的本地 `OfdGoToAction` 外，模型还包括：

- `OfdUriAction`
- `OfdAttachmentAction`（`GotoA`）
- `OfdSoundAction`
- `OfdMovieAction`

URI、附件和媒体动作只负责序列化；库不会请求 URI、打开附件或播放媒体。新建或编辑的 typed 动作必须指向保留的包内本地附件或媒体资源。指向外部附件或缺失载荷的 `GotoA` 保持为不透明内容。

不支持、有歧义或无法证明来源的动作保留为带完整 XML 的 `OfdRawAction`。删除无关页面时，如果附件或媒体目标仍存在，已建模动作可以保留；动作包含未知 XML 或其目标被删除时，页面删除仍会被阻止。

## 来源保留、拆分与合并

导入的 wrapper 保留未知 XML、命名空间和未建模兄弟节点。给未修改的导入 typed wrapper 赋 `null` 会保留来源；撤销已经修改的 wrapper 会失败。不能把属于另一个文档的 typed 节点安装到当前文档。

`OfdDocumentSplitter.Split` 会将编辑后的自定义元数据、权限、视图偏好和版本描述带入所选的新包，并转移所有权。`OfdDocumentMerger.Merge` 采用第一个来源的 `DocUsage`、`Cover` 和 `CreatorVersion`；封面必须是现有且独立的包条目，才能复制其字节。`OfdDocumentMixer.Mix` 同样保留第一个来源的现有封面文件。

这些 API 只编辑具名对象，不提供全文档 XSD 验证、任意跨文档迁移或归档认证。字段形状与顺序核对使用公开 OFD [OFD.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/OFD.xsd)、[Document.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/Document.xsd)、[Definitions.xsd](https://github.com/zc310/ofd/blob/main/internal/schema/xsd/Definitions.xsd)，并以[上游 Java 对象模型](https://github.com/ofdrw/ofdrw/tree/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-core/src/main/java/org/ofdrw/core)作互操作交叉检查。
