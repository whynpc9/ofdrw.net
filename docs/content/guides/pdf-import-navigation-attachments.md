---
title: "PDF 书签与附件导入"
description: "显式启用 PDF → OFD 的书签和嵌入附件复制，理解目标页映射、注释外观、预算和拒绝行为。"
sidebar:
  order: 160
  label: "PDF 书签与附件"
---

默认 PDF 导入和可选矢量导入都支持复制 PDF 大纲与嵌入附件，但 `CopyBookmarks`、`CopyAttachments` 默认均为 `false`，需要调用方显式启用。

## 启用复制

默认双层导入：

```csharp
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;

var options = new PdfToOfdOptions
{
    CopyBookmarks = true,
    CopyAttachments = true
};

await using var input = File.OpenRead("input.pdf");
await using var output = File.Create("output.ofd");
await new PdfToOfdConverter(options).ConvertAsync(input, output);
```

可选 `Ofdrw.Net.Converter.Pdf.Vector` 通过 `PdfVectorToOfdOptions.Compatibility` 暴露同一组配置。矢量扩展在确定原生页面或整页回退后，只对最终包应用一次元数据，不会因回退重复导入附件。

## 书签映射

启用 `CopyBookmarks` 后，PDF 大纲的层次和标题复制到 OFD `Outlines`。本地页目标映射到该源页在输出中第一次保留的出现位置。

支持的目标类型：

- `Fit`
- `XYZ`（两个坐标都有或都没有）
- `FitH`
- `FitV`
- `FitR`

`FitB`、`FitBH`、`FitBV` 会拒绝，因为页面适配的 OFD 类型无法复现其基于内容包围盒的缩放。

坐标使用页面旋转后的显示 CropBox，直接或间接数值均可。`FitH` 的空 `top` 和 `FitV` 的空 `left` 在旋转后保持未设置，以保留阅读器当前位置。栅格化页面也使用同一 CropBox，确保页面图与目标坐标一致。

`XYZ zoom` 在空、省略、零或位于 OFD 范围 0.1–64 内时保留，其他值会失败。目标超出选页、外部或命名目标、未知目标类型、不支持/链式动作，以及畸形或省略的大纲节点都会在写输出前失败。只有没有显式目标的节点才能成为分组大纲。

`MaxBookmarkCount` 和 `MaxBookmarkDepth` 限制复制规模。

## 附件与页面注释

启用 `CopyAttachments` 后：

- PDF 文档级嵌入文件复制为 OFD 嵌入附件。
- 所选页面的 `FileAttachment` 注释复制为附件，并在该页创建带 `GotoA` 动作的本地链接注释。
- 每次页面出现都有独立附件和链接；未选择页面的注释不复制。
- 直接或间接注释 subtype 名称可解析；缺失、无效或循环引用会在输出前失败。

页面注释只接受 PDF flag `/F 4`：可见、可打印且没有其他标志。Hidden、NoView、不可打印或其他组合会失败，因为栅格化后无法把来源图标从页面图中分离。

不同渲染路径的外观处理：

- Native vector 页在来源注释的显示矩形内绘制通用 PushPin。
- 外部 `pdftoppm` 栅格页保留 PDF 渲染图标，并使用不可见链接路径，避免重复图标。
- 默认 PDFium 栅格页绘制通用 PushPin，因为该渲染器不绘制注释图标。

自定义 PDF 注释外观不会复制为可编辑 OFD 对象。外部栅格化使用显示 CropBox，使裁剪或旋转页面的图像和链接坐标一致。

## 文件名、流与预算

- 文件名缩减为 basename；不区分大小写的冲突使用确定性的 ` (2)`、` (3)` 后缀。
- 长名称会缩短以保留后缀和扩展名，总长度不超过 255 个 UTF-16 code unit，且不拆开代理对。
- 支持字面量/十六进制 PDF 字符串及 UTF-16 名称。
- 支持 raw stream，以及直接 filter 名称或单元素 filter 数组声明的 Flate stream。
- 外部文件、其他 stream filter 和 filter 参数明确失败。

`MaxAttachmentCount`、`MaxAttachmentBytes`、`MaxTotalAttachmentBytes` 限制解码载荷。重复页面出现仍分别计入总预算；重复引用复用解码结果，编码字节由 `MaxInputBytes` 约束。输出保存解码后的字节，可在 OFD 读回后按哈希核对载荷身份。

生成的 `GotoA` 必须恰好指向一个保留的本地附件。`51` 与 `051` 这类数值等价 ID 有歧义，无效 ID 会拒绝。普通页面图形不会因此获得附件引用。

关闭任一开关就不会把该类别导入 OFD 模型。其他 PDF 注释类型仍遵循矢量模式的既有策略：失败或整页回退。

此能力的结论是覆盖输入上的模型与包读回。每个实际产物的视觉输出、macOS Preview 和独立 OFD 阅读器行为仍需分别验收。
