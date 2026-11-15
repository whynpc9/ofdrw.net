---
title: "填充命名区域"
description: "在 OFD 模板中预留命名区域，再用 SDK 或 CLI 填入文字或图片。"
sidebar:
  order: 210
  label: "命名区域"
---

命名区域是排版时为待填内容预留的固定矩形。生成模板后，可以按名称填入文字或图片，例如在表单中填写姓名或放入标志，区域之外的内容保持原位。

## 创建模板

`AreaHolderBlock` 占据正文区的整行宽度，构造参数为区分大小写的名称和以毫米为单位的高度。排版完成后，区域所在的页和矩形坐标会随 OFD 保存。

```csharp
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

var flow = new FlowDocument();
flow.Blocks.Add(new Paragraph("HEADER"));
flow.Blocks.Add(new AreaHolderBlock("name", 18));
flow.Blocks.Add(new AreaHolderBlock("logo", 15));
flow.Blocks.Add(new Paragraph("FOOTER"));

await using var template = File.Create("template.ofd");
await new OfdPackageWriter().WriteAsync(flow.Render(), template);
```

## 用 SDK 填充

`GetAreas` 返回名称、从零开始的页索引、毫米坐标与尺寸，以及填充状态。`FillImage` 需要调用方传入原图的像素宽高。例中的 `logo.png` 为 300 × 100 像素。

```csharp
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

await using var template = File.OpenRead("template.ofd");
var package = await new OfdReader().ReadAsync(template);
var editor = new OfdNamedAreaEditor();
foreach (var area in editor.GetAreas(package))
    Console.WriteLine($"{area.Name}: {area.IsFilled}");

editor.FillText(package, "name", "Alice 中文");
editor.FillImage(package, "logo", File.ReadAllBytes("logo.png"), "image/png", 300, 100);

await using var output = File.Create("filled.ofd");
await new OfdPackageWriter().WriteAsync(package, output);
```

## 用 CLI 填充（main，尚未发布）

```sh
ofdrw areas list template.ofd
ofdrw areas list template.ofd --json
ofdrw areas fill template.ofd with-name.ofd --name name --text "Alice 中文"
ofdrw areas fill with-name.ofd filled.ofd --name logo --image logo.png
```

列表每行有七列，以 Tab 分隔，依次是名称、从 1 开始的页码、矩形左上角的 X 和 Y、宽度、高度，以及 `blank` 或 `filled`。四个坐标和尺寸列均以毫米为单位。`--json` 输出 `ofdrw-cli-areas/1`。列出多文档 OFD 中的区域时，可用 `--document N` 选择从 1 开始编号的文档，默认选择第一份。填充只支持包含单个文档的 OFD。

每次调用 `fill` 选择一个区域，以及 `--text` 或 `--image` 中的一项。填文字时可用 `--font` 指定字体名称，用 `--font-size` 指定毫米字号，小数使用点号。

## 使用限制

- 命名区域来自创建模板时加入的 `AreaHolderBlock`，每个区域只能填充一次。
- 文字按近似字宽测量。放不下时填充失败，CLI 不会写入输出文件。OFD 仅引用字体名称，实际显示取决于渲染环境可用的字体。
- 图片支持单个不超过 16 MiB 的 PNG 或 JPEG，保持宽高比并居中放入区域。
- 已签名的 OFD 禁止填充。页重排、合并或拆分若使区域索引失效，填充也会被拒绝。
