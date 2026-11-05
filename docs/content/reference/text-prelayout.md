---
title: "真实字体塑形与预排版查询（R21）"
description: "固定真实字体的 HarfBuzz 塑形、预排版查询、断行、预算及原生运行依赖。"
sidebar:
  order: 100
  label: "真实字体预排版"
---

此 API 提供独立的固定字体查询路径。现有 `FlowDocument.Render` 和 `OfdGraphics.DrawString` 继续保持近似测量或调用方定位的行为。该查询不用于定位导入 OFD 中的文字。

`Ofdrw.Net.Layout.Text` 包含不可变的字面、请求和结果类型，以及可替换的 `ITextShaper` 接口。可选的 `Ofdrw.Net.Text.HarfBuzz` 项目提供 `HarfBuzzTextShaper` 和 `HarfBuzzTextPrelayout.Create()`。它将 HarfBuzzSharp 以及 macOS、Win32 和 Linux 原生资产包固定为 `14.2.1.301`。原生库加载是运行时要求；原生库缺失或不兼容会产生 `TextLayoutError.NativeUnavailable`，绝不退回近似度量。

```csharp
var face = new TextFontFace("Noto Sans", File.ReadAllBytes(fontPath));
var request = new TextPrelayoutRequest("AV A\nTo", new[] { face },
    fontSizeMillimeters: 4, widthMillimeters: 50);
var result = HarfBuzzTextPrelayout.Create().Layout(request);
foreach (var line in result.Lines)
    Console.WriteLine($"{line.SourceStart}:{line.SourceLength} " +
        $"baseline={line.BaselineMillimeters} advance={line.AdvanceMillimeters}");
```

调用方提供许可明确的字体字节。`TextFontFace` 复制这些字节，按需选择 TTC 字面，验证有界的静态单色 TrueType 配置，并记录源字体和所选字面的 SHA-256。它不会发现系统字体，也不会合成字重或样式。显式字面列表可以为整个游程选择回退；`TEXT_FACE_FALLBACK` 记录所选字面。如果没有字面覆盖整个游程，查询就会失败。按簇混合字面回退不在首个配置范围内。

`TextShapeRequest` 接受一个逻辑 Unicode 游程、绝对 .NET UTF-16 源偏移、字面、方向、文字体系、语言和特性区间。结果保存实际 GID、以字体设计单位表示的 X/Y 步进和偏移、簇源区间、墨迹边界以及引擎包版本。簇可以是多对一，其顺序在 RTL 中可能递减。它们不是字符、字素或插入符索引。预排版结果报告总步进宽度和高度、最终行位置与基线、源区间、墨迹矩形、溢出、诊断，以及每个最终行实际使用的塑形游程。Y 只转换为向下的页面毫米一次。支持显式 LF/CRLF 和有界的拉丁文/汉字断行机会，包括截至 Unicode 18 的增补汉字；其他 Unicode 强制换行分隔符会被拒绝。选择断点后会重新塑形最终行。预排版拒绝其他文字体系标签；直接单游程塑形仍由调用方完成分析，不提供段落双向或断词约定。格式控制符、变体选择符以及可变、彩色、CFF 字体会明确失败。有界汉字断行配置允许在拉丁文本后的汉字前断行，使常见闭标点与前一字符保持相连、开标点与后一字符保持相连，并使不换行空格与两侧内容保持相连；它不宣称实现完整 Unicode 断行行为。

软换行处，末尾可断行空格仍保留在按源文本寻址的塑形游程及其报告的总步进中。适配和溢出检查使用这些空格之前的可见前缀，因此不可见的尾部空格不会使本应适配的单词变成溢出行。不换行空格从不是换行点。

`TextPrelayoutOfdWriter.CreateSinglePage` 仅为常规 ASCII 拉丁文消费已完成的结果，其中每个源字符映射到一个标称字形，且没有偏移或垂直步进。它预检整个页面，使用实测 Delta 和基线，在不移动笔位的情况下将每个 TextObject 边界扩展到实测墨迹，并嵌入实际选择的完整字面。连字、组合序列、不支持的样式或页面溢出都会在创建包之前被拒绝；页面预检同时检查最终步进和平移后的墨迹范围，包括负侧边距。首个有界类型化 GID/簇生成器、字形感知子集以及 PDF/SVG 样例消费程序见 [R22 塑形输出](https://whynpc9.github.io/ofdrw.net/reference/text-shaped-output/)。混合方向段落和更广泛的塑形仍属于独立的 R22 阶段；R21 查询成功不代表这些输出配置可用。E2E 还通过一个仅供样例使用、经过预检的“每个 BMP 汉字一个字形”OFD 探针渲染固定汉字样例，以检查回退、实测换行和基线几何。该探针不扩展公开写出器的 ASCII 输出约定。

API 强制实施以下限制：每个字体 32 MiB，每个游程/请求 4096 个 UTF-16 代码单元，256 行，调用方最多选择 8192 次塑形调用，整游程字面选择期间最多估算 800 万次 cmap 查询步骤，以及累计 512 MiB 塑形工作量。每个非空候选在调用塑形器之前，按两份完整字面副本以及每个 UTF-16 代码单元 64 个工作单位计费；因此，包含许多候选边界的宽汉字段落可能在达到调用次数上限前明确失败。原生字形展开另有独立上限。原生塑形前后及排版期间都会检查取消；同步原生调用无法在指令执行中途打断。查询失败不返回部分结果。文件和流写入继续遵循[能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)中独立的原子性约定。

此分支当前已核对的证据连同精确源码 SHA 和生成输出记录在 R21 PR 中；本文定义该配置，本身不证明跨平台原生加载或桌面阅读器验收已经通过。
