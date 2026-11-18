---
title: "页面重组、水印与附件"
description: "使用 CLI 或 SDK 重排、拆分、合并、叠加 OFD 页面，添加水印并理解附件和签名影响。"
sidebar:
  order: 110
  label: "页面与文档工具"
---

安装 [CLI](https://whynpc9.github.io/ofdrw.net/getting-started/installation/) 可完成常见文档操作；需要结果对象和诊断时，使用 `Ofdrw.Net.Layout.Editing` 中的 API。CLI 页号从 **1** 开始，API 页索引从 **0** 开始。

## 常用 CLI 操作

```sh
# 按新顺序输出全部页面
ofdrw reorder input.ofd reordered.ofd --pages 3,1,2

# 按列表顺序选择页面到新文档
ofdrw split input.ofd selected.ofd --pages 3,1

# 依次合并多个文档
ofdrw merge merged.ofd first.ofd second.ofd

# 将两个来源页按顺序叠加，使用第一个来源的页面尺寸
ofdrw mix overlay.ofd background.ofd 1 foreground.ofd 1

# 给第 1、2 页添加文字水印
ofdrw watermark input.ofd marked.ofd --pages 1,2 --text DRAFT

# 添加图片水印到前景层
ofdrw watermark input.ofd marked.ofd --image mark.png --layer-type Foreground

# 显式移除签名声明及可证明无引用的签名载荷
ofdrw clean-signatures signed.ofd unsigned.ofd

# 删除第 2、4、5 页（至少保留一页；自 0.2.0 起）
ofdrw delete-pages input.ofd trimmed.ofd --pages 2,4-5

# 把第 1、3 页的页面框设为 X,Y,宽,高（毫米）；省略 --pages 时作用于全部页面（自 0.2.0 起）
ofdrw crop input.ofd cropped.ofd --box 10,20,180,250 --pages 1,3
```

`delete-pages` 与 `crop` 分别对应 `OfdDocumentEditor.RemovePages`、`CropPage`；`crop` 只改变页面框，不移动或裁剪内容对象。三者与 `reorder` 一样支持原位更新。`reorder` 需要完整页列表，`split` 选择不同页面；`mix` 的参数是输出路径后的一组“输入文件 + 一基页号”。写入新路径后，应重新读取输出并核对页序、模板、资源和原文。

使用 `--skip-unsupported` 合并时，未知 raw 对象可能被跳过，必须检查警告；默认行为是拒绝无法安全重映射的内容，而不是静默丢失。

## SDK 入口

常用类型位于 `Ofdrw.Net.Layout.Editing`：

| 类型 | 用途 |
| --- | --- |
| `OfdDocumentEditor` | 删除、重排或裁剪（`CropPage`）内存包中的页面 |
| `OfdDocumentSplitter` | 按零基页索引列表生成独立包 |
| `OfdDocumentMerger` / `MergeWithResult` | 合并文档，并在显式跳过不支持对象时返回诊断 |
| `OfdDocumentMixer` | 将多个来源页叠加 |
| `OfdWatermark` | 向指定页面和图层添加文字或图片对象 |

编辑后的包使用 `OfdPackageWriter` 写出。基础读写示例见[模型与读取](https://whynpc9.github.io/ofdrw.net/reference/core-model/)，精确编辑与写包行为见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。

## 编辑、附件与签名边界

- 合并和混合需要证明资源引用闭包；无法安全重映射的未知引用会明确失败。
- 附件的**逻辑移除**与**物理载荷清理**是不同操作。共享载荷只有在证明没有其他引用时才能删除，详见[附件清理契约](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/r18-attachment-cleanup-design.md)。
- 删除页面不等于安全擦除信息；仍被模板、附件、扩展或其他页面引用的共享内容会保留。
- 裁剪、签名清理、水印和页面修改可能使原有签名失效。不要把修改后的文件描述为保留原签名效力；需要签名时对最终输出重新签署。
- 多文档容器先阅读 [Multi-DocBody 选择与隔离编辑](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。
- 类型化元数据、权限、版本和动作见[标准对象编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/)。
- 声明 `DocType="OFD-A"` 的输入不能直接编辑，见下节。

## 归档原件（OFD-A）编辑保护

自 0.5.0 起，通用改变内容入口保护归档原件；0.4.0 及更早版本会把源文档的 `DocType` 带到编辑、拆分和合并的输出中。

根节点 `OFD.xml` 声明 `DocType="OFD-A"` 的输入称为归档原件。改变内容的入口在产生任何输出之前拒绝归档原件，抛出 `Ofdrw.Net.Core.Models.OfdArchivalOriginalException`（继承 `InvalidOperationException`，`Operation` 为被拒绝的入口名）。错误消息写明操作名、输入声明 OFD-A，并指向下面的派生路径。CLI 写入命令遇到归档原件时在标准错误输出同一消息，不写输出文件，退出码为 **`5`**。

判断只看原始声明：`OfdDocTypes.IsArchival(value)` 对 `OFD-A` 做精确的序数比较，缺失属性不算归档；它位于 Core，普通 Layout/Converter 使用者不会因此引入归档包。判断不依赖、也不产生任何归档评估结论。

| 操作 | SDK 入口 | CLI 命令 | 输入声明 OFD-A | 普通 OFD | 签名处理 |
| --- | --- | --- | --- | --- | --- |
| 重排页面 | `OfdDocumentEditor.ReorderPages` | `reorder` | 拒绝 | 不变 | 既有行为不变：输出不保留原签名效力 |
| 删除页面 | `OfdDocumentEditor.RemovePages` | `delete-pages` | 拒绝 | 不变 | 同上 |
| 设置页面框 | `OfdDocumentEditor.CropPage(package, page, …)`（新重载） | `crop` | 拒绝 | 不变 | 同上 |
| 拆分选页 | `OfdDocumentSplitter.Split` | `split` | 拒绝 | 不变 | 同重排 |
| 合并 | `OfdDocumentMerger.Merge` / `MergeWithResult` | `merge` | 任一输入声明 OFD-A 即拒绝，与位置无关 | 不变 | 同重排 |
| 叠加 | `OfdDocumentMixer.Mix` | `mix` | 任一输入声明 OFD-A 即拒绝，与位置无关 | 不变 | 不复制原签名 |
| 文字替换 | `OfdTextReplacer.Replace`（全部重载） | `replace-text` | 拒绝 | 不变 | 同重排 |
| 填充命名区域 | `OfdNamedAreaEditor.FillText` / `FillImage` | `areas fill` | 拒绝 | 不变 | 同重排 |
| 水印 | `OfdWatermark.AddText` / `AddImage` | `watermark` | 拒绝 | 不变 | 同重排 |
| 清理签名 | `OfdSignatureCleaner.CleanAsync`、`OfdPackageSignatureCleaner.CleanAsync` | `clean-signatures` | 拒绝；不会无声删除归档原件的签名 | 不变 | 移除声明及可证明无引用的载荷 |
| 移除附件 | 无独立操作入口（见下文模型层说明） | `attachments remove` | 拒绝 | 不变 | 同重排 |
| 口令加密 | `OfdPasswordEnvelope.EncryptAsync` / `EncryptFileAsync`、`OfdStandardEncryption.EncryptAsync` / `EncryptFileAsync` | 无 | 拒绝（加密的 OFD-A 与 GB/T 42133 6.16 相矛盾） | 不变 | 已签名输入仍按既有规则拒绝 |
| 追加签名 | `OfdSignatureService` | 无 | **未决定，行为不变**（6.13 允许对归档文档签章，需单独决定） | 不变 | 不变 |
| 派生普通 OFD | `OfdOrdinaryDerivation.CreateAsync` | `derive-ofd` | 生成新的普通 OFD | 报错：无需派生，不写输出 | 拒绝已签名输入 |
| 归档受限副本 | `ArchivalCopyConverter.ConvertAsync` | 无 | 不加保护：显式的归档层处理器，写新文件并输出审计记录，原件不变 | 不变 | 拒绝已签名输入（既有行为） |
| 只读操作 | `OfdReader`、提取、定位、验签、预检、导出 PDF/PNG/SVG/HTML、`OfdNamedAreaEditor.GetAreas` | `extract-text`、`locate`、`verify-signatures`、`preflight`、`ofd-to-*`、`attachments list`/`extract`、`areas list` | 允许，保留原声明 | 不变 | 不变 |

以下入口不加保护，属于调用方责任：

- 模型层赋值和方法，例如直接修改 `OfdDocumentPackage` 的页面、元素、元数据，`OfdDocumentPackage.RemoveAttachment`，`OfdGraphics` 在已有页面上绘图，以及没有文档上下文的 `OfdDocumentEditor.CropPage(page, …)` 旧重载。它们与“读入后原样写出”共用同一模型，不能区分是否改变内容。
- `OfdPackageWriter` 不做判断。读入后不改内容直接写出会保留显式的 `OFD-A` 声明。
- 调用方在编辑前显式把 `Options.DocType` 设为 `OFD`，视为已作出选择，之后的编辑入口不再拒绝。此时输出不再声明 OFD-A，也不继承任何归档结论。

### 创建普通 OFD 派生副本

需要修改归档原件的内容时，先创建普通 OFD 派生副本，再编辑副本：

```sh
ofdrw derive-ofd archive.ofd working.ofd          # 文本记录
ofdrw derive-ofd archive.ofd working.ofd --json   # ofdrw-cli-derive-ofd/1 JSON
ofdrw delete-pages working.ofd trimmed.ofd --pages 2
```

```csharp
using Ofdrw.Net.Packaging;

var record = await OfdOrdinaryDerivation.CreateAsync("archive.ofd", "working.ofd");
// record.SourceSha256 / SourceLength / SourceDocType、OutputSha256 / OutputLength、
// record.Operation（DeriveOrdinaryOfd）、record.ChangedEntry（OFD.xml）
```

- 输出必须是新文件：与输入相同的路径、已存在的输出（包括指向输入的链接等别名）都会被拒绝，不覆盖任何已有文件。
- 只把 `OFD.xml` 根节点的 `DocType` 值从 `OFD-A` 改为 `OFD`，其余字节不变；其他 ZIP 条目内容逐字节相同。写出后重新读取核对条目集合、未变条目和根声明，并确认源文件哈希未变，才发布输出。
- 已签名输入（声明了签名或含签名载荷）和已加密输入直接拒绝，不删除签名，也不声称签名仍然有效。未声明 OFD-A 的输入报“无需派生”，不复制文件。
- 失败或取消时不留下输出文件，源文件不被修改。CLI 的这些错误返回 `1`，取消返回 `130`。
- 记录中的 `InheritsArchivalConclusion` 恒为 `false`：派生副本不继承原件的任何预检、规则或覆盖结论，也不继承签名效力。

### 边界

- `DocType="OFD-A"` 只是声明，不是符合性结果。本保护只依据声明工作，不验证文档是否真的符合 GB/T 42133；符合性检查见[本地归档预检](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/)。
- 本保护是本 SDK 和 CLI 的行为，不能证明任何第三方阅读器满足 GB/T 42133 第 7.2 条（阅读端禁止编辑 OFD-A 并提示）；阅读器行为须另行实测。
- 只对 `DocType` 精确等于 `OFD-A` 的输入生效；其他取值（包括大小写不同或带空白）按普通 OFD 处理。
