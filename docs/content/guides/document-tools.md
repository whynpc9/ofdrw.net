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
