---
title: "安装与包选择"
description: "按转换、生成、读取、签名或 CLI 场景选择 Ofdrw.Net NuGet 包，并核对运行时、字体和可选扩展要求。"
sidebar:
  order: 10
---

本页安装命令以 **0.5.0** 为基线。普通 SDK 包（含 Archival、Medical）面向 `netstandard2.0` / `netstandard2.1`，CLI 需要 **.NET 10**。目标框架兼容不代表所有操作系统、RID、AOT 或单文件发布方式都已验证。

## 安装转换元包

需要 DOCX、PDF 和 SVG 转换时，安装聚合包即可：

```sh
dotnet add package Ofdrw.Net.Converter --version 0.5.0
```

该元包引用 `Converter.Docx`、`Converter.Pdf` 和 `Converter.Svg`，适用于[快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/)中的转换示例。

## 按任务选择包

只需一个能力时，直接安装对应包。自 0.4.0 起，同系列包依赖精确锁定，请在同一个项目中一起升级所有 `Ofdrw.Net.*` 包。

| 任务 | 包 | 主要入口 |
| --- | --- | --- |
| DOCX ↔ PDF、DOCX → OFD | `Ofdrw.Net.Converter.Docx` | `DocxToOfdConverter`、`DocxToPdfConverter` |
| PDF → OFD、OFD → PDF/PNG/JPEG、图片 → OFD | `Ofdrw.Net.Converter.Pdf` | `PdfToOfdConverter`、`OfdToPdfConverter`、图片转换器 |
| OFD → SVG/静态 HTML | `Ofdrw.Net.Converter.Svg` | `OfdToSvgConverter`、`OfdToHtmlConverter` |
| 流式排版、表格、绘图、页面编辑 | `Ofdrw.Net.Layout` | `FlowDocument`、`OfdDocumentBuilder`、`OfdGraphics` |
| 写出 OFD 包 | `Ofdrw.Net.Packaging` | `OfdPackageWriter` |
| 读取 OFD、提取文本 | `Ofdrw.Net.Reader` | `OfdReader`、`OfdTextExtractor` |
| 签名结构与摘要完整性检查 | `Ofdrw.Net.Signatures` | `OfdSignatureService`、`OfdSignatureVerifier` |
| 固定字体塑形与有界查询 | `Ofdrw.Net.Text.HarfBuzz` | `HarfBuzzTextPrelayout` |
| 只读归档评估（自 0.5.0 起，显式可选） | `Ofdrw.Net.Archival` | `ArchivalAssessmentService` |
| OFD-H 写出 profile 与只读评估（自 0.5.0 起，显式可选） | `Ofdrw.Net.Medical` | `MedicalDocumentProfile`、`MedicalAssessmentService` |

例如，只生成和写出 OFD：

```sh
dotnet add package Ofdrw.Net.Layout --version 0.5.0
dotnet add package Ofdrw.Net.Packaging --version 0.5.0
```

只读取已有 OFD：

```sh
dotnet add package Ofdrw.Net.Reader --version 0.5.0
```

常用类的完整命名空间见 [SDK 入口索引](https://whynpc9.github.io/ofdrw.net/reference/sdk/)。随程序集分发的 XML 文档可用于查看具体公开成员。

## 安装 CLI

先安装 .NET 10，再安装全局工具并检查帮助：

```sh
dotnet tool install --global Ofdrw.Net.Cli --version 0.5.0
ofdrw --help
```

如已安装 CLI，可用以下命令更新至 0.5.0。

```sh
dotnet tool update --global Ofdrw.Net.Cli --version 0.5.0
```

常用命令和页号规则见 [CLI 参考](https://whynpc9.github.io/ofdrw.net/reference/cli/)。

## 字体、原生依赖与可选扩展

- **字体不随主包分发。** 处理中文 DOCX 或要求跨机器一致显示时，准备许可明确的实际字体文件，并按[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)配置。仅设置 `FontFamily` 名称不等于提供字体字节。
- **Native DOCX 不调用 Office。** 直接 DOCX→PDF 和 DualLayer 可能使用 Word、LibreOffice 或 BuiltIn 渲染路径；图片导出还依赖 PDF/PDFium 原生链路。
- **独立源码扩展单独核对。** PDF Vector、Skia、Password、SES interop 不在 0.5.0 的默认 14 包发布集合中，不能假定存在与主包同号的公开扩展包。
- **Archival 与 Medical 自 0.5.0 起加入默认发布集合。** 两者仍需显式引用，`Ofdrw.Net.Converter` 等普通包不会传递带入。Archival 只做只读归档评估，不生成 OFD-A；Medical 提供 GB/T 48666 OFD-H 写出 profile 与只读评估，并依赖 Archival。两者都不等于合规认证，详见[归档评估](https://whynpc9.github.io/ofdrw.net/guides/archival-assessment/)和[电子病历 OFD-H](https://whynpc9.github.io/ofdrw.net/guides/medical-ofd-h/)。

按需安装两个可选包：

```sh
dotnet add package Ofdrw.Net.Archival --version 0.5.0
dotnet add package Ofdrw.Net.Medical --version 0.5.0
```

只需 Medical 时不必另行直接引用 Archival；它会以精确的同批版本依赖带入 Archival、Packaging 和 Core。项目中的其他 `Ofdrw.Net.*` 直接引用仍应一起升级到 0.5.0。

下一步：运行[快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/)，或按任务查看 [DOCX](https://whynpc9.github.io/ofdrw.net/guides/docx/)、[格式转换](https://whynpc9.github.io/ofdrw.net/guides/conversion/)和[流式排版](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/)指南。
