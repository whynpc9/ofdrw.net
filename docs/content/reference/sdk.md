---
title: "SDK 入口索引"
description: "按任务查找 Ofdrw.Net 常用公开类、所属包、完整命名空间及对应使用指南。"
sidebar:
  order: 10
  label: "SDK 入口"
---

本页帮助定位常用入口，不是自动生成的完整 C# API reference。公开成员的参数、重载和异常说明以随程序集分发的 XML 文档及所安装版本为准。

## 转换

| 入口 | 包 / 命名空间 | 用途与指南 |
| --- | --- | --- |
| `DocxToOfdConverter`、`DocxToPdfConverter` | `Ofdrw.Net.Converter.Docx` / `Ofdrw.Net.Converter.Docx.Converters` | [DOCX Native、DualLayer 与分页](https://whynpc9.github.io/ofdrw.net/guides/docx/) |
| `PdfToOfdConverter`、`OfdToPdfConverter` | `Ofdrw.Net.Converter.Pdf` / `Ofdrw.Net.Converter.Pdf.Converters` | [PDF 与 OFD 转换](https://whynpc9.github.io/ofdrw.net/guides/conversion/) |
| `OfdToImageConverter`、`ImageToOfdConverter` | `Ofdrw.Net.Converter.Pdf` / `Ofdrw.Net.Converter.Pdf.Converters` | [PNG/JPEG 导入导出](https://whynpc9.github.io/ofdrw.net/guides/images/) |
| `OfdToSvgConverter`、`OfdToHtmlConverter` | `Ofdrw.Net.Converter.Svg` / `Ofdrw.Net.Converter.Svg.Converters` | [单页 SVG](https://whynpc9.github.io/ofdrw.net/guides/conversion/)、[静态 HTML](https://whynpc9.github.io/ofdrw.net/guides/html-preview/) |
| `PlainTextOfdConverter` | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout` | [严格 UTF-8 纯文本转 OFD](https://whynpc9.github.io/ofdrw.net/guides/plain-text-ofd/) |

## 生成、读取与编辑

| 入口 | 包 / 命名空间 | 用途与指南 |
| --- | --- | --- |
| `FlowDocument`、`Paragraph`、`Span` | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout` | [流式排版](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/) |
| `Table`、`Row`、`Cell` | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout` | [流式表格](https://whynpc9.github.io/ofdrw.net/guides/public-tables/) |
| `OfdGraphics` 及绘图类型 | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout.Graphics` | [原生绘图](https://whynpc9.github.io/ofdrw.net/guides/graphics/) |
| `OfdDocumentBuilder` | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout.Builders` | 显式页面与毫米坐标构建 |
| `OfdReader` | `Ofdrw.Net.Reader` / `Ofdrw.Net.Reader.Readers` | [包读取与 Multi-DocBody 选择](https://whynpc9.github.io/ofdrw.net/reference/core-model/) |
| `OfdTextExtractor` | `Ofdrw.Net.Reader` / `Ofdrw.Net.Reader.Extraction` | [按页文本提取](https://whynpc9.github.io/ofdrw.net/reference/core-model/) |
| `OfdPackageWriter` | `Ofdrw.Net.Packaging` / `Ofdrw.Net.Packaging` | OFD 包写出及写包结果 |
| `OfdDocumentEditor`、`OfdDocumentSplitter`、`OfdDocumentMerger` | `Ofdrw.Net.Layout` / `Ofdrw.Net.Layout.Editing` | [页面重组、拆分与合并](https://whynpc9.github.io/ofdrw.net/guides/document-tools/) |

## 签名与高级文字

| 入口 | 包 / 命名空间 | 用途与指南 |
| --- | --- | --- |
| `OfdSignatureService`、`OfdSignatureVerifier` | `Ofdrw.Net.Signatures` / `Ofdrw.Net.Signatures.Signing`、`Ofdrw.Net.Signatures.Verification` | [签名与摘要完整性](https://whynpc9.github.io/ofdrw.net/guides/signatures/) |
| `HarfBuzzTextPrelayout` | `Ofdrw.Net.Text.HarfBuzz` | [固定字体预排版](https://whynpc9.github.io/ofdrw.net/reference/text-prelayout/) |
| `TextPrelayoutOfdWriter` | `Ofdrw.Net.Layout.Text` | [受限 typed 输出](https://whynpc9.github.io/ofdrw.net/reference/text-shaped-output/) |

先按[快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/)运行最小示例，再查相关主题。转换和编辑通常接收调用方拥有的 `Stream` 与可选 `CancellationToken`；流保持打开、最终复制失败可能留下部分输出，以及各类资源预算的详细规则见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。
