---
title: "Ofdrw.Net 文档"
description: "Ofdrw.Net .NET OFD SDK 与 CLI 的安装、转换、文档生成、编辑、签名和排障指南。"
sidebar:
  order: 0
---

Ofdrw.Net 是预览阶段的 .NET OFD SDK 与命令行工具，支持文档生成、读取、编辑、格式转换，以及可扩展的签名完整性检查。

- **第一次使用：** 从[安装与包选择](https://whynpc9.github.io/ofdrw.net/getting-started/installation/)开始，再运行[快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/)。
- **不想写代码：** 安装 [.NET CLI 工具](https://whynpc9.github.io/ofdrw.net/reference/cli/)，直接转换、提取文本或重组页面。
- **需要评估能力边界：** 查看[功能对照](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)和[已知问题](https://whynpc9.github.io/ofdrw.net/troubleshooting/known-issues/)。

## 选择使用方式

| 使用方式 | 适用场景 | 入口 |
| --- | --- | --- |
| NuGet SDK | 在 .NET 应用中生成、读取、编辑或转换 OFD | [包选择](https://whynpc9.github.io/ofdrw.net/getting-started/installation/)、[SDK 入口索引](https://whynpc9.github.io/ofdrw.net/reference/sdk/) |
| `ofdrw` CLI | 在终端或脚本中完成常见转换和文档操作 | [CLI 命令](https://whynpc9.github.io/ofdrw.net/reference/cli/) |
| 可选源码扩展 | 使用 PDF Vector、Skia、SES interop 或 Password 等独立能力 | [能力契约](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/) |

## 按任务阅读

| 任务 | 推荐阅读 |
| --- | --- |
| DOCX 与 OFD 互转、保留原文 | [DOCX 转换](https://whynpc9.github.io/ofdrw.net/guides/docx/)、[字体部署](https://whynpc9.github.io/ofdrw.net/guides/fonts/) |
| PDF、SVG、图片或 HTML 预览 | [格式转换](https://whynpc9.github.io/ofdrw.net/guides/conversion/)、[图片](https://whynpc9.github.io/ofdrw.net/guides/images/)、[HTML](https://whynpc9.github.io/ofdrw.net/guides/html-preview/) |
| 用代码生成报表 | [流式排版](https://whynpc9.github.io/ofdrw.net/guides/flow-layout/)、[表格](https://whynpc9.github.io/ofdrw.net/guides/public-tables/)、[原生绘图](https://whynpc9.github.io/ofdrw.net/guides/graphics/) |
| 读取文本与编辑已有 OFD | [模型与读取](https://whynpc9.github.io/ofdrw.net/reference/core-model/)、[文档工具](https://whynpc9.github.io/ofdrw.net/guides/document-tools/)、[文字替换](https://whynpc9.github.io/ofdrw.net/guides/keyword-replacement/) |
| 处理签名、加密或归档 | [签名完整性](https://whynpc9.github.io/ofdrw.net/guides/signatures/)、[可选口令加密](https://whynpc9.github.io/ofdrw.net/reference/password-crypto/)、[归档预检](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/) |
| 处理多文档和标准对象 | [Multi-DocBody](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)、[标准对象编辑](https://whynpc9.github.io/ofdrw.net/guides/standard-object-editing/) |

## 当前版本与适用范围

当前公开版本为 **0.2.0**，安装命令和快速开始均以该版本为基线。[发布说明](https://whynpc9.github.io/ofdrw.net/releases/020/)记录该版本的依赖、能力及验收范围。

从 preview.10 升级时，可在[已知问题与版本差异](https://whynpc9.github.io/ofdrw.net/troubleshooting/known-issues/)核对已修复的行为及仍有效的限制。

> 预览版不承诺完整国家/行业标准符合性、厂商互认、签章法律效力或任意复杂 Word 文档保真。请使用自己的输入、字体、部署方式和目标阅读器验证实际结果。

面向仓库维护者的[工程文档](https://github.com/whynpc9/ofdrw.net/tree/main/docs/engineering)、[设计决策](https://github.com/whynpc9/ofdrw.net/tree/main/docs/architecture)与[验收报告](https://github.com/whynpc9/ofdrw.net/tree/main/docs/validation)保留在 GitHub 仓库中。
