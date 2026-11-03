---
title: "安装与包选择"
---

应用最常用的入口是转换元包：

```sh
dotnet add package Ofdrw.Net.Converter --version 0.1.0-preview.10
```

只需一个方向时，可直接安装 Converter.Docx、Converter.Pdf 或 Converter.Svg，使用同一版本。

| 包 | 用途 |
| --- | --- |
| Ofdrw.Net.Layout + Ofdrw.Net.Packaging | 段落、表格、图形生成与写包 |
| Ofdrw.Net.Reader | 读取已有 OFD、提取文本 |
| Ofdrw.Net.Signatures | 签名结构、摘要验证、provider 扩展 |
| Ofdrw.Net.Text.HarfBuzz | 固定字体塑形与有界查询 |
| Ofdrw.Net.Cli | 命令行工具 |

普通 SDK 目标为 netstandard2.0 / netstandard2.1，CLI 需要 .NET 10。具体宿主、RID 和原生依赖仍需核对所用包；目标框架声明不等于所有部署方式均已运行。

```sh
dotnet tool install --global Ofdrw.Net.Cli --version 0.1.0-preview.10
ofdrw --help
```

PDF Vector、Skia、Password、SES 是独立可选源码扩展。使用前核对各自目标框架、许可和本地构建版本，不假定公开源存在与主包同号的扩展包。

字体不随主包分发。中文 DOCX 部署先读[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。Native DOCX 不调用 Office；直接 DOCX→PDF 和 DualLayer 可能使用 Word、LibreOffice 或 BuiltIn 渲染路径。
