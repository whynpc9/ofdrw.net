---
title: "CLI 命令与页号"
description: "安装并使用 ofdrw CLI 完成格式转换、文本提取、图片导出、签名检查和页面操作。"
sidebar:
  order: 20
  label: "CLI 命令"
---

CLI 需要 **.NET 10**。安装并查看所安装版本的完整帮助：

```sh
dotnet tool install --global Ofdrw.Net.Cli --version 0.1.0-preview.10
ofdrw --help
```

## 常用命令

| 任务 | 命令示例 |
| --- | --- |
| 按扩展名推断转换方向 | `ofdrw convert input.docx output.ofd` |
| DOCX → OFD | `ofdrw docx-to-ofd -i input.docx -o output.ofd --docx-ofd-mode native` |
| DOCX → PDF | `ofdrw docx-to-pdf -i input.docx -o output.pdf` |
| PDF → OFD | `ofdrw pdf-to-ofd -i input.pdf -o output.ofd` |
| OFD → PDF | `ofdrw ofd-to-pdf -i input.ofd -o output.pdf` |
| OFD → 单页 SVG | `ofdrw ofd-to-svg -i input.ofd -o page.svg --pages 1` |
| OFD → 静态 HTML | `ofdrw ofd-to-html -i input.ofd -o preview.html --pages 1-3` |
| OFD → PNG/JPEG | `ofdrw ofd-to-image input.ofd page.png --pages 1 --format png` |
| PNG/JPEG → OFD | `ofdrw image-to-ofd page1.png page2.jpg --output images.ofd` |
| UTF-8 文本 → OFD | `ofdrw text-to-ofd notice.txt notice.ofd --font-size-mm 4` |
| 提取文本 | `ofdrw extract-text input.ofd output.txt --include-templates` |
| 检查签名完整性 | `ofdrw verify-signatures -i input.ofd` |
| 重排全部页面 | `ofdrw reorder input.ofd reordered.ofd --pages 3,1,2` |
| 拆分选页 | `ofdrw split input.ofd selected.ofd --pages 3,1` |
| 合并文档 | `ofdrw merge merged.ofd first.ofd second.ofd` |
| 叠加水印 | `ofdrw watermark input.ofd marked.ofd --pages 1,2 --text DRAFT` |
| 清理签名 | `ofdrw clean-signatures input.ofd unsigned.ofd` |

文档操作的更多示例见[页面重组、水印与附件](https://whynpc9.github.io/ofdrw.net/guides/document-tools/)。文字替换需要显式可信原字体身份，见[文字定位与替换](https://whynpc9.github.io/ofdrw.net/guides/keyword-replacement/)。

## 页号和文档号

- **CLI 页号从 1 开始，API 页索引从 0 开始。**
- `--pages` 支持逗号与升序范围，例如 `1,3-5`。
- `reorder` 需要完整页列表；`split` 按列表顺序选择不同页面；转换命令是否允许重复页以具体命令为准。
- 多 `DocBody` 容器使用 `--document N`，同样从 **1** 开始，默认 `--document 1`。
- Native DOCX 的选页基于转换器实际分页，不要直接用 Word UI 页码盲选输出。

```sh
ofdrw ofd-to-pdf -i input.ofd -o second.pdf --document 2 --pages 1,3-5
ofdrw ofd-to-html -i input.ofd -o preview.html --pages 3,1,3 --title "Report"
```

## DOCX 参数

```sh
ofdrw docx-to-ofd -i input.docx -o output.ofd \
  --docx-ofd-mode native \
  --docx-engine built-in \
  --font-directory /path/to/fonts
```

- `--docx-ofd-mode`：`native`（默认）或 `dual-layer`。
- `--docx-engine`：`auto`、`word`（macOS）、`libreoffice` 或 `built-in`；Native 不调用这些视觉渲染引擎。
- `--font-directory` 可重复，用于提供实际字体文件。字体部署见[字体指南](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。
- 当前 `main` 的 `--native-embedded-bold-behavior throw` **尚未进入 preview.10 CLI**；preview.10 的严格真实粗体配置使用 API。

## 退出码与输出安全

| 退出码 | 含义 |
| --- | --- |
| `0` | 成功 |
| `1` | 参数、输入、转换或输出失败 |
| `130` | 用户取消，例如 Ctrl-C |

文件型命令先验证参数，再在目标目录写临时文件，成功关闭后提交；提交前失败或取消会保留原输出。格式转换的输入与输出必须是不同路径，页重排支持原位更新。

失败时保留完整命令、退出码和标准错误诊断。各格式的流写入、临时文件和资源预算边界见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。
