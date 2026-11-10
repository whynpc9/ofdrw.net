---
title: "CLI 命令与页号"
description: "安装并使用 ofdrw CLI 完成格式转换、文本提取、图片导出、签名检查和页面操作。"
sidebar:
  order: 20
  label: "CLI 命令"
---

CLI 需要 **.NET 10**。安装并查看所安装版本的完整帮助：

```sh
dotnet tool install --global Ofdrw.Net.Cli --version 0.3.1
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
| 本地归档预检（自 0.2.0 起） | `ofdrw preflight input.ofd --json --output report.json` |
| 重排全部页面 | `ofdrw reorder input.ofd reordered.ofd --pages 3,1,2` |
| 删除页面（自 0.2.0 起） | `ofdrw delete-pages input.ofd trimmed.ofd --pages 2,4-5` |
| 设置页面框（自 0.2.0 起） | `ofdrw crop input.ofd cropped.ofd --box 10,20,180,250 --pages 1` |
| 列出附件（自 0.2.0 起） | `ofdrw attachments list input.ofd --json` |
| 按 ID 导出附件（自 0.2.0 起） | `ofdrw attachments extract input.ofd --id 7 --output invoice.xml` |
| 只读定位关键字（自 0.2.0 起） | `ofdrw locate input.ofd --find 旧文本 --trusted-font SHA256:BYTECOUNT --json` |
| 按 ID 移除附件（自 0.2.0 起） | `ofdrw attachments remove input.ofd output.ofd --id 7 --physical` |
| 拆分选页 | `ofdrw split input.ofd selected.ofd --pages 3,1` |
| 合并文档 | `ofdrw merge merged.ofd first.ofd second.ofd` |
| 叠加水印 | `ofdrw watermark input.ofd marked.ofd --pages 1,2 --text DRAFT` |
| 清理签名 | `ofdrw clean-signatures input.ofd unsigned.ofd` |

文档操作的更多示例见[页面重组、水印与附件](https://whynpc9.github.io/ofdrw.net/guides/document-tools/)。文字替换需要显式可信原字体身份，见[文字定位与替换](https://whynpc9.github.io/ofdrw.net/guides/keyword-replacement/)。

## 附件

`attachments list` 每行输出 `ID、名称、媒体类型、字节数、internal|external`（制表符分隔），`--json` 时标准输出只有 `ofdrw-cli-attachments/1` JSON；没有附件时也返回 `0`。`attachments extract` 只按 ID 精确匹配，必须用 `--output` 指定目标路径，**从不**用附件内存储的名称作为路径；外部附件没有内嵌载荷，导出会失败。`attachments remove` 默认逻辑移除：从当前附件列表删除引用，保留历史 ZIP 条目；`--physical` 物理移除：在证明来源独占和引用闭合后才删除内部载荷，否则失败且不写输出。`list` 与 `extract` 可配合 `--document N` 读取多 `DocBody` 文档，`remove` 要求单 `DocBody` 文档。以上命令自 0.2.0 起可用。

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
- `0.2.0` 提供 `--native-embedded-bold-behavior throw`，用于要求 Native 嵌入粗体选用真实粗体字面。

## 退出码与输出安全

| 退出码 | 含义 |
| --- | --- |
| `0` | 成功 |
| `1` | 参数、输入、转换或输出失败 |
| `130` | 用户取消，例如 Ctrl-C |

`verify-signatures` 另有专用退出码：

| 退出码 | 含义 |
| --- | --- |
| `0` | 完整验证通过；或未声明签名且未指定 `--require-signatures` |
| `2` | 受保护条目摘要完整，但签名值算法没有已注册的 verifier |
| `3` | 签名无效或验证不完整 |
| `4` | 指定 `--require-signatures` 时文档未声明任何签名 |
| `1` / `130` | 同上表 |

`preflight` 的退出码（自 0.2.0 起）：

| 退出码 | 含义 |
| --- | --- |
| `0` | 运行完整，汇总结果为 `Pass` 或 `NotApplicable` |
| `2` | 至少一条规则为 `Fail` |
| `3` | 汇总为 `Unknown`，或运行未完成（例如输入超出预算或不是 ZIP） |
| `1` / `130` | 同上表 |

`preflight` 只执行 APF-001..006 本地规则，不是认证。`--json` 时标准输出只有 `ofdrw-cli-preflight/1` JSON（`schema`、`input` 的文件名/字节数/SHA-256、`ruleOutcomes` 和完整 `report`）；`--output` 需配合 `--json` 并原子写入文件。输入超过加载预算时不计算哈希（`sha256` 为 `null`）。`--allow-multiple-documents`、`--allow-encryption` 分别放宽单文档与无加密要求。

`locate` 的退出码（自 0.2.0 起）：`0` 结果完整且至少一处命中；`2` 结果完整但无命中；`3` 结果不完整；`1` 参数错误或 SDK 无法度量来源（如未信任/未嵌入字体）；`130` 取消。`--json` 时标准输出只有 `ofdrw-cli-locate/1` JSON（`complete`、一基页号的 `hits` 及其 `bounds`/`canReplace`、`diagnostics`）。`locate` 从不写文件，可作为 `replace-text` 的预览。

用于自动接收文档时，应加 `--require-signatures` 并只把 `0` 视为通过，否则被移除签名的文档也会返回 `0`。`0.2.0` 起，使用该参数时，无签章文档返回 `4`。

文件型命令先验证参数，再在目标目录写临时文件，成功关闭后提交；提交前失败或取消会保留原输出。格式转换的输入与输出必须是不同路径，页重排支持原位更新。

失败时保留完整命令、退出码和标准错误诊断。各格式的流写入、临时文件和资源预算边界见[转换约定](https://whynpc9.github.io/ofdrw.net/reference/conversion-contracts/)。
