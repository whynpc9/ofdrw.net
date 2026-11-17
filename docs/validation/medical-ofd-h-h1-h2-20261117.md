---
title: "OFD-H 专用包 H1/H2 验收记录（2026-11-17）"
---

# OFD-H 专用包 H1/H2 验收记录（2026-11-17）

范围：[路线图](../engineering/plans/ofd-archival-roadmap.md) H1（`Ofdrw.Net.Medical` 只读评估）与 H2（OFD-H 写出 profile、`OfdMetadata.DocId`、`LinearizedEntryOrder`）。源码基线：`main` `9b556418` 上的 `feat/ofd-h-medical` 工作树，最终身份以 PR 合并提交 SHA 为准。环境：Linux（WSL2 5.15.167.4），.NET SDK 10.0.401，poppler `pdftoppm` 24.02.0；无 GUI。

结论：功能测试通过；Linux PNG 视觉复查未发现缺陷；**macOS Preview 验收未完成**（环境无 GUI），需在桌面环境补验。本记录不宣称任何文件符合 GB/T 48666—2026，评估整体结果按设计只有 Fail 或 Unknown。

## 功能测试

| 项目 | 结果 |
| --- | --- |
| `Ofdrw.Net.Medical.Tests` | 39/39 通过（含 PR 评审修复的回归用例；含 DOCX 显式 native 与默认模式各 1 例，需物化的固定 CJK 字体） |
| `Ofdrw.Net.Archival.Tests` | 74/74 通过（opt-in 源码级依赖测试已允许 Medical → Archival，并禁止任何项目引用 Medical） |
| 全解决方案其余测试 | Core 163、Packaging 1785、Layout 443、Cli 400（1 跳过）、Converter.Pdf 1266（17 跳过）等全部通过；`Converter.Docx.Tests` 1 例失败，见下 |
| `python3 -m unittest discover -s scripts/tests` | 通过（包数 14，新增 Medical 依赖闭包接受/拒绝用例） |
| 文档站 `npm ci` / `build` / `typecheck` / `check` | 通过，53 页、链接检查通过；Nimbus 未 lint `.md` 正文、未运行 `astro check`，属于其已声明的部分检查 |

`BuiltInPdfItalicRendererTests.DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam` 在本机期望 98.5816、实际 98.5815（3 位小数舍入边界）。在未修改的 `origin/main` 独立 worktree 上重复两次得到相同失败，与本分支无关，留待 CI 结果确认是否为本机浮点环境差异。

## 视觉验收

样例：`e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（中英文、比例字体、标题颜色、斜体副标题、红色粗体局部样式、表格底色与边框、右对齐日期、两页分页），字体为固定 Noto Sans CJK SC fixture，不使用宿主字体。

链路：`DOCX → native OFD（OfdNamespace=/2016，显式 Native 与默认模式各一）→ OfdReader → MedicalDocumentProfile.Apply → OfdPackageWriter → OFD-H → ofdrw ofd-to-pdf → pdftoppm 100 dpi PNG`；对照组为同一转换未应用 profile 的普通 OFD，经同一 PDF/PNG 链路。

| 检查 | 结果 |
| --- | --- |
| 自动渲染比较 | 两种模式各 2 页，OFD-H 与普通 OFD 的 PNG 像素差 `AE=0`；`pdftotext` 文本哈希一致 |
| OFD 原始文本 | 重新读取 OFD-H 后页数与全部 `TextObject` 文字与普通 OFD 一致（回归断言） |
| PNG 逐页人工复查 | 4 页（native/default × 第 1、2 页）：中文无缺字乱码；英文间距、断行正常；标题色、斜体副标题、红色粗体未扩散；表头底色与表格边框完整；日期右对齐；第 2 页分页位置正确；无裁切、重叠、重影或空白页 |
| macOS Preview | **未完成**：无 GUI；需打开下列 OFD-H 经 PDF 的产物逐页复查 |
| 条目顺序 | OFD-H：`OFD.xml`、`Doc_0/Document.xml`、`Doc_0/PublicRes.xml`、`Page_0/Content.xml`、`Doc_0/Res/Font_*.ttf`、`Page_1/Content.xml`；普通 OFD 保持原有字母序 |
| 产物体积 | OFD-H 290,926 字节，普通 OFD 290,761 字节（增加 165 字节，来自 DocID、CustomDatas 与 `/2016` 命名空间） |

产物保存在未跟踪的 `artifacts/medical-validation/`（`outputs/`、`pdf/`、`png/`），供桌面补验：

| 文件 | SHA-256 |
| --- | --- |
| `outputs/generated-layout-native-ofd-h.ofd` | `4c0a802a911713ab028a95dbbfde28b330f6a24ae8e042ad821d4acc5d0a3a9d` |
| `outputs/generated-layout-default-ofd-h.ofd` | `33e85cc271e001477940e5b3e8d305d69940093e6e94b6716ffdb4ddb83885e5` |
| `pdf/native-ofd-h.pdf` | `7ba540ee2d5c1e4dfc5dccde3b185a43983a4706fec66b7018f7ea84274bba0d` |
| `pdf/default-ofd-h.pdf` | `deec90e15868dc67e21b53bcf0fd7cc710000fd6b8a0a02dbc8f1d7a06443cf8` |
| `png/*-ofd-h-1.png`（两模式相同） | `7a87dc82edce64457445670bdd183e0f47fd186dc64cada848b6e07dfd34a96f` |
| `png/*-ofd-h-2.png`（两模式相同） | `0a91e2d939dd14f4d8cdb50c0991599ed42a3d3762b084611afcc50bcfd17c35` |

## 发现与处理

- 读取得到的包不能改命名空间：`OfdSourceIdentity.Reconcile` 抛出 `Imported source namespace cannot be remapped safely`。未放宽该保护；DOCX 路线改为转换时设置 `DocxConversionOptions.OfdNamespace = /2016`，指南与设计记录已写明。
- 只检查了上述 1 个样例、2 种模式、各 2 页；不推断任意 Word 文档经 profile 后均保真，也不代表 OFD-H 符合性。
