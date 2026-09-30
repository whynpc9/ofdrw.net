# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最终复验：** Review 5 的段末显式换行修复生成了 `artifacts/flow-layout/review6/`；本次已在 macOS Preview 重新打开三份最新 PDF，逐页检查公开 Flow 1–3 页、显式 Native 1–2 页、默认 DOCX 1–2 页。此前锁屏造成的 Preview 验收缺口已关闭。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次产物对应 `codex/public-flow-layout` 的 `4b93eef` 代码状态。后续文档提交不改动生成代码。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review6`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。三份 PDF 均从 `artifacts/flow-layout/review6/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：145 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review6/test-results/`。其中 Layout 新测试 12/12，DOCX 测试 55/55，覆盖公开样式/往返、跨 Span 英文单词、CRLF 与组合字素、段末显式换行、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review6/package-e2e.log` 和 `package-e2e-output/`。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review6/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review6/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review6/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成三份 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review6/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,561 | `59a15fb7f0d7a773f06887c5615c5eeacc881c6bed0ea521aee1892964e90c57` |
| `flow-public.pdf` | 198,572 | `cc68c8cb0bf75be7a62f171b359a37474b9542551bfb55e7c116ca87ab0a3984` |
| `docx-native.ofd` | 15,288,612 | `35b80ff9e86276a50a857608d58ee2ad2a393a83cdb01e032373a9e1a5f55bfe` |
| `docx-native.pdf` | 117,446 | `32e1ed2a45087acd9700c588bccc937fafadccd21976f3a73ef93db02166fe5a` |
| `docx-default.ofd` | 15,288,607 | `09d651186f7ae8049fc24644eb54932863f6f78702f78ec29726ada3a18501ae` |
| `docx-default.pdf` | 117,446 | `cce4b21a26fbfc092f9e160f7ae1f3404a1486f1c88f2467db399fcb98e73130` |

页面图：`review6/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。七张 PNG 与先前 `review2/pages/` 逐字节一致，但最终结论依据本次重新打开的 Preview 页面。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。
