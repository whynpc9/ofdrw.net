# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最新状态：** Review 3 的段落换页修复及 Hangul Jamo 修复已生成 `artifacts/flow-layout/review4/` 产物；七张页面 PNG 与下文已在 Preview 检查的 `review2` 页面逐字节相同。macOS 当前锁屏，`review4` PDF 尚未能在 Preview 重新打开。最新 head 的 Preview 复验仍待完成，不能把下文的 `review2` 视觉结果当作最新 head 已验收。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本记录检查的是 `codex/public-flow-layout` 工作树在 PR 提交前的实现。PR 头提交以 Git 历史为准。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review2`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。三份 PDF 均在第一波 review 修复后从 `artifacts/flow-layout/review2/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：142 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review3/after-jamo-test-results/`。其中 Layout 新测试 10/10，DOCX 测试 54/54，覆盖公开样式/往返、跨 Span 英文单词、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review3/after-jamo-package-e2e.log` 和 `after-jamo-package-e2e-output/`。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review2/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review2/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review2/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成三份 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review2/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,561 | `b8d32c8254d1ec5a9dfaa3fb12295d82673753edc94a4b9f0ae920fb653e6a12` |
| `flow-public.pdf` | 198,572 | `c8486c0aefd30a5d90a571a4656e618d3e5afe3577b1928e96f73212b9fa258a` |
| `docx-native.ofd` | 15,288,612 | `db7f272c0530a8a914d890c0d284cfcee113893ce43a57260a4be2fb5ed271f3` |
| `docx-default.ofd` | 15,288,608 | `afbcb790e2692c31ad45f61edc2b7fc0c3756f6f82d797ea03604e137ca672c0` |

页面图：`review2/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

`review4/` 是当前工作树生成的最新一组样例：`flow-public.ofd` 5,561 字节，SHA-256 `2129a0219058b5cefbafa618085f677269ed5d18ea5b29c7a873ba228b814384`；`docx-native.ofd` 15,288,612 字节，SHA-256 `a40a8b52b4310b3a20128dac9d7ee1b555b8c0176d7bc0c9f621525379a2b785`；`docx-default.ofd` 15,288,607 字节，SHA-256 `ed5e887a0c8f20a09219b57bb903d512312ac8f02f85c655bf9d19e83e22ad27`。其七张 PNG 与 `review2/pages/` 逐字节一致；仍须在 Preview 打开 `review4/flow-public.pdf`（1–3 页）及 `review4/docx-native.pdf`、`review4/docx-default.pdf`（各 1–2 页）。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。
