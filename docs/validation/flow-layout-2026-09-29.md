# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最新复验：** Review 6 指出自动折行的尾空格使居中/右对齐左偏。本轮将自动折行尾空格推进量归零并保留原始字符，公开 Flow 与 DOCX Native 共用修复。已在 macOS Preview 重新打开 `artifacts/flow-layout/review7/` 六份最新 PDF，检查原有 7 页及新增的 3 页对齐样例，未见本次样例的视觉缺陷。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次产物对应 `codex/public-flow-layout` 的 `4cb7b27` 代码状态。后续文档提交不改动生成代码。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review7`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。六份 PDF 均从 `artifacts/flow-layout/review7/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：154 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review7/test-results/`。其中 Layout 测试 15/15，DOCX 测试 61/61，覆盖公开样式/往返、跨 Span 英文单词、CRLF 与组合字素、段末显式换行、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。本轮新增 9 项回归覆盖 Center/Right 的尾分隔空格、跨 Span 空格样式、首行缩进、文本保留、元素右边界，以及自动折行与显式换行/末行的空白推进量差异。Native 往返右边界断言允许 OFD 的 0.001 mm 序列化精度。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review7/package-e2e.log` 和 `package-e2e-output/`。
- 新增样例由同一 E2E 程序生成 `alignment-public.ofd`、`alignment.docx` 及其显式 Native/default OFD。三份对齐产物均为 1 页、95 个非空白字符；E2E 断言两组参考 `Alpha` 与折行 `Alpha  ` 的 X 坐标一致，并保留两个分隔空格。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review7/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review7/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review7/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |
| `review7/alignment-public.ofd` → `alignment-public.pdf` | 1 / 1 | 放大检查居中/右对齐两组：参考 `Alpha` 与折行第一行 `Alpha` 水平位置一致；`information` 正常换行；无右边界裁切或重影。 |
| `review7/alignment-native.ofd` → `alignment-native.pdf` | 1 / 1 | DOCX 显式 Native 的居中/右对齐参考与折行位置一致，无缺字或裁切。 |
| `review7/alignment-default.ofd` → `alignment-default.pdf` | 1 / 1 | 默认 DOCX 的两组对齐外观与显式 Native 一致，无异常空白页或重复文字。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review7/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,559 | `70d96629f3cb9799e2653024875b3120a4a475313af66c6e1c1703236b722749` |
| `flow-public.pdf` | 198,572 | `a83499dd230dd0c818fdfa33ab738f27a6564ae96469d2a0097ceeb60c5f97a6` |
| `docx-native.ofd` | 15,288,612 | `3c11089bf3169c48d39c1ba87dcd29c8b4d7c36638c823aa98e42c258c3bc0a8` |
| `docx-native.pdf` | 117,446 | `fe737b60f980fc5458640425c3021bc001035a9a52a9b58915e641ddcb013ee1` |
| `docx-default.ofd` | 15,288,607 | `bebf1975935cc3ff0ace6c221a2952c1a95f6e8099ff00a9c49d45b32fbf705f` |
| `docx-default.pdf` | 117,446 | `8874209d8b45ded1afb05b920c457b4c67c25801e021e716d33f90e9f092ae75` |
| `alignment-public.ofd` | 1,518 | `632cc4fd61acaed0034990229755ffa5e633263bd71b785004e75ca8c2b05715` |
| `alignment-public.pdf` | 68,677 | `077257ca19b230a5cb62e542353fc8479fbcd94394410818bd7454bacdf0c2c0` |
| `alignment-native.ofd` | 15,286,740 | `8bbb7f0951d1cd65b8c1854b99d13a1fd75c6b92f07fa9c501e6a8605e692509` |
| `alignment-native.pdf` | 68,336 | `35a38d4e24f2458d84b2334993fd43076f25e33290f4b51eba736fafa5ece41a` |
| `alignment-default.ofd` | 15,286,740 | `c00a8406dec65a7e455d67423b820b4a9a368a4469bedd1686b703ed610ea7b1` |
| `alignment-default.pdf` | 68,336 | `60ee181247720203d595524744783d761de1e4ead3a69ce8f7af1b5eb11d1722` |

页面图：`review7/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。原有七张 PNG 与先前 `review6/pages/` 逐字节一致；另保留 `alignment-public-1.png`、`alignment-native-1.png`、`alignment-default-1.png`，但最终结论依据本次重新打开的 Preview 页面。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。
