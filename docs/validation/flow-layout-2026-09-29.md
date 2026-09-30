# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最新复验：** 已修复自动折行尾空格导致的居中/右对齐左偏，以及下一轮 review 指出的不换行空格断行。NBSP、窄 NBSP 和数字空格保留推进量并连接相邻文字；最新 review 进一步指出通用 0.6 em 不适合空格变体，现已为窄 NBSP 使用 0.2 em、数字空格使用所选样式数字字宽，Native 同步规则；Preview 随后检出窄 NBSP 在缺字字体下的 PDF 方框，本轮补了定位空白字素绘制修复。已重新打开 `artifacts/flow-layout/review10/` 六份最新 PDF，逐页复验原有 7 页及 3 页对齐/不换行空格样例，未见本次样例的视觉缺陷。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次产物对应 `codex/public-flow-layout` 的 `f148f8d` 代码状态。后续文档提交不改动生成代码。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review9`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。六份 PDF 均从 `artifacts/flow-layout/review10/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：180 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review10/test-results/`。其中 Layout 测试 25/25，DOCX 测试 71/71，PDF 测试 47/47，覆盖公开样式/往返、跨 Span 英文单词、CRLF 与组合字素、段末显式换行、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。上一轮新增 9 项回归覆盖 Center/Right 的尾分隔空格、跨 Span 空格样式、首行缩进、文本保留、元素右边界，以及自动折行与显式换行/末行的空白推进量差异。Native 往返右边界断言允许 OFD 的 0.001 mm 序列化精度。上一轮另增 18 项回归：三种不换行空格的跨 Span/Native/default 折行与字符保留，中文邻接及首行缩进超宽契约，以及 PDF 普通/模拟粗体空白字素的像素与语义检查。最初四项 PDF 像素回归在临时恢复旧绘制条件时 4/4 失败；恢复修复后通过，最终三种空格的六项像素回归全部通过。旧条件失败日志保存在 `review9/before-fix-pixel-regression.log`。本轮新增 8 项回归覆盖四种粗斜体样式：普通 NBSP = 普通空格、窄 NBSP = 0.2 em、数字空格 = 当前样式数字字宽；公开 API 在修正字宽加 0.001 mm 的临界行宽下能完整渲染，并检查 OFD `DeltaX` 与对象宽度；Native 同时验证 `0` + 数字空格 + `0` 与 `000` 等宽。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review10/package-e2e.log` 和 `package-e2e-output/`。
- 新增样例由同一 E2E 程序生成 `alignment-public.ofd`、`alignment.docx` 及其显式 Native/default OFD。三份对齐产物均为 1 页、132 个非空白字符；E2E 断言两组参考 `Alpha` 与折行 `Alpha  ` 的 X 坐标一致，并保留两个普通分隔空格；还断言 `A` 与 `B` 通过三种不换行空格连接时整组移至下一行，空格推进量大于零。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review10/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review10/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review10/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |
| `review10/alignment-public.ofd` → `alignment-public.pdf` | 1 / 1 | 放大检查居中/右对齐两组：参考 `Alpha` 与折行第一行 `Alpha` 水平位置一致；`information` 正常换行；三种不换行空格的 `A B` 整体换行且保留间距，窄 NBSP 方框已消失且间距比普通 NBSP 更窄，数字空格间距按数字字宽保留；无右边界裁切或重影。 |
| `review10/alignment-native.ofd` → `alignment-native.pdf` | 1 / 1 | DOCX 显式 Native 的居中/右对齐参考与折行位置一致；三种不换行空格连接的文字保持同一行，窄 NBSP 间距小于普通空格，数字空格使用数字字宽；无方框或裁切。 |
| `review10/alignment-default.ofd` → `alignment-default.pdf` | 1 / 1 | 默认 DOCX 的两组对齐及三种不换行空格外观与显式 Native 一致，无方框、异常空白页或重复文字。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review10/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,559 | `a5fc243ae60874bf774f002b2b7fe8c7f1b679dbf6cacefd826d2df17253047d` |
| `flow-public.pdf` | 199,181 | `09047f8c9b233aa97969e5118a6c99e7926d973d637fa20c71c4467012ca4135` |
| `docx-native.ofd` | 15,288,611 | `d58b6c7100fd9cdd03987f12ccd45ec93bb847509ff76ee61f3cc483ba88bdd9` |
| `docx-native.pdf` | 117,510 | `46b7744ff01b99c4e4063bd16e45ecbec6fc663d9f0517015fd43b15038af9c4` |
| `docx-default.ofd` | 15,288,607 | `4ef1dfd5c4b1ea56dcf08a52066807fd445acbcc8118ef9694122ddcf8d0563d` |
| `docx-default.pdf` | 117,510 | `d37122a969311408fc42eabdfc9667ba1a12ef61e3b3cb4c691a5710ec2a22de` |
| `alignment-public.ofd` | 1,676 | `c8ee2a32502f8be2d63b726c549971bf82e692451f74296abcc3208659e6c63b` |
| `alignment-public.pdf` | 70,787 | `eee4a49b84c39578f569985188fa991c9214e251282166c55af54b599fc2ecaf` |
| `alignment-native.ofd` | 15,286,925 | `fb188eb569d6e8bd0e987c817365a05577f97edb04ff992b5d8407aa4ef20baa` |
| `alignment-native.pdf` | 70,261 | `6cc40b34f9d938dbd7e5049579e583690f6337bedc2c04e1b2f8b4973c6cd471` |
| `alignment-default.ofd` | 15,286,925 | `84b22e4f8859155a34c3c8d7d668bb8e576d654db2a7eefdcaabc2e9430c8eeb` |
| `alignment-default.pdf` | 70,261 | `fcc4eb2704af7006b40a552991c21085d2295b21d9c9a76e809f7f378d86e2cf` |

页面图：`review10/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。原有七张 PNG 与先前 `review7/pages/` 逐字节一致；另保留 `alignment-public-1.png`、`alignment-native-1.png`、`alignment-default-1.png`，但最终结论依据本次重新打开的 Preview 页面。对比上一轮，公开 Flow PDF 增加 609 字节（约 0.31%），原 DOCX PDF 增加 64 字节（约 0.05%），来自空白字素的局部剪裁指令；原有页面像素不变。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。

本次 PDF 空白字素修复与像素回归覆盖 Flow/Native 的 `DeltaX/DeltaY` 定位字素及纯空白文本。通用 OFD 无定位的混合文本 run 仍受所绑定字体的缺字影响；该路径未作为本票完整字体保真验收。首版也未实现完整 Unicode 行断算法。
