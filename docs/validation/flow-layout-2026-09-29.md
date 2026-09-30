# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最新复验：** 已修复自动折行尾空格导致的居中/右对齐左偏，以及下一轮 review 指出的不换行空格断行。NBSP、窄 NBSP 和数字空格保留推进量并连接相邻文字；Preview 随后检出窄 NBSP 在缺字字体下的 PDF 方框，本轮补了定位空白字素绘制修复。已重新打开 `artifacts/flow-layout/review9/` 六份最新 PDF，逐页复验原有 7 页及 3 页对齐/不换行空格样例，未见本次样例的视觉缺陷。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次产物对应 `codex/public-flow-layout` 的 `dc60104` 代码状态。后续文档提交不改动生成代码。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review9`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。六份 PDF 均从 `artifacts/flow-layout/review9/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：172 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review9/test-results/`。其中 Layout 测试 21/21，DOCX 测试 67/67，PDF 测试 47/47，覆盖公开样式/往返、跨 Span 英文单词、CRLF 与组合字素、段末显式换行、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。上一轮新增 9 项回归覆盖 Center/Right 的尾分隔空格、跨 Span 空格样式、首行缩进、文本保留、元素右边界，以及自动折行与显式换行/末行的空白推进量差异。Native 往返右边界断言允许 OFD 的 0.001 mm 序列化精度。本轮另增 18 项回归：三种不换行空格的跨 Span/Native/default 折行与字符保留，中文邻接及首行缩进超宽契约，以及 PDF 普通/模拟粗体空白字素的像素与语义检查。最初四项 PDF 像素回归在临时恢复旧绘制条件时 4/4 失败；恢复修复后通过，最终三种空格的六项像素回归全部通过。旧条件失败日志保存在 `review9/before-fix-pixel-regression.log`。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review9/package-e2e.log` 和 `package-e2e-output/`。
- 新增样例由同一 E2E 程序生成 `alignment-public.ofd`、`alignment.docx` 及其显式 Native/default OFD。三份对齐产物均为 1 页、132 个非空白字符；E2E 断言两组参考 `Alpha` 与折行 `Alpha  ` 的 X 坐标一致，并保留两个普通分隔空格；还断言 `A` 与 `B` 通过三种不换行空格连接时整组移至下一行，空格推进量大于零。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review9/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review9/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review9/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |
| `review9/alignment-public.ofd` → `alignment-public.pdf` | 1 / 1 | 放大检查居中/右对齐两组：参考 `Alpha` 与折行第一行 `Alpha` 水平位置一致；`information` 正常换行；三种不换行空格的 `A B` 整体换行且保留间距，窄 NBSP 方框已消失；无右边界裁切或重影。 |
| `review9/alignment-native.ofd` → `alignment-native.pdf` | 1 / 1 | DOCX 显式 Native 的居中/右对齐参考与折行位置一致；三种不换行空格连接的文字保持同一行，无方框或裁切。 |
| `review9/alignment-default.ofd` → `alignment-default.pdf` | 1 / 1 | 默认 DOCX 的两组对齐及三种不换行空格外观与显式 Native 一致，无方框、异常空白页或重复文字。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review9/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,559 | `9466445416024bab9ba48f77a51fa7b135c3d6bef9afaf8d39785e2504fb1d9e` |
| `flow-public.pdf` | 199,181 | `cd295ed3ad08045b4040ef9ae3cc7b3d0f5e9017bc9271bb7fe6361c666e8a04` |
| `docx-native.ofd` | 15,288,611 | `75ae4bc01a1059e51057ec99687aca19977ad3d2f2be7dc82aa83b2bff3463f9` |
| `docx-native.pdf` | 117,510 | `4ea433a4d2fd853cd8eb30e34efa3098c004d598a7447d5d83153a8f976947ef` |
| `docx-default.ofd` | 15,288,607 | `99fa1c8745709ccf7b60b3137fae7309839aa4995fa5833b24bdec2f3e00c332` |
| `docx-default.pdf` | 117,510 | `26119424c474eea81a4ad96cb306f93896ba1748c82fb87d8be10510d88d500b` |
| `alignment-public.ofd` | 1,670 | `13f7ce13b830f40c0426aca7d6ebb88e5dc8e46944fc122cf521eec0c4fcb3fb` |
| `alignment-public.pdf` | 70,786 | `18d0203c9ce544d04b97e9566ba2a08993515764ac9f559639d95dbdcd4ecad6` |
| `alignment-native.ofd` | 15,286,930 | `2083c0eb386e4722b7008f3f09d3a8e719875c155525866222c355b20fa0962a` |
| `alignment-native.pdf` | 70,261 | `dd8198daa1421a397f1b5a83c8a0a8d435587dbfbdd74f0190729aa3049ed3ac` |
| `alignment-default.ofd` | 15,286,926 | `aaa12ba2b0dd3235a831abaa8dd0f9cc44fd9f9f4a783d79514d69b3c52b1561` |
| `alignment-default.pdf` | 70,261 | `139e138620bfff14316a2421e684eea11b342c29f982735ae814f49d65ed2c9d` |

页面图：`review9/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。原有七张 PNG 与先前 `review7/pages/` 逐字节一致；另保留 `alignment-public-1.png`、`alignment-native-1.png`、`alignment-default-1.png`，但最终结论依据本次重新打开的 Preview 页面。对比上一轮，公开 Flow PDF 增加 609 字节（约 0.31%），原 DOCX PDF 增加 64 字节（约 0.05%），来自空白字素的局部剪裁指令；原有页面像素不变。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。

本次 PDF 空白字素修复与像素回归覆盖 Flow/Native 的 `DeltaX/DeltaY` 定位字素及纯空白文本。通用 OFD 无定位的混合文本 run 仍受所绑定字体的缺字影响；该路径未作为本票完整字体保真验收。首版也未实现完整 Unicode 行断算法。
