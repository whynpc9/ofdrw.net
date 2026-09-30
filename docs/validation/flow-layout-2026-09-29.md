# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

**2026-09-30 最新复验：** 已修复自动折行尾空格导致的居中/右对齐左偏，以及下一轮 review 指出的不换行空格断行。NBSP、窄 NBSP 和数字空格保留推进量并连接相邻文字；最新 review 进一步指出通用 0.6 em 不适合空格变体，现已为窄 NBSP 使用 0.2 em、数字空格使用所选样式数字字宽，Native 同步规则；Preview 随后检出窄 NBSP 在缺字字体下的 PDF 方框，本轮补了定位空白字素绘制修复。已重新打开 `artifacts/flow-layout/review12/` 九份最新 PDF，逐页复验原有 7 页、3 页对齐/空格/重音单词样例及 3 页连续尾换行短页样例，共 13 页，未见本次样例的视觉缺陷。

## 基线与复现

- 源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次产物对应 `codex/public-flow-layout` 的 `843b520` 代码状态。后续文档提交不改动生成代码。
- 样例：`e2e/Ofdrw.Net.Layout.E2E/Program.cs` 创建 55 段无坐标的中英混排 Flow；DOCX 使用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。
- 运行：设置根 `AGENTS.md` 的 .NET 环境变量和单节点参数后，构建 E2E 项目，再执行 `dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/flow-layout/review12`。`artifacts/` 被 Git 忽略，保留在此 worktree 供复查。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 `Native` OFD / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。九份 PDF 均从 `artifacts/flow-layout/review12/` 重新打开，不使用直接 DOCX→PDF 的页面。页面 PNG 用 `pdftoppm -scale-to 1300 -png` 从这些 PDF 生成，仅作辅助复查。

## 功能验证

- 最新解决方案 Release 全套测试：197 通过、0 失败、0 跳过；TRX 在 `artifacts/flow-layout/review12/test-results/`。其中 Layout 测试 35/35，DOCX 测试 78/78，PDF 测试 47/47，覆盖公开样式/往返、跨 Span 英文单词、CRLF 与组合字素、段末显式换行、连续显式换页、Hangul（预组与分解 Jamo）/增补汉字（含 U+30000）字宽、段间距跨页、取消与预算、图片首行缩进、Native 窄单元格大字与 section 换页归属，以及原有 Native/default/DualLayer 行为。上一轮新增 9 项回归覆盖 Center/Right 的尾分隔空格、跨 Span 空格样式、首行缩进、文本保留、元素右边界，以及自动折行与显式换行/末行的空白推进量差异。Native 往返右边界断言允许 OFD 的 0.001 mm 序列化精度。上一轮另增 18 项回归：三种不换行空格的跨 Span/Native/default 折行与字符保留，中文邻接及首行缩进超宽契约，以及 PDF 普通/模拟粗体空白字素的像素与语义检查。最初四项 PDF 像素回归在临时恢复旧绘制条件时 4/4 失败；恢复修复后通过，最终三种空格的六项像素回归全部通过。旧条件失败日志保存在 `review9/before-fix-pixel-regression.log`。字宽修复新增 8 项回归覆盖四种粗斜体样式：普通 NBSP = 普通空格、窄 NBSP = 0.2 em、数字空格 = 当前样式数字字宽；公开 API 在修正字宽加 0.001 mm 的临界行宽下能完整渲染，并检查 OFD `DeltaX` 与对象宽度；Native 同时验证 `0` + 数字空格 + `0` 与 `000` 等宽。
- 协调要求的 Astra High 有界契约复核指出连续段末 LF 空行和 GL 空格与跨 Span 组合标记交互两处遗漏，现已修复并复核闭合。另解决最新 bot 指出的预组合重音字母整词折行；本轮新增 17 项回归，覆盖连续 LF/LF+FF 短页无空白尾页、后续正文保留全部三行间距、GL 空格带跨 Span 组合标记的非断行连接/首字符样式、Basic Latin/Latin-1 预组合及分解重音词整词搬行与超长词字素拆分。
- 最终全套测试、样例构建/生成和 11 包消费 E2E 由 Sol Low 独立顺序运行并检查，主任务复核了 7 份 TRX 合计 197/197 和退出结果，随后完成最新 13 页 Preview 验收。
- 本地 11 个 NuGet 包构建、安装和隔离包消费 E2E 通过；公开 Layout 包消费生成 4 页，并完成 OFD 重读及文本检查。最新日志与产物在 `artifacts/flow-layout/review12/package-e2e.log` 和 `package-e2e-output/`。
- 新增样例由同一 E2E 程序生成 `alignment-public.ofd`、`alignment.docx` 及其显式 Native/default OFD。三份对齐产物均为 1 页、167 个非空白字符；E2E 断言两组参考 `Alpha` 与折行 `Alpha  ` 的 X 坐标一致，并保留两个普通分隔空格；还断言 `A` 与 `B` 通过三种不换行空格连接时整组移至下一行，空格推进量大于零；`café` 与 `café` 分别作为完整词移至新行。新增 `terminal-public/native/default` 三份短页样例（公开页高 19 mm、DOCX 页高约 19 mm，正文 `A` 后连续两个换行）均为 1 页、1 个正文字符，OFD 重读后断言没有空白尾页。
- 本次 E2E 对公开 Flow 输入与 OFD 抽取做逐字符（去空白）相等检查：3 页、5865 字符；显式 Native 与默认模式的 OFD 抽取一致：均为 2 页、189 字符。现有 DOCX 测试还逐字比较 OpenXML 原文与 Native OFD，并断言 OFD→PDF 没有重复文字。

## Preview 逐页结果

| 产物与模式 | 实际检查的页 | 结果 |
| --- | --- | --- |
| `review12/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 标题居中；中英、比例英文、局部粗斜体和红色只作用于目标 Span；第 20/21、42/43 段跨页续行完整；未见缺字、乱码、裁切、重影或空白末页。 |
| `review12/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 第一页中文标题、英文斜体、蓝色表头填充与边框完整；第二页红色粗体及日期位置正常；未见重复文字或异常空白页。 |
| `review12/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式与显式 Native 的标题、表格、分页和第二页强调文字外观一致；未见裁切或重复文字。 |
| `review12/alignment-public.ofd` → `alignment-public.pdf` | 1 / 1 | 放大检查居中/右对齐两组：参考 `Alpha` 与折行第一行 `Alpha` 水平位置一致；`information` 正常换行；三种不换行空格的 `A B` 整体换行且保留间距，窄 NBSP 方框已消失且间距比普通 NBSP 更窄，数字空格间距按数字字宽保留；`café` 与 `café` 均完整移至新行，重音显示正常；无右边界裁切或重影。 |
| `review12/alignment-native.ofd` → `alignment-native.pdf` | 1 / 1 | DOCX 显式 Native 的居中/右对齐参考与折行位置一致；三种不换行空格连接的文字保持同一行，窄 NBSP 间距小于普通空格，数字空格使用数字字宽；预组合与分解 café 均整体折行且重音正常；无方框或裁切。 |
| `review12/alignment-default.ofd` → `alignment-default.pdf` | 1 / 1 | 默认 DOCX 的两组对齐及三种不换行空格及两种 café 整词外观与显式 Native 一致，无方框、异常空白页或重复文字。 |
| `review12/terminal-public.ofd` → `terminal-public.pdf` | 1 / 1 | 重新打开后正文 A 完整，Preview 显示仅 1 页，无连续尾换行引入的空白页。 |
| `review12/terminal-native.ofd` → `terminal-native.pdf` | 1 / 1 | 显式 Native 短页正文 A 完整，仍仅 1 页，无空白尾页。 |
| `review12/terminal-default.ofd` → `terminal-default.pdf` | 1 / 1 | 默认短页与显式 Native 一致，正文 A 完整，无额外空白页。 |

首轮公开 Flow 的英文 `DeltaX` 使用粗略字宽，Preview 页面出现明显间距问题；第二轮先改为运行时字体测量，第一波 review 再指出跨机器字宽不稳定。当前版固定 Arial 兼容比例字宽表，重新生成 OFD/PDF 并逐页复验。本记录只对上述样例和页数下结论。样例没有图片、页眉页脚；DOCX 样例的表格不表示公开 Table API 已实现。

## 产物清单与体积

| 文件（均在 `artifacts/flow-layout/review12/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,559 | `d3c9f5bbbbe68ce1934298493302642e18add731885719d53bca4943e8cf1f7b` |
| `flow-public.pdf` | 199,181 | `ffafe255d530a97403f703a12c0f287781cedc0a5deb8463460d7dff186e1c0d` |
| `docx-native.ofd` | 15,288,611 | `03d42dfbca201d4e55132ddd13d30dc6d8ad5d0d25098b8f9bbe7b1b6d55ef64` |
| `docx-native.pdf` | 117,510 | `b2fc71109d2a9952e25822b906b51449ea1fede9e8eac67a165068a4d1fa93d4` |
| `docx-default.ofd` | 15,288,607 | `3a7149b9fcfe0a6417c1dc571478404372813ade5b0572dd9c213dfd8bf52773` |
| `docx-default.pdf` | 117,510 | `a4f97d11e1439cedd85156ae9ba63b104ac8f59419aab4a0ec5f8094ac4cb72f` |
| `alignment-public.ofd` | 1,776 | `ffe19d42267bf51e07337f26b924a9f9999135ad897099a01ea5a1db01edcd6e` |
| `alignment-public.pdf` | 71,413 | `373b490a0aafd600c9aabe0edb49ed73c8c9858335ad831d2c9c0fb71c93950a` |
| `alignment-native.ofd` | 15,287,034 | `d128add1984af520815c3de8c165ef27bd7b0751fecee3e4dce6bc6fd080a8a7` |
| `alignment-native.pdf` | 70,935 | `b09ddbb4a0400363b79976fbdd4422dc5de244415e0ef85a7fa4166022dca193` |
| `alignment-default.ofd` | 15,287,035 | `a4787657031972b04b0768ddde71d879bbc6bda5704377e93b70fcbf211a9d09` |
| `alignment-default.pdf` | 70,935 | `bb45405f72c67ced835fae336acf8bcf29e297f47fb1917d98bbe86138a34f3f` |
| `terminal-public.ofd` | 1,265 | `c5f1af6fddb9dbc67902b021eac2b68705187dc494ec59ef31361bfd93270032` |
| `terminal-public.pdf` | 63,489 | `36faef3e080c32030dde177b18a4f58e6a8ddca33514340adb15c91d29ea7f0e` |
| `terminal-native.ofd` | 15,286,415 | `3a87e3026f5030b04fc724da54a1981a8c65b8f36d18fb927c656eb140544ce0` |
| `terminal-native.pdf` | 63,640 | `16ade56e82f2a36ac07e32f266615e1875018bc353f19e53d26ff470be40642b` |
| `terminal-default.ofd` | 15,286,419 | `67d1cff616de22735ce1eaf15a16bd0834f9b38a067176153fc50f1de8723221` |
| `terminal-default.pdf` | 63,640 | `ed9122aa2223f942d754a57e31381b49fd7ade711830c845d7efd63448b54f02` |

页面图：`review12/pages/flow-public-1.png`～`3.png`、`docx-native-1.png`～`2.png`、`docx-default-1.png`～`2.png`。原有七张 PNG 与先前 `review7/pages/` 逐字节一致；另保留三张 `alignment-*-1.png` 和三张 `terminal-*-1.png`，但最终结论依据本次重新打开的 Preview 页面。对比加入空白字素剪裁前，公开 Flow PDF 增加 609 字节（约 0.31%），原 DOCX PDF 增加 64 字节（约 0.05%），来自空白字素的局部剪裁指令；原有页面像素不变。DOCX OFD 约 15.3 MB，主要由既有字体全量嵌入造成；公开 Flow 只声明字体名，其 5.6 KB OFD 与 DOCX 内容不同，不作同内容压缩率比较。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器的字体。复杂 Word 浮动对象、公开表格、Canvas 和任意字体保真不在本票验收范围。未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。

本次 PDF 空白字素修复与像素回归覆盖 Flow/Native 的 `DeltaX/DeltaY` 定位字素及纯空白文本。通用 OFD 无定位的混合文本 run 仍受所绑定字体的缺字影响；该路径未作为本票完整字体保真验收。首版也未实现完整 Unicode 行断算法。
