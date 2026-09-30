# 公开流式布局与 DOCX Native 页面验收（2026-09-29）

## 最新 review21 正文 PAGE 补充验收（2026-09-30）

- 代码状态：`6ad54807d4d20e387abc925db8fc5b9742046df7`；后续文档提交不改变生成代码。正文 PAGE 段先按原顺序消费前置空行/FF，再使用延期 LF、段间距确定首个正文行的目标页；每次换页都完整重排，确认首行可放下后仅提交一次 gap，再绘制。页号位数改变时字宽和 DeltaX 同步重测。普通无 PAGE 段落仍走既有路径。
- Sol Low 独立最终 Release **263/263 通过，0 失败/跳过**（7 份 TRX，主任务核对）；样例构建 0 警告/错误，29 组 OFD/PDF 共 61 页，11 包隔离消费 E2E 通过。日志、TRX、页面图和包产物在 `artifacts/flow-layout/review21/`。
- 11 项新回归覆盖延期空行后的 PAGE、段后距推动逻辑页号 9→10、完整字段字宽/DeltaX、逻辑页码重启、前置 FF 不重复、同段前置 LF、混合 LF/FF 与 gap 只应用一次。恢复旧 PAGE 路径时 **11/11 失败**；修复后全套通过，负向日志 `review21/before-fix-page-field-tests.log`。Astra 有界复核闭合；Sol 发现的前置 LF 边界也已修复并复核。
- 实际查看链路：DOCX → 显式 Native/default OFD → PDF → **macOS Preview**。六份新增 PDF 均重新打开并核对 review21 路径，检查全部 **16 页**及正文位置/边界。原 45 张页面 PNG 与 review19 **45/45 逐字节相同**，仅作辅助回归；不宣称在 review21 重新打开了全部 61 页。之前实际查看的 review16 28 页、review18 9 页、review19 8 页记录在下方。
- 支持边界：本轮校准首个正文行页号上下文；同段正文在之后自动分页或中途 FF 后的 PAGE 仍沿用该上下文，未实现完整正文动态字段引擎。文档已明确这一限制；公开的页眉页脚逐页计数契约不变。

| 本轮 OFD → PDF | 实际检查页 | 结果 |
| --- | --- | --- |
| `review21/page-field-native.ofd` → `page-field-native.pdf` | 1–4 / 4 | A 在第一页，第二、三页为显式空行，第四页正文正确显示 4，顶部位置与间距正常，没有旧页号或重复文字。 |
| `review21/page-field-default.ofd` → `page-field-default.pdf` | 1–4 / 4 | A 在第一页，第二、三页为显式空行，第四页正文正确显示 4，顶部位置与间距正常，没有旧页号或重复文字。 |
| `review21/page-field-gap-native.ofd` → `page-field-gap-native.pdf` | 1–2 / 2 | 第一逻辑页号从 9 开始，段后距将 PAGE 段移动到第二页，正确显示两位数 10；字间距和边界正常，gap 没有重复应用。 |
| `review21/page-field-gap-default.ofd` → `page-field-gap-default.pdf` | 1–2 / 2 | 第一逻辑页号从 9 开始，段后距将 PAGE 段移动到第二页，正确显示两位数 10；字间距和边界正常，gap 没有重复应用。 |
| `review21/page-field-leading-native.ofd` → `page-field-leading-native.pdf` | 1–2 / 2 | 同段首个 LF 保留空白第一页，第二页正文正确显示 2，位于页顶；无额外空页、裁切或重复文字。 |
| `review21/page-field-leading-default.ofd` → `page-field-leading-default.pdf` | 1–2 / 2 | 同段首个 LF 保留空白第一页，第二页正文正确显示 2，位于页顶；无额外空页、裁切或重复文字。 |

| 文件（`artifacts/flow-layout/review21/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `page-field-native.ofd` | 15,287,453 | `d31c6ed0f0fea1a9d14ffc06aef31d6b529bd525f81c8748ee035b58c7f31c56` |
| `page-field-native.pdf` | 64,897 | `32978d10975e101bf18207025207c9fb5dbedcbc2aab0545778fd2418aaad5f3` |
| `page-field-default.ofd` | 15,287,455 | `297b732c8d7e39934295159fb3c52823b906a89e934e0610c422dbfed1843b44` |
| `page-field-default.pdf` | 64,897 | `dc19bec5aaae8516b1c25f4a6bcefd93dbe138ae15670018ae1b5ab9ef457933` |
| `page-field-gap-native.ofd` | 15,286,834 | `cdef3e860af8fa14c00555b982993afed38233ed15044564137cd53ad115acb9` |
| `page-field-gap-native.pdf` | 64,599 | `ee1465b72ffd1e3343b0a2f5e596f5f50f94ad0e915f7299a28ac197fe97a18c` |
| `page-field-gap-default.ofd` | 15,286,832 | `2de6978db6982973fb05d465a9d4745e525e88a10c8eda4b4602e59bb3bb8c01` |
| `page-field-gap-default.pdf` | 64,599 | `a35b21b26e323915e9d6f416f1306d580cbce3b300fab31bdded1a27bf66f175` |
| `page-field-leading-native.ofd` | 15,286,733 | `705a58121e5d39841bb8bf6df68f17d852ed465df03c9e613e66af53163c2456` |
| `page-field-leading-native.pdf` | 63,915 | `8a8140d88c59d847144f27886ac0e7b42c37747973340c7d2ae33c96b540d914` |
| `page-field-leading-default.ofd` | 15,286,734 | `28e48f369d6341454c1bea25e58b876d131a8c620ebb6ed5ef1a654fff929be6` |
| `page-field-leading-default.pdf` | 63,915 | `8aebb5a26b2139145228e9612a20762fe206beb62ddf03947b98a92611af8f49` |

新增 PDF 为约 64–65 KB，DOCX OFD 保持约 15.3 MB，没有异常体积增长。完整 58 个本轮 OFD/PDF 的清单、大小与哈希位于 `review21/artifact-manifest.tsv`。

## review19 节边界补充验收（2026-09-30）

- 代码状态：`6b821344ab40221af7aafb22bedd6c3a09e59441`，后续验收文档提交不改动代码。非末节结束时，先处理显式 FF，再用所属节的 `LayoutState` 逐行提交延期 LF；末节 EOF 仍延期，不生成空白尾页。Astra 有界复核闭合。
- Sol Low 独立最终全套 Release：**252/252 通过，0 失败/跳过**（7 份 TRX，主任务再次核对）。样例构建 0 警告/错误，23 组 OFD/PDF 共 45 页，11 个本地包隔离消费 E2E 通过；日志/产物在 `artifacts/flow-layout/review19/`。
- 新增 3 项内部参数化回归覆盖 0/1/2 个 LF、空行所属节页宽和 EOF 单页保护，2 项实际 DOCX 回归覆盖 Native/default。恢复旧条件时 4 失败、1 无 LF 保护通过；修复后全套通过，负向证据：`review19/before-fix-section-tests.log`。
- 实际新验收：DOCX → 显式 Native/default OFD → PDF → **macOS Preview**，重新打开 review19 的两份 PDF，检查全部 8 页及完整页面边界。第一节 A 与两页显式空行使用约 106×19 mm；第四页 B 使用第二节约 120×297 mm，位于正文顶部，无裁切、重叠、重复文本或意外尾页。
- review19 的原有 37 张 PNG 与 review18 **37/37 逐字节相同**，仅为辅助回归证据。实际 Preview 查看范围分别保留为 review16 的 28 页、review18 的新增 9 页、review19 的新增 8 页，不宣称在 review19 再次打开了全部 45 页。

| 实际 OFD → PDF | 检查页 | 结果 |
| --- | --- | --- |
| `review19/section-newline-native.ofd` → `section-newline-native.pdf` | 1–4 / 4 | A、空行、空行归属第一节；B 归属第二节，页面尺寸与顶部位置正确，全部页面边界正常。 |
| `review19/section-newline-default.ofd` → `section-newline-default.pdf` | 1–4 / 4 | 默认模式与显式 Native 的四页外观一致，未见缺字、裁切、重复文本或异常尾页。 |

| 文件（`artifacts/flow-layout/review19/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `section-newline-native.ofd` | 15,287,454 | `caa506fde14e0d479246f5cea9e105c476269a5632f8975a9a409c50f2e8ecc3` |
| `section-newline-native.pdf` | 64,986 | `8b0967cb109383d6adfbaf2a15104d7ee2362b938e3147b9a418bf347caab6c8` |
| `section-newline-default.ofd` | 15,287,458 | `2a4adcc0175ade8cf8faf136e2bc6d76a1edb54c04df7fb160076c0b727ce9c7` |
| `section-newline-default.pdf` | 64,986 | `2e8ee54bcdf77f41c902979dd5265ab8e62f54bddc288c9694622e65d23480ae` |

新增两份四页 PDF 均为 64,986 字节；DOCX OFD 仍约 15.3 MB。完整本轮 46 产物清单、字节和哈希保存在 `review19/artifact-manifest.tsv`，原有页面像素没有变化。

## review18 补充验收（2026-09-30）

- 代码状态：`93a6978fd0cba99cd80a64e2ff228ddebccfbcee`。后续文档提交不改变生成代码。修复 Codex 的 LF-only 前段导致 `PageBreakBefore` 忽略，以及 Cursor 的同段/前置混合字号换行使用前一个字号；Astra 对两项有界契约复核闭合。
- Sol Low 独立最终复验：**247/247 通过、0 失败/跳过**（7 份 TRX），E2E 构建 0 警告/错误、21 组 OFD/PDF 共 37 页，11 本地包隔离消费 E2E 通过。主任务核对 TRX 合计。日志、TRX 和包消费产物在 `artifacts/flow-layout/review18/`。
- 新增 24 项回归：6 项公开/Native 显式分页与首块保护，6 项公开与 12 项实际 Native/default 混合字号的同段/前置/纯换行段。临时恢复旧两个条件后 22 失败、2 首块保护通过；恢复修复后全套通过，负向日志 `review18/before-fix-newline-tests.log`。
- 本轮实际重新打开 **6 份新增 PDF、9 页**：公开 API → native OFD → PDF → **macOS Preview**，DOCX → 显式 Native/default OFD → PDF → **macOS Preview**。各文件确认实际 review18 路径，全部页面正文和边界已查看。原有 15 组在 review16 的 28 页实际查看记录保留在下方；review18 重新生成的对应 28 张 PNG 与 review16 **28/28 逐字节相同**，仅是辅助回归证据，不冒充这 28 页在最新目录再次打开的 Preview 检查。
- 空行契约：同段中间或前置空行由结束该行的换行字号决定；正文后的延期尾部换行保存每个控制后的间距；纯换行段的关闭空行重复最后换行字号。已有延期空行使后段 `PageBreakBefore` 开启新页并清除尾间距；首块直接设置它不增加前置空页。

| 本轮新 OFD → PDF | 实际检查页 | 结果 |
| --- | --- | --- |
| `review18/newline-break-public.ofd` → `newline-break-public.pdf` | 1–2 / 2 | 首页为空白（前段显式 LF 的逻辑页面），B 在第二页正文顶部；显式分页清除延期空行，未见裁切、重复文本或额外尾页。 |
| `review18/newline-break-native.ofd` → `newline-break-native.pdf` | 1–2 / 2 | 首页为空白（前段显式 LF 的逻辑页面），B 在第二页正文顶部；显式分页清除延期空行，未见裁切、重复文本或额外尾页。 |
| `review18/newline-break-default.ofd` → `newline-break-default.pdf` | 1–2 / 2 | 首页为空白（前段显式 LF 的逻辑页面），B 在第二页正文顶部；显式分页清除延期空行，未见裁切、重复文本或额外尾页。 |
| `review18/inline-newline-public.ofd` → `inline-newline-public.pdf` | 1 / 1 | ABCD 全部完整；A→B 同段空行、B→C 前置空行、C→D 纯控制段关闭空行的间距依次符合对应毫米/磅字号预期，无重叠或裁切。 |
| `review18/inline-newline-native.ofd` → `inline-newline-native.pdf` | 1 / 1 | ABCD 全部完整；A→B 同段空行、B→C 前置空行、C→D 纯控制段关闭空行的间距依次符合对应毫米/磅字号预期，无重叠或裁切。 |
| `review18/inline-newline-default.ofd` → `inline-newline-default.pdf` | 1 / 1 | ABCD 全部完整；A→B 同段空行、B→C 前置空行、C→D 纯控制段关闭空行的间距依次符合对应毫米/磅字号预期，无重叠或裁切。 |

新增产物哈希（均在 `artifacts/flow-layout/review18/`）：

| 文件 | 字节 | SHA-256 |
| --- | ---: | --- |
| `newline-break-public.ofd` | 1,577 | `a138d8064e023a2bc08abc15c15670b51f27dade33349ce83cd7a3899b12b40e` |
| `newline-break-public.pdf` | 63,708 | `b615871293fe95db828a2cb3c0ef0edc6b3a289df9c07398c2ccba8fc1b131fd` |
| `newline-break-native.ofd` | 15,286,734 | `d5f2f825c87dc5f7d1ee5ef4257dc3b1cc92e7cf2e2485f7699e48904352cf9c` |
| `newline-break-native.pdf` | 63,853 | `45c5e659057d9b5a2cebebb59371202dc5dab5524779f8fcbc849683d2cbaf3f` |
| `newline-break-default.ofd` | 15,286,733 | `bff4861d582244913cca0f50b6faf57b81c5f23e9e28ef19d9903fb24cc916a9` |
| `newline-break-default.pdf` | 63,853 | `d1338c3960d798666d0598b1920bedfc385de7154c0d7891a5a2c8ffa6f14bb7` |
| `inline-newline-public.ofd` | 1,318 | `8a1e425ec1e117a3203149a622a9b99b8a9d575e89a563182e3fea0eef27f655` |
| `inline-newline-public.pdf` | 64,373 | `b96fb2639dbe604352554e1fe8495f209b56770a061bef83c404b2bd179b7b03` |
| `inline-newline-native.ofd` | 15,286,468 | `568d6c54b52631b16bc2ea7760d48d094193c2bda7e65289fcb2951fc7ae389c` |
| `inline-newline-native.pdf` | 64,525 | `7bc2d596f6b472386090b68d02168e57d132a92187eefb83fc972900f535b6d8` |
| `inline-newline-default.ofd` | 15,286,468 | `9cfffc0a8076e977fc40dadf9eee599e25819464462adc86c2a8fbb89c1d54bb` |
| `inline-newline-default.pdf` | 64,525 | `8eeb5ffb14ee195371281dc09e4dddbaba51cb9c343fed7ca306a5bd8e6ea195` |

新增短页 PDF 约 64 KB、ABCD PDF 约 64 KB；原 15 组页面像素与体积保持基准，没有异常增长。原 OFD/PDF 的本轮完整清单与哈希另保存在 `review18/artifact-manifest.tsv`。本轮 9 页与前轮 28 页只对各自实际查看文件下结论。

**2026-09-30 review16 复验：** 本轮修复连续显式空行被段间距夹断、空行字号丢失和组合 í 的字宽/字形差异。已重新打开 `artifacts/flow-layout/review16/` 全部 15 份最新 PDF，实际检查 28 页，未见所列样例的视觉缺陷；其中 6 页是源文显式换行产生的预期中间空页。

## 基线与复现

- 原始源码基线：`be8b74b3cbd699732ec0638189d14037e052a770`；本次生成代码状态：`4a975526de72c2a52491bb796fedf9e1f0ec8416`。后续验收文档提交不改变代码。
- 样例程序：`e2e/Ofdrw.Net.Layout.E2E/Program.cs`；原有 DOCX：`e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`（SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`）。程序另生成对齐、文末连续 LF、连续 LF 后有正文、不同字号 LF 的确定性 DOCX 和公开 Flow 样例。
- 依根 `AGENTS.md` 设置 .NET 环境变量，使用 `--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` 构建/测试；E2E 构建后以 Release `--no-build --no-restore` 生成到 `artifacts/flow-layout/review16`。日志含 `full-tests.log`、`sample-build.log`、`sample.log`、`package-e2e.log`，产物与 TRX 保存在此 worktree 被忽略的 `artifacts/` 中。
- 实际查看链路：公开 API → native OFD → `OfdToPdfConverter` → **macOS Preview**；DOCX → 显式 Native / 默认 OFD → `OfdToPdfConverter` → **macOS Preview**。逐个打开本轮路径并核对 Preview 文件 URL，逐页检查正文、边界和空白。没有用直接 DOCX→PDF 代替 native 产物。`pdftoppm -scale-to 1300 -png` 生成的 28 张页面图仅用于辅助复查。

## 功能验证

- Sol Low 独立顺序运行最新全套 Release 测试：**223 通过、0 失败、0 跳过**；主任务复核 7 份 TRX：Core 5、Packaging 23、PDF 51、Signatures 4、DOCX 90、CLI 5、Layout 45。E2E 项目构建 0 警告、0 错误，15 组 OFD/PDF 生成通过。
- 11 个本地 NuGet 包构建、隔离 feed 安装和消费 E2E 通过；公开 Layout 消费生成 4 页，完成 OFD 重读/文字检查；CLI Native 与 DualLayer 各 2 页。产物在 `review16/package-e2e-output/`。未发布公共 NuGet。
- 既有回归覆盖跨 Span 单词/CRLF/组合字素、CJK 与 Hangul 字宽、比例英文、首行缩进、样式/往返、段间距、取消/预算、连续显式分页、Native section 归属、图片缩进和原有 Native/default/DualLayer 行为。
- 自动折行尾普通空格仅取消推进量，原始文本/样式保留；居中/右对齐及元素右边界回归通过。GL 空格 U+00A0/U+202F/U+2007 连接相邻文字，含跨 Span 组合标记；窄 NBSP = 0.2 em，数字空格 = 当前样式数字字宽。临界宽度与四种粗斜体样式回归通过。定位纯空白 PDF 字素不绘制缺字方框，同时保留语义与推进量。旧绘制条件的四项像素测试 4/4 失败，恢复修复后通过，证据：`review9/before-fix-pixel-regression.log`。
- Astra High 有界复核的连续 LF 尾序列、GL + 组合标记及混合字号 LF 意见已修复并复核闭合；对后续正文和表格逐行应用空行分页，不把它们当作可夹断的段间距。短页 A + 两个 LF 单独为 1 页；后续 B 为 4 页且中间两页为空，B 位于页顶；显式分页覆盖和页数预算回归通过。
- 每个空行保留所属 Span/run 字号及段落最小行高，DOCX `w:br`/`w:cr` 保留 run 格式。大字号连续 LF、先大后小 LF、纯 LF 段及后续正文间距回归通过；E2E 坐标断言容差为 0.005 mm。
- Basic Latin / Latin-1 重音词的整词折行与超长词字素拆分回归通过。受支持单个 Latin-1 字素在测量及 PDF 绘制时使用 NFC 合成，原始 OFD 文本不改变；PDF 复制文字允许规范等价形式。四种样式的 í / i + U+0301 像素一致、OFD 原文保留、PDF 两个语义字母 íB 且无重复层。临时移除合成绘制修复时 4/4 像素测试失败，恢复后 4/4 通过，证据：`review16/before-fix-latin-i-pixel.log`。
- 公开 Flow 3 页、5865 个非空白字符，输入与 OFD 抽取逐字相等；原 DOCX 显式 Native/default 各 2 页、189 字符；对齐三组各 1 页、205 字符；文末 LF 三组各 1 页、1 字符；有后续正文三组各 4 页、2 字符；字号空行三组各 1 页、28 字符。原始全文保持由功能测试验证，视觉结论由以下实际页面检查给出。

## Preview 逐页结果

| 本轮 OFD → PDF | 实际检查页 | 结果 |
| --- | --- | --- |
| `review16/flow-public.ofd` → `flow-public.pdf` | 1–3 / 3 | 中英混排、比例英文、局部粗斜体及红色正常；第 20/21 与 42/43 段续行完整，末页没有裁切、重叠或空白尾页。 |
| `review16/docx-native.ofd` → `docx-native.pdf` | 1–2 / 2 | 中文标题、英文斜体、蓝色表头底色与内外边框正常；第二页红色粗体及右对齐日期正常，无重复文本。 |
| `review16/docx-default.ofd` → `docx-default.pdf` | 1–2 / 2 | 默认模式的正文、样式、表格和两页分页与显式 Native 外观一致。 |
| `review16/alignment-public.ofd` → `alignment-public.pdf` | 1 / 1 | 公共 API 的参考 Alpha 与折行 Alpha 水平位置一致；三种不换行空格保持空白与推进量；café 和 mínimo 的预组/组合形式均整体折行、字形一致，í 没有额外圆点或重影。 |
| `review16/alignment-native.ofd` → `alignment-native.pdf` | 1 / 1 | 显式 Native 的参考 Alpha 与折行 Alpha 水平位置一致；三种不换行空格保持空白与推进量；café 和 mínimo 的预组/组合形式均整体折行、字形一致，í 没有额外圆点或重影。 |
| `review16/alignment-default.ofd` → `alignment-default.pdf` | 1 / 1 | 默认 DOCX 的参考 Alpha 与折行 Alpha 水平位置一致；三种不换行空格保持空白与推进量；café 和 mínimo 的预组/组合形式均整体折行、字形一致，í 没有额外圆点或重影。 |
| `review16/terminal-public.ofd` → `terminal-public.pdf` | 1 / 1 | 公共 API 短页正文 A 完整，连续两个尾换行没有生成空白尾页。 |
| `review16/terminal-native.ofd` → `terminal-native.pdf` | 1 / 1 | 显式 Native 短页正文 A 完整，连续两个尾换行没有生成空白尾页。 |
| `review16/terminal-default.ofd` → `terminal-default.pdf` | 1 / 1 | 默认 DOCX 短页正文 A 完整，连续两个尾换行没有生成空白尾页。 |
| `review16/terminal-followed-public.ofd` → `terminal-followed-public.pdf` | 1–4 / 4 | 公共 API 的 A 在第一页，第二、三页为显式换行所需的预期空页，B 从第四页顶部开始；四页均实际查看。 |
| `review16/terminal-followed-native.ofd` → `terminal-followed-native.pdf` | 1–4 / 4 | 显式 Native 的 A 在第一页，第二、三页为显式换行所需的预期空页，B 从第四页顶部开始；四页均实际查看。 |
| `review16/terminal-followed-default.ofd` → `terminal-followed-default.pdf` | 1–4 / 4 | 默认 DOCX 的 A 在第一页，第二、三页为显式换行所需的预期空页，B 从第四页顶部开始；四页均实际查看。 |
| `review16/styled-newline-public.ofd` → `styled-newline-public.pdf` | 1 / 1 | 公共 API 的大字号 A 后空行间距完整，混合大/小字号换行分别保留行高，B/C/D 位置正常，没有重叠或裁切。 |
| `review16/styled-newline-native.ofd` → `styled-newline-native.pdf` | 1 / 1 | 显式 Native 的大字号 A 后空行间距完整，混合大/小字号换行分别保留行高，B/C/D 位置正常，没有重叠或裁切。 |
| `review16/styled-newline-default.ofd` → `styled-newline-default.pdf` | 1 / 1 | 默认 DOCX 的大字号 A 后空行间距完整，混合大/小字号换行分别保留行高，B/C/D 位置正常，没有重叠或裁切。 |

本记录只对以上样例与页数下结论。样例没有图片、页眉页脚或页码；DOCX 现有表格不代表公开 Table API 已实现。Astra 的有界设计复核与 Sol 的功能验证均已完成，主任务完成本轮 Preview 检查。

## 产物清单与体积

| 文件（`artifacts/flow-layout/review16/`） | 字节 | SHA-256 |
| --- | ---: | --- |
| `flow-public.ofd` | 5,559 | `f84939562dce3696e8b8277ff1fcd8aaaa39c28c9b709ff925f98648c9255f2c` |
| `flow-public.pdf` | 199,181 | `d1def6631eff806697783f28cf4ae5f8608dbb0c34956578840cab289c3887b0` |
| `docx-native.ofd` | 15,288,613 | `d02c163dd187c7b86a345399f8e52666139914e1798164ecf14eaa444f118f64` |
| `docx-native.pdf` | 117,510 | `e92f7658e7bace98ee50675b5423d82c7b2c6f8f925831ed9b9938560e15ccbf` |
| `docx-default.ofd` | 15,288,608 | `f7b090669c2ec4e10727ea027a189c38ecd5e2478a838e2ab036393cf46dec8f` |
| `docx-default.pdf` | 117,510 | `ceb654d1163c2063f41609c1b1c9b0e8e271459ddbaeb4fefbd87427a3fc1dd4` |
| `alignment-public.ofd` | 1,935 | `8bfb4378b5c7d7bcbcd1f43fb71f0da524b0042d5b35985dfa60423cfe89887e` |
| `alignment-public.pdf` | 71,830 | `645e9b073b94427498c62fbe935c061893d2e619e5bb9e253ddd37d3c097079b` |
| `alignment-native.ofd` | 15,287,189 | `bcd87c4d9785875cfe5bbfb760a608494b075cc828c5b6da093eb53da581f389` |
| `alignment-native.pdf` | 71,324 | `e0d057ef786a320f4a6be68b4f5e118fa7a4f7eecbe05c7fd04e9f076ef869b4` |
| `alignment-default.ofd` | 15,287,188 | `8440cc8c90da560627acb372e4f976edfa2e1df94486ebac5147c93174888604` |
| `alignment-default.pdf` | 71,324 | `79c22542e126917a55e12173ea3275585f07094ec73376bec9bdd5a8ac19cdc6` |
| `terminal-public.ofd` | 1,265 | `e14ff8801c5b6887bb49f82761ab2de09e5f253c183c539c32c1645798b1076c` |
| `terminal-public.pdf` | 63,489 | `1e1749fa1b398e9fb379b62144ae17d81e0bdf7892701554ff1bd9fafa294232` |
| `terminal-native.ofd` | 15,286,416 | `b9262e0b64604ead5691f6f86d6e88c859ed31c34565d37a95ccbdfa31770293` |
| `terminal-native.pdf` | 63,640 | `f904659398497f5c15c6baf6302104c16fd9c64a84c673a7e9ea39ba99eedaa9` |
| `terminal-default.ofd` | 15,286,415 | `e2d0579ed90583573049601a379bde5574ea4914a7a7a2e0d4f4431cf8288290` |
| `terminal-default.pdf` | 63,640 | `28653c657c5debd32556d52e1bd3c65f9278c881b17bf1f07ee70ecefbc040c0` |
| `terminal-followed-public.ofd` | 2,276 | `783e707a99085fbc338cd10fb992a824df2c8152d2867a3cc1f2e4835d9c88d9` |
| `terminal-followed-public.pdf` | 64,837 | `f509877555e78f324e8cac6ef4da6e6bd8f4361cc2e0befa390a792d3b4cb505` |
| `terminal-followed-native.ofd` | 15,287,454 | `f6bd12a58176bf0f276f1120b6d035f6738d7a267a8333d207516596aceb6ed7` |
| `terminal-followed-native.pdf` | 64,981 | `7b421c2f422a7ef3b30e86b00bb74dd465c51d1b741068bd26453846671724a3` |
| `terminal-followed-default.ofd` | 15,287,454 | `6459f01fcffd6fbf816589fb46563a288e2317f26f460006ba9ce1feb585b3af` |
| `terminal-followed-default.pdf` | 64,981 | `2f143b9b0cf069e06aa586c6b8682ad133deb2621fa37cd04bda718df592c438` |
| `styled-newline-public.ofd` | 1,431 | `1dbf1dd7d389959b65dc4d79fd866eeeea089315e8393864300f03805f812f10` |
| `styled-newline-public.pdf` | 67,636 | `f16a0ef01206e29940fc06291d3cfaf097fcf2550625dfc3eb2d27b9846c22ad` |
| `styled-newline-native.ofd` | 15,286,584 | `adffe2c5405f6684f11969fe70d08e38dc32ff2ac29b2480de86e42c51e4a7bd` |
| `styled-newline-native.pdf` | 67,695 | `c2ab282f1ccbf809327f69aa70522244df138d8d72be4b4007ea7056e6ff1a06` |
| `styled-newline-default.ofd` | 15,286,587 | `ef8ab2aea11dd93c0f4c9524d49acbb612fe71405455763e61e40ef74dfa0f70` |
| `styled-newline-default.pdf` | 67,695 | `67d9d09215219302dfc18ede5786d08b367c19480738b612632d01526bbe7b14` |

页面图全部保存在 `review16/pages/`（28 张）；原有 `flow-public`、`docx-native`、`docx-default` 的七张 PNG 与 `review7/pages/` 逐字节一致。本轮原有公开 Flow PDF 199,181 字节、DOCX PDF 117,510 字节，与上一轮一致；新增 í 内容的对齐 PDF 约 71–72 KB。四页短页 PDF 约 65 KB，字号页约 68 KB，没有异常体积增长。DOCX OFD 约 15.3 MB 来自既有全量字体嵌入；公开 Flow OFD 仅声明字体，不能跨不同内容比较压缩率。

## 范围与遗留

公开 Flow 的 CJK 字宽采用 1 em，未做字体子集或嵌入；跨机器视觉取决于目标阅读器字体。复杂 Word 浮动对象、公开表格、Canvas、完整 Unicode 行断算法及任意字体保真不在本票范围。通用 OFD 无定位的混合文本 run 仍受绑定字体缺字影响。PDF NFC 修复限于单个受支持 Latin-1 字素，不宣称完整文字塑形。

未使用目标 OFD 桌面阅读器验证原始 OFD 互操作；本次视觉结论仅针对所列 OFD→PDF→Preview 链路。PR 保持未合并，不发布 NuGet，不启动 ticket 16 或依赖票。
