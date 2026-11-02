# N13：中英文常用链路有限盘点与证据入口（2026-11-01）

本轮没有从源码盘点确认需要新引擎或新增产品功能的缺陷。已有公开 Flow、DOCX Native/default 和 OFD→PDF 能力应先验证；本提交仅共享已有测试输入并增加具名 release-only E2E 入口，没有修改 `src/`、默认字体、绘制后端或旧样例预期。**这不是候选通过或可发布结论**：N02 当前核心末尾 Find 失败由其拥有者处理，原始“数量”字形历史 FAIL 保留。

盘点基线为 `ba7682cfceafd6e28d526941e0eb039addfe400f`。协调者告知新 `origin/main` 为 `beb02832fff01e61538d90039f3e1224ed03d702`，其新增 CI/文档不改变本次盘点的产品实现；最终候选仍须绑定合并后的实际源码。冻结范围以 [N00](cn-en-release-n00.md) 与[本轮计划](../next-steps-2026-11-01.md)为准，N00/N13 共用最多 8 小时主动盘点，不通过继续寻找潜在特性重置预算。

## 已有入口与可复用覆盖

下表是源码和现有断言的盘点，不把历史通过改写为本次运行通过。

| 核心用法 | 实际公开入口／实现 | 确定性现有覆盖与边界 |
| --- | --- | --- |
| CN/EN 正文、混排、比例英文、局部粗斜体／颜色 | `FlowDocument.Render`、`Paragraph.Spans`、`Span`；`FlowParagraphLayout` | `FlowDocumentTests.PublicFlow_MultipageStyledText_RoundTripsWithoutCoordinates`；跨 Span 英文、CRLF、组合字素、NBSP、对齐回归。传统 Flow 使用固定 Latin 比例宽和 CJK 1 em，不声称任意物理字体精确测量。 |
| 字重、下划线、字间距、图片／Canvas | `Span.FontWeight/Underline/LetterSpacingMillimeters`、`ImageBlock`、`CanvasBlock`、`OnPage/OnElement` | `FlowFixedBlockTests.SpanSpacingWeightAndUnderline_ShareFinalLineGeometry`、`LetterSpacing_ChangesWrappingWithoutTrailingSpacing`、`ImageAndCanvas_ReserveFinalBoxesAndKeepPaintOrderAcrossPages`；固定块超页明确拒绝。 |
| 普通表格、底色、边框、对齐与整行换页 | `Table/Row/Cell`；`FlowTableLayout`、`FlowTableRenderer` | `TableTests.MergedCells_KeepFixedWidthsStylesAndRawTextThroughPackageRoundtrip`、`WholeRows_MoveToNextPage_AndFollowingParagraphContinuesAfterTable`；不缩小文字以掩盖超高行。 |
| 已有复杂表格 | `RowSpan/ColumnSpan`、`SplitRowsAcrossPages`、`RepeatHeaderRowCount` | `ComplexTableTests` 覆盖合并行组、文字拆行、重复表头、图片及失败策略；这是公开 Table 的有限能力。Native DOCX 的 `vMerge/tblHeader` 不继承该承诺，见 [public-tables](../public-tables.md)。 |
| DOCX→native/default OFD | `DocxToOfdConverter.ConvertWithResultAsync`；`DocxConversionOptions.OfdMode` 默认 `Native` | `GeneratedDocx_ShouldConvertDirectlyToNativeOfdText`、`Native_ShouldPreserveVisualStylesAndProportionalAdvances`；两种入口共用 native 链路，输出正文来自 OpenXML，不经直接 DOCX→PDF 代替。 |
| DOCX 图片、页眉页脚、节页码 | `DocxModelReader`、`BuiltInOfdRenderer` | `Native_ShouldKeepInlineJpegContentDimensionsAndMediaType`、`Native_ShouldRepeatHeadersFootersAndHonorSectionNumbering`；原五页预期 First/Even/Odd/Even/Odd，页码 1/2/3/10/11，各自 `/5`。本提交共享生成器，断言未放宽。 |
| DOCX 常见分页与页选择 | `ConvertWithResultAsync(..., pages)`、`OfdPageSelection` | `Native_ShouldPreservePageOrderDuplicatesAndInvalidSelectionCompatibility`、连续显式分页／尾空行／节边界／正文 PAGE 既有回归。`SourcePages` 是零基下标；CLI `--pages` 是一基。正文 PAGE 的上下文仍为段落首个正文行，不承诺同段任意分页动态字段。 |
| 字体部署、嵌入、缺少粗体策略 | `FontDirectories`、`FontFallbackFamilies`、`FontEmbedding`、`NativeEmbeddedBoldBehavior` | `EmbeddedBoldCompatibilityTests` 测真实 regular 嵌入与明确的合成描边／严格拒绝；`FontIdentityRegressionTests` 的改造 Latin 测试字体只证明 face 选择，不证明真实中文 Bold 可读。Flow 只声明字体名称，跨机部署另验；不扩入高级字体。 |
| OFD→PDF 正文与样式 | `OfdToPdfConverter.ConvertAsync` | DOCX 样例测试核对 OFD 原文、PDF 字符、斜体几何和没有重复字；视觉、Copy/Find 各自单列，不能由这些断言推出 Preview 搜索成功。 |

公开说明可复用 [Flow](../flow-layout.md)、[表格](../public-tables.md)、[纯文本入口](../plain-text-ofd.md)；R22 的查询／复杂脚本边界保留在[剩余范围](r22-remaining-scope.md)。未发现冻结样例中新的中文标点阻断反例，不启动通用禁则／UAX 或 RTL 工作。

## 有限缺口表

| 编号／优先级 | 固定输入、当前证据与预期 | 最小动作／投入上限 | 状态与出口 |
| --- | --- | --- | --- |
| N13-E1 / P0 证据 | 原 `generated-layout.docx` SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6` 仅两页，无图片／页眉页脚／页码，不能单独证明这些核心用法。 | 复用五页节分页和 40×20 mm 红／蓝 JPEG 生成器；native/default 各生成实际 OFD→PDF，正文、原图片字节和页码自动断言。只增本入口，不扩 Word 保真矩阵。 | 本提交提供 `--cn-en-release-only`；当前候选实际运行、逐页 PNG/Preview 由 N12/Terra 记录。 |
| N13-E2 / P0 字体证据 | 可信 bundle 只有真实 CJK Regular，无物理 CJK Bold；合成 Latin face 测试不构成真实中文粗体证据。原始“数量”历史视觉 FAIL 不消失。 | 单独配置同来源、同许可的固定 CJK regular/bold；以严格 `Throw` 及嵌入 face 断言验证，不改默认字体／后端；字体来源、工具链、weight700、字节哈希由 N00 的生成记录固定。只验具名独立 profile。 | 入口接受 `--release-font-directory`；每次记录目录各文件与实际嵌入 face 哈希。没有资料时保持未完成，不能冒称物理粗体通过。 |
| N02 / P1→核心阻断，独立拥有者 | 协调者转交：Terra 对当前 main 核心日期和“待复核”末尾 Find 的真实 Preview 验证失败，API 可查不能替代 UI 可用。 | N02 原有 4 小时／2 假设预算；本项不重复实现或重置额度。 | **未关闭，候选不可宣称整体通过。** 以 N02 的准确文件／版本／页码与实际 UI 复验为准。 |
| N03 / P1，独立拥有者 | 极小正字号被写为 `Size="0"`。 | N03 局部序列化契约和已有边界测试；本项不改 numeric writer。 | 按 N03 独立交付与集成回归判断，不在此重复立票。 |
| N13 产品修复 | 上述以外，没有从本次静态盘点证明新的普通 CN/EN 产品阻断。 | 不为完整性新增功能／字体算法；若本轮确定性页面出现实际丢字、裁切或样式串染，固定反例后再列入。 | **无产品修复提交**；进入准确候选验证，视觉未验不视为无缺陷。 |

## release-only 运行与记录

`tests/TestSupport/DocxReleaseFixtures.cs` 从已有测试原样共享最小 DOCX、五页节分页及 JPEG 生成逻辑。E2E 的 JPEG 标题增加中英文；旧回归的默认标题、图片字节断言、页码断言不变。没有新增二进制 fixture。无参数运行的旧 E2E 路径不变。

按根 AGENTS 设置环境；在资源协调者分配的构建槽运行。以下命令从仓库根执行，`OFDRW_SOURCE_SHA` 必须填实际构建 commit，不能填某个计划文档 SHA。产物只进 ignored `artifacts/`：

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/cn-en-release/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
# 若使用已有全局缓存，请显式设置 NUGET_PACKAGES，避免 DOTNET_CLI_HOME 改变默认缓存位置。
export OFDRW_SOURCE_SHA="$(git rev-parse HEAD)"
dotnet build e2e/Ofdrw.Net.Layout.E2E -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project e2e/Ofdrw.Net.Layout.E2E -c Release --no-build --no-restore -- artifacts/cn-en-release/core e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --cn-en-release-only --release-font-directory /absolute/path/to/pinned-regular-bold
```

不提供 `--release-font-directory` 时，只生成 original profile；这能验核心链路但不能关闭 N13-E2。有目录时生成八份 original PDF（layout 2×2 页、headers 2×5 页、image 2×1 页、selected 2×3 页）及两份 separately named physical-bold layout PDF（2×2 页），共 26 页；加 N00 的 C01/C04/C05 共 8 页构成冻结的 34 页。每个 PDF 都来自同名本次 OFD；不生成直接 DOCX→PDF 供替代。两种字体配置使用完全相同且未修改的 F01：其原有 `eastAsia` 声明已为 Noto Sans CJK SC。original 从可信 bundle 校验并提取 Regular/Latin 字体及许可证到独立目录，physical-bold 使用提供的目录；两份输入身份相同但配置与输出分别记录。

`cn-en-release-manifest.json` 记录源码、环境、输入/OFD/PDF 大小及 SHA-256、页正文、实际 face/子集字节、诊断与执行的断言；PNG、Preview 和 Copy/Find 初始均为 `not-performed`。它们必须由独立验收证据补充，不能把 manifest 的功能字段改写成视觉通过。完整 Flow 和公开表格仍使用已有 E2E `flow-public`／`--tables-only`，不因本入口而删掉 N00 冻结的样例。

## Terra 实际查看清单

重新生成后打开本轮 manifest 指向的 PDF，逐页核对实际路径，验完关闭窗口。不得查看旧缓存。original 的 native/default 各自检查第 1–3 项；physical-bold 仅重复第 1 项：

1. `layout-*.pdf` 第 1–2 页：中文标题、比例英文斜体、局部粗体／红色与普通段恢复、蓝底表头及边框／对齐、分页、右对齐日期；保留 original 的“数量”独立 FAIL 历史，物理粗体结果另记。
2. `headers-*.pdf` 第 1–5 页：Body one/two/three/four/five 各一次；First/Even/Odd/Even/Odd Header；页脚依次 Page1/5、Page2/5、Page3/5、Page10/5、Page11/5；页眉正文不重叠，页脚不裁切，无额外尾页。
3. `image-*.pdf` 第 1 页：中英文图片标题完整；40×20 mm 图片左红右蓝、宽高比 2:1，未被裁切或重复绘制。
4. `selected-original-*.pdf` 第 1–3 页：与同源 headers 原页 3/1/3 的正文、页眉和页脚相同；页序 three/one/three，重复页未丢内容。
5. N00 的 Flow 与表格样例仍逐页核对局部样式、换行、分页、表格底色／边框及末页。普通 CN/EN Copy/Find 独立执行；N02 核心末尾失败由原拥有者提供正例／负控，不能用 PDF 文本可提取替代。

## 本提交验证

SDK 10.0.401 的 focused Release E2E 构建成功（0 错误，1 条原有 Reader CS8604 警告）；共享生成器相关 DOCX 回归 3/3 通过、0 失败／跳过，包含 JPEG、五页节页码及重复／无效页选择。记录在 ignored `artifacts/cn-en-release-n13/`：`e2e-build-local-cache.log`、`shared-fixtures-tests-local-cache.log`、`tests/shared-fixtures.trx`。首次临时 DOTNET_CLI_HOME 导致新 NuGet 默认缓存和慢还原，已停止本任务还原并改为显式 `NUGET_PACKAGES=/Users/wanghongyi/.nuget/packages`、本地层次缓存源后完成；未清缓存或修改许可门。新增入口最终运行结果另在交接补充。完整回归、自动页面渲染、PNG 目视、macOS Preview 和 Copy/Find 均未完成。最终全套回归及精确包消费由合并后的候选 N12 执行，不在多个工作树重复全套构建。

N00 已提供独立物理粗体制品 manifest：SHA-256 `23da8dc97db7c6b73c2f71fc9da15ed1e18fe0701eb498ea5152eab26ad8a3b8`；同源静态 Bold 为 21,638,668 字节、SHA-256 `647a6e375c5a0d8958111173d7b7cec5237e0784cc4490663e43834dfff491be`，weight 700、glyf、无 fvar，工具链 fontTools 4.59.2 / Python 3.14.7。其具体生成脚本、字体及许可在 N00 ignored artifacts 下保留，未重复提交载荷；它们的运行选择和页面外观仍需本次实际验证。

## 2026-11-02：首次实际运行暴露的验收断言修正

集成 `61120ae` 的首次 release-only 运行在 `layout-original-native` 失败：`Core local bold/italic/color intent differs or leaked into ordinary text`。这是本验收入口的错误：`OfdColor` 是未重载 `==`/`!=` 的 class，先前断言比较了对象身份，导致新建的红色和表头底色对象必然不相等。修正仅将三个颜色判断改为 Red/Green/Blue/Alpha 分量比较（黑色 `0/0/0/255`、强调 `192/0/0/255`、表头 `217/234/247/255`）；粗体数量 7、字重 400/700、斜体和底色数量 3 的预期均未变，未改产品实现或放宽阈值。

实际失败 OFD SHA-256 `ec54355d787ebf8153eabfffceedefdb2f0ccd77b2845347cb43353f47c71158` 中，普通正文为 Weight 400／不斜体／黑色，强调为 Weight 700／红色，英文副标题为斜体，粗体对象恰好 7 个，表头填充恰好 3 个。修正后重生 OFD 的这些 source object 值与失败产物逐项相同，说明之前的拒绝来自断言而非该样例样式发生变化；没有由此推断其页面或搜索通过。

提交修正前已实际构建并运行全部 10 份／26 页 release-only 产物（包括原 Regular、真实物理 Bold、页眉页脚、图片和重复页选择），每份的 OFD 原文、样式、页选择／图片与 OFD→PDF 文字断言全部通过。物理粗体两模式严格 `Throw`、真实 Bold face、无合成粗体诊断的断言也通过。此为开发 smoke，源码身份是基线 `4c7cab02423fac47541bcbcc5fe05a4da0cc0a77` 加验收源码 diff SHA-256 `52a41728197ed67a00c03a17b1b3ac5c43438d0c35942ad2bc39e47f6c0d5d26`，不是已冻结的集成候选；最终 main 仍需重新生成和验收。

本次可复查记录在本 N13 worktree 的 ignored `artifacts/cn-en-release-n13/`：`e2e-value-color-build.log`（构建 0 错误，原有 Reader CS8604 警告 1 条）、`value-color-smoke.log`、`value-color-before-after.json`、`value-color-fix.patch`；完整产物在 `value-color-smoke/`，manifest SHA-256 `7a9b4879cc3daf7e5453464fa356c1ea290318932afec5bc73564403e0fe4425`。该记录保留失败前后 OFD 身份、实际控制对象和值、源码差异身份、10 份产物及物理 face 诊断。没有重跑全套回归；PNG 渲染、PNG 目视、Preview 和 Copy/Find 仍未执行。N02 和原 Regular 量字历史 FAIL 均未关闭，候选仍不能据此宣称整体通过。
