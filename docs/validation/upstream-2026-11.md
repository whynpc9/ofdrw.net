# 上游 2026-11 回归验收

依据 [Plan 019](../../plans/019-upstream-2026-11-regression-and-triage.md)，关联 [issue #120](https://github.com/whynpc9/ofdrw.net/issues/120)。本记录覆盖 A 批次的 M01/M02。B、C 批次未运行。

## A · M01/M02 归档和 CLI 文件入口

实现源码为 `7ddc251fa30b765f7cd0d494393699bfdc7a1a91`，测试树为 `55b1dff39959c1143d8f1ba4660b50d747c5ea6e`。原实现完成以下回归，无修复后结果。快照 SHA-256 为 `84944b9e02ea41278a263530764213ad6c81738c4a88ba8b5a7c1d737767b916`。

样例由 [共享 fixture](../../tests/TestSupport/Upstream202611ArchiveFixture.cs) 生成，按固定条目顺序、ZIP 时间戳和无压缩方式写入。完整类名、方法、Theory 行标识及对应条目列于 [TRX 清单](upstream-2026-11-tests.json) 的 `prs.A`。全部输入哈希见下表。

| 条目和缺口 | 已有覆盖 | 输入和独立预期 | 原实现结果 | 视觉或处置 |
|---|---|---|---|---|
| M01 · 混合大小写分支、模板和同名资源 | `Audit002RegressionTests` 检查绝对、相对及重复图像路径，`PackageRoundTripTests.UpstreamComplexSample_ShouldRetainLayersPathsTemplatesAndAnnotations` 检查复杂模板 | `Doc_0/A` 主页面与 `DOC_0/B/Tpls` 模板，引用使用 `a`、`b/TPLS`；两个 `swatch.png` 经跨分支引用，页面蓝色位于左侧，背景模板红色位于右侧 | 公开 SDK 的 seekable 和非 seekable、受控 owned spool 均得到同一模板 ID、背景层次和各自精确资源字节。CLI 在两个子进程 CWD 下均得到 1 页，图像色值和 30 × 30 mm 几何通过 | 下方页面清单 |
| M01 · 完整路径碰撞、必需页面缺失、ZIP 条目穿越 | `BadPackageCorpusTests` 检查 loader 的碰撞、穿越和缺失文档引用，`ReaderInputHygieneTests` 检查重复模板 ID | 完整路径只改变大小写时应拒绝；删除主页面后不得用模板代替；归档项含 `..` 应拒绝 | 公开 SDK 两种流拒绝，抛 `InvalidDataException`。受控碰撞和穿越抛 `InvalidDataException`；缺页抛 `KeyNotFoundException`。CLI 均退出 1 且无 PDF | 拒绝输入无页面。缺页结论限于必需页面 |
| M02 · 普通 ZIP 与前缀 offsets 已调整 | `ControlledReaderAccountingTests` 已检查读入、模型/载荷计量和公有/受控结果一致 | 77 字节前缀在偏移 16 处包含完整空 EOCD；最终 EOCD 和全部 central-to-local offsets 均描述物理位置 | 公开 SDK 两种流、受控 owned spool 和 CLI 均通过；受控预留 1 次元数据并计量 8 个条目；CLI 两个 CWD 的图像位置和色值均符合预期 | 下方页面清单 |
| M02 · 直接 prepend 与非零流位置 | `ControlledReaderAccountingTests.UnsupportedOrInconsistentFooterRejectsBeforeMetadataAndEntryAllocation` 检查 footer 边界 | 不调整 offsets 时中央目录指向旧位置。对同一字节，公开 seekable 流从 Position=77 读取时缓冲尾部完整 ZIP | Position=0 的公开 seekable/非 seekable 均拒绝；受控在元数据预留前拒绝；CLI 退出 1 且无 PDF。公开 SDK 的 Position=77 通过；受控拒绝 Position=77 和非 seekable spool | 记录 .NET 10 实测边界，保持受控 profile |
| M02 · 截断和伪 local header | `BadPackageCorpusTests.RejectsFixedMalformedZipSeeds`、`ControlledReaderAccountingTests.UnsupportedOrInconsistentFooterRejectsBeforeMetadataAndEntryAllocation` | 在完整归档上截去 footer 末尾 8 字节，或把首个 local header 签名改为 EOCD 签名并保留中央目录 | 公开 SDK 两种流均抛 `InvalidDataException`；受控分别抛 `NotSupportedException` 和 `InvalidDataException`，未分配 DOM。CLI 均退出 1 且无 PDF | 结论限于这两个损坏布局 |
| M02 · 前缀计入预算与取消 | `InputBudgetTests` 检查公开预算和取消；`ControlledReaderAccountingTests.CancellationDuringZipMetadataConsumesContextWithoutPayloadOrModel` 检查受控元数据期取消 | 最大输入字节等于无前缀 ZIP 大小，调整 offsets 的前缀 ZIP 必须超限；预先取消应终止读取 | 公开 SDK 两种流及受控均分别抛 `InvalidDataException` 与 `OperationCanceledException` 派生类型，受控未预留元数据或分配条目 | CLI 此命令无对应预算参数，取消未运行 |

复验 CLI 拒绝输出，六个输入均退出 1 且目标 PDF 未发布。碰撞、缺页和穿越分别报告 `duplicate entry`、`references a missing entry` 和 `unsafe path`。直接 prepend 报告最终 EOCD 的条目数与中央目录不符，截断输入报告找不到 EOCD，伪 local header 报告本地头损坏。

### ZIP 偏移记录

| 输入 | 最终 EOCD 位置 | 中央目录字节数 | 声明 / 物理中央目录起点 | 声明 local header offsets（中央目录顺序） |
|---|---:|---:|---:|---|
| 普通 ZIP | 2247 | 516 | 1731 / 1731 | 0, 209, 558, 802, 1026, 1259, 1497, 1614 |
| 调整 offsets | 2324 | 516 | 1808 / 1808 | 77, 286, 635, 879, 1103, 1336, 1574, 1691 |
| 直接 prepend | 2324 | 516 | 1731 / 1808 | 0, 209, 558, 802, 1026, 1259, 1497, 1614 |

### 功能和门禁

环境为 macOS 26.6.2（25G83，arm64），.NET SDK 10.0.401、运行时 10.0.12。四次测试命令均退出 0，完整执行 Packaging 1,710/1,710、CLI 281/281，零跳过；其中 `ControlledReaderAccountingTests` 48/48。独立筛选结果为 Packaging 10/10、CLI 8/8；TRX 检查器在完整与筛选结果中均逐条匹配。合成 TRX 自测 5/5，覆盖失败、未执行、零测试、缺少数据行、错误程序集、缺文件及计数不符。

生成及复验命令，`$RUN` 为每次新建的绝对结果目录，按 [Plan 019](../../plans/019-upstream-2026-11-regression-and-triage.md) 设置 .NET 环境变量。

```sh
for project in Packaging Cli; do
  dotnet test "tests/Ofdrw.Net.$project.Tests/Ofdrw.Net.$project.Tests.csproj" -c Release \
    --logger 'trx;LogFileName=full.trx' --results-directory "$RUN/$project/full" \
    --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
done
for project in Packaging Cli; do
  OFDRW_120_ARCHIVE_ARTIFACT_DIR="$RUN/visual" dotnet test \
    "tests/Ofdrw.Net.$project.Tests/Ofdrw.Net.$project.Tests.csproj" -c Release \
    --filter 'FullyQualifiedName~Upstream202611' --logger 'trx;LogFileName=new.trx' \
    --results-directory "$RUN/$project/new" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
done
python3 scripts/validation/upstream-2026-11/check-test-results.py \
  --manifest docs/validation/upstream-2026-11-tests.json --results-dir "$RUN" --pr A
python3 -m unittest discover -s scripts/tests -p test_upstream_202611_test_results.py
```

文档构建、类型检查、链接和搜索索引生成通过。Nimbus Authoring 跳过 47 个 Markdown 页面。存储检查含已 materialize 的 `e2e` 和 `tests`。全套本地回归、可信字体专项、同批包消费未运行，A 批次的共享实现变更条件未触发。

### 页面清单

样例只使用两块 1 × 1 RGB PNG，没有文字或字体依赖。100 × 70 mm 单页上，左蓝主页面对象位于 (10, 20) mm，右红背景模板对象位于 (60, 20) mm，各 30 × 30 mm。预期其余区域留白，两块不重叠、不裁切。

- **自动渲染**　真实子进程运行 `dotnet <absolute-cli-dll> ofd-to-pdf -i <absolute-input.ofd> -o <absolute-output.pdf>`。CLI 测试用 PdfPig 0.1.15 检查页数、图像包围盒和色值。对两种 ZIP 各两个 CWD 的 PDF 执行 `pdftoppm -png -r 144 -singlefile <absolute-output.pdf> <absolute-page-stem>`（Poppler 26.09.0），4 份单页 PNG 逐字节一致。
- **PNG 视觉复查**　Sol High 查看第 1 页。左蓝右红，方块对齐，边缘完整，间距和留白符合上述预期。其余三个第 1 页的 PNG 与复查图相同。
- **macOS Preview**　独立验收由 `gpt-6.1-terra`（High）执行。核对两份 OFD 和对应 PDF 的哈希后，通过 `OFD → CLI PDF → Preview 11.0`，对普通及调整 offsets 的 PDF 分别按绝对路径重新打开，确认各 1 页。两次均见左蓝右红、等大且对齐，未见裁切、重叠、重影或多余页面。查看后分别关闭窗口并确认 Preview 无开放窗口。两张截图逐字节相同，保留其中一张。
- **保留资料**　[输入 OFD](../evidence/upstream-2026-11/A/input.ofd)、[CLI PDF](../evidence/upstream-2026-11/A/output.pdf)、[第 1 页 PNG](../evidence/upstream-2026-11/A/page-001.png)、[Preview 截图](../evidence/upstream-2026-11/A/preview-page-001.png)。

本地记录根目录为 `/Users/wanghongyi/.codex/worktrees/ea52/ofdrw.net/artifacts/upstream-2026-11/A-55b1dff39959c1143d8f1ba4660b50d747c5ea6e-zIEOy9`。下表路径相对于该目录。

| 文件 | 字节 | SHA-256 |
|---|---:|---|
| `visual/adjusted/alternate-page-001.png` | 1294 | `bab4050a9bff4ef37f69ad997e112732943b8dd9b4299ea2729c7fa9c7deec17` |
| `visual/adjusted/alternate.pdf` | 1815 | `c53b1e9b97d7d3bbd907086be41c78361878a49406c0723263ba3358af4fbcb3` |
| `visual/adjusted/input.ofd` | 2346 | `de3cf9034980c4255fbc6c16572753ea5426650889325927ec0a19f024096ab7` |
| `visual/adjusted/output.pdf` | 1815 | `c89c2f4676eb2c3f1e9621db440cb27da914704d79faea193f18a0a966672a94` |
| `visual/adjusted/page-001.png` | 1294 | `bab4050a9bff4ef37f69ad997e112732943b8dd9b4299ea2729c7fa9c7deec17` |
| `visual/archive-template-case/alternate-page-001.png` | 1294 | `bab4050a9bff4ef37f69ad997e112732943b8dd9b4299ea2729c7fa9c7deec17` |
| `visual/archive-template-case/alternate.pdf` | 1815 | `c8e6ac0c031b3db00a9b6893b47bc52fd1eeb3cc5c2de1e461d084d72cd52338` |
| `visual/archive-template-case/input.ofd` | 2269 | `2a040d1a4fd89d0e57e2c77953997b123862194933dbe86b229822af527c9fad` |
| `visual/archive-template-case/output.pdf` | 1815 | `ed73b7e1b7a336897488da3556b0a8b7eb35c519102fbea8de501599c15f916c` |
| `visual/archive-template-case/page-001.png` | 1294 | `bab4050a9bff4ef37f69ad997e112732943b8dd9b4299ea2729c7fa9c7deec17` |
| `visual/collision/input.ofd` | 2391 | `d454768b61aa07cc2ebfa4a1be73c43b5508bfd0a84dbcd61ec54c76eb4a5f06` |
| `visual/false-local-header/input.ofd` | 2269 | `e1fcde40d475def5c610a2c0579b63c39e71ccb020278c3f867f8ab0f4b027a0` |
| `visual/missing-page/input.ofd` | 1960 | `b6cb1d8bcf7f633ae2971462b8f8342c4e6e9b7018b6e253a4ccf876d03686e8` |
| `visual/traversal/input.ofd` | 2391 | `370fc7e087d26884884fb55af4ee4f9b907c08750c1adb20fc69628957dd2cc1` |
| `visual/truncated/input.ofd` | 2261 | `75a486146ac5b19e07629c5d538bdcf66feedd3deee2102c46b64927268c1630` |
| `visual/unadjusted/input.ofd` | 2346 | `9b199ae88c149ca4a4b3212ec7fbb14885875b8cac3a2a43f49745861c33f0cf` |
| `visual/adjusted/preview-page-001.png` | 66189 | `a8c0824d10165c1829ef3e21d1caffefdfb8b788270aa8f3d6209023abdaf68f` |
| `visual/archive-template-case/preview-page-001.png` | 66189 | `a8c0824d10165c1829ef3e21d1caffefdfb8b788270aa8f3d6209023abdaf68f` |

## PR B 绘制与文字

相关项目为 `Converter.Pdf`。完整用例身份及所属条目见测试清单的 B 段。设计复核使用 `gpt-6-astra` High。以下结论限定于自生成页面与现行渲染 profile。

| 条目 | 既有证据 | 新增判据与入口 | 处置 |
|---|---|---|---|
| M04 Cap、Join、DrawParam LineWidth | `R01PageBlockTests.DrawParamColorsUseDrawingNamespaceAndKeepLocalOverrides`；`Audit015_CapsAndJoins`；`PenCapsAndDash_AffectRenderedInkWithoutChangingDefault` | 读包后分别判定缺失、显式默认、非默认、删除 XML 后重读四种状态。祖先 Square/Bevel/5，最近 Round/Round/4，本地 Butt/Miter/0.353，后续兄弟仍采用最近值。32 项继承可解析；33 项与环经 PDF/SVG/HTML/PNG 拒绝且零输出。 | 继承、覆盖、边界符合当前契约。OFD Border LineWidth 在本仓没有与 Java getter 对应的渲染模型；表格生成器将边线写成 PathObject。 |
| M04 对象 Alpha、Color Alpha、BitsPerComponent | `Audit015_ObjectAlpha`；`Audit015_FreshDefaults` | GraphicUnit Alpha 按对象验证；Color Alpha 按颜色验证，M05 用实际蓝色像素核对乘积。ColorSpace BitsPerComponent 缺省及显式 8 保留原 RGB，显式 4 经四出口拒绝且零输出；移除 4 后重读恢复缺省。 | Alpha 和 BitsPerComponent 按各自模型处理。 |
| M04 文字 Stroke | `StrokedScaledTextUsesInheritedPaintAndKeepsLogicalText`；`OutlineOnlyTextWithoutStrokeColorUsesOpaqueBlack`；`SimulatedBold_ShouldThickenAtTheOriginalBaselineWithoutDuplicateText` | 导入 XML 九格为缺省/false/true × 无色/本地蓝色/继承红色；另附 producer 黑色合成粗体对照。模型与 SVG/HTML paint、PDF 轮廓操作及 PDF/PNG 像素一致；源 OFD、PDF、SVG/HTML 文字按输入顺序各保留一次。各行右侧放置相同 H，用独立墨迹计数验证无意描边及黑色描边粗细。 | 缺省和 false 保持仅填充，true 无色使用默认黑色。 |
| M05 裁剪与 Alpha | `TypedGraphicRenderingTests.SharedNestedVectorPlacementsDrawTwiceWithIntersectingClips`；`Audit015_ObjectAlpha` | 两个 Clip 的交集有红墨，差集白底。嵌套组合内第二个蓝色对象独立于红色 clip，仍受外层约束，后续绿色兄弟独立。Alpha 0/128/255 与颜色 Alpha 128 的乘积经 SVG/HTML 结构、PNG 内部像素及 PDF 页面复查。 | 路径 Alpha 与颜色 Alpha 相乘；白底上的蓝色样点约为 (191,191,255)。维持现有 multi-Area、Text clip 与 group opacity 边界。 |
| M05 图片、路径 CTM | `ImageTransformAndPenStyles_ReachPdfSvgAndHtml`；`SvgGeometryTests` | 对 32×16 四色图以 `B + A·(u,v) + t` 手算变换角点及象限中心，检查整页 Boundary 内小图、旋转、剪切、反射和平移。绝对路径取 `A=I, t=-B`，图形仍在 (56,52)–(69,66) mm。SVG/HTML 查矩阵及路径，PNG 查色块内外，PDF 逐页复查。 | 单次变换。保持既有边界和坐标契约。 |
| M06 Delta、旋转和字形替换 | `SvgGeometryTests`；`ExplicitGlyphTransformTests` | 大平移 `t=(100,200)` 下比较实际 PDF 字符基线与手算 `B+A·p+t`，剪切增量为 (8.5,4.75)，旋转增量为 (-3,7)。第二个 TextCode 起点独立；顶层非恒等映射将 AA 绘制为两个 B（GID 37）；PDF 内嵌字体、逻辑原文和基线均有断言，SVG/HTML 检查对应路径，PNG 检查各位置有墨迹。非零 ReadDirection/CharDirection 原 XML 可写回，四出口拒绝且零输出。 | 支持当前水平字形经 CTM 变换。真竖排维持明确拒绝。 |

页面各 1 页。PNG 经 PDFium 使用 PDF 转换链路，HTML 使用 SVG 渲染链路。几何和色彩预期来自独立坐标推导与内部采样，路径和位图平坦区域每通道允许 3 个量化单位。该样例集包含拉丁字体测试、路径、位图和组合图元。

| 页面 | 输入与预期 |
|---|---|
| `text-stroke` | 100×110 mm。依次为 absent/false/true 三组各三行，末行为 producer。仅 `true-local` 有蓝边、`true-inherited` 有红边，`true-none` 和 producer 为黑描边；每行均可读，右侧同一 H 作为粗细对照。 |
| `clip-scope-alpha` | 100×65 mm。红色交集位于 x=18–28/y=16–26 mm；浅蓝受 x≤64/y≤34 mm 限制；独立绿色兄弟完整。全透明区域保持白底。 |
| `image-path-ctm` | 110×80 mm。上方左至右为常规、90° 旋转、剪切；下方为水平镜像与紫色矩形。各象限色与代码注释所列角点、内部采样相符。 |
| `text-delta` | 230×250 mm。文字集中在页面下方，左侧剪切 AB/CD，中央旋转 EF，右侧剪切后的两个显式 B 字形保留逻辑 AA；原始 Boundary 留白。 |

产物生成时设置 `OFDRW_TEST_ARTIFACTS=<run>/visual` 并执行：

```sh
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release \
  --filter 'FullyQualifiedName~Upstream202611' --logger 'trx;LogFileName=new.trx' \
  --results-directory <run>/Converter.Pdf/new --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
pdftoppm -f 1 -l 1 -singlefile -png -r 144 <case>/output.pdf <case>/poppler-001
```

### PR B 门禁记录

- **基线**　实现基线 `7ddc251fa30b765f7cd0d494393699bfdc7a1a91`，新增用例的测试树 `95c7d90d688e2ce8a88ca50aabdec79d8d980b0b`。可信字体已准备，完整项目 1280/1280，单独筛选新增 8/8，均无跳过。现有实现通过该范围的判据，真竖排和超出 paint profile 的负例按契约拒绝。生产实现未发生变更。
- **功能**　在 A 合入后，以测试树 `89b2aef47016e2fdf12d4f115b126f1a456abeba` 运行同样的两套测试。完整项目 1280/1280，含 `SvgGeometryTests` 14/14；筛选新增 8/8；两次退出 0，无跳过。A 的 `check-test-results.py --pr B` 对本轮 TRX 逐一核对项目、类、方法及完整显示名，全部匹配。
- **自动渲染**　四份 PDF 均为一页。每案生成 SDK PNG、SVG、HTML，另从实际 PDF 用 Poppler 26.09.0 以 144 DPI 渲染。macOS 26.6.2 arm64，.NET SDK 10.0.401，PdfPig 0.1.15，Docnet.Core 2.6.0。内嵌 Noto Sans Regular 的 SHA-256 为 `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5`。
- **PNG 视觉复查**　Sol High 查看四张 Poppler 第 1 页；Terra High 独立查看同四案的 PDFium 和 Poppler PNG。裁剪区域、四色象限、文字位置和颜色均符合上表。粗体对照在初次视觉复查后改用 0.4 mm 描边，正式复查确认仍可与普通字形区分且文字可读。
- **macOS Preview**　`gpt-6.1-terra` High 核对本次输入和输出哈希，从绝对路径重新打开四份最终 PDF，沿 OFD → PDF → Preview 11.0 各看第 1 页。红色交集、浅蓝及绿色兄弟、四图仿射、原点位移与文字粗细均符合判据。文字页未见意外裁切、重叠或重影；文字位移页滚动至底部复查 AB/CD、旋转 EF、可见 BB。四个窗口逐个关闭。
- **存档**　[PR B 页面证据](../evidence/upstream-2026-11/B/drawing-pages.tar.zst) 为 2365356 字节，SHA-256 `9492f6d9c6035ae668b87f1032929239b4d3f1f4f2bb513e02f88a2a1d62cba6`。归档包含四份 OFD、PDF、SVG、HTML、两种 PNG、仅 Preview 窗口截图，以及逐文件大小和 SHA-256、运行命令、环境、TRX 核对及视觉复查记录。证据生成命令与路径对应测试树，验收说明和归档的提交单独记录。
- **本地复现资料**　正式根目录为 `/Users/wanghongyi/.codex/worktrees/69e8/ofdrw.net/artifacts/upstream-2026-11/B-89b2aef47016-formal.rLjasg`；未修复基线根目录为 `/Users/wanghongyi/.codex/worktrees/69e8/ofdrw.net/artifacts/upstream-2026-11/B-95c7d90d688e-baseline.QomkSZ`。各目录的 `sha256-sizes.txt` 记录该轮输入和输出，包括拒绝类 OFD。拒绝类输入由测试现场生成并留在本地。四个正式 PDF 总计 325150 字节。
- **文档与范围**　站点构建、类型、48 个 HTML 页面链接、搜索索引生成均通过；Nimbus 跳过 47 个 Markdown 页面的 Authoring lint。B 未修改共享实现，因此全套本地回归与本地包消费门禁由各自触发条件及 C 的集成验收覆盖。
