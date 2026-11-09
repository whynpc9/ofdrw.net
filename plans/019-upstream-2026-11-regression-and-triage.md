# Plan 019：上游 2026-11 差异验证与有界能力筛选

- 状态：TODO（已处理 R1–R7 评审意见，尚未执行测试或开发）
- 日期：2026-11-09
- 基线：`7ddc251fa30b765f7cd0d494393699bfdc7a1a91`；本地 main 与读取时远端 main 一致；产品版本 0.3.0。
- 来源：[issue #120](https://github.com/whynpc9/ofdrw.net/issues/120)，读取时 open、无评论。
- 优先级：P1 验证；P2 候选设计；P3 延期。
- 工作量：首轮约 5–8 人日（不含新发现缺陷修复、等待 CI 与后续功能开发）；估算需在覆盖去重后修订。
- 风险：补测试低；更改资源解析、几何或字体中高；替换 PDF 后端高。
- 前置：无未完成计划依赖；复用 001–018 已交付能力，尤其 005 诊断、013 依赖门禁、014 字体生命周期结论。

## 必要性结论

值得做一次有界回归和能力处置，不值得按 Java 2.5.x 全量移植，也没有证据要求立即发布修复版。issue 本身明确未复现新的 .NET 缺陷。当前机制和既有测试足以否定“路径大小写、绘制继承、骑缝章、真实字体排版尚未实现”等宽泛立项理由；但不能证明本月的边界组合已经通过。

第一期只交付 M01/M02/M04/M05/M06/M10 的覆盖核对、新增独立判据、实际运行结果与必要修复。M03/M07–M14 必须有明确处置，但不需要全部实现才能完成本期。公开 issue 更新、提交、推送与发布不属于本次计划编写的已执行动作。

本次计划的核查范围是 issue、当前本地实现与相关测试、构建脚本、仓库契约；没有重新执行上游 Java 测试，没有逐项复核所有上游提交/资产许可证，也没有运行新的 .NET 测试或 Preview 验收。下文的“需验证”不能解读成已确认缺陷。

## 当前实现与执行锚点

下列路径相对仓库根目录。开始实施先执行：

```sh
git status --short
git diff --stat 7ddc251fa30b765f7cd0d494393699bfdc7a1a91 -- src tests e2e scripts docs
```

状态或 diff 显示范围内文件变化时，先核对实时代码。此 diff 包含已跟踪文件的暂存/未暂存内容；未跟踪计划、测试、fixture 另由 status 识别并读取，也视为漂移。不得用 checkout/reset 或覆盖用户修改来消除漂移。

| 位置 | 当前事实与影响 |
|---|---|
| `src/Ofdrw.Net.Packaging/Archive/OfdPackageLoader.cs:50` | `new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase)`；完整路径重复拒绝。M01 应验证组合与模板定位，不移植磁盘递归算法。 |
| `src/Ofdrw.Net.Packaging/Archive/OfdAccountedPackageLoader.cs:28` | 受控路径只审计 .NET 10，内部要求可读、可 seek、position=0 的 owned spool；最终无注释 EOCD，且 `centralOffset + centralSize == length - 22`。拒绝某类前缀布局可能符合契约。 |
| `src/Ofdrw.Net.Reader/Readers/OfdPaintResources.cs:88` | 复制源节点；仅对象未声明属性时补 DrawParam；Relative 链就近覆盖、循环/32 层边界拒绝。允许属性有白名单，Alpha 等不能无差别加入继承。 |
| `src/Ofdrw.Net.Core/Models/OfdTextPaint.cs:18` | `Stroke = Boolean(root?.Attribute("Stroke")?.Value, root is null && text.ProducerStrokeColor is not null)`；导入 XML 和 producer 规则不同。 |
| `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:605` | 图片裁切后调整矩阵，再在局部 Save/Restore 内处理 clip、CTM 和 Boundary。约 907–946 行文字在局部累计 Delta。机制存在，仍需独立几何判据。 |
| `src/Ofdrw.Net.Core/Models/OfdRenderPreflight.cs:92` | 非零 ReadDirection/CharDirection 不在当前接受范围；旋转水平文字不能扩展成真竖排承诺。 |
| `src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1128` | 图片按 BaseLoc + MediaFile 经 ResolvePath 查询，未在这条解析路径实现 basename 搜索。M03 会改变资源身份，不适合作默认修复。 |
| `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageOptions.cs` | 已有 PNG/JPEG 编码、`PixelsPerMillimeter` 分辨率、JpegQuality=90 与各层预算；没有 PPM 文件编码。 |
| `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs:97` | `new Image<Rgb24>(..., Color.White)`，PNG 编码为 RGB。透明背景需先验证 PDFium 返回值是否保留所需 alpha，不能只加布尔选项。 |
| `src/Ofdrw.Net.Layout/Internal/Flow/FlowLatinMetrics.cs:76` | 兼容排版仍使用静态宽度表；不代表显式字体引擎缺失。 |

测试沿用 xUnit `[Fact]`/`[Theory]`、内存最小 OFD、结构断言与确定性像素断言。复用 `R01PageBlockTests.cs` 的资源/继承构造，`GraphicsImageAndPenTests.cs`、`TypedGraphicRenderingTests.cs` 的图形样例，以及 `StampClipRenderingTests.cs` 的两页双色章。后者已断言两页各自片区、片外白色、非法 clip 的严格图片导出零输出，不能重复计为新增覆盖。

## 全部条目的处置

| 条目 | 本期决定 | 后续触发条件 |
|---|---|---|
| M01 | 接受差异回归；现有机制复用 | 失败且违反当前路径契约时修复 |
| M02 | 接受分入口、分 ZIP 布局的复现 | 先区分合法兼容缺口、损坏包和受控 profile 限制，再决定修复 |
| M03 | 延期，保留显式“仅图片、当前文档、唯一候选”方案 | 有真实失败文档/调用场景，且定义诊断、写回及签名边界后独立立项 |
| M04 | 接受差异回归 | 七类属性逐一判定适用性；不创建 Java 同名 getter |
| M05 | 接受几何、clip、alpha 独立判据 | 只修证实的支持范围内错误；绝对坐标启发式延期 |
| M06 | 接受 Delta/CTM 回归；真竖排延期 | 真竖排须独立 profile、字体/标点/提取顺序验收 |
| M07 | 复用真实字体引擎；兼容入口量化比较延期 | 有断行/字宽问题再比较同一字体；不切换默认引擎 |
| M08 | 字体格式扩展延期 | 先列格式×读取/轮廓/advance/嵌入/子集/语义×出口支持表；具名格式有实际需求才开发 |
| M09 | 原生 Pattern/渐变/Multiply 延期 | 用具体文档选择一项有界 profile；现有栅格效果不等于原生支持 |
| M10 | 印章差异回归接受；背景色/透明 PNG 为首选功能候选 | 首期回归完成后做 alpha 可行性验证；旋转、BMP、DPI 元数据按需求独立排期；PPM 编码器本期不引入（见后文来源核对）；流输出批量命名暂不引入 |
| M11 | PDF 导出结构核对作为后续调查；不替换后端 | 有书签/附件/交互注释保留需求后制作单一样例，明确保留/展平/诊断/拒绝；Copy/Find 单独验收 |
| M12 | 复用双层/只图/vector 分层；比较与优化延期 | 以代表性 PDF 测大小、耗时、峰值及质量后决定优化，不采用上游样本收益作目标 |
| M13 | MultiMedia typed/播放支持延期 | 先验证未知 XML 保留及拒绝边界；出现明确需求才增加静态模型，不自动执行媒体 |
| M14 | 复用 D10 与计划 013；具体移植前复核许可 | 不改本仓 LICENSE，不移植 Java 依赖版本；本计划不声称 ImageSharp 告警消除 |

## 真实入口矩阵

| 层次 | 方法/命令与测试位置 | 输入与判定 |
|---|---|---|
| 普通 SDK | `OfdReader.ReadAsync(Stream, OfdPackageLoadOptions, CancellationToken)`；Packaging 新增 `Upstream202611ArchiveTests` | seekable、非 seek；非零 Position 单列现行行为及契约，不自动承诺相对起点 ZIP 支持 |
| 内部受控读取 | `OfdReader.ReadControlledAsync(Stream, OfdPackageLoadOptions, OfdControlledEditContext, CancellationToken)`；Packaging 已有 InternalsVisibleTo，沿用 `ControlledReaderAccountingTests` 构造 context | 正例传稳定可读、可 seek、Position=0 的 owned spool；非 seek、非零 Position 等验证内部拒绝。不是公共 API，不新增公共受控读取方法 |
| CLI 文件 | `ofd-to-pdf -i <绝对输入路径> -o <绝对输出路径>`；`tests/Ofdrw.Net.Cli.Tests/Upstream202611ArchiveCliTests.cs`，参考 `OfdToPdfDiagnosticsCliTests.cs` | CLI 打开文件后进入流转换；非 seek 不适用；用子进程设置 WorkingDirectory，绝不修改整个测试进程 CWD |

CLI 具体进程命令：`dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll ofd-to-pdf -i <absolute-input> -o <absolute-output>`（仓库根目录调用示例；子进程改变 WorkingDirectory 前，测试 harness 必须将 DLL 路径也解析为绝对路径。构建后确认 AssemblyName/产物路径，若漂移须更新命令）。内部受控读取和公共 CLI 是两层证据，不能互相替代。

## 实施顺序与 PR 划分

采用**每个 PR 独立交付**。推荐 A→B→C 顺序，每个分支从当时已合并 main 建立；后续分支若提前工作，合并前 rebase 并重跑受影响门禁。A/B 无需等待 C，但必须各自在最新 head 完成下表门禁；C 仅汇总前序真实记录，不能补签前序未完成的验收。

| PR | 合并前门禁 |
|---|---|
| A | Packaging 与 CLI 完整测试项目；新增用例清单/TRX 核对；CLI 生成页面按视觉矩阵验收。修改共享读取/转换实现时增加全套回归、可信字体及同批包消费 |
| B | Converter.Pdf 完整项目（包含 SvgGeometryTests）；若改 Core/Layout，运行相应完整项目；新增用例清单/TRX 核对；本 PR 受影响页面 PNG + Preview。修改共享排版/字体/转换实现时全套回归、可信字体及同批包消费 |
| C | 印章所属 Converter.Pdf 完整项目及受影响项目；新增清单/TRX 核对；印章页面 PNG + Preview；最终集成完整回归和同批包消费，复查前序记录身份 |

所有 PR 都需当前 head 适用 CI 成功。修改实现后必须重新生成/复查其影响产物，旧页面验收不得沿用。发现需要修复时，在 019 下补具名子任务，写清文件、符号、失败测试、修复边界与 STOP 条件。

### PR A：覆盖清单与归档边界（约 1–2 人日）

1. 在 `docs/validation/upstream-2026-11.md` 建验收表：条目、现有测试、缺口、输入/出口、独立预期、基线结果、修复后结果、视觉记录、处置理由；逐项记录新增测试完整名称（Theory 含数据行身份）和项目。M10 分开记录分辨率与编码格式。所有未运行格写“未运行”；仅标记实际测试的范围。
2. 核对既有断言。相同场景直接链接测试，不另造 fixture。为新增案例使用 `Upstream202611` 测试类前缀，方便精确筛选；共享生成器只抽取本轮确实复用的部分。
3. M01：生成 `Doc_0/A` 与 `DOC_0/B` 并存、模板 TPLS/Tpls、资源跨大小写分支的样例；按真实入口矩阵对照；完整路径大小写碰撞、缺失、穿越拒绝。工作目录变化只用独立 CLI 子进程和绝对 fixture 路径，不污染其他测试。
4. M02：普通 ZIP 对照，以及 77 字节前缀包含空 EOCD 的样例。分别生成“物理 offsets 已调整”和“直接 prepend、offsets 未调整”版本，记录中央目录和 local header offsets，不混为同一合法布局。按真实入口矩阵测普通 SDK、内部受控读取和 CLI；受控正例使用 owned spool，非法输入单独验证既有拒绝。
5. 对截断、伪 local header、重复/冲突条目、穿越、预算和取消保留负例。区分明确违反格式的包与只是 ZipArchive 接受的布局；不要求无依据地拒绝所有复杂 ZIP。
6. 先在“未修复实现 + 新测试”的测试树运行。记录实现源码 SHA、测试树提交 SHA、fixture SHA-256；未提交测试树用基准 SHA + 包含新增文件的可重建补丁及其 SHA-256 标识，不能声称原始提交自带新测试。当前实现通过或符合契约的拒绝即可结案，不要求制造失败。只有实际修复实现的条目才要求同一测试/fixture 修复前失败、修复后通过，两次均绑定源码与实际包身份。新用例失败先记录原因；仅违反现行支持契约的失败才升级成缺陷。修复单独成提交，保留前后结果。受控 ZIP 新布局支持必须重算 accounting，不能删除 footer 检查或无限扫描首个 PK。

验证：完成 PR A 门禁表、下文完整项目测试与新增用例 TRX 校验；ControlledReaderAccountingTests 必须实际执行。验收表每个 M01/M02 子场景有入口级结论，并完成 A 的 CLI 页面验收。

### PR B：绘制与文字差异回归（约 2–3 人日）

1. M04：缺失/显式默认/非默认/删除后重新解析四状态；Cap 缺失继承 Round、显式 Butt 覆盖；Relative 最近值与兄弟隔离；循环/既有深度限制。Alpha、LineWidth、BitsPerComponent 分别按所属模型验证，不把不允许属性塞进 DrawParam。
2. XML 文字建立 Stroke 缺省/false/true × StrokeColor 有/无 × 本地/DrawParam 的有效组合；保持 producer 合成粗体对照。检查模型、PDF/SVG 结构与实际像素，文字保持原文且不重复。
3. M05：两对象位置明确不重叠，第一对象小 clip、第二对象内部取样；加父 clip 和下一兄弟，证明既不泄漏也不越过父约束。覆盖多 Clip 相交、nested composite、Alpha=0/128/255 与颜色 alpha 乘积。
4. 图片采用非方形四色块，覆盖整页 Boundary+小 CTM、90°、剪切、反射、平移；手算四角、可见区域与选取的内部像素。路径“绝对数据+CTM 抵消 Boundary”先写出坐标方程，不能用渲染结果自证，也不引入 10% 尺寸猜测。
5. M06：验证 `p'=A·p+t`、`Δ'=A·Δ`。e/f=100/200、旋转/剪切、多 TextCode/支持的 CGTransform 路径；纸张与位置须保证可见，避免采空白导致误通过。真竖排保持明确拒绝和 XML 保留。
6. PDF、SVG/HTML、PNG 各列适用出口；当前不支持的 multi-Area union、Text clip/group opacity 不扩容。PNG 经 PDFium 与 PDF 输出共用链路，不能作为独立实现交叉证明，仍需数学/XML/预期区域判据。

验证：完成 PR B 门禁表、完整项目测试与新增用例 TRX 校验；SvgGeometryTests 必须实际执行。本 PR 合并前按视觉矩阵保存和查看适用产物。像素容差预先定义，内部采样避开抗锯齿边缘；不得靠放宽阈值或替换 golden 消除失败。触及共享排版/字体时执行全套回归、可信字体与包消费门禁。

### PR C：印章与整体验收收尾（约 2–3 人日）

1. M10 复用双色骑缝章，补未覆盖的 nested OFD/位图、页选择/重复页、不可读载荷、预算及取消。
2. PDF 默认容忍路径验证 `ConvertWithResultAsync` 对未绘制原因的诊断；图片严格路径验证失败和适用的无发布输出保证。区分转换前/编码前拒绝与目标流在最终发布期间 I/O 失败，不能承诺任意流出错都能回滚已写字节。SVG 签章投影按自身能力记录。
3. 执行完整回归和同批不可变包消费；相关输出逐页 PNG 复查及 macOS Preview 验收。生成文件后重新打开、结束关闭窗口。查看链路写明 OFD→PDF→Preview；涉及 DOCX 改动时覆盖 `generated-layout.docx` 的默认与显式 native。
4. 按下文视觉矩阵使用实际可用的图片查看工具及 macOS Preview；产物验收不以子代理报告或固定型号为依据。人员/代理选择遵循执行会话的用户指令（包括 GUI 验证偏好），不在可移交计划中固化型号或工具参数。GUI 单一持有者；工具不可用明确记录未完成，不宣称通过。
5. 完成全部 M01–M14 处置记录，更新用户可见能力/诊断变化的文档。最新 PR head 的适用 CI 完成；汇总 issue 的关闭只表示条目有处置，延期能力保持未实现标记。

验证：完成 PR C 门禁表和新增用例 TRX 校验；功能测试、自动渲染、PNG 复查、Preview 四栏分别有真实结果。缺少必要视觉验收、可见缺陷或新测试失败时不得置 DONE。

## 页面级视觉矩阵

PR A 建清单，A/B/C 各自生成和签收自己的产物。输出根目录为 `artifacts/upstream-2026-11/<PR>-<headSHA>-<runID>/visual/<case>/`；每行记录输入 OFD 的绝对路径、大小/SHA-256，输出路径/哈希/页数、生成命令、工具/字体版本、预期、实际检查人或代理、查看工具和逐页结果。重新生成使用新 runID，不覆盖已验收记录。

| case | 所属 PR / 计划页数 | 文件与独立预期 |
|---|---|---|
| archive-template-case | A / 1 | input.ofd、output.pdf、page-001.png；不同目录分支资源和模板出现在明确位置，切换 CWD 结果一致 |
| text-stroke | B / 1（矩阵放不下时增页并预先更新清单） | input.ofd、output.pdf、page-001.svg、page-001.png；缺省/false 不描边，true 与 producer 粗体对照，原文不重复 |
| clip-scope-alpha | B / 1 | 同上；第二对象内部像素不受第一 clip 影响，父 clip 外为空白，alpha 乘积满足预定值 |
| image-path-ctm | B / 1 | 同上；非方形四色块四角及路径单次变换与手算坐标一致 |
| text-delta | B / 1 | 同上（只限接受的 SVG 文字 profile）；位移只乘线性部分，原点平移一次，多 TextCode/CGTransform 的必要组合有具名区域 |
| riding-seal-bitmap | C / 2 | input.ofd、output.pdf、page-001.png、page-002.png；两页仅各自红/蓝片区，片外不污染 |
| riding-seal-nested | C / 2 | 同上；嵌套 OFD 印章同样只显示分配片区 |

HTML 包装入口适用时补 `output.html` 及检查记录；印章 SVG 不作对等保证。不要求每个单元测试输出所有格式；上述是最小必要页面级组合，实际 fixture 变化必须同步清单，不能静默漏页。拒绝类用例记录异常/诊断及未发布输出预期，不虚构视觉页面。

复用现有 env-gated 产物写出或包消费 harness；PR A/B/C 在验收表写明能生成这些文件的实际命令和环境变量后才执行签收，不另建发布系统。PDF 页面图可用 `pdftoppm -png -r 144 <absolute-output.pdf> <absolute-case-dir>/page`（记录实际文件名；工具不可用则记录替代渲染器版本）。PNG 逐页查看；macOS 用 `open -a Preview <absolute-output.pdf>` 打开本次最终 PDF，逐页核对正文与预期。记录 OFD→PDF→Preview 链路、关闭旧窗口/重新打开方式、页码和截图，结束关闭对应窗口。图片工具复查不代替 Preview，缺少 Preview 时保持未完成。

## 后续功能候选：图片背景与透明 PNG

优先于换 PDF 后端，但不作为本期必须交付项。先做小型 spike：确认空白、半透明图元、重叠区域经当前 PDF/PDFium 链路能恢复正确 RGBA；颜色预乘、背景合成、重采样均有独立预期。如果底层已丢失 alpha，报告边界后重新评估，不通过“白色变透明”伪造功能。

可行后再写独立计划：构造时复制选项、默认白底不变；PNG 显式 RGBA；JPEG 必须合成到明确的不透明背景；不支持组合显式拒绝；调整 buffer 预算；添加透明棋盘/四色象限、编码后 alpha、取消/超限和现有默认快照回归。建议先 SDK 后按调用需求接 CLI，旋转/BMP/批量命名不捆绑。工作量待 spike 后估，不在本计划承诺版本号。

PPM 文件编码器本期不引入，也不计为上游新增能力遗漏：固定上游 [CanvasImageExporter.java](https://github.com/ofdrw/ofdrw/blob/c633a61fb37d300392a51f67b101f46738866c83/ofdrw-converter/src/main/java/org/ofdrw/converter/export/CanvasImageExporter.java) 的 `setPPM` 表示每毫米像素；`normalizeFormat` 只接受 PNG/JPG/BMP（JPEG 归一为 JPG）。评审 R6 的术语澄清采纳，但“上游新增 PPM 编码器”的前提不成立。未来有真实 Portable Pixmap 消费需求才独立评估，不与透明 PNG 捆绑。

## 文件范围与边界

- 首先修改：`tests/Ofdrw.Net.Packaging.Tests/`、`tests/Ofdrw.Net.Converter.Pdf.Tests/` 中相应测试及最小共享 fixture；`tests/Ofdrw.Net.Cli.Tests/` 文件入口测试；必要的 Core/Layout 测试（修改时运行完整项目）；`docs/validation/upstream-2026-11.md`、`docs/validation/upstream-2026-11-tests.json`、本计划和索引。
- 只有复现证明后才修改：对应 Packaging、Reader、Core 或 PDF/SVG 渲染实现；将具体路径与理由加入验收表，不顺手重构整个管线。
- 必要时增加 `scripts/validation/upstream-2026-11/` 的确定性生成/检查脚本与包消费样例。优先既有 harness，不建立第二套发布体系。
- 不修改默认严格策略、默认布局/PDF 导入策略、真竖排范围、PDF 后端、根 LICENSE、旧版验收签收与历史 manifest。
- 不复制来源/许可不明确的上游字体或样本；自生成样例附命令、工具版本、大小、SHA-256 和预期。二进制遵守 LFS；日常结果留 ignored `artifacts/upstream-2026-11/`，不提交冻结源码、全量日志或重复字体。
- 使用 `codex/upstream-2026-11-*` 分支；按归档、绘制、验收拆 PR。每次修复重新生成受影响产物并验收，避免先验收旧产物再改实现。

## 验证命令

在仓库根目录执行；这些是实施命令，本次编写未执行。SDK 由 `global.json` 的 .NET 10.0.100/latestFeature 约束。先确认所需 LFS fixtures 已 materialize，按需获取，不拉历史证据。

```sh
export DOTNET_CLI_HOME="$(mktemp -d "${TMPDIR:-/tmp}/ofdrw-120-dotnet.XXXXXX")"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

# 每个 PR 使用全新结果根目录，不得读取旧 TRX；以下完整项目命令按 PR 门禁表选取。
# 保守执行全部五个项目也可；若已有同一 head 的全套结果且覆盖这些项目，可复用该次完整项目记录。
mkdir -p artifacts/upstream-2026-11
export OFDRW_120_RESULTS="$(mktemp -d "$PWD/artifacts/upstream-2026-11/run.XXXXXX")"
for project in Packaging Converter.Pdf Cli Core Layout; do
  dotnet test "tests/Ofdrw.Net.$project.Tests/Ofdrw.Net.$project.Tests.csproj" -c Release --logger 'trx;LogFileName=full.trx' --results-directory "$OFDRW_120_RESULTS/$project/full" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false || exit 1
done

# 新增用例另跑：仅列出本 PR 确实新增测试的项目，清单须在运行前固定。
# PR A 至少 Packaging/Cli；B、C 至少 Converter.Pdf；若新增 Core/Layout 用例必须列入。
# 以下以 PR A 为例，不能原样作为 B/C 门禁。
for project in Packaging Cli; do
  dotnet test "tests/Ofdrw.Net.$project.Tests/Ofdrw.Net.$project.Tests.csproj" -c Release --filter 'FullyQualifiedName~Upstream202611' --logger 'trx;LogFileName=new.trx' --results-directory "$OFDRW_120_RESULTS/$project/new" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false || exit 1
done
# 下列检查器是 PR A 必须先实现的门禁脚本，未实现/未校验前不能验收任何 PR。
python3 scripts/validation/upstream-2026-11/check-test-results.py --manifest docs/validation/upstream-2026-11-tests.json --results-dir "$OFDRW_120_RESULTS" --pr A

# 集成门禁：环境准备遵循 .github/workflows/ci.yml。
dotnet test Ofdrw.Net.sln -c Release --logger trx --results-directory "$OFDRW_120_RESULTS/integration" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/run-trusted-font-tests.py --no-build --results-directory "$OFDRW_120_RESULTS/trusted-fonts"

# 从本次唯一目录生成本地候选版本；不覆盖旧候选；先生成、冻结 SHA-256 清单，再消费同批包。
export OFDRW_120_VERSION="0.3.1-issue120.${OFDRW_120_RESULTS##*/}"
./scripts/run-converter-package-e2e.sh "$OFDRW_120_VERSION" --pack-only --packages-dir "$OFDRW_120_RESULTS/candidate-packages"
./scripts/run-converter-package-e2e.sh "$OFDRW_120_VERSION" --consume-only --packages-dir "$OFDRW_120_RESULTS/candidate-packages" --output-dir "$OFDRW_120_RESULTS/package-output"
python3 scripts/check-documentation.py

# 暂存后、提交前必须运行；fixture 涉及的提交使用完整检查。
python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests
```

PR A 先实现上述检查器并用最小合成 TRX 验证：正常通过、空结果/零测试、失败、未执行、缺失预期测试、错误项目、缺失文件都必须得出正确结论。脚本参考 `scripts/run-trusted-font-tests.py:45`，但额外核对测试身份。检查器必须接受 `--pr A|B|C`；上面示例为 A，B/C 执行时显式改选，缺失/未知 PR 或空清单一律失败，不从结果目录猜测 PR。JSON 清单按 PR 记录需要的完整项目、带新测试的项目、每个测试的项目名+完整 class/method+Theory 数据行 display name，以及该用例对应的 M 编号；不得从执行结果反推“预期清单”。脚本要求：

- 每个清单指定的 full/new TRX 都存在，使用本次新建目录；完整项目至少匹配该 PR 门禁表，不能靠删除项目缩小范围。
- 解析 ResultSummary/Counters 的 total/passed/failed/notExecuted；新增套件 total>0、passed=total、failed=notExecuted=0；完整项目 total>0、failed=0，跳过必须逐项在验收表解释并经审阅，必需用例不准跳过。
- 将 UnitTestResult 的 testId 关联到 TestDefinitions 的 TestMethod className/name，并核对 Theory 的数据行身份；每个预期新增用例都必须出现且 Passed。不能用 TRX 随机生成的 GUID 作跨运行固定 ID。
- 校验摘要和实际结果条数一致；有缺失/非 Passed 的必需结果即退出非零。检查器输出匹配/缺失清单与计数供复核。
- Shell 测试命令失败本身就是门禁失败，不能靠后续检查器通过覆盖；保存每条命令和退出码。所有命令均是未来实施步骤，本次仅修订计划，没有声称此检查器已存在。

命令预期退出 0，测试不得有未解释失败/跳过；可信字体和 native 依赖缺失属于环境未完成，不能视为已通过。不能为了通过测试清空全局 NuGet 缓存。需要 SVG/HTML、可选 vector 等包验证时补其现有 harness，基础包脚本不能代表所有扩展包。

修改站点正文后，在 `docs/site/` 执行 `npm ci`、`npm run build`、`npm run typecheck`、`npm run check`，检查生成页面和实际支持的导航/搜索范围。只有后续另行执行发布时才进入发布门禁：同步 release note/GitHub Release、首页/安装/能力说明，确认对应 main 提交的 Pages 构建部署及线上内容；本计划不触发发布。

## 完成条件与异常处置

- [ ] M01/M02/M04/M05/M06/M10 每个实际适用场景有当前通过、已复现并修复或符合契约的拒绝结论；不能用“暂时不做”冒充回归通过。
- [ ] 所有新增用例绑定未修复实现 SHA、测试树身份和 fixture/hash；未改实现的条目记录当前通过或契约内拒绝，不要求先失败。只有修复实现的条目要求相同测试/fixture 修复前失败、修复后通过，并绑定两次源码及包身份。
- [ ] M03/M07–M14 有复用/接受有界方案/延期/不引入的理由；延期不计为已支持。
- [ ] 负例、取消、预算、原文和资源身份保持；共享路径改动通过完整回归与包消费。
- [ ] 功能与视觉证据分别齐备；产物、页面、环境、字体、源码及包哈希可追溯；CI 对应最新 head。
- [ ] 文档和存储门禁通过，索引更新状态；不重签历史验收。

若发现必须扩大默认策略、重写 accounting/字体/后端才能继续，先保留复现和契约冲突，拆成独立设计任务；不要将扩大支持伪装成小修复。发现工具/环境缺失则标明具体未完成项。维护者重点审查独立 oracle、受控 ZIP 范围、XML 与 producer 差异，以及新断言是否只是重复既有测试。

## 本次评审处置记录（2026-11-09）

来源：`plans/019-upstream-2026-11-regression-and-triage-review.md`（用户提供的评审原文保留，不改写其历史意见）。

| 意见 | 处置 |
|---|---|
| R1 | 采纳：区分新测试树与原实现；只有修复项要求前失败/后通过 |
| R2 | 采纳并核实：公开流、内部受控、真实 CLI 路径入口分别列出，不增加公共 API |
| R3 | 采纳：完整相关项目门禁 + 独立新增测试执行 + TRX/预期身份核对；检查器明确列为实施前置任务 |
| R4 | 采纳：每个 PR 独立验收后合并，C 不补签 A/B |
| R5 | 部分采纳：移除固定型号/工具调用参数，增加页面清单；保留对执行会话用户 GUI 偏好的遵循，不由评审原文撤销用户指令 |
| R6 | 术语修正采纳；上游新增 PPM 编码器的前提经固定源码核实不成立，本期不引入该格式 |
| R7 | 采纳：漂移比较覆盖已跟踪工作树，未跟踪内容另查；禁止覆盖已有工作 |

修订后独立只读复核确认 R1–R7 主要问题已解决；额外发现的固定集成目录和 PR 清单选择歧义也已修正为逐次唯一目录及显式 `--pr`。

本次只修订计划/索引；功能、自动渲染、PNG、Preview 均未执行。未更新远端 issue 或发布。
