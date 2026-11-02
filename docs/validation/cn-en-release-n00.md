# N00：中英文发布范围与最小验收矩阵（2026-11-01）

本报告执行 [2026-11-01 计划](../next-steps-2026-11-01.md)的 N00，只冻结当前源码、可判定场景、输入身份和交接。**本次未构建、测试、转换、查看页面或发布；所有新候选的功能、PNG、Preview、Copy/Find 结论仍待执行。** 静态源码表明已有常用 Flow、DOCX Native、表格、图片、页眉页脚和页码入口，不能因完整 R22 未完成就判定这些能力缺失。

## 1. 本次身份与范围

| 身份 | 本次只读核对结果 |
| --- | --- |
| 执行源码基线 | 远端 `git ls-remote origin refs/heads/main` 与本地 `origin/main` 均为 `beb02832fff01e61538d90039f3e1224ed03d702` |
| 执行源码树 | `d5ef5b9e0a4b9e28cfe1c6caed0b7418344c6dc3` |
| PR48 必需祖先 | `1c604d28c67f3c08f4dd16d6163a9bf7d444e6c1` 是该 main 的祖先，`merge-base --is-ancestor` 退出 0 |
| N00 工作树 | `codex/cn-en-release-n00`；编辑前 HEAD `ba7682cfceafd6e28d526941e0eb039addfe400f`，树 `7982d41578364109366c86c43f3c1a3d36436887`，无未提交差异 |
| 分支差异 | N00 HEAD 带计划/复盘文档；main 另有 CI 去重提交。两者在 `src/`、`tests/`、`e2e/`、`scripts/` 和 tag workflow 上无差异；CI 约定按 `git show origin/main:docs/ci-performance.md` 读取 |
| 历史验证身份 | `b02162bef61abc17761349d891237fc2518d3989` / 树 `4b88afedc907151a6443d013d1239ca9eea15a04`，仅作先验证据，不转移为新 main/新包通过 |
| 版本 | 仓库默认 `0.1.0-preview.9`；CI 消费戳为 `0.1.0-preview.8`；远端已有 `v0.1.0-preview.9`。这不是新候选版本选择，N12 必须核对未占用版本和公开源后另冻 |

冻结范围是中文、英文或两者混排的常规文档：完整原文、比例英文、局部粗斜体/颜色、换行/对齐、常用表格、图片、页眉页脚/页码、分页、OFD 读回与 PDF/PNG 导出、普通复制和搜索、精确包安装。现有库目标为 `netstandard2.0/netstandard2.1`，CLI 与现有 E2E 为 `net10.0`；运行范围沿用现有 Linux/Windows/macOS CI 和既有 RID/消费组合，不从目标框架声明推断任意宿主可用。

PR78 合并不是前置；N01 可独立限时处置。完整 R22、新通用文字引擎、RTL/复杂脚本、新字体格式、新 RID/musl/AOT/trimming 和完整归档符合性不扩入；已有回归与明确拒绝边界保留。N09 维持取消。N04/N05 只有自然新失败及新证据才触发有界工作。

## 2. 确定性输入与字体身份

本次只材料化下列 DOCX 和共享字体 ZIP，没有拉取历史证据归档或全量 LFS。核对 ZIP 与每个成员的大小/SHA-256，许可证与清单一致；没有改动 fixture。

| ID | 复用输入 | 字节 / SHA-256 |
| --- | --- | --- |
| F01 | `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` | 3,347 / `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6` |
| F02 | `tests/fixtures/trusted-fonts/sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip` | 11,821,583 / `40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59` |
| F02-CJK | ZIP `fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf` | 21,659,544 / `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa` |
| F02-Latin | ZIP `fonts/NotoSans-Regular.ttf` | 569,208 / `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5` |
| F02-CJK-license | ZIP `licenses/Ofdrw-CI-Noto-OFL.txt` | 4,301 / `6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2` |
| F02-Latin-license | ZIP `licenses/NotoSans-LICENSE` | 4,377 / `0dab92d0544f7b233403f14b84a663bdbfa746982eda629e7f4f9ffe1b036feb` |

来源、真实名称表及许可见 [trusted manifest](../../tests/fixtures/trusted-fonts/manifest.json)、[bundle manifest](../../tests/fixtures/trusted-fonts/bundle.json)和 [THIRD-PARTY-NOTICES](../../THIRD-PARTY-NOTICES.md)。CJK 是固定 `wght=400` 静态实例，真实 family `Noto Sans CJK SC`、subfamily `Regular`、PostScript `NotoSansCJKSC-Regular`；Latin 为 `Noto Sans` / `Regular` / `NotoSans-Regular`。不是上游恰好同字节的静态发行字体，也不是生产输入的替代物。

固定 CJK 配方为 [install-ci-fonts.py](../../scripts/install-ci-fonts.py)：上游 noto-cjk revision `f8d157532fbfaeda587e826d4cd5b21a49186f7c`，源 VF SHA-256 `990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f`，`fonttools==4.59.2`。本次核验既有 ZIP 字节，没有重新生成字体。字体 fixture 不进入生产 NuGet；子集、嵌入后的实际 face/内容身份需由本次输出再核对，不能要求子集哈希等于原字体。

**真实 CJK Bold 身份尚未冻结，属于 N13 的有限验收资料缺口。** 共享 ZIP 只有 Regular；[FontIdentityRegressionTests](../../tests/Ofdrw.Net.Converter.Docx.Tests/FontIdentityRegressionTests.cs) 的 physical-bold 测试使用修改 flag 的合成 Latin fixture，不能当成中文粗体部署/可读性证据。N13 先复用原源字体的静态粗体配方或有合法来源的既有物理面，记录实际哈希、face index、名称、weight/italic、许可证和配置；若采用同源 `wght=700`，须真实生成核验，N00 不预填哈希或宣称它已通过。

固定两个独立配置，不能互换结果：

- **原 Regular 配置**：F01 不修改，通过 `FontDirectories` 提供 F02-CJK，`UseInstalledMicrosoftOfficeFonts=false`；显式 Native 仅设置 `OfdMode=Native`，default 不设置 `OfdMode`（当前默认 Native）。保持 `NativeEmbeddedBoldBehavior=TextStroke` 原回归，记录 `DOCX_EMBEDDED_BOLD_SIMULATED` 和“量”字结果；不改原文、阈值或 golden。
- **物理 Bold 部署配置（待 N13 固定）**：同一 F01，隔离目录包含已核验真实 Regular/Bold；优先验证 `NativeEmbeddedBoldBehavior=Throw` 下粗体实际选择，检查字形/原文/常用样式。此结果只说明明确配置的部署路径，不抹去原 Regular 配置失败。

原始 `flow-public` 只声明默认 `SimSun` 名称字体，使用固定比例字宽表，不自动嵌入。其分页可确定，跨机器像素不能单凭名字确定。保留原样例；N13/Terra 记录实际导出所用 face 并增加明确字体配置的独立部署记录，不能声称已为原 Flow 增加新布局引擎或自动嵌入。

## 3. 固定核心矩阵与逐页预期

以下均为**待本轮执行的预期**。页数来自现有代码断言或明确标记的历史样例，不能替代本轮真实页面。输入为生成器的项目须先保存本次 DOCX/OFD 和哈希，生成器身份绑定本次 SHA，不提交重复大载荷。

| ID / 消费者与来源 | 固定输入、模式与页数 | 逐页预期及功能裁判 |
| --- | --- | --- |
| C01 SDK Flow | [Layout E2E Program](../../e2e/Ofdrw.Net.Layout.E2E/Program.cs) 原 `flow-public`；FlowDocument→OFD→PDF；历史确定样例 3 页 | 页 1 居中 6 mm 粗体双语标题，局部 Bold/Italic/红色 `192,0,0`，后文恢复普通样式；55 段双语、`information/minimum` 自然折行。历史边界第 20/21、42/43 段续页；本次逐页核对。非空白原文由生成器输入与 OFD 逐字相等（历史 5,865 字符），无丢重、裁切或异常尾页。 |
| C02 DOCX 原 Regular | F01→显式 native OFD→PDF；F01→default OFD→PDF，各 2 页，包消费者 `ValidateDocxSampleAsync` | 页 1 中文标题、居中灰色斜体英文、第一节正文、3×3 表（表头“项目/数量/状态”，Alpha/128/正常、Beta/256/待复核）；底色 `D9EAF7`、外框 `4472C4`、内框 `A5A5A5`。页 2 第二节、普通正文后 `C00000` 红色粗体、右对齐 10 pt 日期 `测试日期：2026-01-01`。正文 189 非空白字符，原文完整，两种模式页序/外观一致，无重复文本层。“量”字须明确结论。 |
| C03 DOCX 物理 Bold | 同 F01 与 C02 两模式，各 2 页；待 N13 记录唯一字体配置 | 与 C02 同页内容/表格/分页；真实 Bold 资源不能错用 Regular 或追加重复补绘。检查粗斜体/颜色没有串染。保留 C02 结果并单独判定 C03。 |
| C04 对齐与 Latin | 同 Program 的 `alignment-public/native/default`，各 1 页 | 参考 Alpha 与折行 Alpha 的可见位置分别居中/右对齐一致；NBSP/U+202F/U+2007 保持空白推进量、不画方框；café、mínimo 预组/分解形式整体折行且无重音重影。OFD 原始序列保留，PDF 允许规范等价；既有坐标断言不放宽。 |
| C05 SDK 固定图片与表格 | [FlowFixedSamples](../../e2e/Ofdrw.Net.Layout.E2E/FlowFixedSamples.cs) `--flow-fixed-only`；`flow-fixed` 固定 2 页 | 页 1 双语标题、红色加字距/下划线和 italic、蓝/黄底表格及蓝黄双色图片；图片只在页 1。页 2 绿色 Canvas、蓝斜线、`Canvas 画布`、保留区与画布后文字；Canvas 只在页 2。既有回调、边界、唯一图片、named area 断言保持。额外 `letter-spacing-wrap` 是既有负控，单独记录。 |
| C06 DOCX 页眉页脚与页码 | [DocxRevisionTests](../../tests/Ofdrw.Net.Converter.Docx.Tests/DocxRevisionTests.cs) `CreateMultiSectionDocx` / `Native_ShouldRepeatHeadersFootersAndHonorSectionNumbering`；复用生成器经 native/default，各 5 页 | 页 1–5 正文依次 Body one/two/three/four/five；页眉 First/Even/Odd/Even/Odd Header；居中页脚 Page 1/5、2/5、3/5、10/5、11/5。页眉在正文上方，页脚 y>260 mm 且在页面内。现有 default Native 单测已覆盖结构；本轮须补双模式持久产物/PNG/Preview，不把它立为缺失实现。 |
| C07 DOCX 内嵌图片 | 复用 [DocxConversionTests](../../tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs) `Native_ShouldKeepInlineJpegContentDimensionsAndMediaType` 的 40×20 mm inline JPEG 生成器；native/default 各 1 页 | 2026-11-01 执行校正：N13 共享已有 Native JPEG 生成器，保存左红右蓝、宽高比 2:1 的清晰图样，核验原图字节/尺寸与 OFD 资源，再验实际 PDF/PNG/Preview。BuiltIn 直接 PDF 的透明 2×3 PNG 测试不是此项证据，不扩大为浮动图片。 |
| C08 分页选择 | C06 同源，选择 `[2,0,2,-1,99]`，native/default 各 3 页 | `SourcePages=[2,0,2]`，正文顺序 three/one/three，重复页完整相同；仅无效 `[-1,999]` 按既有兼容规则回退全部 5 页。覆盖 `Native_ShouldPreservePageOrderDuplicatesAndInvalidSelectionCompatibility` 现行断言。 |

C06/C07 的功能基础已经存在；待补的是核心链路证据与可复用产物。常用表格样式由 C02/C05 覆盖；[PublicTableSamples](../../e2e/Ofdrw.Net.Layout.E2E/PublicTableSamples.cs) 的合并表头、11 行、填色/对齐、表后正文，以及 [ComplexTableSamples](../../e2e/Ofdrw.Net.Layout.E2E/ComplexTableSamples.cs) 的重复表头、跨页拆行、28 个唯一段号和唯一图片，继续按既有 E2E/全套回归执行，不重新设计表格，也不把“至少三页”的现行断言编造为固定页数。

终端换行、样式空行、正文 PAGE 延期计算等已有 Program/单测回归全保留；固定 4 页的 `terminal-followed-*` 中页 2/3 是预期空页，不能把所有空页都当缺陷。F01 不含图片/页眉页脚/页码，C02 的两页通过绝不能代替 C06/C07。

## 4. 有限缺口与阻断判据

N00/N13 共用最多 8 小时主动盘点，以下是可执行有限表；N00 未发现需要重写产品引擎的证据。修复只对应已证实核心问题；到限仍有核心失败则候选不可发布，不重置预算接力。

| 项 / 优先级 | 证据、实际消费者与预期 | 最小后续与上限 / 当前分类 |
| --- | --- | --- |
| G01 量字 / P0 判定 | F01 表头“数量”；原 main `b02162be` 最终交接 `quantityCounter=FAIL`、`overallVisualPass=false`。C02 是必验页面 | N13 先冻结真实物理 Bold 部署配置；Terra 对 C02/C03 独立看页，不追加光学算法。初盘纳入 8 小时共享预算；原 Regular 失败保留。若现行必验仍 FAIL，N12 不得 accepted；新配置通过不自动解决旧配置失败的发布取舍。 |
| G02 字体部署资料 / P0 | F02 仅 Regular；合成 Latin physical-bold 单测不是中文实际字重证据 | N13 固定真实中文 Bold、Latin 比例 face 与配置/许可/哈希，复用同源资产；仅补样例/配置/文档，有实际失败才就地修复。纳入共享初盘预算。 |
| G03 图片/页眉页脚的双模式链路证据 / P0 | C06 已有 Native 结构断言；C07 已有 Native JPEG 资源断言及 BuiltIn 直接 PDF 测试；都不能替代本次两模式页面证据 | N13 复用生成器、有限 E2E 持久化与断言；Terra 看 C06/C07 全部页面。属于验收覆盖缺口，不是已证明的产品缺失；不得新建通用 DOCX 功能项目。 |
| G04 极小正 Size / N03 P1 | [OfdNumericFormat](../../src/Ofdrw.Net.Core/Models/OfdNumericFormat.cs) 使用 `0.###`；计划与封存交接保留正数被写成零的问题 | N03 独立最小修复/回归：合法正 Size 或输出提交前明确拒绝；常规字号与多 UPEM 不回归。不全局改变坐标精度。N00 没有新动态复现。 |
| G05 页末 Find / N02 P1 待判别 | 历史问题未关闭，PR90 七页复查明确没有重做完整 Copy/Find | Terra 在 C01/C02/C03 的实际 PDF 分别复制完整普通中英文，查询首/中/末文本；保留 UI 原始命中/选择/剪贴板证据。N02 上限 4 小时、2 个单因素假设；若核心普通中文/英文尾字确实找不到，升级发布阻断，不能改用 AX/PDFKit 结果冒充 UI 成功。 |
| G06 发布记录/说明漂移 / N12 P0 | `release-preview-acceptance.json` 是旧 preview.9 accepted；工程门文档仍写 not-reviewed，依赖数 38/22 与当前 baseline 40 条不同；文字教程仍写“当前不做子集化” | N12 按准确候选实际清单同步说明和本次记录；不能先把 JSON 改 accepted 或重生成许可 baseline 过门。纯文档校正不等于候选通过。 |
| N04/N05 / 被动观察 | Skia 历史真实失败、Windows 分配历史异常仍未归因 | 保留现有诊断、预算和全套测试；无新自然失败不主动压力复现。新失败读取分类，命中核心/现行门禁即阻断，不能默认噪声。 |

阻断条件固定为：核心样例原文丢重/乱码/缺字、普通正文不可读、严重样式串染、裁切重叠/错误分页、无效输出、普通复制/搜索不满足承诺、候选包不可消费，或现行强制测试/CI/发布门失败。历史光学优化不自动立为通用研究前置，但**不能用“延期 PR78”豁免 C02 或其它现行必验页面的真实 FAIL**。字体来源/嵌入许可未知也不能给出可部署结论。

## 5. N12 不可省略的现行门禁

按准确最终 main 的 [CI 约定](../ci-performance.md)、[发布门](../release-engineering-gates.md)和 [tag workflow](../../.github/workflows/publish-nuget.yml)执行；本轮不跑这些门，不创建 tag/Release/公开包。

1. **源码与测试**：锁定最终 main SHA/tree、差异和 PR48 祖先；完整 Release 回归、三平台适用完整测试、mandatory trusted-font、坏包/资源预算、API 文档，记录真实 pass/fail/skip。Linux 保留 no-Office 过滤；Windows/macOS 去重不删完整覆盖。共享布局/字体/转换有改动时，完整回归及本地精确包 E2E 必须重做。
2. **包身份**：当前 validator 的 12 包为 Core、Packaging、Layout、Text.HarfBuzz、Reader、Converter.Abstractions、Converter.Docx、Converter.Pdf、Converter.Svg、Signatures、Converter、Cli（均 `Ofdrw.Net.` 前缀）。从一次不可变 feed/manifest 隔离消费；source mapping、选中 TFM、包/缓存/输出及实际加载 DLL/native 各自核对，消费前后哈希不变。未来包数从实际清单取，不能把旧 11/12 数字当证据。
3. **Q01–Q04 与 CI**：还原后许可/直接与传递依赖清单、同批包声明、坏包负控、公开 API 新增文档门、四个单页视觉 golden 的全页/标题/正文 RMSE≤0.035 原门不变；CI complete 与独立 SES 必过。Cursor 豁免不变，不触发/等待重试，也不声称 bot clean；有效新意见需处置。
4. **本次页面**：C01–C08 输出清单写入输入/字体/模式/工具/页数/OFD/PDF/逐页 PNG 哈希和大小，canonical 144 DPI 图与诊断变体分路径。功能断言、自动渲染、PNG 人工目视、实际 macOS Preview、Copy/Find 分别给结果。适用 SVG 的已有门继续跑，实际 SVG 离线外观另验；不由 PDF 推断 SVG/浏览器语义。
5. **候选 Q05**：两模式 F01→OFD→PDF→Preview 均逐页无遗留视觉缺陷，才填写新候选 `release-preview-acceptance.json`，用 `verify-preview-record.py --print-fingerprint` 取源码指纹并校验版本。该指纹涵盖跟踪文件；报告/文档整合后变更也须重新核验绑定，不能沿用本次 N00 HEAD。脚本检查字段不证明查看行为。
6. **存储与文档**：日常输出 ignored `artifacts/`；必要长期证据按 [存储约定](../repository-storage.md)以 LFS 归档并带清单，不复制源码/缓存/字体。准确候选的支持范围/部署方法/限制同步 README、教程、能力契约/对照/路线图，保留 R02/R20/R22 有限边界。暂存后提交前存储检查必跑，修改 fixture 时加 require-materialized e2e/tests。

实际 NuGet 推包、tag 工作流成功和公开源安装回读是后续独立结果；“可发布”不得写成“已发布”。

## 6. 精确交接与本次核对结果

- **N13**：从 §1 执行 main 基线复核/整合最新差异，只做 G02/G03 的有限 harness/配置/文档与实证缺陷修复。先固定 C03 的实际字体哈希和 C07 的清晰图样（N13 已共享原 Native JPEG 生成器）；复用 C06 五页及 C08 选择规则。C02 原始 fixture/Regular 失败样例保持独立。生成器输入/输出放本次 artifacts，有长期必要才 LFS，不复制整套历史证据。N00/N13 初盘总上限仍为 8 小时。
- **Terra High 本地验证**：由协调者明确交付本次 SHA、唯一 artifacts 根、包 manifest、C01–C08 文件/字体哈希和预期后领取 GUI；不得提前沿用旧 PDF。逐页 PNG + Preview：C01 3 页、C02/C03 各双模式 2+2 页、C04 三路径各 1 页、C05 2 页、C06 双模式各 5 页、C07 双模式各 1 页、C08 双模式各 3 页；这是本轮冻结最小清单 34 页，实际生成不符预期时先查明，不跳页。普通 Copy/Find 与页末判别独立记录，验后关闭任务窗口并释放 GUI。原 Regular 量字、物理 Bold 量字分别结论。
- **N12**：可提前做许可/依赖/版本/打包/文档检查，最后在修复合并的准确 main 上完成 §5。若 G01 或其它必验页面仍失败，输出不可发布及明确阻断，不填 accepted；当前授权不包含自动公开发布。

历史视觉证据入口（原桌面本地路径，不随仓库分发）：`/Users/wanghongyi/.codex/worktrees/r22-span-context-query/ofdrw.net/artifacts/r22-shared-position-consumer/main-b02162be/`。本次重新读取并计算：

| 文件 | SHA-256 / 可用结论边界 |
| --- | --- |
| `actual-main-final-functional-and-preview-handoff.json` | `497befad5ddba9dde635f73156c0c69a293888414a0dfd5bcd49bd556f82ef41`；历史源码与功能记录，`quantityCounter=FAIL`、`overallVisualPass=false`、Copy/Find 未重做 |
| `seven-page-preview-input-manifest.json` | `b521015d7f7bfb68ea4e6a9efaa26035fd0421b93a11817fe544c27bac5f8539`；列旧 native/default 各 2 页及 3 张 typed PDF，不是新候选输入 |
| `primary-preview-reference-audit.json` | `815db588b9942fde69773a89e9b93d83cb58c90dfb28aed24bdfd830c6266c43`；明确旧七页检查完成且量字仍 FAIL，不转述为视觉通过 |

本次完成：远端基线/祖先/差异、准确 main 门禁源码、确定性 DOCX 内容和字体/许可证字节核对、现有 API/测试/生成器盘点。功能测试未执行；自动渲染未执行；PNG 视觉复查未执行；macOS Preview/Copy/Find 未执行（N00 文档盘点范围，无 GUI）。未取得 GUI 所有权、未改产品代码、未创建远端 PR/tag/Release/公开包。
