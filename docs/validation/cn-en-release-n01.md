# 中英文发布 N01：PR78 一次性存量处置

日期：2026-11-01。执行范围为[当前计划 N01](../next-steps-2026-11-01.md#n01pr78-一次性存量处置替代量字优化任务)规定的最多 2 小时主动只读审查。本次在单个短会话内完成，未用满额度，不续接研究。只交付本文；没有构建、运行产品测试、生成候选/字体矩阵、操作 GUI、修改产品或操作远端 PR。

**结论：PR78 本轮不整体合并，建议关闭并保留原分支和证据。保留两个独立正确性修复候选 C1/C2，均只是一份有限提案，N01 不实施、不自动开工、不成为发布前置。** Skia 首次失败诊断已在 main 独立覆盖；其余光学增强、增强自身的兼容性修正及相关配套本轮不采用。N01 到此完成，发布继续按 N00/N13/N03/N02/N12 推进。

原字体“量”字可读性 **FAIL**、页末 Find 失败及两个未解决 review 均保留；本处置不将它们改写为通过。当前发布必验页面如果仍有缺陷，N12 仍须按自身门禁判定，不能凭 N01 的非阻断属性豁免。

## 1. 核对身份与方法

| 对象 | 本次核对结果 |
| --- | --- |
| 发布 main | 本地 `origin/main` 与 `git ls-remote` 均为 `beb02832fff01e61538d90039f3e1224ed03d702` |
| 旧规划/验收基线 | `b02162bef61abc17761349d891237fc2518d3989`；至上述 main 只有 `.github/workflows/ci.yml`、`docs/ci-performance.md` 变化，产品源码未变；旧验收仍只属于旧提交，并非新候选验收 |
| 本文工作树基线 | `76b612677a65ed2c589915b351383600025e433c`，分支 `codex/cn-en-release-n01`；开工时干净 |
| PR78 | [PR #78](https://github.com/whynpc9/ofdrw.net/pull/78)，API 返回 open、未合并、`mergeable=false`、19 commits；远端分支 head 同下 |
| 审查区间 | base/与 main 的 merge-base `a05badc212a0b653a0a10afc9648e24c50c3f11a` → head `2c4b48fbcd93618c131744c11be66cde83c84615` |
| 变更体量 | 41 文件，3,917 行新增、82 行删除；19 个提交逐项处置见 §2，不等同 19 项独立缺陷 |

方法：读取每提交文件清单/相关差异、最终 planner/row builder/PDF geometry/union 实现、现有固定反例和测试、[复盘](pr78-and-stability-retrospective-2026-11-01.md)、原环境封存记录；用当前 main 对应实现和 Git blob 比较判定覆盖。没有构造新反例，没有重放旧测试或旧页面。

对候选涉及的 main 文件另作精确比较：`OfdPackageWriter.cs` 在 PR78 base 与 main 的 blob 均为 `ac1816b0732fa82760a268856573564c4970ec6d`；`OfdToSvgConverter.cs` 均为 `dfd3e4d856ea788616fb1a287bdaf98df0ee7273`。因此下面的静态反例路径仍存在，不能声称已被后续 PR 修复；这也不替代候选实施后的动态验证。

## 2. 19 个提交的处置表

“本轮不采用”包括光学增强、依赖增强的修正或缺少独立契约证明；不是删除原分支、失败证据或主线已有覆盖。每行范围以准确提交为准，拆出候选时不得直接 cherry-pick 整个提交。

| # | 提交 | 内容与既有依据 | 处置 |
| --- | --- | --- | --- |
| 1 | `a54aa2fe75f4b0ec592e4a40fbd58cdaccaae8a4` | `FilledStrokeCoverageTests` 标题接缝/原字形路径对照。后续报告纠正了原 opaque-interior mask 的边界归类，不能将初版全部计数直接当最终正确性证明 | 本轮不采用：保留复现和更正记录，不单独恢复旧阈值或光学修复 |
| 2 | `e84af7d6e437cea2543a55b4b895c8d5ec4e7914` | `DrawTextRunsMapped` 在同色不透明描边时重复 fill；紧随其后的固定小表头反例失败 | 本轮不采用：该最初修复已知有回归，不能作为保守补丁提取 |
| 3 | `80432418743d6b975c942e333fa531c32d23a80d` | 默认 144 DPI 的“数量”两字腔及非空上横断言；e84 产品上的两窗口均值 22.3/18.4，原门槛 100 | 本轮不采用独立补丁：保留历史反例，不改门槛，不把它升级为新算法任务 |
| 4 | `1b75337fe389c3237a46a4c5143ccafed954a4ef` | 新 `PdfGlyphStrokeGeometry`、几何 fill/stroke 与空 clip 语义文字路线，改变共享 PDF paint | 本轮不采用：与新绘制算法不可分；不能仅提取“保字腔”名称 |
| 5 | `38b464f23dbad5c6808be9ae1c287bc2efb060b7` | PDF 包消费 E2E 额外保留 native/default 的默认公开图片；断言仅为页数/非空 | 本轮不采用独立正确性补丁：是该候选的证据扩充，不证明 main 的独立缺陷；不删除旧证据 |
| 6 | `0f36cc3be28b95bd6cc84d87e5812a4663a9e34e` | 新 native 描边规划器、模型宽度和 writer 传递、PDF 双描边合成；最终演化为 3.5%-em 光学路线 | 本轮不采用：需新规划/几何及其验收，不能作为无依赖的重复描边修复 |
| 7 | `0ef49ebe42c87cb3cbd40eb7932fdfec8e6c8cd6` | 规划预算、SVG 抑制再次 bold、静态轮廓探测，并修正 CSS 字体族选择 | **仅字体族选择保留候选 C2**；其余本轮不采用，理由见 §3 |
| 8 | `acec1025d66ac4500c13c8fc80534f1d7ec02a1f` | 几何路线使用 `sourceBoldStrokeWidth` 而非单个 Latin 字形重新取 2.5%；`量A` 控制保留约 2.04% 规划宽度 | 本轮不采用：修正 PR78 新规划/几何路线自身问题；main 没有该新增分支，不能改名为 main 独立缺陷 |
| 9 | `5cadb1f83e2e77bdefd41e9a91e6271ba653ccaa` | writer 保留导入 paint/CTM；另有 fresh missing-Noto 的 name-only 选脸标记、克隆及 producer 行为 | **仅导入 paint/CTM 保留候选 C1**；missing-Noto 策略本轮不采用，不能随 C1 一并带入 |
| 10 | `677205e15eacd648500dbd6b7bb437d0c8c72069` | `FontReadTrace`、raw-style preflight、测试注入和 CI 留痕 | **主线已覆盖**：`bd86b4db7617ddd11494cff84f72208f2979f289` 独立交付同一产品诊断；`c82b0c34623015f802603a77b294bb760592dbe2` 进一步保留首份 collector 记录。产品诊断文件与 PR78 对应版本无差异；不重复提取，不声称修复 Skia 根因 |
| 11 | `ffc8cf436b40b6a702c0af0eaf86fb5d4a85a185` | 表格同行同字体/尺寸/颜色取最小已规划宽度；原 OFD delta 报告中“项目/状态”由 0.129 降至 0.037 | 本轮不采用：同行粗细差异来自新增按字腔取不同宽度；移除增强后该依赖消失 |
| 12 | `0c63fd3c60dd179152f38fb5b50012d96b7ed29a` | Clipper2 2.0.0 全轮廓拓扑、原 GID/预算、包装许可 | 本轮不采用：新字腔分析及依赖，仍非完整光学证明，不引入 Clipper2 |
| 13 | `f9b5c94939fec1792283eb82e0b6401f595d93dc` | Clipper 许可缺失/改变测试和测试包 fixture | 本轮不采用：专属于未采用依赖的配套，main 无此依赖缺口 |
| 14 | `0324cbc3036cbd0f4106c14dc39b56f71527db7a` | Clipper 许可 LF 属性及 autocrlf 测试 | 本轮不采用：同上，不单独引入无消费者的许可文件/属性 |
| 15 | `10f4ae205b1894f7a891ea723ea6e50d073d4c5e` | `NativeProtectedRowStrokeBuilder` 原曲线 3.5%-em 补绘、同行字腔排除 clip/预算及顺序 | 本轮不采用：明确光学增强，不提取预算/clip 修正来反向引入整条增强路线 |
| 16 | `69da66cc58542bbfff4a8f71158afa5144a53174` | wrapper fixture 避开新增 `Clips/Clip/Area`；verifier 强制 shared fixture 含 protected-stroke clips | 本轮不采用：原失败由第 15 项给 shared fixture 引入 clips 暴露；整补丁的前置在 main 不存在，不把它当产品缺陷。通用 fixture 防串改价值保留在历史记录，不自动开测试改造任务 |
| 17 | `3d8b7ce0bfefbd927b3698b63bc517ab2e08184c` | 已完成光学分析缓存、ForcePPEMToInt 尺度还原及 WenKai 三表头/11 clips 验证 | 本轮不采用：main 无 planner/protected-curve replay，相应缓存和尺度缺陷不成立 |
| 18 | `63ce2d8395b0c86185c82a01143552d6f2d4d48e` | 规划器 `MathF.Round` 改 `(float)Math.Round` 以跨 TFM | 本轮不采用：仅为第 17 项兼容补丁，不是 main 的跨 TFM 缺陷 |
| 19 | `2c4b48fbcd93618c131744c11be66cde83c84615` | 直线字形 fill/stroke 单一并集边界、准入/取消/预算测试 | 本轮不采用：新增几何算法且准确 head 原字体“量”可读性仍 FAIL，不能靠几何等价或 CI 成功批准整体合并 |

## 3. 两个有限正确性候选

### C1：导入的 name-only 文本重写保留原 paint/CTM

- 来源：`5cadb1f83e2e77bdefd41e9a91e6271ba653ccaa` 的 `src/Ofdrw.Net.Packaging/OfdPackageWriter.cs` → `ApplyEmphasis`，仅提取 `SourceXml` 存在时不再注入兼容 paint/CTM 的 guard，以及与之相符的注释。最小依赖只有 main 已有的 `SourceXml`、现有 reader/writer；不需要 `UseNameOnlyBoldSelection`、模型/克隆/producer 修改或任何 planner。
- 现有确定反例：`ImportedNameOnlyPaintIsPreservedThroughRepeatedWrites` 的 name-only `回A`、Size=4、Weight=700、显式 LineWidth=0.023。main 在保留 SourceXml 的分支仍调用 `ApplyEmphasis`，非嵌入字体不受早退保护，直接把线宽写成 `max(0.08, 4×0.045)=0.18`，并强制 Stroke/Fill=true；`no-stroke`、`outline-only` 因而也被覆盖。`italic-no-ctm` 还会获得原来没有的 CTM。此路径不依赖 PR78 新宽度或字腔算法。
- 独立预期：导入 XML 的 paint/CTM 保持；允许现有 ID/Font 重绑定及模型 Weight/Italic 的既有处理。显式细描边、颜色/alpha、仅描边、无描边和省略默认属性不能因为读写一次被改为 faux bold/italic。
- 直接复用的回归：上述测试的 9 个既有 `kind`（thin、different-color、alpha、outline-only、no-stroke、italic-ctm、sim-sun-explicit、omitted-paint、italic-no-ctm），比较两次读写后的 XML；其函数独立，不依赖该文件其他新增光学测试。main 已有 `EmbeddedBoldCompatibilityTests` 中嵌入字体 roundtrip、fresh 通用 name-only 行为应保持。
- 如以后选定实施：局部 writer/reader 回归、受影响完整回归与精确包消费；以相同导入输入及 native/default 控制检查实际 OFD→PDF 页面、PNG/Preview、原文完整和无重复文本。不得直接沿用 PR78 的测试/页面验收。是否影响冻结发布样例由 N13 判定，本处置没有宣布它是新发布阻断。

### C2：SVG CSS 使用声明字体族，避免泄漏资源身份键

- 来源：`0ef49ebe42c87cb3cbd40eb7932fdfec8e6c8cd6` 的 `src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs` → `AddText`，仅提取初始化 family 时优先非空 `resource.FamilyName`、否则保持 `text.FontName` 的选择。已嵌入字体随后仍覆盖为现有的 hash family。最小依赖只有 main 已有资源字段，不需要 `OpenTypeFace.HasStaticTrueTypeOutlines`、tuple 扩充或 bold 抑制条件。
- 现有确定反例：`NameOnlyText_ShouldUseDeclaredFamilyWithoutLeakingStyleResourceKey` 中 `FontName="Fixture Resource|bold|italic"`、`FamilyName="Fixture Family"`。main 将前者输出为 CSS `font-family`；预期是后者，粗斜体仍由 `font-weight`/`font-style` 输出。main native producer 的 `FontKey/DeclaredFontName/GetOrAddFontResource` 确实分开保存带 `|bold|regular|italic` 的身份键与原 FamilyName，所以该反例不依赖 PR78 规划器。
- 直接复用的回归：上述测试的有 FamilyName/无 FamilyName 两例，保留文本 `Ab`、bold、italic，并验证无 FamilyName 时原 fallback；没有新字体载荷、参数或算法。
- 如以后选定实施：SVG/HTML 受影响回归及精确包出口，核对同一现有 name-only 输入在声明字体可用/不可用时的实际 SVG 文本属性与宿主选脸，保留嵌入字体和无 FamilyName 负控，按仓库规则检查页面。SVG/HTML 是本轮保留边界，不因候选而扩成新的发布出口承诺。

以下三个相近名称不保留为候选，防止范围再次扩张：

| 范围 | 本轮不采用的理由 |
| --- | --- |
| SVG 再次 bold 合成 | 旧 same-input 诊断确认 CSS bold+stroke 确实造成额外加重，但该观察不自动证明任意输入的 Weight 应被忽略。PR78 修正另外引入静态轮廓准入和按现有比例判定“足够”的策略；现有单元断言主要验证所选策略。没有在本次额度内取得独立于外观策略的完整输入契约证明，归为证据不足，不为补齐证明追加新字体/参数矩阵；不混入 C2 |
| fresh missing-Noto 不注入 4.5%-em stroke | 这是 name-only 物理选脸/宿主 fallback 策略变化，仍需区分实际宿主能否选到 bold。与 C1 导入 XML 无关，本轮不采用该 producer/model/clone 扩张 |
| 非历史宽度 PDF 双描边 | 确有旧 1.25%-em 控制与约 25.22% opacity 差异的历史记录，但现有修复绑定几何 paint/空 clip 和后续并集算法，未得到符合 N01 的无新光学算法最小补丁；本轮不采用，不延长探索 |

## 4. Review 与遗留项映射

本次通过 GitHub 只读查询确认两个线程仍 `is_resolved=false`：

| 线程 | 真实状态与本次处置 |
| --- | --- |
| `PRRT_kwDORQ4ZKs6x9Ipp`，新渲染路线完整验证 | 非 outdated。历史最终 head 有大批功能/包证据，但实际可读性/Find 尚有失败。整条新路线不采用，该意见随该范围退出；没有 resolve，也不宣称已修复。C1/C2 如选定，须绑定其自己最终提交和受影响验收 |
| `PRRT_kwDORQ4ZKs6yGlhO`，混排 Latin 不覆盖规划宽度 | outdated 但未 resolved；`acec1025`/最终分支虽修改对应表达式，不能据此称审阅已关闭。新增规划/几何分支本轮不采用，该意见随范围退出；C1/C2 不触及该表达式 |

页末 Find 仍由 N02 按其预算和核心样例判断，不能借 N01 接回完整文字引擎研究。Skia 只保留主线已有诊断/普通 CI，历史根因 INCONCLUSIVE，不追加压力实验。历史量字失败作为已知限制保留；当前必验样例的最终可发布性由 N12 实际页面检查决定。

## 5. 原证据与本次验证边界

原始资料仍在原环境，没有复制运行时、字体、源码快照、页面或整批日志到本报告工作树。以下绝对位置只用于该桌面环境复取；长期转存须按[仓库存储规则](../repository-storage.md)按需归档。Git 提交与 GitHub PR 是可复取的源码/讨论身份。

| 本次实际读取的原记录 | 本次读取字节的 SHA-256 / 限定用途 |
| --- | --- |
| `/Users/wanghongyi/Projects/ofdrw.net/artifacts/roadmap-v12-20261021/PR78-svg-duplicate-emphasis-coordinator-review-20261030.json` | `fffdf7cbfd17d0fcd648f05ff55631de60f691f21cc954741991b0b54d4bec4b`；same-input 额外 bold 诊断，不是当前 main 动态验收 |
| 同目录 `PR78-ffc8cf43-source-paint-delta-coordinator-audit-20261030.json` | `f80c5d0d0c2853b2ec759dd14966463f2bb8c33d9fddd0b90925ec2635a03110`；同行宽度变化，不能误作 C1 导入 paint 的动态证明 |
| 同目录 `PR78-69da66cc-fixture-fix-coordinator-audit-20261031.json` | `eae5d9b056a22f8c92125ac175762fe517fcce67317a680271acdfe70e52ccbf`；fixture/clip 依赖及原工作轮次的证据范围 |
| 同目录 `PR78-2c4b48fb-exact-ci-feed-coordinator-audit-20261031.json` | `64efc164508574244607f28346e6b129c951f04f02f3749a451016ac206bb8ef`；准确 head、同树 CI checkout、12 包身份；不代表视觉接受 |
| 同目录 `PR78-2c4b48fb-svg-font-evidence-coordinator-audit-20261031T1139Z.json` | `ccddf04fb94d0f325a1281022e0eef7149223ba645454b57ea977394092f0272`；原 SVG 字节/字体引用审计，报告自身说明没有新 GUI 验收 |
| `/Users/wanghongyi/.codex/worktrees/9c7f/ofdrw.net/artifacts/r22-cjk-filled-stroke/pr78-2c4b48fb/current-status.json` | `6e0ec96490749a783c7923da5a66a1fd1b695328459cf074ede365da4d7b0283`；10-31 11:54 UTC 阶段封存，仍写 GUI pending；晚于它的实际 FAIL 以 PR 描述/复盘为准，不把 pending 当最终 GUI 记录 |
| `/Users/wanghongyi/.codex/worktrees/r22-span-context-query/ofdrw.net/artifacts/r22-shared-position-consumer/main-b02162be/actual-main-final-functional-and-preview-handoff.json` | `497befad5ddba9dde635f73156c0c69a293888414a0dfd5bcd49bd556f82ef41`；旧 main 功能/7 页 Preview 记录，其中 `quantityCounter=FAIL`、`overallVisualPass=false`；不转移为 beb02832 新验收 |

PR78 的 45 PDF/59 页、27 Copy、76 Find（66 通过/10 正例失败）是原轮次记录，本次未重验。上述小型记录本次重算了自身 SHA-256；未递归重验它们引用的全部页面、包和模块，不能扩大为本次完整证据完整性审计。

| 本次项目 | 结果 |
| --- | --- |
| 功能测试/构建/新包消费 | 未执行；N01 明确只读，不构建新候选 |
| 自动渲染检查 | 未执行，没有新产物 |
| PNG 视觉复查 | 未执行，本次不提出页面通过结论 |
| macOS Preview 验收 | 未执行，无 GUI 所有权/窗口操作；历史 FAIL 保留 |
| 文档与 Git 检查 | 19 个提交、当前 main、远端 PR/head/review 只读核对；仅本文暂存，提交前运行仓库存储检查 |

后续动作只有：协调者评审本处置，决定是否在既有发布缺口中选择 C1/C2 的明确最小范围；原 PR78 建议关闭并保留证据。本文没有创建实施任务、PR、自动化或后续调查，也没有执行远端关闭。
