# Ofdrw.Net.Archival：只读归档评估包（A1a–A1c）

本文记录[归档路线图](../engineering/plans/ofd-archival-roadmap.md) A1a 的设计决定，以及 A1b、A1c 的补充决定。适用版本为 main（0.4.0 之后，尚未发布）。用户说明见[归档评估指南](../content/guides/archival-assessment.md)。

## 定位与依赖方向

`Ofdrw.Net.Archival` 是可选包，目标框架 `netstandard2.0;netstandard2.1`，只引用 `Ofdrw.Net.Packaging`（经它引用 Core）。依赖方向固定为 `Archival → Packaging → Core`，不引用 Reader、Layout 或任何 Converter；反过来，Converter、Layout、Reader、Signatures、Cli 也不得引用 Archival。

- 包随默认发布集合打包（`verify-package-artifacts.py` 的 `PACKAGES`，共 13 个），但列入 `OPT_IN_PACKAGES`：任何其他包的 `Ofdrw.Net.*` 传递依赖到达它即校验失败；Archival 自身只允许依赖 Packaging。
- 包消费 E2E 检查 Converter 消费者的 `deps.json`，不得解析出 `Ofdrw.Net.Archival`。

Packaging 的 csproj 对 `Ofdrw.Net.Archival` 声明了 `InternalsVisibleTo`。理由：A1b/A1c 需要沿用 Packaging 的已知 CustomTags 遍历（`KnownCustomTagScope`）和资源基址解析（`OfdResourceBase`），共享同一份实现比在 Archival 中复制一份更不容易出现两边语义漂移；这些代码仍是内部实现，不成为公开 API。代价是 Archival 绑定 Packaging 的内部细节，因此 Archival 对 Packaging 采用**精确版本**依赖（两者同版本一起发布），不能单独升级其中一个。Core 的字体代码不走这条路，见下文“字体验证接口”。

新包组合已发布的 Preflight API（`ArchivalCoverageService` 及其中的 AR/AC/APF 结果、`ArchivalClauseCatalog`），不搬移公开类型，不复制检查器。Packaging 中的类型、程序集位置和行为不变；`ArchivalCoverageRuleCatalog` 仍是 15 项，账本修订仍为 `ofdrw-archival-ledger/1`。

## 公开入口与只读契约

`ArchivalAssessmentService.AssessAsync(Stream, ArchivalAssessmentOptions, CancellationToken)` 返回 `ArchivalAssessmentReport`。

- 输入须可读、可定位且位于零点；只读不写，保持打开，任何出口都恢复到零。参数错误抛 `ArgumentException`。
- 在 `LoadOptions.MaxInputBytes` 内一次读入快照并计算 SHA-256 和长度，此后所有检查都基于这份字节，报告的 `Input` 与被检查字节一致。超出字节预算时不计算摘要，`Input` 为 null。
- 预算沿用 `ArchivalPreflightOptions`（选项先复制再读输入）。本包自己的范围检查另按同一组 XML 字节、节点、深度和引用预算计数，不与 Preflight 服务共享计数。
- **取消抛出 `OperationCanceledException`，不返回部分报告**，与现有 Preflight 服务一致，因此执行状态没有 Cancelled。预算耗尽和输入被拒以报告返回：`ExecutionState` 为 `BudgetExhausted` 或 `Error`，并带稳定原因（`InputBudgetExceeded`、`InspectionBudgetExceeded`、`ScopeInspectionBudgetExceeded`、`InputRejected`）。普通 I/O 错误继续抛出。
- 公开加载器对 ZIP 限额和格式错误抛同一种 `InvalidDataException`，本包以消息中的 “limit” 区分预算耗尽与输入被拒；这是已知的弱区分。
- 非 `Complete` 的运行中，条款结果除确定的 Fail 外一律为 Unknown，输入支持维度追加 `AssessmentIncomplete`，不会呈现为可交付。

## 报告维度

| 维度 | 成员 | 来源与约束 |
| --- | --- | --- |
| 声明 | `Declaration` | 状态和结果直接映射 AC-001；`RawDocType` 取原始 `OFD.xml` 根的无限定属性，另报告是否存在同名外来属性。不使用 Reader 的缺省回退。 |
| 输入支持 | `InputSupport` | 所选 profile、`Supported`/`OutOfScope`/`Unknown` 及逐项原因。超范围不是违反国标，不生成条款 Fail。 |
| 条款 | `Clauses` | 105 项，账本顺序；每项带强度、条件标记、首批 profile 适用类别、矩阵置信度、映射规则观察、结果、证据位置、剩余证据和注记。 |
| 文档整体 | `Document` | 已有检查与无检查的条款、命名缺口、标准版本与来源哈希；`StandardOutcome` 只为 Fail 或 Unknown，`CoverageStatus=Incomplete`，`Certification=NotAssessed`。 |
| 软件要求 | `SoftwareRequirements` | 第 7 章四项，`NotAssessed`，列出写入端/阅读端缺口。 |
| 外部认可 | `ExternalRecognition` | 政策、引用版本、PKI、独立核准，`NotAuthenticated`。 |
| 执行状态 | `ExecutionState` 等 | 状态、原因，以及评估（`ofdrw-archival-assessment/1`）、profile、账本、矩阵修订号。 |

不存在 `IsCompliant` 一类的笼统成员。`Coverage` 原样保留底层 `ArchivalCoverageReport`，APF 六项仍是本地策略，不映射到条款。

条款结果按映射观察聚合：Fail > Unknown > Pass > NotApplicable；没有可执行检查的条款为 Unknown。Pass 只表示已映射的有界检查成立，不关闭该条的 `RemainingEvidence`；NotApplicable 只来自规则在完整作用域内证明的无对象。

## 首批 profile

`ofdrw-archival-profile/self-generated-simple/1`：单一 DocBody；未签名；未加密；无附件；无 Versions；无文档、大纲、页面或对象动作；无批注、扩展和 CompositeObject；简单 Layout 输出或使用受控字体的 Native DOCX→OFD，保留来源 CustomTags。

不以“由本库生成”为信任依据（`ProvenanceVerified` 恒为 false），每项都检查包内容。发现排除内容记为 OutOfScope（如 `MultipleDocuments`、`SignaturesDeclared`、`EncryptionMarkerPresent`、`AttachmentsDeclared`、`VersionsDeclared`、`ObjectActionsDeclared`）；无法识别或无法解析的结构记为 Unknown（如 `UnrecognizedDocumentContent`、`PageContentUnavailable`）。OutOfScope 优先于 Unknown。既有规则在超范围输入上发现的确定违规（多 DocBody 的 AC-002、根加密标记的 AC-015）照常 Fail。

字体载荷（A1c）和 CustomTags 引用闭包（A1b）由 AC-016～AC-020 判断。`FontPayloadNotAssessed`、`CustomTagClosureNotAssessed` 缺口只有在对应检查于一次 `Complete` 运行中全部作出判定（没有 Unknown 观察）时才去掉，否则保留。

## 规范强度策略

[适用性矩阵](../engineering/baselines/gbt42133-2022-clause-applicability.md)显示标准正文没有“应”：68 项 `Statement`、24 项 `Should`、9 项 `May`、4 项 `Shall`。按维护者授权的决定实施：

- `Statement` 与 `Shall` 条款上的规则 Fail 保持 Fail。
- `Should` 与 `May` 条款不产生条款级 Fail，也不使 `StandardOutcome` 为 Fail；规则发现未满足时，该观察列入 `Advisories`，条款结果为 Unknown。
- 每个条款都公开 `Strength`，调用方可自行采用更严或更宽的策略。
- 矩阵标为 `Unverified` 的 6.11、6.13.1(a)–(c) 原样公开，并附 `ModalScopeUnverified` 注记。
- AC-009、AC-012 以批注作为注释对象保留（6.10.1 方式 a）为前提；采用方式 b 把外观合并到页面时不适用。6.10.1(a)(2)、(a)(3) 条款附 `AnnotationObjectRetentionAssumed` 注记。

17 项 AR/AC 和 AC-017～AC-020 都映射到 `Statement` 条款。AC-016 同时映射 6.11（`Should`，`Unverified`）和 6.1（第一段，`Statement`）：标签闭包问题在 6.11 上只作 advisory；引用目标缺失或越出包外在 6.1（第一段）上为 Fail。

## 矩阵数据

矩阵以生成的 C# 表 `ArchivalClauseMatrix.Data.cs` 内置，修订号 `ofdrw-archival-matrix/1`，由 `scripts/generate-archival-clause-matrix.py` 从 `docs/engineering/baselines/gbt42133-2022-clause-applicability.json` 生成。不嵌入 JSON，避免为 netstandard2.0 引入 JSON 依赖。`ArchivalClauseMatrixTests` 逐项比对 JSON 全部字段，并要求条款序列等于 `ArchivalClauseCatalog.Clauses`。JSON 有任何行变化时重新生成并提升修订号。

## 规则编号与扩展点

AC-001–AC-015 归 Packaging 所有。本包从 **AC-016** 起继续 `AC-` 序列，编号在本包内定义，不修改 `ArchivalCoverageRuleCatalog`，因此 Packaging 的规则数量和既有报告不变。观察的 `RuleOwner` 标明 `Ofdrw.Net.Packaging` 或 `Ofdrw.Net.Archival`。

扩展点为内部接口 `IArchivalClauseCheck`（非公开插件 API）：每项检查带 `ArchivalOwnedRuleDefinition`（规则号、条款、条件、修订），从 `ArchivalCheckContext` 取得有预算的 XML 读取器、范围检查已解析的根、文档和页面，返回至少一条观察。服务构造时校验编号为 AC-016 及以上、互不重复、不与 Packaging 冲突且条款在账本内。检查耗尽预算时该检查记 Unknown，运行记为 BudgetExhausted。一条规则可映射多个条款（`AdditionalClauses`），服务校验每条观察的条款都在规则声明的条款内。当前注册：

| 规则 | 条款 | 内容 |
| --- | --- | --- |
| AC-016 | 6.11；6.1（第一段） | 已知 CustomTags 引用闭包（A1b） |
| AC-017 | 6.2.6(b)(1) | 填充/描边文字所用字体的包内载荷覆盖实际字符和字形编号（A1c） |
| AC-018 | 6.2.6(b)(2) | 同上，裁剪区文字 |
| AC-019 | 6.2.6(c)(1) | 嵌入载荷保留轮廓表，所用字形轮廓可读 |
| AC-020 | 6.2.6(c)(2) | 嵌入载荷保留八张必需表 |

## 已知 CustomTags 与 EMR 的归类（A1b）

APF 新增 `ArchivalPreflightOptions.RecognizeKnownCustomTags`，默认关闭，关闭时输出逐字节不变（回归测试固定）；开启时 `ProfileVersion` 为 `ofdrw-local-preflight/0.1-draft+known-custom-tags.1`。评估入口在内部开启它。

遍历按 XML 字段语义进行（元素/属性名、命名空间、FileLoc 的元素或属性形式），不用正则。识别两种形状：

- **中性键值通道**：`TypeID="ofdrw-kv"`、`NameSpace="urn:ofdrw-net:custom-tags:kv:1"`，载荷根为同一命名空间的 `Tags`。本库当前写出的形状。
- **EMR 通道**：`TypeID="EMR"`、`NameSpace="urn:ofdrw-net:custom-tags:emr"`，载荷根为 OFD 命名空间的 `EMRTags`。**归类为“本库历史上写出的已知形状”，没有对照 GB/T 33190 核验**；报告的覆盖缺口如实记录这一点。识别它只为了让历史产物不被无差别 Unknown，不表示 EMR 载荷符合任何电子病历标准。

其余（未知 TypeID/命名空间、外来属性、SchemaLoc、歧义引用、循环、超预算）为 Unknown；空、外部、越出根目录或缺失的目标为 Fail。

## 字体检查（A1c）

字体按声明它的资源表（Document 的 PublicRes/DocumentRes、页面的 PageRes）解析，`FontFile` 相对该表的 `BaseLoc`，复用 Packaging 的 `OfdResourceBase`。已用字符与字形编号取自页面和模板中的 TextCode 与 CGTransform；无法建模的映射、范围不完整的页面为 Unknown。宿主字体从不作为证据；`fsType` 只记录，可解析不代表有许可。没有 FontFile 时按 `ArchivalAssessmentOptions.FontExemptionPolicy` 处理：默认 `ExemptionBasisUnverified` 为 Unknown，`NoExemptionsAccepted` 为定位到 Font 元素的 Fail。

## 字体验证接口的决定

`src/Ofdrw.Net.Core/Fonts/` 的 OpenType 解析与子集代码为 internal，本包不在 Core 的 `InternalsVisibleTo` 名单内。**决定：在 Core 增加面向字节验证的窄公共接口，不为 Archival 增加内部可见性。** 理由：归档包不绑定尚未稳定的内部类型；依赖方向仍为 `Archival → Packaging → Core`；接口可被其他只读检查复用。

A1c 定稿的形态：

- 位置：`Ofdrw.Net.Core.Fonts`，无状态静态入口 `OfdFontPayloadInspector.Inspect(byte[] payload, OfdFontPayloadRequest request, CancellationToken)`。建在既有内部 `OpenTypeFace`、cmap 和字形轮廓代码上，没有第二套解析器。
- 请求：需要覆盖的 Unicode 码点和 GID 集合，以及载荷字节、cmap 工作量和字形检查数预算。
- 结果：`OfdFontPayloadReport`，含 `Status`（`Parsed`/`Unsupported`/`Malformed`/`BudgetExceeded`）、`Format`（`TrueType`/`OpenTypeCff`/`Unknown`）、稳定原因码、表清单、缺失的轮廓表和必需表、缺失码点/GID、不可读字形和 `EmbeddingPermissions`。TTC、WOFF、CFF2、可变和无法识别的格式为 `Unsupported`，归档层映射为 Unknown 而非 Fail；畸形字体返回状态，不抛出格式异常；取消照常抛出。
- 约束：不查找宿主字体，不回传字体字节，不判断嵌入许可是否构成授权；许可位只作为单独字段报告。

## 有意不声明的内容

- 不存在整部标准的 Pass，也不提供 `IsCompliant`。
- 本包不生成、不修复、不标记 OFD-A；不删除任何内容来扩大支持范围。
- 第 7 章阅读器与写入端行为未被证明；单个文件无法证明。
- 机构政策、PKI 与独立核准未经认证；引用标准版本未固定。
- 输入为本库生成不构成信任依据；简单输入 Supported 只表示在 profile 内，不表示符合标准。
