# Plan 030 只读归档预检设计

- 交付状态：设计 spike 已完成；生产规则引擎、公开 API、CLI、功能测试、视觉验收与发布均未实施。本文不构成 GB/T 42133 或其他标准的符合性、认证或法律效力结论。
- 源码核查基线：`ec121b087e095db4b4325cd4d61e61cbb4372142`；原计划编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`；核查日期：2026-10-15。
- 关联：[原计划](/Users/wanghongyi/Projects/ofdrw.net/plans/030-archival-preflight-spike.md)、[归档方向](../docs/capability-roadmap.md)、[开发约定](../docs/capability-roadmap-development-plan.md)。指定源码基线没有 `plans/` 目录，原计划从上述协调者路径完整读取；本 worktree 只新增本文，索引由协调者维护。下文“建议”类型、命令与样例均未实现或执行。

## API 与边界

### 目标与输入所有权

候选场景是业务系统在归档入库前，对**一个已有 OFD ZIP**做静态排查，获得有限规则的观察证据，再由调用方自己的政策决定是否接收。仓库路线图提供了方向，但本次没有具体档案馆、业务系统或供应商提供的正式需求、接收政策、合格样本；不能把场景假设写成已确认客户需求。

建议先在 `Ofdrw.Net.Packaging` 内做独立预检模块，复用已有 Core 引用；公开命名及是否另发可选包待调用方验证后确定。库保持 netstandard2.0/2.1，不依赖 Reader 的页面渲染、Converter、密码包、SES 或网络服务。建议入口如下，**只是契约草案，所有类型和入口均不是现有 API**：

```csharp
// 建议，未实现。
Task<ArchivalPreflightReport> CheckAsync(
    Stream input,
    ArchivalPreflightOptions options,
    CancellationToken cancellationToken = default);
```

契约要求：

1. 只读输入流；不调用 Writer、修复、解密、转换、保存或外部命令，不修改 ZIP、XML、附件及原始 byte[]，不重排包条目。输入流保持打开；可寻址流在成功/失败/取消后恢复原 Position，无法恢复则显式报告 I/O 错误。非可寻址流从当前位置消费到末尾，位置不可恢复，调用方应传专用流。
2. 拒绝不可读流；可寻址输入必须从 Position 0 开始。对非可寻址流做有上限的内存暂存，不写明文临时文件，读取受取消和 `MaxInputBytes` 约束。调用方在借用期间不得并发修改输入，生产调用应提供稳定快照；前后文件哈希相同不能证明并发读取过程一致。
3. 不抓取 HTTP(S)、FTP、`file:`、UNC 或其他 URI，不读取宿主字体/绝对路径/外链，不注册 XML resolver。XML 显式禁止 DTD/外部实体并限制字节、字符、节点、深度和访问次数。URI 分类发生在路径解析/任何目标访问之前。已知 OFD 路径字段里的 `/Doc_0/...` 是包根路径，不是宿主绝对路径。
4. 选项快照包含加载预算、XML/遍历预算、规则集版本和本地政策。建议初值：XML 总字节 32 MiB、单 XML 8 MiB、深度 64、总节点 200,000、已知引用边 100,000、输出结果 10,000 条；这些是待样本校准的本地预算，不是标准限值。在分配或下降遍历前执行限制，用已访问条目/节点及预算截断循环，不做无界递归。
5. 调用错误抛出参数异常；取消保留 `OperationCanceledException`；普通 I/O 错误抛出原异常。底层 `InvalidDataException`（坏/危险 ZIP、已知预算超限）映射为 `RunState=Incomplete` 的输入诊断，未评估规则为 Unknown。已知位置局部 XML 畸形可形成确定失败；无法读取、加密遮蔽或未知编码则 Unknown，不尝试解密。不吞掉意外编程异常。
6. 库返回报告；未来 CLI 仅向显式指定的另一条流写报告，不提供自动覆盖输入的路径重载。CLI 文件别名检查另列切片；禁止隐式远程校验、在线证书查询及认证文案。

### 四态结果与证据

建议独立结果契约，不修改现有 `ValidationReport` 的公开成员或既有布尔含义：

| 建议字段 | 语义 |
| --- | --- |
| `ProfileVersion`、`Scope` | 本地 profile `ofdrw-local-preflight/0.1-draft`，已观察包/DocBody/条目与覆盖范围；不是标准版本。记录启用规则与政策，默认 `RequireSingleDocument=true`、`RequireNoEncryption=true`。 |
| `RunState` | Complete / Incomplete，表示请求范围是否执行完成，不代替规则结论。 |
| `RuleResults[]` | 每条规则/观察对象一项，包含 RuleId、RuleVersion、Severity、Outcome、ApplicabilityReason、ReasonCode、Location、Evidence、Recommendation、Source；无对象也需明确结果。 |
| `Outcome` | Pass / Fail / NotApplicable / Unknown。Unknown 用于必要对象/范围不清、未支持语义、证据不足或预算耗尽；NotApplicable 必须有证据证明无对象，或明确本地政策不要求该规则。 |
| `Severity` | 沿用 Info / Warning / Error；下表列失败严重度，Unknown 默认 Warning，Pass/NotApplicable 为 Info。Warning 的 Fail 仍是 Fail。 |
| `Evidence[]` | 建议包含 EvidenceId、Entry、namespace-aware Selector、Observed、Expected、ReasonCode、关联结果 ID；包级读取错误用 Entry=`<input>` 与 ZIP 项/偏移等实际可得标识。不能从异常消息猜路径。 |
| `Source` | `Kind=LocalPolicy`、本地文档/代码来源、`StandardStatus=Unverified`、`Clause=""`。未获得标准原文、版本和映射审阅，不填写条款号。 |
| `CoverageGaps[]`、`Observations[]` | 未支持命名空间、未知扩展、未建模语义、孤立疑似签名/加密条目、截断与加密遮蔽单独记录；存在观察不等于规则失败或验签有效。 |
| `SummaryOutcome` | 有确定 Fail 则 Fail，仍保留 Unknown；无 Fail 但有 Unknown 或 Incomplete 则 Unknown；范围完整、至少一个适用结果且适用结果全 Pass 才 Pass；全部 NotApplicable 时为 NotApplicable。 |

规则多对象聚合按 Fail > Unknown > Pass；只有无适用对象时为 NotApplicable。已有确定失败但部分对象未知时，保留对象级 Unknown 和缺口。结果按 RuleId、Entry、Selector 做 ordinal 稳定排序；EvidenceId 是报告内 ID，不能当跨报告对象身份。Summary Pass 只针对声明的本地范围。

建议 Location 为 `Doc_0/PublicRes.xml#/{http://www.ofdspec.org}Res[1]/{http://www.ofdspec.org}Fonts[1]/{http://www.ofdspec.org}Font[1]/{http://www.ofdspec.org}FontFile[1]`；索引在同名同命名空间兄弟中计数。证据同时记录 Font ID、原引用、规范化目标和长度。目标缺失时定位**声明节点**，目标另列；缺 OFD.xml 时定位包根所需条目。输出只含定位所需元数据并限制长度；外链凭据/查询值脱敏，不输出正文、密码、密钥、附件或字体内容。

需要旧消费者时另做显式适配器：Fail/Unknown 映射 `ValidationFinding`，Clause 留空；Evidence 以结果 ID 为键前缀。仅新 Summary 为 Pass 才令旧 `Passed=true`；Unknown、Incomplete、全 NotApplicable 均 false 并附解释。适配不得丢弃四态主报告，也不修改已有调用方的含义。

### 六条有限本地规则

以下 ID 均为本仓库本地政策 ID，无已验证标准条款。每项只承诺“所列对象在所列范围满足条件”，不得扩写成全包或标准认证。

| RuleId（版本1） | 失败严重度 / 适用性 | 可确定条件与定位证据 | Unknown 与 NotApplicable 边界 |
| --- | --- | --- | --- |
| APF-001 `single-document-structure` | Error；RequireSingleDocument=true | 支持命名空间的 OFD/DocBody 直属子节点恰一项，其中恰一非空 DocRoot 指向可解析 Document，至少一个 Pages/Page 的非空 BaseLoc 指向已有条目。缺失、零/多个、坏 XML、重复 DocRoot、缺页则 Fail。记录节点计数及根/页声明与目标；不验证页面绘制内容。 | 未支持根/Document 命名空间、内容加密或加载未完成为 Unknown。RequireSingleDocument=false 为政策 NotApplicable，仍记录文档数量；其他规则清点全部可识别 DocBody，不默认只取首个。 |
| APF-002 `known-payloads-in-package` | Error；存在已知资源字段 | 清点 CommonData/PublicRes、DocumentRes，Pages/Page 与 TemplatePage 的 BaseLoc，页面/模板相邻且实际存在的约定 PageRes.xml，以及 Res 中 FontFile/MediaFile：必须是明确包内路径且目标存在。外链、确定根逃逸、空的必需位置、缺目标则 Fail。记录源/selector、原引用、BaseLoc 组合、规范化目标与存在性。 | 范围可判定且无这些字段为 NotApplicable。所有权/BaseLoc/命名空间无法识别，作用于这些字段的未知扩展，含歧义 query/fragment/转义路径为 Unknown。只检查已知边，不能宣称未知 XML 不藏外链。 |
| APF-003 `declared-font-payload-present` | Error；已识别 Res/Fonts/Font 声明存在 | 每个声明恰一非空 FontFile，按资源表 BaseLoc 与 FontFile 包根覆盖规则解析，目标存在且非零长；缺 FontFile、缺/空文件、确定外链则 Fail。记录资源表、Font ID、BaseLoc、FontFile 与长度。 | 无字体声明且所有已覆盖内容均无 TextObject 为 NotApplicable；有文字但字体引用不明、重复 ID、未知资源覆盖、未支持组合/注释作用域则 Unknown。仅证明载荷存在，不判断字体格式、许可、嵌入位、子集/字形覆盖和外观。 |
| APF-004 `known-encryption-absent` | Error；RequireNoEncryption=true | 完整条目清点无根 Encryptions.xml、无 PasswordCrypto/ 保留条目，已检查控制 XML 无未支持加密声明时，在**已知标记**范围 Pass。存在根 Encryptions.xml 即 Fail，不必先解密/解析其内容。记录精确条目及可得 profile namespace，不判断密码强弱。 | 孤立 PasswordCrypto 条目、未识别加密扩展、内容不可读或清点不完整为 Unknown，不能证明无加密。RequireNoEncryption=false 为政策 NotApplicable，仍保留加密观察及其他规则的遮蔽 Unknown。 |
| APF-005 `attachment-locations` | Error；Document 声明 Attachments 或有疑似索引 | 非空 Attachments 声明指向可解析索引；每条 Attachment 恰一非空 FileLoc，明确包内且目标存在。确定外链、External=true、缺索引/载荷、空 FileLoc 为 Fail。记录附件 ID/Name、FileLoc、目标存在性。零字节附件允许存在。 | 完整检查的 Document 无声明且无疑似索引为 NotApplicable；未知编码、加密、冲突的属性/子元素 FileLoc、无声明的孤立 Attachments.xml 为 Unknown。不能把 Reader 的目标缺失即 external 语义当真实 URI 证据。 |
| APF-006 `signature-inventory` | Warning；DocBody 声明 Signatures 或有疑似签名项 | 清点声明→签名列表→Signature/@BaseLoc→Signature.xml→SignedValue 与 References/Reference/@FileRef；已知字段非空且目标存在。空声明/列表、缺列表/载荷、确定 FileRef 缺目标或外链为 Fail。记录签名 ID、原声明、目标及被列出的引用集合，不计算摘要。 | 无声明且无疑似签名项为 NotApplicable；不支持结构/变换、坏 XML/加密无法清点、孤立签名项或歧义位置为 Unknown。Pass 只表示可清点及目标存在，不表示摘要、签名值、证书链、身份、时间戳或法律效力有效。 |

APF-002 的 PageRes.xml 清点参照 Reader 的页面/模板同目录约定，不虚构 XML 中已有 PageRes 声明；不存在时不凭约定制造缺失错误。APF-002 不重复处理附件/签名字段，由005/006负责。压缩图像、字体等载荷内部引用、附件嵌套包、脚本、动作网页链接、CustomTags/Extensions 任意内容、XSD 全量验证、完整资源 ID 闭包属于 CoverageGaps。网页动作不是本规则的资源外链失败，仍应说明未覆盖。没有解析资源内容时，不能用“没发现”代替“明确不存在”。

## 现有能力复用

以下为本次基线静态核查，均为真实现有代码，与建议新增类型区分：

| 现有入口 | 已有能力 / 复用方式 | 限制 |
| --- | --- | --- |
| [ValidationReport.cs](../src/Ofdrw.Net.Core/Validation/ValidationReport.cs) | Finding 有 RuleId/Clause/Severity/Message/Location/Recommendation；Evidence 是 Key/Value；Report 是 Passed 与两个列表。沿用严重度及可选字段映射。 | 没有逐规则状态、适用性、证据关联和执行范围，不能以无 Error 判通过。当前 src/tests 搜索未发现该模型的生产使用点，但公开类型仍不能随意改变兼容性。 |
| [structure checker](../src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs)、[issue](../src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureIssue.cs) | 检查 OFD.xml、DocRoot、页面目标、若干 Document 引用；issue 只有 Code/Message/IsError。保留为宽松结构诊断和回归参照。 | 按 LocalName 泛扫、允许多 DocRoot，私有 Resolve 对越根 `..` 会截断；不是安全路径解析器。新模块自行确定 namespace/所有权/位置，不能从 Message 拼证据，也不改历史行为。 |
| [loader](../src/Ofdrw.Net.Packaging/Archive/OfdPackageLoader.cs)、[options](../src/Ofdrw.Net.Packaging/Archive/OfdPackageLoadOptions.cs) | 约束压缩输入、条目数、单项/总展开量、压缩比；拒绝绝对/`..` ZIP 项与忽略大小写重复项；Read 模式保留输入打开。复用公开 LoadAsync 与选项快照。 | 输入稳定性/报告映射由新入口负责；PageCount/Annotation 限额由 Reader 路径执行，不会自动约束预检 XML/边遍历。需独立预算，不能把加载成功当语义完整。 |
| [accounted loader](../src/Ofdrw.Net.Packaging/Archive/OfdAccountedPackageLoader.cs) | 内部受控加载器有累计预算上下文，并限制 .NET10、EOCD、ZIP64/注释等轮廓。 | 本设计不直接使用内部受控重载，不把它的 .NET10 特定保证当 netstandard 公开入口已有保证。若需要同等累计内存证明，另拆实现与平台验证。 |
| [archive](../src/Ofdrw.Net.Packaging/Archive/OfdPackageArchive.cs)、[path](../src/Ofdrw.Net.Core/IO/OfdPackagePath.cs) | Archive 枚举/定位条目；Core 内部路径工具解析相对路径并拒绝根逃逸。先识别 URI 再复用包路径规则。 | Archive 构造器借用 Dictionary，GetBytes 暴露 byte[]；新入口自己加载并独占，只读数组，首版不接收调用方可变 Archive。路径工具未识别 URI；archive 使用 case-insensitive 语义，报告明示，标准大小写语义待核实。 |
| [constants](../src/Ofdrw.Net.Core/Constants/OfdConstants.cs)、[reader](../src/Ofdrw.Net.Reader/Readers/OfdReader.cs) | 两个 namespace 为 `http://www.ofdspec.org`、`http://www.ofdspec.org/2016`；Reader 读取 public/document/page/template 资源字体和 BaseLoc、附件索引。作为字段依据。 | Reader 部分按 LocalName 匹配并将缺失附件目标当 external；预检按 namespace 白名单清点源 XML，区分真实外链与包内缺失，不靠有损模型推断全包状态。 |
| [resource catalog](../src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs) | 编辑组件有 namespace-aware 字体/媒体清点及重复 ID 拒绝逻辑，可作字段/作用域回归依据。 | 不实例化可计划修改的编辑组件进行只读检查。 |
| [password envelope](../src/Ofdrw.Net.Crypto.Password/OfdPasswordEnvelope.cs) | 私有 profile v2 使用根 Encryptions.xml、PasswordCrypto/decryptseed.dat、entriesmap.dat、auth.dat，明确无长期保存/认证保证。 | 只清点标记，不引用可选包、不收口令/解密；孤立标记 Unknown，未调用密码包不代表无加密。 |
| [signature models](../src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs)、[verifier](../src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs) | 已区分 HasSignatureDeclarations、HasSignatures、ReferenceIntegrityValid、CryptographicStatus、FullyValid；支持显式验证器。保留这些文案边界。 | 预检不调用 VerifyAsync、不注册测试验证器，只读原始声明/目标；自签容器和外观不是有效签章证明。 |
| [bad package tests](../tests/Ofdrw.Net.Packaging.Tests/BadPackageCorpusTests.cs)、[roundtrip tests](../tests/Ofdrw.Net.Packaging.Tests/PackageRoundTripTests.cs) | 已有小 ZIP 构造、坏 ZIP/路径/重复/预算测试、字体附件断言及 checker 调用。后续复用构造方式。 | 本次未执行这些测试、未生成归档包，源码存在不是本次运行证据。 |

## 样例与预期

### 确定性种子与六规则四态矩阵

以下均为**设计样例，未生成/未运行**。建议 `Seed-Minimal`：UTF-8、固定 ZIP 时间戳与条目顺序；OFD.xml 使用支持 namespace，仅一 DocBody/DocRoot=`Doc_0/Document.xml`；Document 有 Pages/Page ID=1、BaseLoc=`Pages/P0/Content.xml`；页面为只有 PathObject 的已知 Page。无资源、文字、加密、附件、签名。预期001 Pass、002 Pass（Page BaseLoc）、003 NotApplicable、004 Pass、005/006 NotApplicable；Summary Pass 只针对这个范围。

`Seed-Font` 追加 CommonData/PublicRes=`PublicRes.xml`；资源表 Res/@BaseLoc=`Res`，Font ID=1、FontFile=`font.bin`；`Doc_0/Res/font.bin` 为固定4字节非空哨兵；页面 TextObject Font=1、TextCode=`归档 Archive`。这个**刻意非字体载荷**可让003的存在性 Pass，证明规则不冒充字体合法性/覆盖性检查；不用它渲染，也不作视觉验收样例。真正字体解析规则必须另用合法字体。

每个病例从种子作所述最小变化。每格是本条规则 Outcome，不推断其他规则/总报告；Actual 均“未执行，仅设计审阅”。

| RuleId | Pass 示例及证据 | Fail 示例及证据 | NotApplicable 示例及理由 | Unknown 示例及缺口 |
| --- | --- | --- | --- | --- |
| APF-001 | Seed-Minimal：DocBody=1、DocRoot=1、Page=1，目标均存在。 | 复制第二 DocBody：OFD.xml DocBody[2]、count=2，`MultipleDocuments`；另设缺 Document、缺 Page、坏 Document XML 分离病例。 | 同一双文档包显式 RequireSingleDocument=false，`PolicyAllowsMultipleDocuments`；仍记录 count=2，其他规则遍历两文档。 | OFD 根 namespace=`urn:vendor:unknown`：`UnsupportedNamespace`，不能靠 LocalName 的一个 DocBody 判 Pass。 |
| APF-002 | Seed-Font：Document→PublicRes、Page→Content、Res→font.bin 的原引用、BaseLoc 组合及已有目标齐全。 | FontFile=`https://example.invalid/font.bin`：定位该节点，`ExternalReference`，请求数必须0；另测缺目标及越根 `../../../x`。 | 已知 OFD/Document、Pages为空、无资源/模板字段：002 无对象 NA；001 必须 Fail，整体不能 Pass。 | PublicRes 所有权/格式由未知扩展替代：`UnsupportedResourceScope`，不能把空枚举当无资源。 |
| APF-003 | Seed-Font：FontFile 指向4字节哨兵，记录 PayloadLength=4，明确不验证字体格式。 | 删除 FontFile：定位 Font ID=1，`FontPayloadNotDeclared`；另测零字节、缺载荷。 | Seed-Minimal 无字体且完整已知内容无 TextObject：`NoFontDeclarationsOrText`。 | 页面 Font=99、资源仅Font=1：`UnresolvedTextFont`；或同作用域两个 ID=1；不退回系统字体或猜引用。 |
| APF-004 | Seed-Minimal：完整清点无已知加密标记，列 marker set 与条目数，仅对已知标记范围 Pass。 | 添加根 Encryptions.xml，即使其 XML 错误也为 `EncryptionDeclared`；其他被遮蔽规则 Unknown。 | 同包显式 RequireNoEncryption=false：`PolicyDoesNotRequirePlainPackage`；保留加密存在观察。 | 只有 PasswordCrypto/entriesmap.dat 无 Encryptions.xml：`OrphanEncryptionMarker`，不能报确定有/无加密。 |
| APF-005 | Attachments=`Attachs/Attachments.xml`，索引 Attachment ID=A1、FileLoc=`a.txt` 且载荷存在，可为零字节。 | 声明索引缺失：定位 Document/Attachments，`MissingAttachmentIndex`；另测载荷缺失、外链、External=true。 | Seed-Minimal 无声明/疑似索引：`NoAttachmentDeclarations`。 | 无声明但有 `Doc_0/Attachs/Attachments.xml`：`OrphanAttachmentIndex`；或 FileLoc 属性/子元素不同：`AmbiguousFileLoc`。 |
| APF-006 | DocBody/Signatures=`Doc_0/Signs/Signatures.xml`，列表 Signature ID=S1、BaseLoc=`Sign_0/Signature.xml`，后者 SignedValue=`SignedValue.dat`、Reference FileRef=`/Doc_0/Document.xml`，所列目标存在；固定无效 SignedValue 哨兵也可清点 Pass，验签未判断。 | 空 Signatures 或缺列表：定位原声明，`EmptySignatureDeclaration` / `MissingSignatureList`；另测空列表及缺 SignedValue 目标。 | Seed-Minimal 无声明及 Signs/Signatures.xml、Signature.xml、SignedValue.dat 等疑似条目：`NoSignatureDeclarations`；不代表验签通过。 | 只有孤立 Signs/Sign_0/SignedValue.dat：`OrphanSignaturePayload`；或签名未知变换：`UnsupportedReferenceTransform`。 |

签名路径草案：DocBody/Signatures 按包根解释；Signature/@BaseLoc 相对列表；SignedValue 相对 Signature.xml；以 `/` 开头覆盖为包根；Reference/@FileRef 作为包根引用，不与签名目录拼接。与未来核实的标准冲突时先修订本地 profile/种子，不用碰巧找到同名项掩盖歧义。

附件属性 FileLoc 仅建议兼容无子元素、单属性、无冲突形状并记录兼容观察，复杂形状 Unknown。未知 namespace 的同名 DocRoot、Font、Attachment、Signature 不当已知声明；若它处于规则必要范围则相应规则 Unknown。孤立条目只作观察，不能按文件名推断其为生效的资源身份。

### 跨规则及只读不变量

| 场景 ID | 设计输入/动作 | 必须验证的预期 | 本次实际 / 结论 | 未来复跑入口（均为建议，未存在） |
| --- | --- | --- | --- | --- |
| RuleMatrix | 上表6×4基本病例及多对象失败+未知混合。 | 各RuleId四态均有原因和证据；Fail与Unknown同时保留，政策NA不变Pass，无空结果假通过。 | 未运行；逐格静态审阅完成。 | `ArchivalPreflightTests.RuleMatrix` |
| ReadOnly | 禁止Write的只读流，保存前后完整字节、位置及生命周期；原文件只读打开。 | SHA-256/长度不变，Write/SetLength/Dispose（借用流）计数0；可寻址位置恢复，非可寻址按契约消费；异常/取消同样验。哈希仅证明本次输入不变，不是认证签名。 | 未生成输入、未执行。 | `ArchivalPreflightTests.ReadOnly` |
| NoFetch | 已知字段写HTTP/HTTPS/file/UNC/未知scheme、DTD外实体、包根路径；宿主文件与本地服务器放可观察哨兵。 | 网络请求、宿主读取次数0，DTD显式拒绝，`/Doc_0/x`只在ZIP内解析；结合静态依赖审阅与I/O哨兵，不能用不存在域名未报错证明无请求。 | 未执行，仅定义未来断言。 | `ArchivalPreflightTests.NoFetch` |
| EvidenceLocation | 外链、缺FontFile、缺附件、缺签名列表。 | 每个Fail/Unknown都有RuleId、Entry、Selector/包级标识、ReasonCode、Observed、Expected、EvidenceId；关联无悬挂，能回到声明。 | 已审阅每规则证据设计，未生成产品报告。 | `ArchivalPreflightTests.EvidenceLocation` |
| NoCertification | 001 Fail+003 Unknown，其余混合；另测全NA政策/空对象结果。 | 无Certified、标准符合或FullyValid结论；有范围、LocalPolicy/空Clause/未核实来源、Unknown计数和缺口；全NA不得Passed=true。 | 仅报告草图审阅，未运行。 | `ArchivalPreflightTests.NoCertification` |
| BoundedInput | 小坏ZIP、大小写重复、危险路径、深度65、边/结果预算+1、取消和循环引用。 | 有界样例触限产生Incomplete与Unknown；取消仍抛取消异常；不得丢未评估项后报Pass，不用无界极值输入。 | 未执行，无性能/内存实测。 | `ArchivalPreflightTests.BoundedInput`及现有BadPackageCorpusTests |

报告草图（手写契约示例，未由产品生成；Selector前缀ofd仅为缩写，正式报告携带namespace映射或展开名）：

```json
{
  "profileVersion": "ofdrw-local-preflight/0.1-draft",
  "runState": "Complete",
  "summaryOutcome": "Fail",
  "certification": "NotAssessed",
  "ruleResults": [
    {
      "ruleId": "APF-003",
      "ruleVersion": 1,
      "severity": "Error",
      "outcome": "Fail",
      "applicabilityReason": "One font declaration in the observed resource table",
      "reasonCode": "FontPayloadNotDeclared",
      "location": "Doc_0/PublicRes.xml#/ofd:Res[1]/ofd:Fonts[1]/ofd:Font[1]",
      "evidence": [{"evidenceId": "E1", "observed": "Font ID=1; FontFile count=0", "expected": "One local non-empty font payload"}],
      "source": {"kind": "LocalPolicy", "standardStatus": "Unverified", "clause": ""}
    }
  ],
  "coverageGaps": ["Font validity, glyph coverage and licensing not assessed", "Standard clauses not verified"]
}
```

该片段只展示一条失败形状，**不是完整报告 fixture**；完整报告须输出所有启用规则及NotApplicable/Unknown。

## 验收命令

### 本次实际静态验证

worktree：`/Users/wanghongyi/.codex/worktrees/ofdrw-plan-030/ofdrw.net`；分支：`codex/plan-030-archival-preflight-spike`。仅新建本文及忽略的自有 artifacts，未改源代码、工程、依赖、版本、原计划或索引。本地提交授权来自本次执行指令，覆盖原计划“不授权提交”旧句；没有push、PR、merge、publish。

以下为本次实际静态检查入口，日志/脚本在 `artifacts/plan-030/static-checks.txt` 与 `artifacts/plan-030/check-design.py`（按仓库惯例被忽略，不进产品包）：

```bash
git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- plans/030-archival-preflight-design.md
git status --short
git diff -- plans/030-archival-preflight-design.md
git diff --cached -- plans/030-archival-preflight-design.md
python3 artifacts/plan-030/check-design.py
git diff --check
```

初始漂移与两种diff为空，worktree初始干净，HEAD确认为指定基线；该基线无plans目录不是已有文件被删。本次静态验收核对六节唯一/非空、六规则四态样例、四个核心场景、JSON可解析、相对链接按本文目录可解析、无空白错误及仅改本文。本次检查结果：6节非空且唯一、6条规则共24个四态样例、19个本地链接全部存在、1份JSON草图可解析、作用域仅本文，均通过；git diff --check退出0。最终结果同时交接于上述静态日志与提交记录；这些检查不证明规则实现。

### 后续实施验证（建议，全部未执行）

建议新增的 `ArchivalPreflightTests` 尚不存在；先创建行为测试、统计发现/执行数，再跑过滤。零命中不得当通过。使用真实已有Packaging测试工程，避免假设已有归档项目：

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-030/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --filter 'FullyQualifiedName~ArchivalPreflightTests' --logger 'trx;LogFileName=archival-preflight.trx' --results-directory artifacts/plan-030/future-tests
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --logger 'trx;LogFileName=packaging.trx' --results-directory artifacts/plan-030/future-tests
dotnet build src/Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
```

宿主遵循global.json的SDK10.0.100+latestFeature；build应覆盖工程两个netstandard目标，不能把.NET10运行结果当其他运行时证据。若独立可选包落地，改为真实项目路径并补同批本地包消费，证明不引密码/签名包也能运行。CLI实现后再按真实解析器定义命令与输入/输出路径别名门；目前没有 `archive-check` 产品命令，不给出假装可执行的调用。

本次功能测试：未执行（无实现）；自动渲染检查：未执行（无生成页面）；PNG视觉复查：未执行；macOS Preview验收：未执行。静态设计无转换/排版/渲染改动及视觉产物要求，不把这些门标为通过。后续若修改转换、字体或布局，额外按AGENTS.md生成并逐页验收，DOCX遵循DOCX→native OFD→PDF→Preview。

## 风险与未决项

| 项目 | 已知事实 / 本次边界 | 实施前决定或证据 |
| --- | --- | --- |
| 标准条款 | 无GB/T42133正文、版次确认及授权材料，路线图中的标准名只是方向，本地政策不等于标准要求。 | 调用方提供可合法使用的版本/条款、正反样本，由领域人员逐规则审阅；此前Clause为空、Unverified，不虚构条款号/必选性/认证。本spike无需获取收费标准。 |
| 真实调用方 | 尚无对多文档、多签名、附件、加密、字体、误报容忍度的正式接收规则。 | 确认单文档/无加密是否硬性、签名是否必须、附件能否外置、报告是否存档、规模/延迟预算。六规则只是候选MVP。 |
| Unknown传播 | 未知XML、资源作用域、加密可遮蔽证据；LocalName泛扫易误认。 | 字段所有权/namespace白名单、覆盖清单先行，区分无对象与无法枚举，不把未知形状强行兼容成通过。 |
| 字体与外观 | FontFile非空不代表字体可解、合法嵌入、中文覆盖或渲染正确。 | 升级能力时新增独立规则、合法字体及真实功能/视觉证据，不悄悄扩写003含义。 |
| 签名与密码 | 不算摘要、不验签名值/证书/密码算法；保留输入，cleanup/decrypt越界。 | 信任验证另由调用方接入并分开报告；生产链、OCSP、时间戳及法律效力不在本票。 |
| 路径/版本 | archive忽略大小写；resolver不识别URI；两个namespace并存。 | 根/相对路径、BaseLoc覆盖、点段、盘符/UNC/转义/大小写固定样例；标准语义待核实，不能Resolve后碰巧Contains即Pass。 |
| 内存/性能 | 公开loader展开字节；建议预算不等价于受控路径累计内存证明。 | 真实包测峰值/延迟/取消/结果爆量；需要流式/所有权模式则另列实现切片，不声称已有指标。 |
| 只读与信息泄露 | archive可变数组、输入可并发变更、外链可能带凭据。 | 自有快照、禁止写探针、字节不变、无网络/宿主访问哨兵、证据限长脱敏共同验收。 |

## 后续实施切片

| 顺序与交付 | 最小范围（全部为建议） | 验收与停止边界 |
| --- | --- | --- |
| S0 调用方/规则目录确认 | 对六条本地政策确认需求、版本、字段范围、Unknown原因码；无生产实现。 | 获得真实接收政策或继续实验授权；无标准来源只能交本地预检，不因命名输出认证。 |
| S1 四态/只读壳 | Packaging内部预检模块、独立结果类型，复用loader/path；测试在现有Packaging.Tests。 | 可寻址/非可寻址/取消/坏ZIP/预算/DTD/NoFetch断言，未评估结果完整保留；不改现有ValidationReport/checker。预算无法在分配前执行时先补边界，再扩大样本。 |
| S2 001/002/004 | namespace-aware控制图、已知资源边与标记清点，不走Reader/Writer。 | 单/多文档、缺项、外链/根路径/逃逸、未知namespace、加密/孤立标记四态，报告可定位。先小种子后真实包。 |
| S3 003/005/006 | BaseLoc/字体载荷、附件和签名清点、字段所有权/作用域测试。 | 四态、重复ID/冲突/孤立项、零附件与非零字体差异、假SignedValue仅清点Pass；不新增验签/解密/字体修复。 |
| S4 报告/可选入口 | JSON序列化、必要的显式旧模型适配；经调用方确认再公开API/可选CLI。 | 全启用规则、范围/缺口、空Clause、Unknown/全NA不Passed、稳定排序；CLI落地后再定义真实命令及退出码，Warning Fail不当成功。公开XML文档及两个netstandard编译门齐全。 |
| S5 实际包/交接 | 确定性fixture、经授权的较干净/刻意违规真实包、消费示例与本地报告。 | 测试发现数非零、Packaging回归、同批包消费、输入不变/无外访证据、确认误报/Unknown；如有页面/转换变化按AGENTS.md视觉验收。标准认证与发布不自动完成。 |

各切片独立提交并按实际行为收证；实现后再根据真实源码/运行更新本文。当前设计完成只表示规则、样例、边界及未来门可供实施者审阅，不表示切片、整体产品、标准认证或发布已完成。
