---
title: "本地归档预检"
description: "六项本地归档预检规则、有限标准映射、只读覆盖检查和独立副本转换边界。"
sidebar:
  order: 170
  label: "归档预检"
---

`Ofdrw.Net.Packaging.Preflight.ArchivalPreflightChecker` 使用六项有限的本地规则检查现有 OFD ZIP。其配置为 `ofdrw-local-preflight/0.1-draft`，规则修订号为 1。它不是 OFD-H、OFD 或 GB/T 42133 符合性认证，也不是归档认证。标准映射尚未验证，每个 `Clause` 均为空。

```csharp
using Ofdrw.Net.Packaging.Preflight;

// 调用方提供稳定快照；可定位流应从位置零开始。
using var input = File.OpenRead("candidate.ofd");
var report = await new ArchivalPreflightChecker().CheckAsync(
    input, new ArchivalPreflightOptions(), cancellationToken);
var fontPresence = report.GetRuleOutcome("APF-003");
// 应用自己的策略前，检查 RunState、RuleResults 和 CoverageGaps。
```

检查器保持输入打开，并在成功、拒绝、取消或失败时将可定位输入恢复到位置零，且从不写入输入。不可定位流会从当前位置读入有界内存，不使用临时文件。调用方必须防止输入被并发修改。位置恢复失败和普通 I/O 失败会继续抛出；参数错误和取消也会继续抛出。格式错误或不安全的 ZIP 输入、预算耗尽会产生 `RunState=Incomplete`，所有未完成规则均为 Unknown，同时保留此前观察结果。此 API 不接受调用方持有的可变包归档。

加载器沿用 `LoadOptions` 的现有 ZIP 限制。其他默认值为：不同 XML 累计 32 MiB、每个 XML 8 MiB、读取深度 64、200,000 个 XML 节点、100,000 次引用访问和 10,000 条结果。至少需要十二个结果槽位，其中六个为中断诊断保留。XML 字节数、深度和节点数会在构建树之前检查。重复引用消耗引用和结果预算；不同 XML 的解析结果会缓存。这些是本地运行限制，不是标准限制，也不保证进程峰值内存。加载器的中央目录分配限制仍由现有公开加载器决定。

| 规则 | 适用的本地条件 | 重要限制 |
| --- | --- | --- |
| APF-001 | 恰好一个可识别的 DocBody、一个 DocRoot、可读取的 Document，以及至少一个存在的 Page 目标 | `RequireSingleDocument=false` 时此规则为 NotApplicable；其他规则仍遍历所有可识别文档。不验证页面绘制内容。 |
| APF-002 | 可识别的资源/索引声明、页面/模板 BaseLoc、资源 FontFile/MediaFile 均解析到包内且存在 | 只检查已知边。包含现有的相邻 PageRes.xml；缺少按约定命名的 PageRes.xml 不是错误。 |
| APF-003 | 每个可识别 Font 声明都有一个非空包内 FontFile；可见的已知 TextObject 字体 ID 能在受检资源作用域中解析 | 非空字节仍可能不是有效字体。不检查格式、字形覆盖、子集正确性、嵌入权限、许可和外观。重复 ID 或未建模资源作用域仍为 Unknown。 |
| APF-004 | 根 Encryptions.xml/PasswordCrypto 标记不存在，且受检控制文件未显示不支持的加密范围 | 发现根 Encryptions.xml 会使此规则失败，且不会解析或解密它；孤立 PasswordCrypto 和不可读/未知控制文件仍为 Unknown。`RequireNoEncryption=false` 时此规则为 NotApplicable，同时保留标记证据。 |
| APF-005 | 已声明的附件索引以及每个无歧义的 Attachment/FileLoc 均存在于包内 | 允许空载荷；空声明索引仍为 Unknown。支持单独的 FileLoc 属性；属性/元素冲突为 Unknown。External=true/1 会失败。孤立的 Attachments.xml 为 Unknown。 |
| APF-006 | 已声明的签名列表、Signature BaseLoc、SignedValue 和列出的 Reference FileRef 目标均可清点 | 空声明/列表或缺失目标会失败；格式错误或不支持的签名结构或变换为 Unknown。不验证哈希、签名值、身份、证书链、时间戳或法律效力。Fail 的严重级别为 Warning，但仍会使汇总结果为 Fail。 |

只识别 `http://www.ofdspec.org` 和 `http://www.ofdspec.org/2016`。子字段必须具有所属命名空间和预期父节点。未知控制结构、扩展、批注/组合作用域和孤立的约定资源会保守地报告为覆盖缺口。任意未知 XML 可能包含其他引用：此检查不是完整外链审计、XSD 验证器或资源 ID 闭包证明。

DocRoot 和 DocBody/Signatures 从包根解析；页面/模板、附件索引、签名 BaseLoc 和 SignedValue 从其所在 XML 解析。资源 BaseLoc 相对于其资源表解析，FontFile/MediaFile 相对于该基址解析。前导 `/` 覆盖为包根。签名 Reference/@FileRef 从包根解析。包条目查找沿用现有加载器的大小写不敏感语义；按大小写不敏感规则判定为重复的 ZIP 名称会被拒绝。这些本地路径策略尚未通过标准认证。

URI 方案、UNC 和双斜线路径会在包解析前分类，且从不抓取。查询、片段、百分号转义、控制字符和有歧义的反斜杠引用仍为 Unknown；根逃逸会失败。不调用网络客户端、宿主机字体发现、外部进程、写出器、密码或签名验证器。XML 禁止 DTD，解析器为空。外部或有歧义的位置元数据会被脱敏，避免暴露凭据或查询值；已知且受支持的命名空间选择器保留展开名称。未知命名空间字符串替换为 `<unsupported-namespace>`，避免复制不受信任 URI 中的凭据。载荷文本、字体字节、附件字节和签名字节从不复制到报告证据中。

`RuleResults` 保存各对象的观察结果，包括报告本地的结果和证据 ID、来源条目、选择器、稳定原因、观察值/预期元数据、严重级别、建议和本地来源。`GetRuleOutcome` 按 Fail、Unknown、Pass 的顺序聚合；没有适用观察结果时为 NotApplicable。`SummaryOutcome` 使用相同顺序，运行不完整时同样保持 Unknown。全部为 NotApplicable 时，汇总仍为 NotApplicable。确定失败可以与 Unknown 观察结果共存；零对象结果不能证明成功。`RunState=Complete` 表示检查已完成，即使不支持的语义产生 Unknown。作出任何应用决策时都应保留完整报告。

报告使用公开只读属性，可由调用方的序列化器序列化（例如带 JsonStringEnumConverter 的 System.Text.Json）。库不引入序列化或网络依赖。CLI 命令 `ofdrw preflight`（自 [0.2.0](https://whynpc9.github.io/ofdrw.net/releases/020/) 起可用）封装同一 API 并输出 `ofdrw-cli-preflight/1` JSON；它仍是本地预检，不是认证：

```sh
ofdrw preflight input.ofd                          # 人类可读摘要；退出码 0 通过、2 失败、3 未知/不完整
ofdrw preflight input.ofd --json > report.json     # 标准输出仅包含 JSON 报告
ofdrw preflight input.ofd --json --output report.json --allow-multiple-documents
```

现有 `ValidationReport` 和宽松的 `OfdPackageStructureChecker` 保持原有约定；不存在隐式的有损布尔适配。

可执行边界示例位于 `ArchivalPreflightTests`。其中四字节字体和 SignedValue 哨兵仅用于证明存在性，绝不可用作渲染或签名有效性证据。预检不生成或修改页面，因此页面渲染、PNG 检查和 Preview 验收不适用于此实现。

## 可选归档规则子集与独立副本

`ArchivalRuleService.CheckAsync` 在保持 `LocalPreflight` 六项 APF 结果不变的同时，增加两项独立且带标准映射的观察结果。映射目标为 **GB/T 42133-2022 第 6.2.2(a) 条**（移除 `Document.xml/Permissions`）和**第 6.2.2(b) 条**（移除 `Document.xml/VPreferences`）。版本和条款已对照[江汉大学图书馆托管的标准 PDF](https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf) 核对，其 SHA-256 为 `7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82`，位置为 PDF 第 8 页（印刷页第 4 页）。`VerifiedSubset` 仅描述此映射，不代表完整 GB/T 42133 符合性或认证；报告的 `Certification` 仍为 `NotAssessed`。其他条款尚未实现，也不能根据规则数量推定。GB/T 33190-2016 允许一般 OFD 包含 Permissions，因此其文档根条款不能作为归档移除该节点的依据。

新增检查要求可定位的稳定输入从位置零开始，并在结束后保持输入打开且回到起点。如果本地检查不完整、文档有歧义或外来命名空间影响作用域，新增规则返回 `Unknown`。出现可识别声明时为 `Fail`；在此有限规则下，声明不存在为 `Pass`。调用方必须分别检查 `LocalPreflight` 和新增结果。归档子集的 Pass 不能证明图像、字体、动作、签名、附件或所有资源引用均符合归档要求。

`ArchivalCopyConverter.ConvertAsync(inputPath, outputPath, options)` 是独立的显式操作。它只接受单文档、本地预检无问题且支持的输入；拒绝签名和疑似孤立签名载荷，并报告 `SignedInputRequiresResigning`；同时拒绝任一目标声明中的未知字段。内部处理器只从 `Document.xml` 移除受支持的 `Permissions` 和 `VPreferences` 节点。R16 类型化对象用于指导识别结构，但副本处理器不会通过通用写出器重写出页面或其他包载荷。签名输入不会被修改、静默去签或表示为编辑后仍然有效。无变更输入不会产生输出。

转换器对有界源文件创建快照，在目标目录写入临时 ZIP，重新运行本地和子集检查，并比较条目集合及每个非目标条目的字节后再发布。它还输出 `<output>.audit.xml`，记录源/输出 SHA-256、已变更和未变更条目哈希、移除的声明、条款、信息损失告知及签名处置。失败会删除临时文件。输出和审计路径必须事先不存在，输入保持不变。这些检查证明此操作逐字节保留现有页面、资源和附件载荷边；但不能证明原始输入在有限 APF 范围之外具有完整语义资源闭包。处理器会明确失败，而不是隐藏此边界。

```csharp
var options = new ArchivalPreflightOptions();
using (var input = File.OpenRead("source.ofd"))
{
    var observations = await new ArchivalRuleService().CheckAsync(input, options);
    // 检查 observations.LocalPreflight、observations.RuleResults 和 Scope。
}
var copy = await new ArchivalCopyConverter().ConvertAsync("source.ofd", "archive-copy.ofd", options);
// 同时保留 copy.OutputPath 和 copy.AuditPath。
```

## 按版本维护的覆盖清单

[GB/T 42133-2022 条款账本](https://whynpc9.github.io/ofdrw.net/reference/archival-standard-coverage/)跟踪第 5–7 章和规范性附录 A，包括仍未覆盖或需要外部证据的条款。`ArchivalCoverageService.CheckAsync` 增加可选的只读观察：`AC-001`–`AC-003` 覆盖归档 DocType、单文档和文档 Goto 类型；`AC-004`–`AC-006` 覆盖页面/大纲 Goto 类型和大纲目标是否存在；`AC-007` 检查已用图像插值；`AC-008` 检查文本 Size；`AC-009` 检查保留批注的三个标志；`AC-010` 检测保留的对象动作；`AC-011` 检查重复页面图层名称；`AC-012` 检查保留批注外观中的嵌套 PageBlock；`AC-013` 检测可识别的外部 URI 动作；`AC-014` 检测带目标的保留文档扩展声明；`AC-015` 检测保留的根加密标记。`AC-010`、`AC-013`、`AC-014` 和 `AC-015` 在没有违规时仍为 Unknown，因为无法从一个输出包证明来源迁移、外部资源闭包、所有扩展位置或完整明文。该服务保留不变的 R17 报告，不扩展其副本处理器，也不改变六项本地 APF 规则。

```csharp
using var file = File.OpenRead("candidate.ofd");
var coverage = await new ArchivalCoverageService().CheckAsync(file, new ArchivalPreflightOptions());
// 检查 coverage.RuleResults 和 coverage.FoundationalRules。
// 条款覆盖不完整时，coverage.StandardOutcome 不能为 Pass。
```

`CoverageStatus` 始终为 `Incomplete`，`Certification` 为 `NotAssessed`。确定的违规使 `StandardOutcome=Fail`；否则完整标准结果为 `Unknown`，即使所有已映射条件通过也是如此。输入被拒绝、不支持的根命名空间、无法解析的页面/批注引用和不透明 XML 作用域均为 Unknown。缺失或明确错误的已知强制声明为 Fail，即使存在同名的外来声明也不例外。PDF 来源和 SHA 固定在每个规则定义及账本中；包中不附带标准 PDF 字节。R17 的审计副本可以在此服务中分别检查转换前后，但该副本只移除其文档说明的两种声明类型。新增检查不会重写页面、批注或导航。

## F07 来源与证据档案（自 0.3.0 起）

`ArchivalDossierService.CheckAsync` 将两个稳定的 OFD 快照及外部证据编入同一份只读档案。`ArchivalClauseCatalog` 的 `ofdrw-archival-ledger/1` 枚举账本第 5–7 章的 92 个项目和附录 A 的 13 行；它关联现有 17 项 AR/AC 条件和两个已有副本处理器。APF 继续是独立本地政策。所有条款都会出现在报告中，包括没有对象、没有实现或没有提供材料的条款。

```csharp
using var source = File.OpenRead("source.ofd");
using var result = File.OpenRead("archive-copy.ofd");
var profile = new ArchivalDossierProfile(
    "my-archive/1", null, // 尚未指定归档机构
    "未签名单文档的有限归档范围",
    new Dictionary<string, string>()); // 提供经评估的引用标准版本/修改单；缺项会列出
var dossier = await new ArchivalDossierService().CheckAsync(
    source, result, profile, Array.Empty<ArchivalEvidenceInput>(),
    new ArchivalPreflightOptions(), cancellationToken: cancellationToken);
// 分别保存 SourceCoverage、ResultCoverage、Clauses 和 Evidence，不能仅检查摘要。
```

`ArchivalEvidenceInput` 指定唯一 ID、账本条款、证据类别、可定位流、预期 SHA-256、来源描述，以及材料所针对的源/结果 SHA-256。类别包括正例、反例、不确定例、资源验证、附件、转换、视觉复查、独立阅读器、政策、PKI、规范引用及独立核准。证据字节、源/结果身份必须同时匹配；旧来源的材料不能填补本次缺项。调用方应分别提供每条款所需材料；`MissingControlEvidence` 只列基础正/反/不确定例和核准文件的登记缺项，条款特定要求还须阅读 `RemainingEvidence`。摘要相符和填写标准版本不证明材料真实、充分、适用或核准有效。`ApprovalStatus=NotAuthenticated` 如实保留这一边界。

报告保存源/结果文件及证据的大小与摘要、全部条款、映射规则、处理器、来源描述及缺失引用标准。可用 System.Text.Json 序列化，只读属性不包含证据正文或流。`EvidenceIntegrity` 仅汇总证据摘要和源/结果身份匹配；无材料为 Unknown，错配为 Fail。它与标准结论独立。源文件的违规保留在 `SourceCoverage`，不会冒充输出文件仍有同一违规。结果中已知违规使 `StandardOutcome=Fail`，否则仍为 Unknown；全部材料登记完也不能自动变为完整符合性 Pass。不存在隐式适用豁免或认证。

默认累计读取预算为 1 GiB，包含源文件、结果及每次证据读取；最多 256 项证据，可设置更低限额（数量硬上限 10,000）。两个 OFD 各自还受 `LoadOptions.MaxInputBytes` 和原检查器预算约束，包快照保存在内存中；证据只流式计算摘要。枚举、读取和检查支持取消；超过预算、I/O 错误或取消抛出异常，不返回部分成功。所有借用流必须可读、可定位、从零开始，读取后恢复零并保持打开，调用方须防止并发修改。

此 API 不执行附件转换、物理删除、归档副本编辑、解密或签章去技术化，也不自动获取 URL。签名输入和 R18 删除所有权继续由原操作处理。完整包消费示例见 `scripts/validation/f07-dossier/`，设计和未完成范围见 [F07 实施记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/f07-archival-dossier.md)。

仅持有官方标准、尚未指定归档机构时，`ArchivalAuthority` 可为 null，无需填写占位机构。`MissingReferencedStandards` 列出未登记版本；`MissingExternalEvidence` 列出尚无非空且摘要/身份匹配材料的规范引用、接收政策及独立核准三类基础角色；`ExternalAssessmentGaps` 明确保留机构未指定、引用版本/政策未评估、核准未认证的状态。即使角色文件和版本声明均已登记，`ExternalAssessmentOutcome` 仍为 Unknown，条款专属材料仍按账本逐项核对。PKI 仅在适用签章情形补充，不假定所有未签名包都需要证书材料。
