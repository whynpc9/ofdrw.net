---
title: "归档评估（只读）"
description: "可选包 Ofdrw.Net.Archival 的只读 GB/T 42133-2022 评估：能评估什么、不能证明什么，以及如何解读 Unknown、OutOfScope 和 advisory。"
sidebar:
  order: 175
  label: "归档评估"
---

`Ofdrw.Net.Archival` 是 main 分支上新增的可选包（0.4.0 之后，**尚未发布到 NuGet**）。它组合现有的[本地归档预检](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/)和 AR/AC 条款检查，对一个 OFD 快照给出分维度的只读报告。

先说清三件事：

- **没有整部标准的 Pass。** 文档整体结果只会是 `Fail` 或 `Unknown`，也没有 `IsCompliant` 一类的成员。
- **本包不生成 OFD-A。** 它不写、不修复、不改标签，也不删除任何内容。
- **第 7 章的阅读器和写入端行为没有被证明。** 单个文件证明不了软件行为。

## 引用与调用

普通转换包不会带入本包，需要显式引用。在发布到 NuGet 之前，从源码以项目引用使用（路径按实际位置调整）：

```xml
<ProjectReference Include="src/Ofdrw.Net.Archival/Ofdrw.Net.Archival.csproj" />
```

```csharp
using Ofdrw.Net.Archival;

await using var input = File.OpenRead("candidate.ofd");
var report = await new ArchivalAssessmentService().AssessAsync(
    input, new ArchivalAssessmentOptions(), cancellationToken);

if (report.ExecutionState != ArchivalExecutionState.Complete)
{
    // 预算耗尽或输入被拒：报告不可用于交付判断。
}
```

依赖方向为 `Archival → Packaging → Core`。`Ofdrw.Net.Converter`、Layout、Reader 和 CLI 都不引用它；发布校验会拒绝任何把它带入传递依赖的包。

输入须可读、可定位且位于零点。服务只读不写，保持流打开，结束后把位置恢复到零。它先在 `LoadOptions.MaxInputBytes` 内读入一份快照并计算 SHA-256，之后的所有检查都基于这份字节，`report.Input` 记录的就是被检查字节的长度和摘要。预算沿用 `ArchivalPreflightOptions`。取消会抛出 `OperationCanceledException`，不返回部分报告，与现有预检服务一致。

## 报告的七个维度

| 维度 | 成员 | 说明 |
| --- | --- | --- |
| 声明 | `Declaration` | 原始 `OFD.xml` 根上的 `DocType`：`Archival`、`NotArchival`、`Missing`、`Ambiguous` 或 `Unavailable`。结果直接来自 AC-001，不使用 Reader 的缺省回退。 |
| 输入支持 | `InputSupport` | 所选 profile，以及 `Supported`、`OutOfScope` 或 `Unknown` 和逐项原因。 |
| 条款 | `Clauses` | 账本 105 项，按账本顺序，每项带规范强度、适用类别、映射规则观察、结果、证据位置和剩余证据。 |
| 文档整体 | `Document` | 有检查和无检查的条款、命名缺口、标准版本和来源哈希；`StandardOutcome` 只为 `Fail` 或 `Unknown`。 |
| 软件要求 | `SoftwareRequirements` | 第 7 章，固定为 `NotAssessed`。 |
| 外部认可 | `ExternalRecognition` | 归档机构政策、引用标准版本、PKI、独立核准，固定为 `NotAuthenticated`。 |
| 执行状态 | `ExecutionState`、`ExecutionReason` | `Complete`、`BudgetExhausted` 或 `Error`，以及评估、profile、账本和矩阵修订号。 |

底层 `ArchivalCoverageReport` 原样保留在 `report.Coverage`。六项 APF 仍是本地策略，不映射到条款。

## 首批输入 profile

profile 为 `ofdrw-archival-profile/self-generated-simple/1`：单一 DocBody，未签名，未加密，无附件，无 Versions，无动作、批注、扩展或 CompositeObject；对应简单 Layout 输出或使用受控字体的 Native DOCX→OFD，保留来源 CustomTags。

是否在范围内只看包内容，不因“由本库生成”而信任（`ProvenanceVerified` 恒为 false）。A1a 版本还没有检查字体载荷和 CustomTags 引用闭包，报告分别以 `FontPayloadNotAssessed` 和 `CustomTagClosureNotAssessed` 缺口标出。

## 能评估什么

- 声明：原始 `DocType` 是否为 OFD-A、缺失还是有歧义。
- 现有 17 项 AR/AC 有界检查，映射到 17 个条款，例如单一 DocBody（6.2.1(c)）、Permissions/VPreferences（6.2.2(a)(b)）、文本 `Size`（6.6(c)）、图像插值（6.5(b)）、动作类型和目标、批注标志、根加密标记。
- 输入是否在首批 profile 内，以及原因的位置。
- 每个条款的规范强度和适用类别，来自内置的条款适用性矩阵（`ofdrw-archival-matrix/1`）。

## 不能证明什么

- 整部标准符合。其余 88 个条款没有可执行检查，保持 `Unknown`。
- 字体是否有效、是否覆盖所用字形（计划在 A1c）；CustomTags 引用是否闭合（计划在 A1b）。
- 来源到结果的处理是否正确，例如动作迁移、版本保留、扩展外观冻结。
- 阅读器能完整显示、报告未处理项或禁止编辑 OFD-A（第 7 章）。
- 机构是否接收、PKI 是否有效、核准是否真实。

## 如何解读结果

**条款 `Unknown`。** 证据不足：要么没有可执行检查，要么检查范围不完整，要么运行没有完成。对象不存在不会变成 `Pass`。

**条款 `Pass`。** 只表示映射到该条款的有界检查成立；`RemainingEvidence` 仍然有效，条款并未因此闭合。

**条款 `NotApplicable`。** 规则在完整作用域内证明没有适用对象。

**`OutOfScope`。** 输入包含首批 profile 排除的内容，例如多个 DocBody、签名、加密标记、附件或 Versions。这不是违反国标，本包不会因此生成条款 Fail。但已有规则在这类输入上发现的确定违规照常保留：多个 DocBody 使 6.2.1(c) 为 Fail，根 `Encryptions.xml` 使 6.16 为 Fail。

**输入支持 `Unknown`。** 存在无法识别的结构（如 Document 下的未知元素、外来命名空间）或无法解析的页面，或者运行没有完成。

**advisory。** 标准正文没有“应”：矩阵记为 68 项陈述（`Statement`）、24 项“宜”（`Should`）、9 项“可”（`May`）、4 项“应”（`Shall`，均在附录 A）。`Statement` 和 `Shall` 条款上的规则 Fail 保持 Fail；`Should` 和 `May` 条款即使检查未满足也不产生条款 Fail，也不使 `StandardOutcome` 为 Fail，未满足的观察放在该条款的 `Advisories` 中，条款结果为 `Unknown`。每个条款都公开 `Strength`，需要更严或更宽的策略时由调用方自行处理。现有 17 项检查都映射到陈述类条款，所以当前报告中不会出现 advisory。

**矩阵未确认项。** 6.11、6.13.1(a)–(c) 的情态词作用范围有不止一种读法，`Confidence` 为 `Unverified`，并带 `ModalScopeUnverified` 注记。

**批注前提。** AC-009、AC-012 假定批注以注释对象保留，即 6.10.1 的方式 a；如果改用方式 b，把批注外观合并到页面，这两项检查不适用。对应条款带 `AnnotationObjectRetentionAssumed` 注记。

**执行状态。** 只有 `Complete` 表示检查做完。`BudgetExhausted`（如 `InputBudgetExceeded`、`InspectionBudgetExceeded`）和 `Error`（`InputRejected`）的报告中，除确定的 Fail 外所有条款都是 `Unknown`，输入支持带 `AssessmentIncomplete`。字节预算在读入前就超出时 `Input` 为 null。

## 报告示例

以下摘自对一份 `FlowDocument` 生成的普通 OFD 实际运行的报告（Linux WSL2、.NET SDK 10.0.401，main 开发版本，System.Text.Json 序列化，枚举以字符串输出）。OFD 内含创建时间，每次生成的摘要不同。为节省篇幅，各片段省略了部分字段（如 `Profile` 只保留 `Id`）。

```json
{
  "AssessmentRevision": "ofdrw-archival-assessment/1",
  "LedgerRevision": "ofdrw-archival-ledger/1",
  "MatrixRevision": "ofdrw-archival-matrix/1",
  "ExecutionState": "Complete",
  "ExecutionReason": "Complete",
  "Input": { "Length": 1337, "Sha256": "4326983c3360aa27659061d101dcd5026fe9e2999dd07fefbd5b9cdf04a5d2cd" },
  "Declaration": {
    "Status": "NotArchival",
    "RawDocType": "OFD",
    "ForeignDocTypePresent": false,
    "RuleId": "AC-001",
    "Outcome": "Fail",
    "ReasonCode": "NonArchivalDocType"
  },
  "InputSupport": { "Profile": { "Id": "ofdrw-archival-profile/self-generated-simple/1" }, "Support": "Supported", "Findings": [] }
}
```

其中一个有检查的条款（省略了 `Definition`、`Matrix`、重复的 `Evidence` 和部分证据字段）：

```json
{
  "Clause": "6.6(c)",
  "Strength": "Statement",
  "Conditional": false,
  "Applicability": "PackageDecidable",
  "Confidence": "Verified",
  "HasExecutableCheck": true,
  "Observations": [
    {
      "RuleId": "AC-008",
      "RuleOwner": "Ofdrw.Net.Packaging",
      "Outcome": "Pass",
      "ReasonCode": "ConditionSatisfied",
      "Evidence": { "Entry": "Doc_0/Pages/Page_0/Content.xml", "Observed": "Recognized objects=1; ConditionSatisfied" }
    }
  ],
  "Advisories": [],
  "Outcome": "Pass",
  "RemainingEvidence": "**精确的只读属性检查：**`AC-008` 要求可识别页面/模板和保留批注外观内容中的文本对象具有正的无限定 `Size`。不透明图形作用域为 Unknown；实际字形几何仍需外部证据。",
  "Notes": []
}
```

文档整体：

```json
{
  "StandardOutcome": "Fail",
  "CoverageStatus": "Incomplete",
  "Certification": "NotAssessed",
  "Gaps": [
    "NoFullStandardAssessment", "ReferencedEditionsNotFixed",
    "MatrixScopeUnverified:6.11,6.13.1(a),6.13.1(b),6.13.1(c)",
    "LocalPreflightRulesNotMappedToClauses", "CustomTagClosureNotAssessed", "FontPayloadNotAssessed"
  ]
}
```

这份文件的 105 个条款中：1 个 `Fail`（6.2.1(a)，DocType 为 OFD），5 个 `Pass`，7 个 `NotApplicable`，92 个 `Unknown`。同一主机上用默认 Native 模式转换 `generated-layout.docx` 得到的结果相同：输入在 profile 内，条款计数一致。只把根标签改为 OFD-A 后，声明变为 `Archival`，6.2.1(a) 变为 `Pass`，其余条款不变，`StandardOutcome` 变为 `Unknown`，而不是 Pass：标签本身不改变任何实质结论。

## 相关文档

- [GB/T 42133-2022 归档覆盖账本](https://whynpc9.github.io/ofdrw.net/reference/archival-standard-coverage/)：105 项条款、映射规则与剩余证据。
- [本地归档预检](https://whynpc9.github.io/ofdrw.net/guides/archival-preflight/)：APF、AR、AC 和证据档案的原有接口。
- [设计记录](https://github.com/whynpc9/ofdrw.net/blob/main/docs/architecture/archival-assessment-package.md)：依赖方向、强度策略、AC-016 起的编号归属和字体验证接口的决定。
- [条款强度与适用性矩阵](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/baselines/gbt42133-2022-clause-applicability.md)。
