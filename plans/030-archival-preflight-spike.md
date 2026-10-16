# Plan 030: 界定可解释的只读归档预检

> 后续授权实施已完成并统一本地合并。实际交付见[030-followup-implementation.md](030-followup-implementation.md)，合并后验证见[最终验收](followup-integration-acceptance-2026-10-16.md)。下文保留最初调查/设计范围和当时状态。

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/030-archival-preflight-design.md"`；再运行 `git status --short`、`git diff -- "plans/030-archival-preflight-design.md"` 与 `git diff --cached -- "plans/030-archival-preflight-design.md"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（仅设计）
- 来源：D03
- 优先级：P3；工作量：L；修复/实施风险：MED
- 类别：direction
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（仅设计；19链接复核）；冻结交付 `fe01184f`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。 调查/设计完成只表示本计划指定交付完成，不代表附件修复或新功能已经实现。

## 为什么要做

当前规则模型与结构检查器提供可复用基础；归档方向需要可解释规则和真实需求，不能以预检代替标准认证。



## 当前代码与约定

[docs/capability-roadmap.md:126](/Users/wanghongyi/Projects/ofdrw.net/docs/capability-roadmap.md:126)

```
| P2-05 | OFD-A 检查器 | 对标 ofdrw-archive 的规则引擎（GB/T 42133） | 先做只读检查与违规报告；转换管道另议 | 宣称档案合规认证 |
| P2-06 | GM/T 0099 证书加密、完整性协议、多重加密 | 仅当 P2-01 稳定且有明确调用方 | 有互操作样本 | 无证书来源时用自签冒充 PKI |

生产验章、证书链、OCSP、合格时间戳、制章系统 **不设 TODO ID**。需要时通过已有扩展点由调用方接入，不在本仓库排期。

## 7. 工程债（与功能并行）

来自对照表里「进入生产评估前」仍有效的项，不挡 P0 开发，但发布前要过：
```

[docs/capability-roadmap-development-plan.md:99](/Users/wanghongyi/Projects/ofdrw.net/docs/capability-roadmap-development-plan.md:99)

```
| [14 OFD-A 检查器](../.scratch/capability-roadmap/issues/14-ofd-a-checker.md) | `archive/check/OFDArchiveChecker.java`、`ArchiveRule.java`、各 `rule/*` | 独立只读规则集，先列可静态判断的单文档、外部资源、字体嵌入、加密、附件/签名范围等规则及适用条件；每项返回规则 ID、严重度、包内路径和证据。坏包与故意违规/较干净样本出结构化报告；不修改包、不改变转换默认、不宣称 GB/T 42133 认证。 |
| [15 发布闸门](../.scratch/capability-roadmap/issues/15-release-engineering-gates.md) | 上游 `ofdrw-archive` 可作规则样本参考；本票以本仓库工程链路为准 | Q-01 许可/第三方声明清单及同批 NuGet 消费；Q-02 票据、发票、模板、异常包像素差阈值；Q-03 固定坏包/模糊输入与资源预算；Q-04 新增公开 API 的 CS1591 门禁及旧债清理顺序；Q-05 Native OFD→PDF→Preview 的记录模板。五项全做才可宣称发布闸门完成。 |

## 5. 各批次验收与发布判定

### 5.1 测试矩阵

| 层级 | 必查内容 | 证据 |
```

其余核查入口：[src/Ofdrw.Net.Core/Validation/ValidationReport.cs:12](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Validation/ValidationReport.cs:12)；[src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs:12](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs:12)。

**沿用范例**：[src/Ofdrw.Net.Core/Validation/ValidationReport.cs:12](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Validation/ValidationReport.cs:12)

```csharp
public sealed class ValidationFinding
{
    public string RuleId { get; set; } = string.Empty;

    public string Clause { get; set; } = string.Empty;

    public ValidationSeverity Severity { get; set; }

```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `plans/030-archival-preflight-design.md`

此外可写本次任务的 `artifacts/plan-030/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。本计划不修改生产实现；临时探针只能写任务 artifacts。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-030-archival-preflight-spike`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-030/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

从已有 ValidationFinding/PackageStructureChecker 设计有限规则集：单文档结构、外部资源、可判定字体嵌入、加密/附件/签名存在性。逐条给 RuleId、严重度、适用条件、包内位置和证据；未知/无法判断必须独立于通过。

**验证**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/030-archival-preflight-design.md"` → 已确认基线；在指定交付文件写齐样例、证据和预期，不依赖会话上下文。

### 2. 形成可审阅闭环

用正常与刻意违规的小型包定义报告样例与只读不改包断言；记录标准条款来源待核实项和真实调用方需求，拆出后续可实现 MVP。仅设计 spike，不宣称 GB/T 42133 认证，不新增自动修复/转换。

**验证**：完成下列样例矩阵，并执行文档检查命令；缺失事实或未判断项逐项标明，不把设计完成写成实现通过。

### 3. 记录完成边界

状态更新只表示设计已交付；生产实现、功能测试、视觉验收和发布均未完成。可选方向不阻塞修复批次。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `RuleMatrix`：每规则有通过/失败/不适用/无法判断示例
- `ReadOnly`：输入包保持相同字节且不抓取外链
- `EvidenceLocation`：所有失败都有规则ID和可定位证据
- `NoCertification`：报告明确范围和未判断项，不输出认证结论

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。


## 设计交付检查

交付文件必须有以下标题，各节填写实质内容而非占位：`## API 与边界`、`## 现有能力复用`、`## 样例与预期`、`## 验收命令`、`## 风险与未决项`、`## 后续实施切片`。

```bash
python3 - <<'PY'
from pathlib import Path
p = Path('plans/030-archival-preflight-design.md')
s = p.read_text()
for heading in ['## API 与边界','## 现有能力复用','## 样例与预期','## 验收命令','## 风险与未决项','## 后续实施切片']:
    assert heading in s, heading
print('design structure OK; still requires content review')
PY
```

这是文档结构检查，不能代替对契约与样例的人工审阅。设计里的未来验证命令必须标为未执行。


## Done criteria

- [ ] 设计具备 API/边界/样例/验收/切片，未把设计写成产品已实现。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 若需要收费/未提供标准才能确定条款，不编造条款号；记录缺失来源。生产证书链/OCSP/法律效力不在本计划。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

每个新增规则维护可观察证据和标准版本；预检通过不等于档案认证。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
