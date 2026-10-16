# Plan 027: 先验证附件描述符的无操作往返边界

> 后续授权实施已完成并统一本地合并。实际交付见[027-followup-implementation.md](027-followup-implementation.md)，合并后验证见[最终验收](followup-integration-acceptance-2026-10-16.md)。下文保留最初调查/设计范围和当时状态。

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/027-attachment-investigation-result.md"`；再运行 `git status --short`、`git diff -- "plans/027-attachment-investigation-result.md"` 与 `git diff --cached -- "plans/027-attachment-investigation-result.md"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（调查CONFIRMED，修复未实施）
- 来源：I01
- 优先级：P2；工作量：S；修复/实施风险：LOW
- 类别：investigation
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（调查CONFIRMED；修复未实施）；冻结交付 `19e58deb`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。 调查/设计完成只表示本计划指定交付完成，不代表附件修复或新功能已经实现。

## 为什么要做

附件描述符未知信息往返尚未完成最小验证，不能由 CustomTags 的已确认缺陷推导同样结论。



## 当前代码与约定

[src/Ofdrw.Net.Reader/Readers/OfdReader.cs:410](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Readers/OfdReader.cs:410)

```csharp
                    var attachment = new OfdAttachment
                    {
                        Id = node.Attribute("ID")?.Value,
                        Name = node.Attribute("Name")?.Value ?? "Attachment",
                        MediaType = node.Attribute("MediaType")?.Value
                            ?? node.Attribute("Format")?.Value
                            ?? "application/octet-stream",
                        Format = node.Attribute("Format")?.Value,
```

[src/Ofdrw.Net.Packaging/OfdPackageWriter.cs:1188](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging/OfdPackageWriter.cs:1188)

```csharp
        var attachments = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Attachments",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                package.Attachments.Select((attachment, index) =>
                {
                    var fileName = $"Attach_{index + 1}_{SanitizeFileName(attachment.Name)}";
                    if (!attachment.IsExternal)
```

其余核查入口：无额外入口。

**沿用范例**：[tests/Ofdrw.Net.Packaging.Tests/PackageRoundTripTests.cs:19](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Packaging.Tests/PackageRoundTripTests.cs:19)

```csharp
    [Fact]
    public async Task Writer_ShouldGenerateRequiredEntries_AndReaderCanParse()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `plans/027-attachment-investigation-result.md`

此外可写本次任务的 `artifacts/plan-027/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。本计划不修改生产实现；临时探针只能写任务 artifacts。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-027-attachment-preservation-investigation`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-027/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-027/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

在任务自有临时探针中构造包含未知属性/子节点、多命名空间、原始位置、外部附件和 payload 的小包；通过公开 Reader→Writer 无修改保存，比较 descriptor 与载荷的语义/字节差异。只确认有明确契约的部分，保留原输入。

**验证**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "plans/027-attachment-investigation-result.md"` → 已确认基线；在指定交付文件写齐样例、证据和预期，不依赖会话上下文。

### 2. 形成可审阅闭环

增加对照：修改一个建模字段、显式移除附件、新建附件；追踪模型来源与 Writer 重建路径。输出最小复现、实际失败/成功范围、删除语义与是否需要修复的结论；确认后在本计划内追加精确修复范围供维护者审阅，当前不改生产实现。

**验证**：完成下列样例矩阵，并执行文档检查命令；缺失事实或未判断项逐项标明，不把设计完成写成实现通过。

### 3. 记录完成边界

交付结果为 CONFIRMED / REJECTED / INCONCLUSIVE，并附命令、输入、实际差异与可复查路径；确认缺陷后再给最小修复建议，本轮不修改生产源码。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `OpaqueDescriptorNoOp`：未知属性/子节点和 FileLoc 是否保留
- `PayloadNoOp`：载荷字节不变
- `ExplicitRemoval`：删除动作不会被 opaque 恢复
- `ExternalAttachment`：不发网络请求、不擅自内嵌外部资源

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。


## 可重复调查探针

探针仅放 `artifacts/plan-027/probe/`。创建 `Probe.csproj`：SDK Microsoft.NET.Sdk，OutputType Exe，TargetFramework net10.0；ProjectReference 指向 `../../../src/Ofdrw.Net.Reader/Ofdrw.Net.Reader.csproj` 与 `../../../src/Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj`。Program.cs 构造上述 ZIP/XML 样例，经公开 API 往返，输出逐项 JSON/Markdown 差异并保存输入/输出包，不能访问网络附件。

先按上面的环境设置，再执行：

```bash
dotnet restore artifacts/plan-027/probe/Probe.csproj --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build artifacts/plan-027/probe/Probe.csproj -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project artifacts/plan-027/probe/Probe.csproj -c Release --no-build --no-restore
```

预期：探针退出 0 表示调查完整运行，输出必须逐项写 MATCH / LOSS / UNDECIDABLE；LOSS 是调查证据，不是允许掩盖的测试成功。若 API/构建失败则先区分环境问题，不能据此认定附件缺陷。交付 `plans/027-attachment-investigation-result.md` 与这些本地产物路径；不提交临时 probe。


## Done criteria

- [ ] 最小探针可复跑且逐项有结论，未确认项保持未知。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 没有可复现差异或契约无法证明时标 INCONCLUSIVE/REJECTED，不把 F08 直接套成附件缺陷。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

结果应成为后续独立修复的输入；调查完成不等于问题已修复。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
