# Plan 012: 限制验签累计工作并复用摘要

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationOptions.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationBudget.cs" "docs/tutorials/10-signatures.md" "tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationOptions.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationBudget.cs" "docs/tutorials/10-signatures.md" "tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationOptions.cs" "src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationBudget.cs" "docs/tutorials/10-signatures.md" "tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE
- 来源：F14
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：安全/性能
- 依赖：[001](/Users/wanghongyi/Projects/ofdrw.net/plans/001-signature-completeness.md)
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（独立验签40项、CLI78项）；冻结交付 `a0fa1da3`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。

## 为什么要做

F14：合法 ZIP 预算内可有大量重复引用，对同一大条目反复计算摘要；SM3 还复制完整填充数组。Reference 循环缺少取消检查；未注册密码插件也先计算这些摘要。



## 当前代码与约定

[src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:114](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:114)

```csharp
        foreach (var record in list
            .Descendants()
            .Where(element => element.Name.LocalName == "Signature"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var baseLocation = record.Attribute("BaseLoc")?.Value;
            var result = new OfdSignatureVerificationResult
            {
```

[src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:182](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:182)

```csharp
        foreach (var reference in references?
                     .Elements()
                     .Where(element => element.Name.LocalName == "Reference") ??
                 Enumerable.Empty<XElement>())
        {
            var fileReference = reference.Attribute("FileRef")?.Value ?? string.Empty;
            var referenceResult = new OfdReferenceVerificationResult
            {
```

其余核查入口：[src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:224](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:224)；[src/Ofdrw.Net.Signatures/Crypto/Sm3Digest.cs:105](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Signatures/Crypto/Sm3Digest.cs:105)。

**沿用范例**：[tests/Ofdrw.Net.Signatures.Tests/SignatureTests.cs:27](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Signatures.Tests/SignatureTests.cs:27)

```csharp
    public async Task Verifier_ShouldValidateUpstreamSm3References()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs`
- `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationModels.cs`
- `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationOptions.cs`
- `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerificationBudget.cs`
- `docs/tutorials/10-signatures.md`
- `tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-012/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-012-signature-verification-work-budget`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-012/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --filter "FullyQualifiedName~Audit012RegressionTests" --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

安全复现优先使用小输入、小额度和可控 internal 接缝。新选项或注入点尚不存在时，先添加兼容的最小接缝，或将依赖新 API 的测试安排在步骤2后；不能把编译失败当重现。无法安全运行的旧行为以已核实静态路径记录，不强制执行巨量分配/慢循环以求红灯。

在 `tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --filter "FullyQualifiedName~Audit012RegressionTests" --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 可安全运行的用例被发现且正常控制通过，缺陷断言或静态未执行依据明确；依赖新 API 的用例列为后置验证。记录后再修复。

### 2. 限制验签累计工作并复用摘要

在保持现有构造/VerifyAsync 签名可用的前提下增加 options 入口，定义签名记录数、引用访问数、累计实际散列字节的正值上限。缓存只在一次 VerifyAsync 内按规范化条目身份+算法复用；每个引用访问仍计数，cache 命中不重复收取实际散列字节。默认值在现有签名语料上统计后定，记录合理余量。

**验证**：`dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --filter "FullyQualifiedName~Audit012RegressionTests" --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

在循环和 Compute 前检查 token/扣预算；预算耗尽产生明确不完整失败且不能经 F01 聚合变成有效。先排除调用方 OperationCanceledException 再捕获普通插件/摘要错误，避免取消被吞。缓存不得跨输入包或信任上下文。

**验证**：`dotnet test tests/Ofdrw.Net.Signatures.Tests/Ofdrw.Net.Signatures.Tests.csproj -c Release --filter "FullyQualifiedName~Audit012RegressionTests" --logger trx --results-directory artifacts/plan-012/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：上列受影响项目测试与 `git diff --check` 退出 0；实际测试数非零、新增用例全部执行。

## 回归/样例矩阵

- `Audit012_ReferenceVisitBudget`：重复引用仍受访问次数限制
- `Audit012_DigestByteBudget`：新条目在散列前扣字节，缓存命中不重复散列
- `Audit012_SignatureRecordBudget`：重复签名记录受总数约束
- `Audit012_CancellationPropagation`：引用循环和插件取消抛给调用方
- `Audit012_UnsupportedStillBounded`：未注册算法插件的路径也受限
- `Audit012_ValidBoundary`：恰好额度与正常签名通过，超额报告不完整

主测试入口为 `tests/Ofdrw.Net.Signatures.Tests/Audit012RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。



## Done criteria

- [ ] 新用例编译/被发现，缺陷修复后全部通过，既有项目回归通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 如果实现要求改写 SM3 算法或全局缓存摘要，停止；本票仅编排/额度/取消。不得为了复用结果跳过各引用 CheckValue 比较。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

任何新增摘要算法必须有规范化缓存键；新 warning 不得让预算失败被聚合为成功。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
