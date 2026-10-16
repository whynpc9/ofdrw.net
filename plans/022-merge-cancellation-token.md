# Plan 022: 让带诊断的合并入口传播取消

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Packaging.Tests/DocumentToolTests.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Packaging.Tests/DocumentToolTests.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Packaging.Tests/DocumentToolTests.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE
- 来源：F22
- 优先级：P2；工作量：S；修复/实施风险：LOW
- 类别：正确性/响应性
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（独立库4项、CLI2项）；冻结交付 `1d7e8ff7`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。

## 为什么要做

F22：CLI 选择带诊断的 MergeWithResult，但该重载把 default token 传给 MergeCore。Ctrl-C 必须等整个模型合并结束后才能到达 writer 的取消点。



## 当前代码与约定

[src/Ofdrw.Net.Cli/Program.cs:395](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Cli/Program.cs:395)

```csharp
        var merged = OfdDocumentMerger.MergeWithResult(
            sources,
            new OfdDocumentMergeOptions
            {
                SkipUnsupportedRawElements = skipUnsupported
            });
        await using var output = new AtomicOutput(outputPath);
        foreach (var diagnostic in merged.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
```

[src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs:43](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs:43)

```csharp
    public static OfdDocumentMergeResult MergeWithResult(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options = null) => MergeCore(sources, options, default);

    internal static OfdDocumentPackage MergeForSplit(OfdDocumentPackage source, OfdModelMaterialization material, CancellationToken token)
        => MergeCore(new[] { source }, null, token, material).Package;

    internal static OfdDocumentPackage MergeForMix(IEnumerable<OfdDocumentPackage> sources, OfdDocumentMergeOptions options, OfdModelMaterialization material, CancellationToken token)
```

其余核查入口：[src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs:67](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs:67)。

**沿用范例**：[tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs:15](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs:15)

```csharp
    [Fact]
    public async Task Commands_UseOneBasedPagesAndAtomicInPlaceOutput()
    {
        var input = PathFor("source.ofd"); var output = PathFor("output.ofd");
        var source = new OfdDocumentPackage();
        for (var i = 0; i < 2; i++)
        {
            var page = new OfdPage { Index = i, WidthMillimeters = 100 + i, HeightMillimeters = 150 };
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs`
- `src/Ofdrw.Net.Cli/Program.cs`
- `tests/Ofdrw.Net.Packaging.Tests/DocumentToolTests.cs`
- `tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs`
- `tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-022/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-022-merge-cancellation-token`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-022/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit022RegressionTests" --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit022RegressionTests" --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 让带诊断的合并入口传播取消

保留现有 MergeWithResult(IEnumerable, options=null) 二进制/源码签名，新增明确的三参数 token 重载并委托现有 MergeCore。不要增加导致调用歧义的可选重载组合。CLI 传递自身 cancellationToken。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit022RegressionTests" --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

利用现有 MergeCore 检查点测试在 sources 枚举/资源遍历中取消，保证 diagnostics 行为不变，stage/提交不发生。保持 typed ownership/unsupported 拒绝策略。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit022RegressionTests" --logger trx --results-directory artifacts/plan-022/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：上列受影响项目测试与 `git diff --check` 退出 0；实际测试数非零、新增用例全部执行。

## 回归/样例矩阵

- `Audit022_PreCanceled`：预取消立即传播
- `Audit022_CancelBetweenSources`：枚举第二个来源前触发取消，不继续访问
- `Audit022_DiagnosticsControl`：未取消时 skipUnsupported 诊断保留
- `Audit022_OldOverload`：既有调用形式继续编译执行
- `Audit022_CliNoPublication`：取消返回130且既有目标不变

主测试入口为 `tests/Ofdrw.Net.Packaging.Tests/Audit022RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。



## Done criteria

- [ ] 新用例编译/被发现，缺陷修复后全部通过，既有项目回归通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 若需要改变 Merge 对 typed annotation/navigation 的拒绝范围，停止；它与取消传播无关。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

新增 with-result API 应与不带诊断入口同样接受并传递 token。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
