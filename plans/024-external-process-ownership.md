# Plan 024: 在包装进程退出后仍清理其子进程

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs" "src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs" "tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs" "tests/Ofdrw.Net.ProcessFixture/Program.cs" ".github/workflows/ci.yml" "plans/024-process-ownership-design.md" "tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs" "src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs" "tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs" "tests/Ofdrw.Net.ProcessFixture/Program.cs" ".github/workflows/ci.yml" "plans/024-process-ownership-design.md" "tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs" "src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs" "tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs" "tests/Ofdrw.Net.ProcessFixture/Program.cs" ".github/workflows/ci.yml" "plans/024-process-ownership-design.md" "tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（已验证64位三平台，Unix32明确拒绝）
- 来源：F24
- 优先级：P2；工作量：L；修复/实施风险：MED
- 类别：正确性/进程资源
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（独立15项；三平台实测）；冻结交付 `c5520dd3`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。

## 为什么要做

F24：Runner 明确等父进程退出后仍未关闭的继承管道，但 KillTree 看到父进程退出立即返回。调用方超时/取消后，后代仍可能占用资源或写文件；现有 fixture 父进程一直等子进程。



## 当前代码与约定

[src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:61](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:61)

```csharp
            // Include pipe completion in the deadline: a descendant may inherit
            // the handles even when its parent has already exited.
            while (!process.HasExited || !pipes.IsCompleted)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (elapsed.Elapsed >= timeout)
                {
                    throw new TimeoutException($"{startInfo.FileName} exceeded {timeout}.");
```

[src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:99](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:99)

```csharp
            KillTree(process);
            // Termination is bounded. Closing the process/streams below releases
            // inherited pipes even if an external program cannot be cleaned up.
            process.WaitForExit(5000);
            try
            {
                if (await Task.WhenAny(pipes, Task.Delay(1000)).ConfigureAwait(false) == pipes)
                {
```

其余核查入口：[src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:140](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs:140)；[tests/Ofdrw.Net.ProcessFixture/Program.cs:18](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.ProcessFixture/Program.cs:18)。

**沿用范例**：[tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs:22](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs:22)

```csharp
    public async Task Runner_ShouldCancelStartedProcessAndDescendant()
    {
        var pidFile = Path.Combine(_directory, "pid");
        using var cancellation = new CancellationTokenSource();
        var running = ExternalProcessRunner.RunAsync(CreateStart("tree", pidFile, RuntimeConfig),
            TimeSpan.FromSeconds(20), cancellation.Token);
        await WaitForFileAsync(pidFile + ".child", running);
        var parent = int.Parse(await File.ReadAllTextAsync(pidFile));
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs`
- `src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs`
- `tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs`
- `tests/Ofdrw.Net.ProcessFixture/Program.cs`
- `.github/workflows/ci.yml`
- `plans/024-process-ownership-design.md`
- `tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-024/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-024-external-process-ownership`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-024/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~Audit024RegressionTests" --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

先扩展 ProcessFixture：父进程生成写入自有 PID 文件的子进程后退出，子进程保留继承管道。测试 finally 仅清理本 fixture 明确拥有的进程，避免预期红灯遗留进程；随后在 `tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs` 建立下面的用例。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~Audit024RegressionTests" --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 在包装进程退出后仍清理其子进程

基于步骤1的安全失败证据，写跨平台所有权设计：Unix group/session 与 Windows Job 的启动竞争、已有 Office wrapper 和 netstandard 限制；用设计记录选择可实现策略。

**验证**：`dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~Audit024RegressionTests" --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 此阶段不要求转绿：回归仍按基线失败且 fixture 被安全清理；另运行 `test -s plans/024-process-ownership-design.md`，确认设计记录包含平台策略、启动竞争、清理边界和兼容性，完成设计复核后才进入步骤3。

### 3. 接入调用链并验证兼容行为

在启动时取得持久所有权，取消/超时对 owned group/job 有界终止及收尾，即使初始 Process.HasExited=true。保持正常退出、stdout/stderr 有界捕获和清理异常诊断；CI 在 Windows/macOS/Linux 原生平台执行联合过滤器 `FullyQualifiedName~ExternalProcessTests|FullyQualifiedName~Audit024RegressionTests`，不能扫描杀掉同名应用。

**验证**：`dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~Audit024RegressionTests" --logger trx --results-directory artifacts/plan-024/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：上列受影响项目测试与 `git diff --check` 退出 0；实际测试数非零、新增用例全部执行。

## 回归/样例矩阵

- `Audit024_ParentExitsFirstTimeout`：父退出后超时，子 PID 真正消失
- `Audit024_ParentExitsFirstCancellation`：取消同样清理继承管道持有者
- `Audit024_NormalExit`：成功输出和退出码保持
- `Audit024_ExistingTreeControl`：父等待子原测试继续通过
- `Audit024_NoCollateralProcess`：独立同名 fixture 不受影响
- `Audit024_BoundedCleanup`：异常清理有界并不遗留已知 PID

主测试入口为 `tests/Ofdrw.Net.Core.Tests/Audit024RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。



## 原生平台验证命令

各平台设置本计划的 .NET 环境（Windows 可使用 CI job env）后执行：

```bash
dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~ExternalProcessTests|FullyQualifiedName~Audit024RegressionTests" --logger trx --results-directory artifacts/plan-024/native-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
```

## Done criteria

- [ ] Windows/macOS/Linux 均记录新 Audit024 用例非零执行及结果；缺失平台不能由旧 tree 测试代替。

- [ ] 新用例编译/被发现，缺陷修复后全部通过，既有项目回归通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 不要求新增视觉产物的本计划没有凭静态/文本结果声称视觉通过。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 如果所选平台无法在启动时可靠建立所有权，停止并报告该平台边界；不能删除 HasExited 判断就声称解决孤儿进程。不能杀用户正在运行的 Office。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

生命周期所有权必须早于子进程派生，增加外部 wrapper 时复测父先退出场景。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
