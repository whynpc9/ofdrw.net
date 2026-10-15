# Plan 021: 只把选项位置的 help 识别为帮助

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Cli/DocumentTools.cs" "src/Ofdrw.Net.Cli/Cli.Images.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Cli.Tests/ImageCommandTests.cs" "tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Cli/DocumentTools.cs" "src/Ofdrw.Net.Cli/Cli.Images.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Cli.Tests/ImageCommandTests.cs" "tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Cli/DocumentTools.cs" "src/Ofdrw.Net.Cli/Cli.Images.cs" "src/Ofdrw.Net.Cli/Program.cs" "tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs" "tests/Ofdrw.Net.Cli.Tests/ImageCommandTests.cs" "tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：IMPLEMENTED；功能测试通过，solution build / 文档门 / 视觉验收进行中
- 来源：F21
- 优先级：P2；工作量：S；修复/实施风险：LOW
- 类别：正确性/DX
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：HIGH 静态路径核实；本计划未实现、未执行其 .NET/视觉门。

## 为什么要做

F21：args.Any(IsHelp) 先于解析，水印文字为 help/--help 时无操作却退出 0；文件名等值也可能触发。



## 当前代码与约定

[src/Ofdrw.Net.Cli/DocumentTools.cs:12](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Cli/DocumentTools.cs:12)

```csharp
        if (args.Length == 0 || args.Any(IsHelp))
        {
            PrintHelp();
            Console.WriteLine("Watermark options: --pages 1,2 --text TEXT | --image FILE --layer ID --layer-type Background|Body|Foreground --x MM --y MM --width MM --height MM --font NAME --font-size MM --alpha 0..255");
            return 0;
        }
        if (command == "mix")
        {
```

[src/Ofdrw.Net.Cli/Program.cs:614](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Cli/Program.cs:614)

```csharp
    private static bool IsHelp(string arg)
    {
        return arg is "-h" or "--help" or "help";
    }

    private static bool IsHelpRequested(CliOptions options)
    {
        return options.Help;
```

其余核查入口：[src/Ofdrw.Net.Cli/Cli.Images.cs:11](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Cli/Cli.Images.cs:11)。

**沿用范例**：[tests/Ofdrw.Net.Cli.Tests/TextReplacementArgumentsTests.cs:12](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Cli.Tests/TextReplacementArgumentsTests.cs:12)

```csharp
    [Theory]
    [InlineData("")]
    [InlineData("--help")]
    [InlineData("-中文")]
    public void LiteralOptionLookingValuesAndDeletionAreUnambiguous(string replacement)
    {
        var result = TextReplacementArguments.Parse(Valid(replacement));
        Assert.Equal("中文A", result.Keyword); Assert.Equal(replacement, result.Replacement);
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Cli/DocumentTools.cs`
- `src/Ofdrw.Net.Cli/Cli.Images.cs`
- `src/Ofdrw.Net.Cli/Program.cs`
- `tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs`
- `tests/Ofdrw.Net.Cli.Tests/ImageCommandTests.cs`
- `tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-021/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-021-cli-help-token-parsing`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-021/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --filter "FullyQualifiedName~Audit021RegressionTests" --logger trx --results-directory artifacts/plan-021/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-021/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --filter "FullyQualifiedName~Audit021RegressionTests" --logger trx --results-directory artifacts/plan-021/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 只把选项位置的 help 识别为帮助

删除参数全数组 Any(IsHelp) 的提前扫描，按命令 parser 消费顺序判断 help；--text、--image、输出路径等带值选项消费的下个 token 即使等于 help/--help 也属于值。保留顶层 help 和实际选项位置 -h/--help 的帮助功能。

**验证**：`dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --filter "FullyQualifiedName~Audit021RegressionTests" --logger trx --results-directory artifacts/plan-021/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

为 DocumentTools 与 Images 路径补一致的缺参数/未知参数处理；若支持 -- 分隔符则显式记录且测试，不为本票重写整套 CLI parser。

**验证**：`dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --filter "FullyQualifiedName~Audit021RegressionTests" --logger trx --results-directory artifacts/plan-021/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：上列受影响项目测试与 `git diff --check` 退出 0；实际测试数非零、新增用例全部执行。

## 回归/样例矩阵

- `Audit021_WatermarkLiteralHelp`：--text help 和 --text --help 真正写水印
- `Audit021_ImageOrOutputNamedHelp`：值位置为 help 的合法路径不变成空操作
- `Audit021_ActualHelp`：选项位置 --help 返回帮助且不要求输入存在
- `Audit021_MissingValue`：缺值报错不静默返回0
- `Audit021_ReplacementControl`：既有 replacement 字面值约定不变

主测试入口为 `tests/Ofdrw.Net.Cli.Tests/Audit021RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。



## 视觉与产物验收

视觉范围限定为实际生成的文字水印页：help 与 --help 必须显示为字面水印内容；帮助和缺参数用例无需生成页面。

本计划包含实际转换/渲染控制样例，必须遵守仓库 AGENTS.md。在 `artifacts/plan-021/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

- DOCX 相关场景同时运行默认与显式 Native，检查原文完整和选页；查看本次 **DOCX→Native OFD→PDF→macOS Preview**，直接 DOCX→PDF 不能替代 Native 产物。
- OFD/PDF 场景检查本次受影响产物全部页面；SVG 另在浏览器查看本次 SVG，不能仅凭 PDF 判定 SVG 正确。
- 将最终 PDF 用 `pdftoppm -png -r 144 <实际PDF绝对路径> <任务产物目录/前缀>` 逐页渲染；这里尖括号为待记录的实际产物路径，不是字面命令参数。图像查看工具逐页查缺字、样式、基线、表格/图片、页码、裁切重叠和异常空页。
- macOS 执行 `open -a Preview <实际PDF绝对路径>`，重新生成后重新打开以避开缓存，逐页查看正文，完成后关闭相关窗口。
- 验收记录分四项：功能测试、自动 PNG 渲染、PNG 目视、Preview。无 GUI 时完成前三项可执行部分并保留文件；Preview 标未完成，完整验收前不能标 DONE。比较明显的文件体积变化，确认样式补绘未重复文本。

## Done criteria

- [ ] 新用例编译/被发现，缺陷修复后全部通过，既有项目回归通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 正常控制产物的功能、自动 PNG、PNG 目视和 Preview 已分别记录，完整视觉验收未完成时不标 DONE。
- [ ] 未扩大到无关源代码/依赖/发布修改。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 若为了修复需升级命令行框架或改变全部命令语法，停止；只收敛当前参数值与控制 token 的歧义。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

先消费选项参数再判断控制 token，新命令不要恢复全局字符串扫描。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。


## 执行记录（2026-10-15）

- 独立 worktree：`/Users/wanghongyi/.codex/worktrees/ofdrw-plan-021/ofdrw.net`；分支 `codex/plan-021-cli-help-token-parsing`；继承基线 `f79c98dceba1fb3bfb186a233183ec235981c05d`（004/006/020）。操作者已明确授权本地提交，覆盖原计划“不授权提交”句；未授权 push/PR/merge/publish。索引由 root 协调者维护。
- 漂移已核对：Program.cs 的路径身份、页范围上限、提交后通知保护与 Cli.Images.cs 的通知变化均保留；未发现未提交改动。
- 正常控制先绿 8/8；初始矩阵基线 38 项为 16 红/22 绿（字面水印、合法 help 路径、未知选项被 help 吞掉）；修复后 38/38。补充通用输出路径独立验证为 2 红/1 绿（`--output --help` / `--output -h` 被误判缺值），统一取值后最终新增 41/41；CLI 项目 169/169，0 skipped。
- 修改：移除 DocumentTools、Images、Merge 的 `Any(IsHelp)` 全局扫描；带值选项逐个消费字面 token；顶层 bare `help` 和实际选项位置 `-h` / `--help` 保留。文档工具固定输入/输出位置的 help-looking token 属于路径。未新增 `--` 分隔符；现有未知选项仍报错，不为本票扩展语法。
- 本次产物/日志位于 `artifacts/plan-021/`，红绿 TRX 位于 `tests/` 子目录。最终构建与视觉门仍进行中，未标 DONE。
