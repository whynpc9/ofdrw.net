# Plan 003: 在有效 PDF 候选与最终输出提交之间建立明确边界

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：IMPLEMENTED；实际视觉 PARTIAL/FAIL（基线已存在的直接 BuiltIn PDF 副标题斜体缺失）
- 来源：F03、F17
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：正确性/输出契约 / 正确性/契约漂移
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：源码156541ec；root完整diff与独立10/10通过，实际7页Preview中Native/default/选页5页通过、direct BuiltIn 2页保留副标题斜体缺陷。d0f88c7同环境PNG逐字节相同，已有缺陷；artifacts/plan-003/style-investigation.md记录PdfSharp1.3.67 simulation flags及宿主fallback根因与未实施方案。未在本计划内扩大字体依赖修复，不宣称全部视觉通过。

## 为什么要做

F03：最终 CopyToAsync 位于 Auto 引擎回退 catch 内；IOException 可触发下一引擎。如果流部分写入后再次可写，可能在残缺 PDF 后追加另一份 PDF 并返回成功；持续失败则浪费重试且误报原因。

F17：外部进程成功退出且文件有 PDF 前缀即可提交；截断 PDF 会令直接 DOCX→PDF 报成功，并阻止 Auto 尝试后续引擎。已定设计要求可打开检查。

这两个发现共用同一个交付边界，但回归断言分别保留，不把一个用例通过当作两个问题都解决。

## 当前代码与约定

[src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:141](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:141)

```csharp
                    using var pdfFile = File.OpenRead(outputPath);
                    await pdfFile.CopyToAsync(pdfOutput, 81920, cancellationToken).ConfigureAwait(false);
                    diagnostics.AddRange(engineDiagnostics);
                    return new DocxConversionResult(engine, attempted, diagnostics);
                }
                catch (OperationCanceledException)
                {
                    throw;
```

[src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:150](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:150)

```csharp
                catch (Exception ex) when (CanFallback(engine, ex))
                {
                    lastFailure = ex;
                    diagnostics.Add(new DocxConversionDiagnostic(
                        "DOCX_ENGINE_FAILED",
                        $"{GetEngineName(engine)} failed; Auto is trying the next available engine.",
                        DocxConversionDiagnosticSeverity.Warning));
                }
```

[src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:271](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:271)

```csharp
    private static void ValidatePdfOutput(DocxConversionEngine engine, string outputPath)
    {
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length < 5)
        {
            throw new InvalidOperationException($"{GetEngineName(engine)} did not produce a PDF.");
        }

        using var stream = File.OpenRead(outputPath);
```

[src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:139](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:139)

```csharp
                    ValidatePdfOutput(engine, outputPath);

                    using var pdfFile = File.OpenRead(outputPath);
                    await pdfFile.CopyToAsync(pdfOutput, 81920, cancellationToken).ConfigureAwait(false);
                    diagnostics.AddRange(engineDiagnostics);
                    return new DocxConversionResult(engine, attempted, diagnostics);
                }
                catch (OperationCanceledException)
```

其余核查入口：[src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:288](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:288)；[docs/builtin-docx-conversion-engine-plan.md:85](/Users/wanghongyi/Projects/ofdrw.net/docs/builtin-docx-conversion-engine-plan.md:85)；[docs/builtin-docx-conversion-engine-plan.md:89](/Users/wanghongyi/Projects/ofdrw.net/docs/builtin-docx-conversion-engine-plan.md:89)。

**沿用范例**：[tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs:483](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs:483)

```csharp
    public async Task Auto_ShouldFallbackWhenLibreOfficeFailsToStart()
    {
        var converter = new DocxToPdfConverter(
            new DocxConversionOptions
            {
                Engine = DocxConversionEngine.Auto,
                LibreOfficePath = Path.Combine(Path.GetTempPath(), "missing-ofdrw-soffice")
            },
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs`
- `tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs`
- `tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-003/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-003-docx-pdf-output-transaction`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-003/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --filter "FullyQualifiedName~Audit003RegressionTests" --logger trx --results-directory artifacts/plan-003/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-003/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-003/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan003 --output-dir artifacts/plan-003/package-output
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --filter "FullyQualifiedName~Audit003RegressionTests" --logger trx --results-directory artifacts/plan-003/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 在有效 PDF 候选与最终输出提交之间建立明确边界

在 ValidatePdfOutput 对候选临时文件执行已有 PdfPig 可打开性/页结构检查；异常转换为 Auto 可回退的候选失败，不误把输入安全拒绝变成可回退。固定假渲染器与合法小 PDF 控制样例；仅添加 internal 测试接缝，不扩展公共引擎系统。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --filter "FullyQualifiedName~Audit003RegressionTests" --logger trx --results-directory artifacts/plan-003/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

候选引擎循环仅生产并校验暂存文件；选定后在循环的 fallback catch 之外一次 CopyToAsync。写入失败或取消原样传播，不再执行另一引擎，不尝试回滚任意用户 Stream。最终清理临时文件且保留正确 AttemptedEngines/diagnostics。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --filter "FullyQualifiedName~Audit003RegressionTests" --logger trx --results-directory artifacts/plan-003/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan003 --output-dir artifacts/plan-003/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 回归/样例矩阵

- `Audit003_PartialOutputFailure`：接收部分字节后抛 IOException 的输出流，无后续引擎调用
- `Audit003_PdfPrefixOnly`：仅 %PDF- 文件必须拒绝，Auto 尝试下一候选
- `Audit003_BrokenStructure`：截断交叉引用/无有效页面不能提交
- `Audit003_ValidFallback`：首引擎启动失败后合法候选仅提交一次
- `Audit003_CanceledOutput`：取消不转换成引擎回退

主测试入口为 `tests/Ofdrw.Net.Converter.Docx.Tests/Audit003RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-003/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

- DOCX 相关场景同时运行默认与显式 Native，检查原文完整和选页；查看本次 **DOCX→Native OFD→PDF→macOS Preview**，直接 DOCX→PDF 不能替代 Native 产物。
- OFD/PDF 场景检查本次受影响产物全部页面；SVG 另在浏览器查看本次 SVG，不能仅凭 PDF 判定 SVG 正确。
- 将最终 PDF 用 `pdftoppm -png -r 144 <实际PDF绝对路径> <任务产物目录/前缀>` 逐页渲染；这里尖括号为待记录的实际产物路径，不是字面命令参数。图像查看工具逐页查缺字、样式、基线、表格/图片、页码、裁切重叠和异常空页。
- macOS 执行 `open -a Preview <实际PDF绝对路径>`，重新生成后重新打开以避开缓存，逐页查看正文，完成后关闭相关窗口。
- 验收记录分四项：功能测试、自动 PNG 渲染、PNG 目视、Preview。无 GUI 时完成前三项可执行部分并保留文件；Preview 标未完成，完整验收前不能标 DONE。比较明显的文件体积变化，确认样式补绘未重复文本。

## Done criteria

- [ ] 新用例编译/被发现，缺陷修复后全部通过，既有项目回归通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 功能、PNG 自动渲染、PNG 目视和 macOS Preview 四项分别记录；完整视觉验收已完成。
- [ ] 全套回归与本次本地包消费 E2E 均通过。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 若所选解析器要求新增依赖或改变 netstandard 目标，停止评审；当前 Docx 已引用 Converter.Pdf，优先复用其现有 PdfPig。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

新增引擎必须只写暂存文件；输出流的失败不属于引擎可恢复失败。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
