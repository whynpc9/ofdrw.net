# Plan 011: 对 cmap 解析去重并在解析前计费

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs" "src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs" "src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs" "src/Ofdrw.Net.Core/Models/OfdFontEmbeddingOptions.cs" "src/Ofdrw.Net.Core/Fonts/FontParseBudget.cs" "tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit011RendererRegressionTests.cs" "src/Ofdrw.Net.Core/Models/OfdOwnedEditCandidate.cs" "src/Ofdrw.Net.Core/Models/OfdSerializedInventory.cs" "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs" "tests/Ofdrw.Net.Core.Tests/OfdControlledEditTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit011DocxRegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs" "src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs" "src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs" "src/Ofdrw.Net.Core/Models/OfdFontEmbeddingOptions.cs" "src/Ofdrw.Net.Core/Fonts/FontParseBudget.cs" "tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit011RendererRegressionTests.cs" "src/Ofdrw.Net.Core/Models/OfdOwnedEditCandidate.cs" "src/Ofdrw.Net.Core/Models/OfdSerializedInventory.cs" "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs" "tests/Ofdrw.Net.Core.Tests/OfdControlledEditTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit011DocxRegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs" "src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs" "src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs" "src/Ofdrw.Net.Core/Models/OfdFontEmbeddingOptions.cs" "src/Ofdrw.Net.Core/Fonts/FontParseBudget.cs" "tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit011RendererRegressionTests.cs" "src/Ofdrw.Net.Core/Models/OfdOwnedEditCandidate.cs" "src/Ofdrw.Net.Core/Models/OfdSerializedInventory.cs" "src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs" "src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs" "tests/Ofdrw.Net.Core.Tests/OfdControlledEditTests.cs" "tests/Ofdrw.Net.Converter.Docx.Tests/Audit011DocxRegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（独立40项、Preview14页、SVG9图）
- 来源：F13
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：安全/性能
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：源码 cac33dee；全套2375零跳过、trusted65及11包通过；root独立40项通过；本次Preview14页、SVG9图实际验收通过。详见 execution-status.md 和原worktree artifacts/plan-011/revision1/desktop-acceptance-2026-10-16.md。

## 为什么要做

F13：每条编码记录都重新处理其子表；format14 每次分配并复制，variation 引用也重复遍历。构造器没有工作预算/取消参数，字形闭包 Spend 在之后才启动。输入字节上限因此不约束累计解析工作。



## 当前代码与约定

[src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs:20](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs:20)

```csharp
        var count = U16(table, 2); Require(table, 4, count * 8);
        for (var i = 0; i < count; i++)
        {
            var record = 4 + i * 8; var platform = U16(table, record); var encoding = U16(table, record + 2);
            var offset = Size(U32(table, record + 4)); var format = U16(table, offset);
            if (format == 14)
            {
                var length = Size(U32(table, offset + 2)); Require(table, offset, length);
```

[src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs:43](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs:43)

```csharp
    private void ReadVariations(byte[] variation)
    {
        Require(variation, 0, 10); var count = Size(U32(variation, 6)); Require(variation, 10, checked(count * 11));
        for (var i = 0; i < count; i++)
        {
            var record = 10 + i * 11; var defaults = Size(U32(variation, record + 3)); var nonDefaults = Size(U32(variation, record + 7));
            if (defaults != 0)
            { var number = Size(U32(variation, defaults)); Require(variation, defaults + 4, checked(number * 4)); }
```

其余核查入口：[src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs:93](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs:93)；[src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs:37](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs:37)。

**沿用范例**：[tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs:242](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs:242)

```csharp
    [Theory]
    [InlineData("bytes")][InlineData("scalars")][InlineData("operations")]
    public async Task SubsetBudgetsActuallyStopWork(string kind)
    {
        var package = Package(Latin.Value);
        if (kind == "bytes") package.Options.FontEmbedding.MaximumFontBytes = 10;
        if (kind == "scalars") package.Options.FontEmbedding.MaximumUsedScalars = 2;
        if (kind == "operations") package.Options.FontEmbedding.MaximumGlyphClosureOperations = 1;
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Core/Fonts/OpenTypeCmap.cs`
- `src/Ofdrw.Net.Core/Fonts/TrueTypeSubset.cs`
- `src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs`
- `src/Ofdrw.Net.Core/Models/OfdFontEmbeddingOptions.cs`
- `src/Ofdrw.Net.Core/Fonts/FontParseBudget.cs`
- `tests/Ofdrw.Net.Packaging.Tests/FontSubsetTests.cs`
- `tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs`
- `src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs`
- `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
- `src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit011RendererRegressionTests.cs`
- `src/Ofdrw.Net.Core/Models/OfdOwnedEditCandidate.cs`
- `src/Ofdrw.Net.Core/Models/OfdSerializedInventory.cs`
- `src/Ofdrw.Net.Layout/Editing/OfdDocumentMerger.cs`
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs`
- `tests/Ofdrw.Net.Core.Tests/OfdControlledEditTests.cs`
- `tests/Ofdrw.Net.Converter.Docx.Tests/Audit011DocxRegressionTests.cs`

此外可写本次任务的 `artifacts/plan-011/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-011-cmap-parse-work-budget`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-011/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit011RegressionTests" --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-011/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan011 --output-dir artifacts/plan-011/package-output
python3 scripts/run-trusted-font-tests.py --results-directory artifacts/plan-011/trusted-font-tests
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

安全复现优先使用小输入、小额度和可控 internal 接缝。新选项或注入点尚不存在时，先添加兼容的最小接缝，或将依赖新 API 的测试安排在步骤2后；不能把编译失败当重现。无法安全运行的旧行为以已核实静态路径记录，不强制执行巨量分配/慢循环以求红灯。

在 `tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit011RegressionTests" --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 可安全运行的用例被发现且正常控制通过，缺陷断言或静态未执行依据明确；依赖新 API 的用例列为后置验证。记录后再修复。

### 2. 对 cmap 解析去重并在解析前计费

按物理 subtable offset/format 去重，不按平台记录身份拒绝合法共享映射；format14 variation 引用也按实际数组位置缓存验证。保持现有 format12 权威选择与 symbol/CFF/full-preserve 规则。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit011RegressionTests" --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

在首次复制或展开前计费 encoding 记录、子表字节和 variation 记录，传播 token 到 cmap 构造及调用方。独立新增有限可配置 parse budget，勿假借 glyph closure operations 覆盖此前未计费阶段；重复解析路径明确共享或分别有界。默认值按现有字体语料成本统计确定并记录，测试使用很小额度。所有 OpenTypeCmap 旧构造路径也必须取得有限默认预算，禁止保留无限制兼容重载；PDF DocumentFontContext 的 embedded、name-only、selected 与 fallback 四个调用点，以及 SVG AddEmbeddedFonts 都纳入同一约束。token 从 PDF 普通/签章嵌套导出和 SVG ConvertAsync 传入解析与覆盖链；可选字体发现 catch 不能吞调用方取消或让累计预算失败换条路径继续耗费。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit011RegressionTests" --logger trx --results-directory artifacts/plan-011/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan011 --output-dir artifacts/plan-011/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 导出入口补充验证

生产调用方位于 `DocumentFontContext.cs:40/64/128/206`、`OfdToPdfConverter.cs:81/144` 和 `OfdToSvgConverter.cs:188`。新增 `Audit011RendererRegressionTests` 分别经 PDF 普通字体、fallback、嵌套签章与 SVG 入口用小预算断言超额/取消，不只测试 embedding/subset。执行：

```bash
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit011RendererRegressionTests" --logger trx --results-directory artifacts/plan-011/renderer-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
```

预期：新渲染用例非零且通过；`rg -n 'new OpenTypeCmap' src` 返回的每个构造点均已确认有限预算和 token 来源。

## 回归/样例矩阵

- `Audit011_SharedOffset`：多 encoding 指同 subtable 结果等价且只解析一次
- `Audit011_VariationAliases`：共享 variation 数组不会反复展开
- `Audit011_SmallBudget`：额度前置扣减，超额在复制前失败
- `Audit011_CancelDuringParse`：生产嵌入/子集入口传播取消
- `Audit011_RealFontControls`：format4/12/14、TTC、symbol 保守回退不回归

主测试入口为 `tests/Ofdrw.Net.Packaging.Tests/Audit011RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-011/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

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

- 若去重改变 format12 权威或合法 Unicode variation 结果，先停；不得以拒绝所有 alias 代替解析去重。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

未来新增 cmap 格式需声明解析工作单位，禁止构造函数偷偷脱离预算。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。


## 执行期协调修订（2026-10-15）

独立复核发现新增 MaximumCmapParseWork 在受控克隆、Merge、Native DOCX 的三个 options复制点重置为默认值，且来源inventory未覆盖该字段。已授权以上最小扩围；必须以非默认73/29及变更73→74验证克隆/限额/防变更，而非只测默认值。执行者提交cac33dee修复并补回归；来源快照与公开预算文档也覆盖variation查询工作量。最终测试/视觉门仍须实际完成。
