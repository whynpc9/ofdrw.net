# Plan 013: 让包消费栅格往返验证逐页内容

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs"`；再运行 `git status --short`、`git diff -- "e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs"` 与 `git diff --cached -- "e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（独立包E2E、Preview11页）
- 来源：F27
- 优先级：P2；工作量：M；修复/实施风险：LOW
- 类别：测试/视觉回归
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：源码5583a584；root完整diff与实际校准复核、独立consume-only包E2E通过；本次Test/Test3/ff共11页Preview通过，root核验全部URL/页码和3PDF原SHA。仍保留Test第5页栅格软化与无文本层边界，不宣称像素等同。

## 为什么要做

F27：当前 smoke 不比较页面内容，只看首对 PNG 非空和同尺寸；后续页面未渲染。能够生成同尺寸但内容错误的结果仍可通过。现有 RMSE helper 可直接复用。



## 当前代码与约定

[e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:607](/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:607)

```csharp
    // This smoke checks raster appearance and page size. Imported private-use
    // CID text has no source font bytes; semantic re-export is tested separately
    // and must refuse unverified glyphs rather than drawing fallback boxes.
    var pdfToOfd = new PdfToOfdConverter(new Ofdrw.Net.Converter.Pdf.PdfToOfdOptions
    {
        TextLayerMode = Ofdrw.Net.Converter.Pdf.PdfTextLayerMode.None
    });
    await using (var pdfInput = File.OpenRead(samplePdf))
```

[e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:644](/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:644)

```csharp
    var sourcePng = await RenderFirstPageAsync(samplePdf, sampleOutputDir, "source");
    var roundtripPng = await RenderFirstPageAsync(pdfPath, sampleOutputDir, "roundtrip");
    await AssertImageHasContentAsync(sourcePng);
    await AssertImageHasContentAsync(roundtripPng);
    await AssertImagesSameSizeAsync(sourcePng, roundtripPng);

    Console.WriteLine($"[E2E] PDF visual smoke passed: {Path.GetFileName(samplePdf)} -> {ofdPath}");
}
```

其余核查入口：[e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:695](/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:695)。

**沿用范例**：[e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:695](/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs:695)

```csharp
static async Task<double> AssertImagesCloseAsync(string expected, string actual, double tolerance, string? diffPath = null)
{
    await AssertImagesSameSizeAsync(expected, actual);
    var result = await RunProcessAsync("compare", $"-metric RMSE \"{expected}\" \"{actual}\" \"{diffPath ?? "null:"}\"");
    var match = System.Text.RegularExpressions.Regex.Match(result.Error, @"\(([0-9.eE+-]+)\)");
    if (result.ExitCode > 1 || !match.Success ||
        !double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var error) || error > tolerance)
        throw new InvalidOperationException($"Rendered pages differ beyond {tolerance}: {expected} vs {actual}; {result.Error}");
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `e2e/Ofdrw.Net.Converter.Pdf.E2E/Program.cs`

此外可写本次任务的 `artifacts/plan-013/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-013-package-raster-all-pages`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-013/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-013/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan013 --output-dir artifacts/plan-013/package-output
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 建立范围与可验证样例

将既有 raster roundtrip smoke 改为读取源/输出页数并要求一致；同 DPI 渲染每一页到唯一目录，校验尺寸、非空，再用已有 AssertImagesCloseAsync 输出每页 RMSE 与差异图。不要重写整个 runner 或新建大规模 golden 库。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan013 --output-dir artifacts/plan-013/package-output` → 记录全部页数和每页比较结果；当前产物若暴露缺陷则记录失败，不修改容差掩盖。

### 2. 形成可审阅闭环

用现有小样本实测抗锯齿差异后固定并说明容差；增加保留尺寸/非空但改变显著内容区域的负对照，必须被同一比较器拒绝。现有 Test.pdf/Test3.pdf/ff.pdf 基线分别5/4/2页，执行前回读实际页数，发生变化须解释而非硬编码11放行。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan013 --output-dir artifacts/plan-013/package-output` → 正常样例全页通过，内置负对照被检测为失败；主进程把预期负对照失败识别为测试成功而非忽略异常。

### 3. 记录完成边界

记录实际验证结果，未通过的门保持未完成。

**验证**：`git diff --check` 退出 0，`git status --short` 只新增/修改本计划允许的文件（保留原有用户改动）。

## 回归/样例矩阵

- `EveryPage`：源和往返页数一致且所有页均有比较记录
- `NegativeControl`：明显内容区域变化导致比较失败
- `DimensionMismatch`：尺寸变化明确失败
- `StableControls`：全部现有样本正常比较通过且差异图可复查

每个场景必须记录预期、实际、结论及复跑入口；未知项不可改写成通过。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-013/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

- DOCX 相关场景同时运行默认与显式 Native，检查原文完整和选页；查看本次 **DOCX→Native OFD→PDF→macOS Preview**，直接 DOCX→PDF 不能替代 Native 产物。
- OFD/PDF 场景检查本次受影响产物全部页面；SVG 另在浏览器查看本次 SVG，不能仅凭 PDF 判定 SVG 正确。
- 将最终 PDF 用 `pdftoppm -png -r 144 <实际PDF绝对路径> <任务产物目录/前缀>` 逐页渲染；这里尖括号为待记录的实际产物路径，不是字面命令参数。图像查看工具逐页查缺字、样式、基线、表格/图片、页码、裁切重叠和异常空页。
- macOS 执行 `open -a Preview <实际PDF绝对路径>`，重新生成后重新打开以避开缓存，逐页查看正文，完成后关闭相关窗口。
- 验收记录分四项：功能测试、自动 PNG 渲染、PNG 目视、Preview。无 GUI 时完成前三项可执行部分并保留文件；Preview 标未完成，完整验收前不能标 DONE。比较明显的文件体积变化，确认样式补绘未重复文本。

## Done criteria

- [ ] 完整包消费、所有页比较和负对照检查通过。
- [ ] 所有本计划要求的命令都有实际结果；未执行/跳过项明确列出，无零测试伪通过。
- [ ] 功能、PNG 自动渲染、PNG 目视和 macOS Preview 四项分别记录；完整视觉验收已完成。
- [ ] 全套回归与本次本地包消费 E2E 均通过。
- [ ] `git diff --check` 退出 0，改动只在 Scope（原有用户改动保持），索引更新真实完成范围。

## STOP conditions

- 若未经修复的现有产物已有真实视觉缺陷，报告并保持此门未通过；不提高阈值掩盖问题，也不在测试票内修改转换实现。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

更新 rasterizer 或 DPI 后重新校准容差并运行负对照；PNG 比较不替代 Preview。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
