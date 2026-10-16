# Plan 014: 按源 TextCode 语义生成渲染游程

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs" "src/Ofdrw.Net.Core/Models/OfdTextRenderProjection.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs" "src/Ofdrw.Net.Core/Models/OfdTextRenderProjection.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs" "src/Ofdrw.Net.Core/Models/OfdTextRenderProjection.cs" "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs" "src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（独立97项、Preview9页、SVG2图）
- 来源：F11
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：正确性/共享语义
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：最终源码4790e5c；全套2411零跳过、trusted65与11包通过，root独立Audit01436及Packaging61通过；本次6PDF9页Preview与2SVG实际目视通过，所有原PDF SHA不变。证据在原worktree artifacts/plan-014/verified/acceptance.md。

## 为什么要做

F11：Reader 原始投影给每个省略 X/Y 的 run 独立默认值，PDF/SVG 直接使用；定位器却正确沿用前一 TextCode 起点。源文本十六进制转义同样在定位器解码而渲染器按字面输出。现有 fixture 明确断言 raw Y=4、源 Y 继承 8。



## 当前代码与约定

[src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1094](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1094)

```csharp
                            Text = ReadDirectText(textCode, material, cancellationToken),
                            XMillimeters = ParseDouble(textCode.Attribute("X")?.Value, 0d),
                            YMillimeters = ParseDouble(textCode.Attribute("Y")?.Value, text.FontSizeMillimeters),
                            DeltaX = textCode.Attribute("DeltaX")?.Value,
                            DeltaY = textCode.Attribute("DeltaY")?.Value
                        });
                    }

```

[src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs:126](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs:126)

```csharp
                var x = code.Attribute("X") is { } xa ? Number(xa.Value) : previousX ?? throw new InvalidDataException("First TextCode requires X.");
                var y = code.Attribute("Y") is { } ya ? Number(ya.Value) : previousY ?? throw new InvalidDataException("First TextCode requires Y.");
                previousX = x; previousY = y; // Table46 inherits the previous START, never its last glyph.
                var raw = string.Concat(code.Nodes().OfType<XText>().Select(node => node.Value));
                var decoded = Decode(raw, budget);
                if (decoded.Text.Length == 0)
                {
                    ValidateCanonicalEmptySlot(code, budget);
```

其余核查入口：[src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs:458](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs:458)；[src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:566](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:566)；[tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs:152](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs:152)。

**沿用范例**：[tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs:152](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs:152)

```csharp
            $"<o:TextCode X=\"{next}\" DeltaX=\"{spaceAdvance}\">\\0020{last}</o:TextCode>" + EmptyAt("9");
        var source = await Read(bytes, face, matrix, 1, overrideCodes: codes);
        Assert.Equal(4, ((OfdTextElement)source.Pages[0].Elements[0]).Runs[4].YMillimeters); // Reader raw projection, source Y inherits8.
        var firstProvider = new ObservedProvider(bytes); var before = new OfdKeywordLocator(firstProvider).Find(source, keyword);
        using var destination = new MemoryStream(); await new Ofdrw.Net.Packaging.OfdPackageWriter().WriteAsync(source, destination);
        destination.Position = 0; var reread = await new OfdReader().ReadAsync(destination);
        var secondProvider = new ObservedProvider(bytes); var after = new OfdKeywordLocator(secondProvider).Find(reread, keyword);
        Assert.True(before.IsComplete); Assert.True(after.IsComplete); Assert.Empty(after.Diagnostics);
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Reader/Readers/OfdReader.cs`
- `src/Ofdrw.Net.Reader/Keywords/Internal/OfdStrictTextGeometry.cs`
- `src/Ofdrw.Net.Core/Models/OfdTextRenderProjection.cs`
- `src/Ofdrw.Net.Core/Models/OfdCompositeProjection.cs`
- `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
- `src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs`
- `src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs`
- `src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs`

此外可写本次任务的 `artifacts/plan-014/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-014-source-text-render-projection`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-014/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit014RegressionTests" --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-014/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan014 --output-dir artifacts/plan-014/package-output
python3 scripts/run-trusted-font-tests.py --results-directory artifacts/plan-014/trusted-font-tests
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit014RegressionTests" --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 按源 TextCode 语义生成渲染游程

先锁定现有 raw Reader 投影和受控定位/替换身份测试，不修改 Runs 原始含义。为具有可信来源 XML 的文本构建独立渲染投影，分别继承上个 TextCode START X/Y，按源语义解码反斜杠四位编码并保持 Delta 对解码后字素的索引。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit014RegressionTests" --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

有效游程必须在字体选择、控制字符检查与字体覆盖校验之前建立：PDF `OfdToPdfConverter` 调用 `DocumentFontContext.Resolve` 前、SVG `EmbeddedFontCoverage.Validate` 前，传入同一有效视图而非继续校验 raw text。不得修改源模型来实现此视图。PDF/SVG 共用有效渲染游程，程序新建模型没有来源时保留字面反斜杠。检查 source 与模型变更的一致性，避免以过时 SourceXml 覆盖用户明确修改。对不支持源语义明确报错，不把 strict locator 的整套编辑限制无差别移到通用 renderer。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit014RegressionTests" --logger trx --results-directory artifacts/plan-014/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan014 --output-dir artifacts/plan-014/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 回归/样例矩阵

- `Audit014_InheritedAxes`：X-only/Y-only/连续省略各轴继承前一 START
- `Audit014_EscapedSpace`：源 \\0020 按空格及对应 Delta 输出
- `Audit014_FreshLiteralSlash`：新建模型的字面反斜杠不被解码
- `Audit014_EmptySlots`：空 TextCode 与 Unicode 组合字符仍遵守现有定位契约
- `Audit014_RawAndOwnedControl`：原始投影断言、定位和受控替换身份不变
- `Audit014_DecodedCoverageSucceeds`：字体仅覆盖解码字符、不覆盖转义字面字符时，PDF/SVG 都成功。
- `Audit014_DecodedMissingGlyphFails`：字体仅覆盖转义字面字符、缺解码字符时，PDF/SVG 都拒绝，不能绕过缺字保护。
- `Audit014_BothRenderers`：PDF 与 SVG 实际基线位置和文本语义一致

主测试入口为 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit014RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-014/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

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

- 若共享投影需要 Reader↔Core 循环依赖或破坏受控来源鉴权，停止设计；不能直接全局重写 Reader 坐标默认值。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

渲染视图、raw 模型和受控编辑来源各有责任，后续特性必须明确使用哪一层。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。


执行范围校准：已核实组合对象投影会在renderer之前清除clone的SourceXml；允许在既有source admission之后对text clone应用同一渲染游程视图，不能修改存储definition或放宽未知语义准入。Audit014须覆盖实际组合出现与两renderer文字/轴/覆盖校验。

Revision1经实际回归校准，允许精确增加Packaging/OfdPackageWriter.cs（仅新建typed文本literal斜杠编码并预先计copy/work和取消）、Reader/Extraction/OfdTextExtractor.cs（共享逻辑解码，保留direct-XText扩展容忍契约）、Pdf.Tests.csproj（引用既有Docx库供实际Native/default回归）。来源XML分支不编码、raw模型与鉴权不变，不添加私有标记。

字体重存闭环校准：允许仅Packaging/OfdFontEmbeddingPlan.cs的coverage/usage调用共享逻辑视图，原CGTransform guard、Full/许可/unsafeReason/subset算法与Cmap预算保持。writer传既有material，兼容保留原精确2参constructor并委托必需3参。Auto/literal与imported escape重复保存正反控制必须验证，不能以Full控制替代默认策略。
