# Plan 015: 在 PDF/SVG 中应用有效路径样式

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Core/Models/OfdPathStyle.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Core/Models/OfdPathStyle.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Core/Models/OfdPathStyle.cs" "src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs" "src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（独立37项、Preview15页、SVG9图）
- 来源：F12
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：正确性/视觉
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：source2d74f22，继承最终014后c2db7ee2；root完整diff与独立37通过，最终FontSubset59、01401573及新11包通过；Preview15页（path9旧源+DOCX6最终源）与9SVG实际通过，root核验15条URL/页码/原PDF及截图SHA，复用依据单独记录。保留PDF后端显式拒绝边界，不声称隐式默认逐像素同。

## 为什么要做

F12：已知属性保留在 SourceXml，但 PDF pen 只使用颜色和宽度；SVG 只额外处理 cap/join/miter。对象 Alpha、dash/offset 没有生效，PDF 也忽略非默认端帽/连接。IsVisible 只处理 Visible，不能补救 Alpha。



## 当前代码与约定

[src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1143](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1143)

```csharp
                if (string.Equals(localName, "PathObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                    var strokeColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "StrokeColor"), material, cancellationToken);
                    var fillColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "FillColor"), material, cancellationToken);
                    var path = new OfdPathElement
```

[src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs:55](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs:55)

```csharp
            if (element.Stroke)
            {
                pen = new XPen(
                    ToXColor(element.StrokeColor),
                    MillimetersToPoints(element.LineWidthMillimeters > 0 ? element.LineWidthMillimeters : 0.01d));
            }

            if (element.Fill)
```

其余核查入口：[src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:149](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs:149)；[src/Ofdrw.Net.Core/Models/OfdGraphicXmlContract.cs:47](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core/Models/OfdGraphicXmlContract.cs:47)。

**沿用范例**：[tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs:13](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs:13)

```csharp
    [Fact]
    public async Task SharedNestedVectorPlacementsDrawTwiceWithIntersectingClips()
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var leaf = new OfdCompositeResource { WidthMillimeters = 20, HeightMillimeters = 20 };
        leaf.Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C M 5 5 L 15 5 L 15 15 L 5 15 C", Fill = true, Stroke = false, FillRule = OfdFillRule.EvenOdd, WidthMillimeters = 20, HeightMillimeters = 20 });
        var parent = new OfdCompositeResource { WidthMillimeters = 30, HeightMillimeters = 30 };
        parent.Elements.Add(new OfdCompositeElement { Resource = leaf, XMillimeters = 3, YMillimeters = 4, WidthMillimeters = 20, HeightMillimeters = 20 });
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Core/Models/OfdPathStyle.cs`
- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdPathRenderer.cs`
- `src/Ofdrw.Net.Converter.Svg/Converters/OfdToSvgConverter.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/SvgGeometryTests.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-015/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-015-path-effective-style`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-015/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit015RegressionTests" --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-015/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan015 --output-dir artifacts/plan-015/package-output
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit015RegressionTests" --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 在 PDF/SVG 中应用有效路径样式

扩展既有 OfdPathStyle 为一次解析、有限数值校验的 effective style，覆盖对象 Alpha、Cap/Join/MiterLimit、DashPattern/DashOffset；保留新建模型默认值和已有子颜色 Alpha，并明确组合关系。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit015RegressionTests" --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

映射 PDF XPen/brush 与 SVG 属性，核对毫米转 point 与 dash 单位；不支持的非法/超界样式明确失败，不能静默变成实线/不透明。不要改 typed graphics 序列化所有权。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit015RegressionTests" --logger trx --results-directory artifacts/plan-015/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan015 --output-dir artifacts/plan-015/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 回归/样例矩阵

- `Audit015_ObjectAlpha`：0/半透明/255 与子色 alpha 组合正确
- `Audit015_DashOffset`：非零 dash offset 与模式在 PDF/SVG 一致
- `Audit015_CapsAndJoins`：butt/round/square 和 miter/bevel/round 控制图
- `Audit015_InvalidNumericStyle`：NaN/负值/异常模式拒绝
- `Audit015_FreshDefaults`：SDK 新建普通路径视觉不变

主测试入口为 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit015RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-015/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

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

- 若后端不能表达所要求样式，先给出明确限制或设计替代，不默认忽略属性；不通过整页光栅化隐藏差异。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

新增 PathObject 属性先进入有效样式契约，再逐后端映射。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。


执行范围校准：组合出现的path clone清SourceXml会丢有效样式，允许精确增加Core/Models/OfdCompositeProjection.cs，在已准入后为私有render clone挂不可变有效style snapshot，原definitions/SourceXml/Writer所有权不变。PdfSharpCore 1.3.67无法精确表达的有效样式须明确NotSupported；SVG可表达则保留，不静默改dash或截miter，需实际后端输出/拒绝回归。
