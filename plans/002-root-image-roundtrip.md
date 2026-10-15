# Plan 002: 保留根目录图片的真实载荷

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Reader/Readers/OfdReader.cs" "tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE
- 来源：F02
- 优先级：P1；工作量：S；修复/实施风险：LOW
- 类别：正确性/数据保留
- 依赖：无硬前置；见下方文件协作边界。
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：本次执行结论：DONE（独立复核及12页Preview）；冻结交付 `010db666`。源码基线、独立复核、实测项目、视觉范围、遗留门及产物路径见 [执行调度记录](execution-status.md) 和 [机器清单](execution-state.json)。

## 为什么要做

F02：mediaMap 已存规范化包内路径；无斜杠的根目录图片名又相对 PageRes 解析，读取结果 Data 为空。Writer 将该 Data 作为原资源路径的新载荷，读后保存可能破坏原图。



## 当前代码与约定

[src/Ofdrw.Net.Reader/Readers/OfdReader.cs:975](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Readers/OfdReader.cs:975)

```csharp
            mediaMap[id!] = material.ResolvePath(resPath, mediaLoc, cancellationToken);
            mediaTypeMap[id!] = ToMediaType(media.Attribute("Format")?.Value);
        }
    }

    private static Dictionary<string, OfdCompositeResource> ReadCompositeResources(OfdPackageArchive archive, OfdDocumentPackage package,
        string documentPath, XNamespace ns, IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, OfdModelMaterialization material, CancellationToken token)
```

[src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1130](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader/Readers/OfdReader.cs:1130)

```csharp
                        var mediaPath = mediaFile.Contains('/') ? mediaFile : material.ResolvePath(pageResPath, mediaFile, cancellationToken);
                        if (archive.TryGetBytes(mediaPath, out var bytes))
                        {
                            image.Data = bytes;
                            image.FileName = Path.GetFileName(mediaPath);
                        }
                    }

```

其余核查入口：[src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs:123](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs:123)；[src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs:223](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging/OfdResourceCatalog.cs:223)。

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

- `src/Ofdrw.Net.Reader/Readers/OfdReader.cs`
- `tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-002/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-002-root-image-roundtrip`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-002/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit002RegressionTests" --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-002/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan002 --output-dir artifacts/plan-002/package-output
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

在 `tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit002RegressionTests" --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新用例被发现；缺陷断言可重现，正常控制通过。记录失败名和原因后才进入修复。

### 2. 保留根目录图片的真实载荷

保持 ReadResources 填入 mediaMap 的值为已规范化包内路径；ReadPage 图像分支直接查询该路径，删除基于是否包含斜杠再次解析的启发式。不要改资源写入器的所有权规则。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit002RegressionTests" --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

通过公开 Reader→Writer→Reader 覆盖根目录和嵌套图片。检查原始 ZIP 图片字节和模型 Data 一致，正常相对路径不回退；在回归产物中导出 PDF 显示实际图像。

**验证**：`dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --filter "FullyQualifiedName~Audit002RegressionTests" --logger trx --results-directory artifacts/plan-002/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan002 --output-dir artifacts/plan-002/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 回归/样例矩阵

- `Audit002_AbsoluteRootImage`：资源以 /image.png 引用根图片，读写字节不变
- `Audit002_ParentRelativeRootImage`：嵌套 Res 经 ../ 指向根图片，保存仍可读
- `Audit002_NestedImageControl`：正常文档目录图片无回归
- `Audit002_RepeatedResource`：两处引用一个根图片均显示且不写空载荷

主测试入口为 `tests/Ofdrw.Net.Packaging.Tests/Audit002RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-002/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

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

- 若 mediaMap 在其他调用处实际包含未解析值，先逐调用确认契约并更新计划，不能同时尝试多条路径掩盖错误。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

资源定位始终区分 XML 相对位置和规范化 ZIP 条目名。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
