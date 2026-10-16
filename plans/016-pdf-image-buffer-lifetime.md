# Plan 016: 释放 PDF 普通图片解码缓冲并复用资源

> 执行者须按步骤完成验证；发生 STOP 条件时报告证据，不自行扩范围。本计划独立包含所需背景，无需读取会话。完成后更新 `plans/README.md` 对应行；如果协调者维护索引，只提交状态证据。
>
> **先查漂移**：`git diff --stat d0f88c7321a57443a4cb0ec92631664f66ed8cc0..HEAD -- "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs"`；再运行 `git status --short`、`git diff -- "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs"` 与 `git diff --cached -- "src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs" "src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs" "tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs"`。不要覆盖他人的未提交改动。前置计划造成的变化也必须重读相关符号；若当前语义与下列依据冲突，先修订计划再实施。

## 状态

- 状态：DONE（独立15项、Preview11页）
- 来源：F15
- 优先级：P1；工作量：M；修复/实施风险：MED
- 类别：性能/依赖适配
- 依赖：[010](/Users/wanghongyi/Projects/ofdrw.net/plans/010-image-decoder-admission.md)
- 编写基线：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`，2026-10-15
- 当前证据：sourced9e68a3；全套2526、PDF468、11包/docs通过，root完整source/tests与独立15通过；本次7PDF11页Preview实际通过且原SHA不变，root全11条URL/页码/原PDF与JPEG核验，图片解码后释放、16→1资源、alpha/旋转/nested均有实际断言及页面证据。原worktree artifacts/plan-016/validation.md。

## 为什么要做

F15：每次 placement 新建 XImage.FromStream。固定 PdfSharpCore 版本的 Dispose 只标记 wrapper，默认 image source 留住已解码像素，ImageTable 以随机身份缓存；重复图也不能复用资源。现有签章专用 EncodedBitmapImageSource 已解决同类生命周期。



## 当前代码与约定

[src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:456](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs:456)

```csharp
            using var xImage = XImage.FromStream(() => new MemoryStream(data, writable: false));
            var interpolate = OfdImageRenderingHints.ReadPdfInterpolate(image);
            if (interpolate.HasValue) xImage.Interpolate = interpolate.Value;
            graphics.DrawImage(xImage, 0, 0, 1, 1);
        }
        finally
        {
            graphics.Restore(state);
```

[src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs:13](/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs:13)

```csharp
// PdfSharpCore retains IImageSource through document.ImageTable. Keep only encoded bytes and dimensions there.
// Its PdfImage constructor realizes the stream synchronously; each encoder call owns and releases decoded pixels.
internal sealed class EncodedBitmapImageSource : ImageSource.IImageSource
{
    private readonly byte[] _encoded;
    private readonly Func<byte[], Image<Rgba32>> _decode;
    private readonly CancellationToken _token;
    internal EncodedBitmapImageSource(byte[] encoded, long maximumPixels, CancellationToken token,
```

其余核查入口：[tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs:543](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs:543)。

**沿用范例**：[tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs:543](/Users/wanghongyi/Projects/ofdrw.net/tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs:543)

```csharp
    public void RealPdfImageTableRetainsNoDecodedPixelsAfterRealization(bool jpeg)
    {
        using var document = new PdfSharpCore.Pdf.PdfDocument(); var page = document.AddPage();
        using (var graphics = PdfSharpCore.Drawing.XGraphics.FromPdfPage(page))
        {
            for (var index = 0; index < 12; index++)
            {
                using var image = new Image<Rgba32>(32, 32, new Rgba32((byte)(index * 20), 30, 220));
```

库主要面向 netstandard2.0/2.1，CLI/测试运行 .NET10；不得因测试宿主较新就在库中调用不支持的 API。沿用 file-scoped namespace、现有 async/ConfigureAwait(false) 与 xUnit Fact/Theory 风格；恶意/损坏输入沿用 InvalidDataException，调用方取消保留 OperationCanceledException。公共新增成员补 XML 文档，保持现有重载兼容。只添加能证明缺陷与边界的回归，不写镜像实现的测试。

## 修改范围

允许修改（新建 helper/回归文件按需）：

- `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
- `src/Ofdrw.Net.Converter.Pdf/ImageIO/EncodedBitmapImageSource.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs`

此外可写本次任务的 `artifacts/plan-016/` 验证产物、更新本计划与 `plans/README.md` 的状态。新文件仅在实际需要时创建；列表中不存在的 helper 是建议新增位置，不是已存在 API。

**范围外**：未列出的生产文件、密码/信任体系、包版本/目标框架、发布流程和无关重构。引用的范例/证据文件不自动获得修改授权。

## 工作方式

- 在操作者选定的 checkout 执行；需要新分支时使用 `codex/plan-016-pdf-image-buffer-lifetime`，不要把分支名当作已创建。
- 按回归/实现/验收等逻辑单元组织差异；提交风格参考 `fix: align optional package metadata with password profile v2`。本计划不授权提交、推送、创建 PR 或发布。
- 同一文件的其他计划串行整合或使用隔离 checkout；独立实现可以并行，但不要在共享目录互相覆盖。未知用户改动不清理，不常规清 NuGet 缓存。

## 命令与预期

以下代码块是环境设置与最终门清单，不应在基线阶段无差别整块执行。先按实施步骤运行对应命令；涉及额度/极值的用例只在保护代码具备后启用。

从仓库根目录执行。SDK 由 global.json 的 10.0.100 + latestFeature 约束，不能把旧 TFM roll-forward 当兼容证明。以下命令来自现有工程/runner；当前尚未运行。出现 IPC SocketException(13) 时按 AGENTS.md 关闭 build server 并以上述单节点标志重试，不能据此删除包缓存。

```bash
export DOTNET_CLI_HOME="$PWD/artifacts/plan-016/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1

dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet build Ofdrw.Net.sln -c Release --no-restore --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit016RegressionTests" --logger trx --results-directory artifacts/plan-016/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --logger trx --results-directory artifacts/plan-016/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

dotnet test Ofdrw.Net.sln -c Release --no-build --logger trx --results-directory artifacts/plan-016/full-tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan016 --output-dir artifacts/plan-016/package-output
python3 scripts/check-public-api-docs.py

git diff --check
git status --short
```

预期：所列 restore/build 命令退出 0；如包含测试命令，必须发现非零用例并逐项记录结果，过滤器零命中不算通过。首次审查的 Python 20 项中15通过、2失败3报错，后5因缺 SES restore assets；需要跑脚本测试时先完成 solution restore，不能把此先决条件缺失当新回归。

## 实施步骤

### 1. 固定可观察回归

安全复现优先使用小输入、小额度和可控 internal 接缝。新选项或注入点尚不存在时，先添加兼容的最小接缝，或将依赖新 API 的测试安排在步骤2后；不能把编译失败当重现。无法安全运行的旧行为以已核实静态路径记录，不强制执行巨量分配/慢循环以求红灯。

在 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs` 建立下面的用例，优先通过公开 API 构造/回读。先确认正常控制通过，再确认至少一个缺陷用例在未修复基线上按预期断言失败；不能用编译错误、零测试或缺字体当作重现。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit016RegressionTests" --logger trx --results-directory artifacts/plan-016/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 可安全运行的用例被发现且正常控制通过，缺陷断言或静态未执行依据明确；依赖新 API 的用例列为后置验证。记录后再修复。

### 2. 释放 PDF 普通图片解码缓冲并复用资源

普通 body 及嵌套 appearance 图片改用现有 encoded-only image source，使每次编码回调拥有并释放解码像素；保留 F10 格式准入和 pixel limits。沿用严格签章已验证的同步 realization 所有权，不依赖 XImage.Dispose 释放后端像素。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit016RegressionTests" --logger trx --results-directory artifacts/plan-016/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 对应第一阶段用例转绿，尚未解决的用例明确记录，不扩大修复范围。

### 3. 接入调用链并验证兼容行为

在每个 PDF 文档内按有效图片身份复用资源：编码字节内容、格式与 interpolation 等影响输出的属性纳入键。不同 alpha/插值状态不得错误合并；保持图像变换、边界与严格签章现有策略。

**验证**：`dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --filter "FullyQualifiedName~Audit016RegressionTests" --logger trx --results-directory artifacts/plan-016/tests --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → 新增用例全部通过；上列其他项目命令执行既有回归。对补在现有测试类中的跨入口用例，单独记录其完整测试名。

### 4. 完成最终验收与交接

执行下面的 Done criteria，写明实际运行平台、测试名称/数量、跳过项和产物位置。检查差异仅限 Scope；需要的视觉项未完成时保留状态，不把功能成功当整体通过。

**验证**：`./scripts/run-converter-package-e2e.sh 0.1.0-preview.plan016 --output-dir artifacts/plan-016/package-output` 和上列全套测试命令均退出 0；实际测试数非零、新增用例全部执行，真实包来自本次输出。

## 回归/样例矩阵

- `Audit016_NoLiveDecodedPixels`：每次 realization 后已解码图像被释放
- `Audit016_RepeatedImageResources`：相同图像多次放置资源数稳定
- `Audit016_InterpolationVariants`：同字节不同插值不合并
- `Audit016_AlphaAndNestedAppearance`：透明图和嵌套图片视觉无回归
- `Audit016_CancelAndDecodeFailure`：异常路径也释放已创建像素

主测试入口为 `tests/Ofdrw.Net.Converter.Pdf.Tests/Audit016RegressionTests.cs`；参考上述范例，必要的跨入口测试写入 Scope 已列出的测试文件并记录实际名称。各 case 至少对应一个独立事实断言，可用 Theory 参数化；不按字符串行数或随意哈希当功能验收。


## 固定依赖的所有权依据

当前 PdfSharpCore 1.3.67 对应提交 `d6ac8b092129a4f797365bbbf3eea2723d6c3ebd`。其 [XImage](https://raw.githubusercontent.com/ststeiger/PdfSharpCore/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Drawing/XImage.cs) 的 Dispose 不释放 source 解码像素，[PdfImageTable](https://raw.githubusercontent.com/ststeiger/PdfSharpCore/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Pdf.Advanced/PdfImageTable.cs) 保存资源。不要把调用 using 视为缓冲已释放；复用仓库 EncodedBitmapImageSource 的显式所有权测试。

## 视觉与产物验收

本计划涉及转换/字体/布局或转换链路测试，必须遵守仓库 AGENTS.md。在 `artifacts/plan-016/` 保存源码基线、输入/模式、生成命令、哈希、环境、结果和已检查页列表。优先复用 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，再增加本计划的确定性小样例。

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

- 若后端固定版本不再同步消费 source，需要重新确认所有权后设计；不能用强制 GC 或拍脑袋内存阈值替代释放断言。
- 当前代码、测试或文档的关键语义与摘录冲突，或需要修改范围外文件时，先报告差异并调整计划。
- 同一验证失败两次且合理的局部修正仍不能解释原因，停止并报告实际命令/错误；不删除断言或降低标准放行。
- 所需平台、字体或 GUI 不可用时，完成可执行项并明确未完成门；不能用历史记录替代本次验收。

## 维护与复核重点

升级 PdfSharpCore 时必须重新核实资源缓存和 realization 生命周期。

计划完成只表示本计划的交付和验收完成，不自动意味着整体审查关闭、PR 合并或包已发布。
