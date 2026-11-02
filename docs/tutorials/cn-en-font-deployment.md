# 中英文 DOCX 字体部署 / CN/EN font deployment

本教程用调用方提供的真实 Regular 与 Bold 字体，运行 **DOCX → native/default OFD → PDF**。目标是固定可复现的字体选择、保留原文，并避免把 Regular 描边当成物理 Bold。以下 API/CLI 对应当前仓库源码；使用 NuGet 时须先确认所装版本包含这些接口。本教程不改变 SDK 默认值。未发布 `0.1.0-preview.10` 的受限候选 profile 为 `cn-en-physical-cjk-400-700-v1`，范围及当前验收状态见[2026-11-02 计划](../next-steps-2026-11-02.md)。

This recipe supplies physical Regular/Bold faces to native DOCX conversion. It covers the current source API and CLI; a successful conversion is separate from visual, Copy/Find, and release acceptance.

## 1. 先选字体与字体面

- **Regular（常规面）**与 **physical Bold（真实粗体面）**有各自的字形轮廓和度量。文件名含 `Bold`、DOCX 的加粗开关或字体资源的 Bold 标志，都不能把 Regular 变成真实 Bold。下面从同一个已固定的可变源分别生成 `wght=400` 与 `wght=700` 的静态 TrueType 实例；运行时不使用变量字体。
- **Synthetic bold（合成粗体）**在 Regular 上加文字描边。Native 嵌入文字缺少真实 Bold 时，默认 `NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.TextStroke`，并返回 `DOCX_EMBEDDED_BOLD_SIMULATED`。需要严格真实粗体时，调用方设置 `Throw`；它只检查 Native 中实际使用的嵌入粗体，不能作为全局缺字体或 name-only 字体检查。
- 本配方含 `Noto Sans CJK SC` Regular/Bold，另含比例英文 `Noto Sans` Regular。它**不包含 Noto Sans Bold 或真实 Italic**。源文档如显式请求 `Noto Sans` 的粗体，还需提供该族合法的真实 Bold，否则严格模式可能失败；斜体可能由现有路径模拟。不能由中文 Bold 推断所有家族/样式齐全。
- DOCX 的字体声明仍参与选面。建议在源文档中明确设置中文 `Noto Sans CJK SC` 和英文所需家族，并确认覆盖实际用字。`FontFallbackFamilies` 用于没有解析出显式字体的场合，不会全局改写 DOCX 所有字体。宋体/黑体等现有 name-only 兼容路径仍由阅读器解析，不能据此承诺跨机一致。

A physical bold face has its own outlines. `Throw` rejects missing physical bold only for selected embedded native text. Supply the corresponding face for every explicitly requested family/style; fallback names alone do not provide font bytes.

## 2. 在隔离目录准备固定字体

从仓库根目录执行。Python/fontTools 只用于这一步离线准备资料；它们不是 SDK 或应用的运行时依赖。配方沿用 [install-ci-fonts.py](../../scripts/install-ci-fonts.py) 的来源、版本、静态实例化参数和校验方式，并增加同源 `wght=700` 实例。输出在已忽略的 `artifacts/`，不安装系统字体、不复制字体进源码/NuGet。

```sh
python3 -m venv artifacts/font-deployment/venv
artifacts/font-deployment/venv/bin/python -m pip install 'fonttools==4.59.2'
artifacts/font-deployment/venv/bin/python - <<'PY'
from pathlib import Path
from urllib.request import urlopen
import hashlib
import json
import fontTools
from fontTools.ttLib import TTFont
from fontTools.varLib.instancer import instantiateVariableFont

assert fontTools.__version__ == '4.59.2'
root = Path('artifacts/font-deployment')
fonts = root / 'fonts'
fonts.mkdir(parents=True, exist_ok=True)
cjk = ('https://raw.githubusercontent.com/notofonts/noto-cjk/'
       'f8d157532fbfaeda587e826d4cd5b21a49186f7c/Sans')
latin = ('https://raw.githubusercontent.com/notofonts/noto-fonts/'
         'ffebf8c1ee449e544955a7e813c54f9b73848eac')

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def fetch(url, path, expected):
    # Existing downloads are reused only after checking their pinned identity.
    if not path.exists():
        with urlopen(url, timeout=120) as response:
            path.write_bytes(response.read())
    if digest(path) != expected:
        raise ValueError(f'SHA-256 mismatch: {path}')

source = root / 'NotoSansCJKsc-VF.ttf'
fetch(cjk + '/Variable/TTF/NotoSansCJKsc-VF.ttf', source,
      '990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f')
for weight, style in [(400, 'Regular'), (700, 'Bold')]:
    with TTFont(source, recalcTimestamp=False) as variable:
        font = instantiateVariableFont(variable, {'wght': weight}, inplace=True,
                                       updateFontNames=True)
        assert 'fvar' not in font and 'glyf' in font
        assert font['OS/2'].usWeightClass == weight
        assert bool(font['OS/2'].fsSelection & 0x20) == (weight == 700)
        assert set(map(ord, '中文数量 Alpha')).issubset(font.getBestCmap())
        font.save(fonts / f'Ofdrw-CI-NotoSansCJKsc-{style}.ttf')
fetch(cjk + '/LICENSE', fonts / 'Ofdrw-CI-Noto-OFL.txt',
      '6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2')
fetch(latin + '/hinted/ttf/NotoSans/NotoSans-Regular.ttf',
      fonts / 'NotoSans-Regular.ttf',
      'b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5')
fetch(latin + '/LICENSE', fonts / 'NotoSans-LICENSE',
      '0dab92d0544f7b233403f14b84a663bdbfa746982eda629e7f4f9ffe1b036feb')
expected = {
    'Ofdrw-CI-NotoSansCJKsc-Regular.ttf':
        '3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa',
    'Ofdrw-CI-NotoSansCJKsc-Bold.ttf':
        '647a6e375c5a0d8958111173d7b7cec5237e0784cc4490663e43834dfff491be',
    'Ofdrw-CI-Noto-OFL.txt':
        '6a73f9541c2de74158c0e7cf6b0a58ef774f5a780bf191f2d7ec9cc53efe2bf2',
    'NotoSans-Regular.ttf':
        'b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5',
    'NotoSans-LICENSE':
        '0dab92d0544f7b233403f14b84a663bdbfa746982eda629e7f4f9ffe1b036feb',
}
records = []
for name, sha256 in expected.items():
    path = fonts / name
    if digest(path) != sha256:
        raise ValueError(f'Output differs from pinned recipe: {path}')
    records.append(dict(file=name, bytes=path.stat().st_size, sha256=sha256))
(root / 'font-manifest.json').write_text(json.dumps({
    'fontTools': fontTools.__version__, 'weights': [400, 700],
    'sourceSha256': digest(source), 'files': records,
}, indent=2) + '\n', encoding='utf-8')
print(fonts.resolve())
PY
```

关键身份：CJK Regular 为 21,659,544 字节；Bold 为 21,638,668 字节。两者 family 为 `Noto Sans CJK SC`，subfamily 分别为 `Regular` / `Bold`，由源的独立轮廓实例化得到；没有通过修改 Regular 的样式标志造粗体。源、fontTools 版本、`recalcTimestamp=False`、`updateFontNames=True` 和输出哈希都固定。哈希不符时先查工具链和配方，不把新字节自动改成新的验收基准。

字体采用 **SIL OFL 1.1**。交付字体时保留对应版权和许可证，并遵守许可中的保留字体名称等条件；不要将专有 Office/系统字体复制成公共部署包。本配方不是嵌入权的通用授权，参见[字体子集与许可约定](../font-subset-and-reuse.md)。许可证和既有输入身份也见[可信字体清单](../../tests/fixtures/trusted-fonts/manifest.json)。

The pinned source is downloaded once and verified before use. Keep both OFL license files with the prepared fonts. No original font payload is added to this tutorial or the SDK package.

## 3. 完整 API 示例：严格物理粗体，两种 Native 入口

在仓库根下创建 `artifacts/font-deployment-app/FontDeployment.csproj`：

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="../../src/Ofdrw.Net.Converter.Docx/Ofdrw.Net.Converter.Docx.csproj" />
    <ProjectReference Include="../../src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj" />
  </ItemGroup>
</Project>
```

保存以下完整代码为同目录的 `Program.cs`。使用发布包的应用可引用对应版本的 `Ofdrw.Net.Converter.Docx` 和 `Ofdrw.Net.Converter.Pdf`，代码相同；这里用项目引用明确验证当前源码。

```csharp
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;

if (args.Length != 3)
    throw new ArgumentException("Usage: input.docx font-directory output-directory");
var inputPath = Path.GetFullPath(args[0]);
var fontDirectory = Path.GetFullPath(args[1]);
var outputDirectory = Path.GetFullPath(args[2]);
if (!Directory.Exists(fontDirectory))
    throw new DirectoryNotFoundException(fontDirectory);
Directory.CreateDirectory(outputDirectory);

// Initialize before any other PDFsharp font operations in this process.
PdfFontRegistry.EnsureInstalled();
foreach (var mode in new[] { "native", "default" })
{
    var options = new DocxConversionOptions
    {
        UseInstalledMicrosoftOfficeFonts = false,
        NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw
    };
    options.FontDirectories.Add(fontDirectory);
    options.FontFallbackFamilies.Clear();
    options.FontFallbackFamilies.Add("Noto Sans CJK SC");
    options.FontFallbackFamilies.Add("Noto Sans");
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    // "default" deliberately leaves OfdMode unset; current default is Native.
    var ofdPath = Path.Combine(outputDirectory, $"font-{mode}.ofd");
    using (var input = File.OpenRead(inputPath))
    using (var output = File.Create(ofdPath))
    {
        var result = await new DocxToOfdConverter(options)
            .ConvertWithResultAsync(input, output);
        Console.WriteLine($"{mode}: {result.Mode}; original text: {result.OriginalTextPreserved}");
        foreach (var diagnostic in result.Diagnostics)
            Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
        if (!result.OriginalTextPreserved || !result.PageTextMappingAccurate)
            throw new InvalidDataException("Inspect DOCX preservation diagnostics.");
    }
    using (var input = File.OpenRead(ofdPath))
    using (var output = File.Create(Path.ChangeExtension(ofdPath, ".pdf")))
        await new OfdToPdfConverter().ConvertAsync(input, output);
}
```

运行命令（当前仓库使用 .NET 10 SDK；已有 NuGet 缓存可通过 `NUGET_PACKAGES` 显式指定）。示例 DOCX 使用 Git LFS；若尚未材料化，先按 [README](../../README.md#clone-and-large-files) 拉取该 fixture，不把 LFS pointer 作为文档输入：

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/font-deployment/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
dotnet build artifacts/font-deployment-app/FontDeployment.csproj -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet artifacts/font-deployment-app/bin/Release/net10.0/FontDeployment.dll e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx artifacts/font-deployment/fonts artifacts/font-deployment/output
```

`FontDirectories` 加载真实字体字节，而非注册一个任意别名；配置目录为空时 Native 扫描平台字体目录。目录缺失/字体加载失败的诊断应检查，不能把目录存在当成选中了预期面。`UseInstalledMicrosoftOfficeFonts=false` 关闭的是 **macOS LibreOffice 暂存已安装 Word 私有字体**；Native 不调用该暂存流程，也不调用 Word/LibreOffice。该开关不禁用所有系统字体解析、不提供全局沙箱或嵌入保证。

`Throw` 会在遇到没有真实 Bold 的已嵌入粗体文字时、写包前拒绝；API 示例自身会创建输出文件，失败时应清理/隔离未完成文件。`OriginalTextPreserved` 和诊断是功能结果，不能证明实际阅读器显示、复制或搜索成功。启用默认描边对照时，用独立 Regular-only 目录和输出目录，显式改为 `TextStroke`；保留 `DOCX_EMBEDDED_BOLD_SIMULATED`，不要把该对照标成物理粗体。

## 4. CLI：目录选项与 API 的差别

已安装 CLI 的对应源码选项如下，`--font-directory` 可以重复。`--docx-ofd-mode` 省略时当前同样走 Native：

```sh
ofdrw docx-to-ofd input.docx artifacts/font-native.ofd --docx-ofd-mode native --font-directory /absolute/path/to/fonts
ofdrw docx-to-ofd input.docx artifacts/font-default.ofd --font-directory /absolute/path/to/fonts
ofdrw ofd-to-pdf artifacts/font-native.ofd artifacts/font-native.pdf
ofdrw ofd-to-pdf artifacts/font-default.ofd artifacts/font-default.pdf
```

CLI **没有** `NativeEmbeddedBoldBehavior`、`FontFallbackFamilies` 或 `UseInstalledMicrosoftOfficeFonts` 的对应开关；其嵌入粗体策略保持 `TextStroke`。需要上述严格配置时使用 API 示例。`--docx-engine` 选择的是直接 PDF/双层渲染器，不能借此把 Native 改成 Word 排版。CLI 与 API 选项以 [Program.cs](../../src/Ofdrw.Net.Cli/Program.cs) 和 [DocxConversionOptions](../../src/Ofdrw.Net.Converter.Docx/DocxConversionOptions.cs) 为准。

For strict physical bold and an explicit fallback list, use the API. The CLI accepts repeated `--font-directory` arguments but retains its default embedded-bold policy.

## 5. 检查实际输出及当前限制

先保留输入、源码 SHA、字体清单和 OFD/PDF 哈希，再核对原 OFD 的 TextCode、实际嵌入 face/子集及诊断。默认可能生成子集，**嵌入子集哈希不应直接等同于源字体哈希**。粗体资源的物理样式、文字的 Weight、是否发生额外描边要分别检查；仅没有模拟诊断也不能证明所有文本已嵌字。

随后检查两种模式生成的每一页：中英文是否完整、比例英文间距、局部粗斜体/颜色、表格底色/边框、分页、裁切及重叠。使用本次链路 **DOCX → OFD → PDF → macOS Preview**，重新生成后重新打开 PDF；直接 DOCX→PDF 不代替本检查。PNG 渲染/目视、Preview、原始 Unicode Copy 与首/中/末 Find 分别记录。

已有[本地独立验证](../validation/cn-en-release-local-validation.md)绑定 `1e6f5eefb60b18ea9951769273b749a27ac9c174`，不是本教程新运行结论：

| 项目 | 已有结果与范围 |
| --- | --- |
| 显式物理 Regular/Bold 配置 | N13 两模式严格 `Throw` 选面及原文/样式断言完成；这两份 PDF 第 1 页“数量”的“量”字在该次 PNG/Preview 中清晰。仅适用于该具名样例/配置。 |
| 原 Regular-only 默认合成粗体 | 同一未修改样例 native/default 第 1 页“量”字在 144 DPI 下仍有字腔堵塞，**FAIL 保留**；物理 Bold 的结果不覆盖它。 |
| Preview Find | 新生成 original native/default 的完整页末日期仍无可操作命中，**FAIL，根因未确定**；字体配方不宣称修复搜索。 |
| 原始 Unicode Copy | 原始剪贴板字节验收**未完成**；粘贴到搜索字段的结果不能替代该项。 |
| 发布 | 旧运行的原 Regular FAIL、Find FAIL 与 Copy 未完成保留。2026-11-02 用户将新候选限定到本教程的真实 CJK Regular/Bold 配置，后两项明列预览版限制；按[当前计划](../next-steps-2026-11-02.md)完成新候选验证后才可评估 profile-qualified accepted。教程不是新验收记录，也不表示已发布。 |

The unpublished preview.10 candidate is limited to this named physical-font configuration. Regular-only failure remains recorded. Preview Find compatibility and unverified raw Unicode Copy are disclosed limitations, not passed checks. For Flow-generated OFD, see [Flow API](../flow-layout.md) and [font resources/subsetting](../font-subset-and-reuse.md): this DOCX option does not automatically embed fonts into arbitrary Flow documents.
