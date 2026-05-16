# Ofdrw.Net

Ofdrw.Net is an early-stage .NET SDK and CLI for OFD document packaging, reading, and PDF/OFD conversion.

## 免责声明（请在使用前阅读）

本项目当前处于早期阶段（MVP/预览版），实现目标是为 `.NET` 生态提供 OFD 文档生成、读取与格式转换能力。

- 本项目并非官方标准实现，也不构成对任何国家/行业标准合规性的承诺。
- 在医疗、档案、司法、政务等高风险场景中使用前，请自行完成充分的功能验证、兼容性验证与安全评估。
- 与签名、签章、加密、长期保存、合规审计相关的能力当前仅部分实现或预留扩展点，不应默认视为满足生产要求。
- 因使用本项目造成的直接或间接损失，使用者应自行评估并承担相应风险。

## Acknowledgements

- 本项目在模块划分、能力边界与目标方向上参考了原项目 [ofdrw/ofdrw](https://github.com/ofdrw/ofdrw)（OFD Reader & Writer, Java）。
- 感谢原项目为 OFD 生态提供的开源实践与 API 设计思路。本项目为 `.NET` 生态下的独立实现，不隶属于原项目。

## Capabilities

- OFD core models and document builder API.
- OFD ZIP packaging with `OFD.xml`, `Doc_0`, `Pages`, resources, attachments, and custom tags.
- OFD reader/parser for pages, text elements, image elements, attachments, and custom tags.
- PDF to OFD conversion by embedding rendered PDF pages as OFD image objects.
- OFD to PDF conversion with text/image drawing and raster fallback.
- Command-line conversion tool packaged as `Ofdrw.Net.Cli`.

PDF rasterization uses `pdftoppm` when converting PDF pages into OFD image resources. Install Poppler and make sure `pdftoppm` is available on `PATH` before using PDF to OFD conversion.

## For Developers

Install the high-level conversion package:

```bash
dotnet add package Ofdrw.Net.Converter --version 0.1.0-preview.1
```

For a narrower dependency surface, install the PDF converter package directly:

```bash
dotnet add package Ofdrw.Net.Converter.Pdf --version 0.1.0-preview.1
```

Convert PDF to OFD:

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("input.pdf");
await using var output = File.Create("output.ofd");

var converter = new PdfToOfdConverter();
await converter.ConvertAsync(input, output);
```

Convert OFD to PDF:

```csharp
using Ofdrw.Net.Converter.Pdf.Converters;

await using var input = File.OpenRead("input.ofd");
await using var output = File.Create("output.pdf");

var converter = new OfdToPdfConverter();
await converter.ConvertAsync(input, output);
```

Convert selected pages with 1-based page numbers in your application code converted to zero-based indexes:

```csharp
var pages = new[] { 0, 2, 3 }; // pages 1, 3, and 4
await converter.ConvertAsync(input, output, pages);
```

Build an OFD document programmatically:

```csharp
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;

var builder = new OfdDocumentBuilder();
builder.AddPage(new OfdPage
{
    Index = 0,
    WidthMillimeters = 210,
    HeightMillimeters = 297,
    Elements =
    {
        new OfdTextElement
        {
            Text = "Hello OFD",
            FontName = "SimSun",
            FontSizeMillimeters = 4,
            XMillimeters = 20,
            YMillimeters = 20,
            WidthMillimeters = 80,
            HeightMillimeters = 10
        }
    }
});

await using var output = File.Create("hello.ofd");
await new OfdPackageWriter().WriteAsync(builder.Build(), output);
```

Read an OFD package:

```csharp
using Ofdrw.Net.Reader.Readers;

await using var input = File.OpenRead("input.ofd");
var package = await new OfdReader().ReadAsync(input);

Console.WriteLine(package.Pages.Count);
```

Package overview:

- `Ofdrw.Net.Converter`: convenience meta-package for conversion use cases.
- `Ofdrw.Net.Converter.Pdf`: PDF/OFD converter implementation.
- `Ofdrw.Net.Cli`: command-line PDF/OFD conversion tool.
- `Ofdrw.Net.Converter.Abstractions`: converter interfaces.
- `Ofdrw.Net.Core`: shared models and constants.
- `Ofdrw.Net.Packaging`: OFD package writer and archive utilities.
- `Ofdrw.Net.Reader`: OFD package reader.
- `Ofdrw.Net.Layout`: document builder helpers.

## For CLI Users

Install the CLI as a .NET tool:

```bash
dotnet tool install --global Ofdrw.Net.Cli --version 0.1.0-preview.1
```

Convert by file extension:

```bash
ofdrw convert input.pdf output.ofd
ofdrw convert input.ofd output.pdf
```

Specify the conversion direction explicitly:

```bash
ofdrw pdf-to-ofd --input input.pdf --output output.ofd
ofdrw ofd-to-pdf --input input.ofd --output output.pdf
```

Convert selected pages:

```bash
ofdrw pdf-to-ofd -i input.pdf -o selected.ofd --pages 1,3-5
ofdrw ofd-to-pdf -i input.ofd -o selected.pdf --pages 2
```

Show help:

```bash
ofdrw --help
```

Run the CLI from a source checkout:

```bash
dotnet run --project src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj -- convert input.pdf output.ofd
```

Build a standalone executable from source:

```bash
dotnet publish src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj -c Release -o artifacts/ofdrw-cli
./artifacts/ofdrw-cli/Ofdrw.Net.Cli convert input.pdf output.ofd
```

## Repository Layout

- `src/Ofdrw.Net.Core`
- `src/Ofdrw.Net.Packaging`
- `src/Ofdrw.Net.Layout`
- `src/Ofdrw.Net.Reader`
- `src/Ofdrw.Net.Converter.Abstractions`
- `src/Ofdrw.Net.Converter.Pdf`
- `src/Ofdrw.Net.Converter`
- `src/Ofdrw.Net.Cli`
- `tests`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E`

## Development

Restore, build, and test:

```bash
dotnet restore Ofdrw.Net.sln
dotnet build Ofdrw.Net.sln
dotnet test Ofdrw.Net.sln
```

Pack locally:

```bash
dotnet pack Ofdrw.Net.sln -c Release -o artifacts/nuget
```

Run the package-consumer E2E check:

```bash
scripts/run-converter-package-e2e.sh
```

Use a specific local package version for E2E:

```bash
scripts/run-converter-package-e2e.sh 0.1.0-preview.local
```

## Notes

- SDK libraries target `netstandard2.0` and `netstandard2.1`.
- The CLI targets `net10.0`.
- Signature, signature verification, encryption, and long-term archival workflows are extension-stage features and are not complete production capabilities.
