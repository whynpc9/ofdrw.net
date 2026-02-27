# Ofdrw.Net

`.NET 10` based OFD SDK + REST service for EMR fixed-layout document workflows.

## 免责声明（请在使用前阅读）

本项目当前处于早期阶段（MVP/预览版），实现目标是为 `.NET` 生态提供 OFD 文档生成、转换与电子病历技术规范校验能力。

- 本项目并非官方标准实现，也不构成对任何国家/行业标准合规性的承诺。
- 在医疗、档案、司法、政务等高风险场景中使用前，请自行完成充分的功能验证、兼容性验证与安全评估。
- 与签名、签章、加密、长期保存、合规审计相关的能力当前仅部分实现或预留扩展点，不应默认视为满足生产要求。
- 因使用本项目造成的直接或间接损失，使用者应自行评估并承担相应风险。

## Acknowledgements

- 本项目在模块划分、能力边界与目标方向上参考了原项目 [ofdrw/ofdrw](https://github.com/ofdrw/ofdrw)（OFD Reader & Writer, Java）。
- 感谢原项目为 OFD 生态提供的开源实践与 API 设计思路。本项目为 `.NET` 生态下的独立实现，不隶属于原项目。

## Implemented MVP

- OFD core models and builder API.
- OFD ZIP packaging (`OFD.xml`, `Doc_0`, `Pages`, `Res`, `Attachs`, `Tags`).
- OFD reader/parser for pages, text/image elements, attachments, custom tags.
- PDF <-> OFD conversion:
  - PDF -> OFD: primary path with `PdfPig` text extraction.
  - PDF -> OFD fallback: `pdftoppm` raster import when extraction fails.
  - OFD -> PDF: primary vector draw path with `PdfSharpCore`.
  - OFD -> PDF fallback: raster page image embedding.
- EMR technical specification validator service (`emr-ofd-h-202x` profile).
- ASP.NET Core REST API (`net10.0`).
- Unit/integration tests for packaging, conversion, and EMR validation.

## Project Layout

- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Core`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Packaging`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Layout`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Reader`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Abstractions`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Converter.Pdf`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.EmrTechSpec`
- `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Service`
- `/Users/wanghongyi/Projects/ofdrw.net/tests/*`

## Target Frameworks

- SDK libraries: `netstandard2.0;netstandard2.1`
- Service: `net10.0`

## Run

```bash
dotnet test /Users/wanghongyi/Projects/ofdrw.net/Ofdrw.Net.sln
dotnet run --project /Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.Service/Ofdrw.Net.Service.csproj
```

## NuGet Packaging

Recommended package granularity:

- `Ofdrw.Net.Converter` (meta-package, easiest entry for format conversion use cases)
- `Ofdrw.Net.Converter.Pdf` (PDF <-> OFD converter implementation)
- `Ofdrw.Net.Converter.Abstractions` (converter interfaces)
- `Ofdrw.Net.Core` / `Ofdrw.Net.Packaging` / `Ofdrw.Net.Reader` / `Ofdrw.Net.Layout` (lower-level SDK building blocks)
- `Ofdrw.Net.EmrTechSpec` (EMR OFD-H validation)

Pack locally:

```bash
dotnet pack /Users/wanghongyi/Projects/ofdrw.net/Ofdrw.Net.sln -c Release -o /Users/wanghongyi/Projects/ofdrw.net/artifacts/nuget
```

Consumer project (recommended for conversion):

```bash
dotnet add package Ofdrw.Net.Converter --version 0.1.0-preview.1
```

If you want only PDF conversion implementation:

```bash
dotnet add package Ofdrw.Net.Converter.Pdf --version 0.1.0-preview.1
```

## Package E2E Check

Run an isolated package-consumer E2E flow (pack -> install -> convert):

```bash
/Users/wanghongyi/Projects/ofdrw.net/scripts/run-converter-package-e2e.sh
```

The E2E subproject is:

- `/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E`

## REST Endpoints

- `POST /api/v1/ofd/generate`
- `POST /api/v1/ofd/convert/pdf-to-ofd`
- `POST /api/v1/ofd/convert/ofd-to-pdf`
- `POST /api/v1/ofd/validate/emr-tech-spec`
- `GET /api/v1/ofd/validate/profiles/emr-ofd-h-202x`

## Notes

- Signature and encryption are extension-stage features; not fully implemented in MVP.
- `pdftoppm` is used as external fallback engine for PDF rasterization.
- The EMR profile file is at:
  `/Users/wanghongyi/Projects/ofdrw.net/src/Ofdrw.Net.EmrTechSpec/Profiles/emr-ofd-h-202x.json`
