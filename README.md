# Ofdrw.Net

`.NET 10` based OFD SDK + REST service for EMR fixed-layout document workflows.

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
