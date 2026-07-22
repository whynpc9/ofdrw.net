# Ofdrw.Net.Converter.Pdf Package E2E

This project validates the package-consumer flow for `Ofdrw.Net.Converter.Pdf`,
`Ofdrw.Net.Converter.Docx`, and related packages. It includes visual smoke checks
for generated OFD packages:

1. Pack SDK projects to local NuGet artifacts.
2. Install the PDF and DOCX converter packages from that package source in an isolated app.
3. Execute `OFD -> PDF -> OFD` roundtrip.
4. Parse roundtrip OFD and assert page count.
5. Parse known-good upstream OFD samples from `ofdrw/ofdrw`.
6. Convert upstream PDF samples through `PDF -> OFD -> PDF`, render the first source/roundtrip PDF pages with `pdftoppm`, and assert the rendered images are non-blank and have the same dimensions.

The upstream samples are copied from `ofdrw/ofdrw` `ofdrw-converter/src/test/resources`:

- `Test.pdf`
- `Test3.pdf`
- `ff.pdf`
- `helloworld.ofd`
- `999.ofd`

Visual validation requires `pdftoppm` plus ImageMagick's `convert` and `identify` commands on `PATH`. On macOS these are typically provided by `poppler` and `imagemagick`; the same commands are available from Ubuntu's ImageMagick 6 package used by CI.

## Run

```bash
scripts/run-converter-package-e2e.sh
```

Use a specific package version:

```bash
scripts/run-converter-package-e2e.sh 0.1.0-preview.4
```

## Output

Generated files are under:

- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/source.ofd`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/converted.pdf`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/roundtrip.ofd`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/generated-docx.pdf`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/generated-docx.ofd`
- `e2e/Ofdrw.Net.Converter.Pdf.E2E/output/upstream-ofdrw/**`

For manual Preview validation, open one of the generated `.ofd` files from the output directory and compare the generated `source.png` and `roundtrip.png` images beside it.
