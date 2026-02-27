# Ofdrw.Net.Converter.Pdf Package E2E

This project validates the full package-consumer flow for `Ofdrw.Net.Converter.Pdf`:

1. Pack SDK projects to local NuGet artifacts.
2. Install `Ofdrw.Net.Converter.Pdf` from package source in an isolated app.
3. Execute `OFD -> PDF -> OFD` roundtrip.
4. Parse roundtrip OFD and assert page count.

## Run

```bash
/Users/wanghongyi/Projects/ofdrw.net/scripts/run-converter-package-e2e.sh
```

Use a specific package version:

```bash
/Users/wanghongyi/Projects/ofdrw.net/scripts/run-converter-package-e2e.sh 0.1.0-preview.1
```

## Output

Generated files are under:

- `/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/output/source.ofd`
- `/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/output/converted.pdf`
- `/Users/wanghongyi/Projects/ofdrw.net/e2e/Ofdrw.Net.Converter.Pdf.E2E/output/roundtrip.ofd`
