# Text engine comparison evidence (Issue #109)

The optional LFS archive [`engine-switch-evidence.tar.zst`](engine-switch-evidence.tar.zst) holds
the outputs reviewed for the [engine switch comparison](../../validation/text-engine-switch-comparison.md):

- `comparison/side/*.png` — twelve 144 dpi page pairs (left: compatibility layout, right: paragraph
  layout engine) for `generated-layout.docx` (2 pages), the CN/EN multi-section header/footer sample
  (5), the inline JPEG sample (1), the Flow sample (2) and the plain-text sample (2);
- `comparison/rmse.txt` — normalized ImageMagick RMSE per page pair;
- `comparison/engine-comparison-manifest.json` — inputs, fonts, page counts, compact text per page
  and output hashes written by `Ofdrw.Net.Layout.E2E --engine-comparison-only`;
- `comparison/ofd/*.ofd` — the ten compared OFDs (PDFs are regenerated from them);
- `f08/` — the F08 visual sample (tabs, synthetic bold/oblique, decorations, letter spacing,
  Ext-B, Arabic, CID-keyed CFF) and the Flow switch sample, with their 144 dpi PNGs.

| Identity | Value |
| --- | --- |
| Source commit | `29e8b7cabaec55b3b0d5485d5d6b1482f79b2a7c` |
| Fonts | `Ofdrw-CI-NotoSansCJKsc-Regular.ttf` and `NotoSans-Regular.ttf` from the trusted-font bundle `sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip` |
| Archive bytes | `3,612,192` |
| Archive SHA-256 | `fabf337ae8fffcaf9fe3fe0c998031f6dcf815eea9591d30ba99baac389aadbb` |
| Archive entries | 29 files; sizes and SHA-256 in `manifest.json` |

Fetch and extract under the ignored `artifacts/` directory:

```sh
git lfs pull --include="docs/evidence/text-engine-switch/**" --exclude=""
mkdir -p artifacts/r22-engine-switch
zstd -dc docs/evidence/text-engine-switch/engine-switch-evidence.tar.zst | tar -xf - -C artifacts/r22-engine-switch
```

Regenerate the comparison with
`OFDRW_SOURCE_SHA=<commit> dotnet run --project e2e/Ofdrw.Net.Layout.E2E -- <out> e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx --engine-comparison-only`,
then `pdftoppm -r 144 -png` each PDF.

## Issue #115 Preview re-check bundle

[`issue115-preview-recheck.tar.zst`](issue115-preview-recheck.tar.zst) holds the outputs regenerated
after the [Issue #115 fixes](../../validation/text-engine-switch-comparison.md),
for a macOS Preview re-check: `f08-visual-sample.{ofd,pdf}` (single highlight path),
`plain-text-{engine,legacy}.{ofd,pdf}` (minimum tab gap), their 144 dpi page-1 PNGs, a 400 dpi
highlight crop and `SHA256SUMS`.

| Identity | Value |
| --- | --- |
| Source commit | `133d03c2739c7867124a97b26748ce63983cfae6` (fonts as above) |
| Archive bytes | `2,882,772` |
| Archive SHA-256 | `a09a44da57dd1144ab9221760b1fc679426b3065f7c7cb95c394335c887c119c` |
| Archive entries | 11 files under `bundle/`; SHA-256 in `bundle/SHA256SUMS` |

Extract with `zstd -dc docs/evidence/text-engine-switch/issue115-preview-recheck.tar.zst | tar -xf - -C artifacts/r22-engine-switch`.
