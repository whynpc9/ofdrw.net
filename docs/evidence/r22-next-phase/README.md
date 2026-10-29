# R22 bounded text-reference evidence

The optional LFS archive [`reference-evidence.tar.zst`](reference-evidence.tar.zst)
contains the **nine one-page reference PDFs actually inspected in macOS
Preview**, their nine 144 dpi page PNGs, and the failed two-word Type 3
diagnostic PDF/PNG. It also contains the selected PdfPig/PDFKit matrix,
generator scripts, the three Preview receipts and the Type 3 result. The
archive contains no build output, NuGet packages, font duplicate, or source
tree snapshot.

| Identity | Value |
| --- | --- |
| Source `main` baseline | `4d9f003a43d0bfc8a1be0ccce298b40080329d75` |
| Input font | `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf` |
| Font SHA-256 | `d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc` |
| Archive bytes | `72,186` |
| Archive SHA-256 | `81f8f6f07371614b7e99d931b4aa11bacc46bac8123e8a36796949cb619f716f` |
| Archive entries | 31, with individual byte sizes and SHA-256 values in `manifest.json` |

Fetch this optional directory with `git lfs pull
--include="docs/evidence/r22-next-phase/**" --exclude=""`, then extract the
archive under ignored `artifacts/r22-next-phase/reference-20261029/`:

```sh
mkdir -p artifacts/r22-next-phase/reference-20261029
zstd -dc docs/evidence/r22-next-phase/reference-evidence.tar.zst |
  tar -xf - -C artifacts/r22-next-phase/reference-20261029
```

The scripts use the repository's existing LFS font fixture. `fonts.conf`
records the original absolute Fontconfig path, so update it to the current
checkout before reproducing Pango generation. The PDFs and PNGs in the
archive are the exact previously hashed outputs; extraction does not require
regeneration. `selected-reader-matrix.json` lists the pinned software
versions, PDF structure and PdfPig/PDFKit results for the nine references.

The reference comparisons do **not** admit adjacent Hebrew words in the
product writer. Actual Preview Find failed on the fpdf2 between-word wrap
and the Pango controls; the tool did not expose raw clipboard Unicode.
The Type 3 diagnostic repaired the pinned PdfPig word tokens but failed
PDFKit whole-page selection. Its one-character-per-word representation also
loses individual character geometry. The full architecture and limitation
record is [`r22-mixed-text-decision.md`](../../architecture/r22-mixed-text-decision.md).
