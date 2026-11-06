# F06 RTL import visual evidence

[`f06-evidence.tar.zst`](f06-evidence.tar.zst) (LFS, 528,680 bytes, SHA-256
`9cb03d30f5f93df87c753abde8e7b6227020d278b18380368aaba2a32391d958`) holds 22
entries listed with sizes and SHA-256 in its `manifest.json`:

- for four fixtures (`fpdf2-two-word-cell`, `fpdf2-arabic-rtl-cell`,
  `fpdf2-hebrew-rtl-wrap`, `libreoffice-rtl`): the imported OFD, its re-exported
  PDF, 144 dpi PNGs of the source PDF and the re-export, and a stacked
  source/re-export comparison PNG;
- `f06-import-receipt.json` from the installed-package consumer (loaded
  `Ofdrw.Net.Converter.Pdf` DLL SHA-256, input/OFD hashes, imported text);
- `environment.json` (source baseline, OS, SDK, renderer, viewing chain).

Chain: source PDF → `PdfToOfdConverter` (`VisualLineProfile`) → OFD →
`OfdToPdfConverter` → PDF → `pdftoppm -r 144`. macOS Preview was not used.

```sh
git lfs pull --include="docs/evidence/f06-pdf-rtl-import/**" --exclude=""
mkdir -p artifacts/f06-validation && zstd -dc docs/evidence/f06-pdf-rtl-import/f06-evidence.tar.zst | tar -xf - -C artifacts/f06-validation
```
