# D3 reverse-import controls

These two one-page PDFs are reused byte-for-byte from the existing R22
[reference evidence](../../../../docs/evidence/r22-next-phase/README.md).
They contain the same DejaVu subset, mapped characters and declared source
`abc שלום מקום def`; their positioned visual order differs.

| File | SHA-256 | Role |
| --- | --- | --- |
| `fpdf2-two-word-cell.pdf` | `7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8` | Correctly shaped visual reference; default PdfPig words reverse both Hebrew words and their phrase order. |
| `fpdf2-two-word-fixed.pdf` | `7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5` | Unshaped negative; its matching raw word string must not count as correct import. |

The imported raster preserves each input's appearance. Neither input has a
verified logical structure consumed by the importer, so semantic extraction
rejects both. Explicit raster-only conversion remains available.

Font licensing: [DejaVu/Bitstream/Arev notices](../../../../e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-LICENSE).
PDFs use Git LFS. They add no duplicate full font file or source snapshot.
