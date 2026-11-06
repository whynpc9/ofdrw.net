# F06 right-to-left import fixtures

One-page PDFs from three independent generators, each with a declared logical
source. They are frozen inputs; tests verify every SHA-256 before use.

| File | SHA-256 | Generator | Declared source / role |
| --- | --- | --- | --- |
| `pango-two-word.pdf` | `c20d0d39456727ea7f967b679120cb6ffc0f52bb89cfc7816b746f48d7a99513` | Pango 1.58.2 / Cairo 1.18.4 | `abc שלום מקום def`; glyphs in logical content order |
| `pango-after-rtl-wrap.pdf` | `3c7d4b6a973097e14c2761f0045d6ed86bfcc8d178f131f44d2f5ecdafe7e7bc` | Pango / Cairo | same source wrapped after the RTL words; `Auto` is ambiguous, `LeftToRight` recovers it |
| `pango-between-rtl-wrap.pdf` | `e61b35acc0a14847253b68fd6a01728917ab025392b441637116cf9afbe417b5` | Pango / Cairo | negative: carries `/ActualText` |
| `fpdf2-between-rtl-zero-margin.pdf` | `9c40fc51de911841fecf8dc44dafe8c88bd392bf525a4f9f85112df8648eb3c1` | fpdf2 2.8.9 | same source wrapped between the RTL words |
| `fpdf2-arabic-ltr-cell.pdf` | `a3ec912dff07c255dd6c04e9094b8834a924f254a7f64c15a4bff609104c1b68` | fpdf2 2.8.9 + uharfbuzz 0.57.0 | `abc مرحبا لا def`; lam-alef glyph maps to two characters |
| `fpdf2-arabic-rtl-cell.pdf` | `8abf6f54ae326e7bb7fbd597a09b88e37ab0b73e3959a80d176c6f88861052c0` | fpdf2 + uharfbuzz | `مرحبا بالعالم لا`, RTL paragraph |
| `fpdf2-hebrew-rtl-wrap.pdf` | `39678a380122d6e62940723cdf23a481ac467130ae14bc34254407dec7b5c02f` | fpdf2 + uharfbuzz | `שלום עולם abc מקום טוב`, RTL paragraph on two lines |
| `fpdf2-hebrew-digits-cell.pdf` | `c9a087c9856f297caf8d4f38cf8545cc0db34a9b5e23dc9dedfd17d5e0f36f26` | fpdf2 + uharfbuzz | negative: European digits (bidi EN) |
| `fpdf2-hebrew-niqqud-cell.pdf` | `f4ac0949f4b78e9298a03c01788d0b628676df3535736d100063eda322dd053f` | fpdf2 + uharfbuzz | negative: Hebrew points (bidi NSM) |
| `libreoffice-rtl.pdf` | `3db83d540e7e6d9772a5a01861550541ac69c74669e928d02754b601ad5c2f07` | LibreOffice 24.2.7.2 Writer/Web | `libreoffice-rtl.html`: RTL Hebrew, RTL Arabic and an LTR paragraph |

The four Pango and `fpdf2-between-…` files are byte-identical copies of PDFs in
the [mixed-text reference evidence](../../../../docs/evidence/text-reader-reference/README.md)
archive. The fpdf2 files were generated with `make_fpdf2.py` from the repository
historical font fixture (SHA-256 `d7441d42…a075b8fbcc`), preserved by the
source commit that generated the PDFs. The current semantically named fixture
`e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-TextFixture.ttf`
has the same outlines and metrics but renamed font metadata (SHA-256 `3362132d…877a8212`):

```sh
python -m venv venv && venv/bin/pip install fpdf2==2.8.9 uharfbuzz==0.57.0
venv/bin/python -I make_fpdf2.py <repo>/e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-TextFixture.ttf <out-dir>
```

`libreoffice-rtl.pdf` was produced with
`soffice --headless --convert-to pdf:writer_pdf_Export libreoffice-rtl.html`; it
embeds a subset of the host's Debian `fonts-dejavu-core` 2.37-8 DejaVu Sans
(SHA-256 `ae7b7855…837280`). Regenerated files can differ in bytes; tests use
the frozen ones.

Font licensing: [DejaVu/Bitstream/Arev notices](../../../../e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-LICENSE).
PDFs use Git LFS.
