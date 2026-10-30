# R22 reader-method validation evidence

The optional LFS archive [`reader-method-evidence.tar.zst`](reader-method-evidence.tar.zst)
retains the preregistered six-PDF headless reader and geometry experiment at
`2f45eaa20990eff29cfbea4d1f4866d231517e50`. It contains the bounded
request, preregistration, diagnostic probe sources, raw PdfPig/PDFKit/Poppler
and PDF-structure JSON, fixed-criteria verdict, supplementary CID/GID map and
an entry-by-entry SHA-256 manifest. It contains no duplicated input PDFs,
fonts, packages, binaries, build output or full logs. The six exact input PDFs
are already in the [earlier reference archive](../r22-next-phase/README.md).

| Identity | Value |
| --- | --- |
| Archive bytes | `29,922` |
| Archive SHA-256 | `87891bdf3799a46534c17689f71309e455eea6dc500c3f143bd2c13f5a63d553` |
| Entries | 16, including `manifest.json` |
| Preregistration SHA-256 | `62a0c89682f30d52fb03c73ce0c6af8c1634c668e9b5dbd600435f1a56947243` |
| Fixed-criteria verdict | Proposed PDFKit character-index to PdfPig visible-letter association **falsified**; 0/8 same-scalar Hebrew positions matched in each shaped positive |
| Other observations | PDFKit logical text/Find and Poppler physical word rectangles passed the bounded positive controls; all negatives retained their failure classification |

Fetch this optional archive with `git lfs pull
--include="docs/evidence/r22-reader-method/**" --exclude=""`. Extract it
under ignored `artifacts/r22-reader-contract/method-validation/` before
running the scripts. The saved request contains absolute paths from the
original checkout; point them to the six hash-verified PDFs in the earlier
archive when reproducing elsewhere. The scripts do not generate PDFs.

The [validation record](../../validation/r22-reader-replacement-candidate.md)
explains the method failure and why the supplementary structural observation
does not establish a product or Preview pass. The heavy lease was released
after the bounded batch. No replacement acceptance method is frozen by this
archive.
