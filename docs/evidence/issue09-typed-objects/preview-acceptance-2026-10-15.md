# macOS Preview visual acceptance — 2026-10-15

**PASS for the listed five PDFs and eight pages.** Opened the original generated files in macOS Preview, verified native window path/title and page count, and inspected every page using native screenshots in Single Page view. This completes the previously blocked Preview gate. [Exact file identities and observations](preview-acceptance-2026-10-15.json).

| Actual file | Pages inspected | Visual result |
| --- | --- | --- |
| `low-60163cc/consumer-output/before.pdf` | 1–2 | Chinese/Latin, two shared nested placements, styles, image, table, clip/fill examples and annotation correct |
| `low-60163cc/consumer-output/after.pdf` | 1–2 | Intended left-placement and Area CTM edits visible; right placement unchanged; no duplication or unexpected layout change |
| `low-60163cc/consumer-output/original-v1.pdf` | 1–2 | Exact original V1 OFD newly exported by candidate package; full table fill/borders, no diagonal bridge or triangular hole; local annotation, date and second page correct |
| `pr20/document-tools-repair/binary-xml-split.pdf` | 1 | Page-two text, local red bold style, date and whitespace preserved |
| `pr20/document-tools-repair/split-template-liveness.pdf` | 1 | Same expected retained page, no stray template artwork or extra page |

The primary chain is **public PackageReference API → native OFD → PDF → Preview**. The two affected DocumentTools cases follow the actual native DOCX OFD through Split → OFD → PDF → Preview. No direct DOCX → PDF surrogate was used. The intended 38 mm clipped text and viewport overflow masks were checked against the generator's explicit intent; these are not unreported layout defects.

No missing glyph, unwanted clipping, overlap, ghosting, duplicate annotation, diagonal table rule, triangular fill gap or anomalous blank page was observed on these eight pages. NonZero remains solid and EvenOdd has a hole. Explicit footer labels and V1 dates are present; there is no claim to have tested automatic header/footer or page-number features.

Product source remains **60163cc**, candidate packages remain `0.1.0-issue09.c.60163cc`, and the branch head inspected before this documentation update is **a1cebe3**. The existing OFD/PDF bytes match their preserved source records; no product was regenerated and no font payload, functional test, package or budget gate was changed/repeated. Functional results remain independently scoped: 1320/1320 required-font local tests, 20 dual-TFM builds and 11 local packages. PR review/CI was closed on a1cebe3; later documentation-only head checks are recorded in the PR.

Native Preview screenshots were examined live and are retained in this chat. The checked-in page PNGs remain auxiliary Poppler renders, accurately labelled; they are not substituted for this Preview session. The other 53 generated DocumentTools matrix pages and SVG browser output are outside this manual verdict. All windows opened for this session were closed, ending in the native `noWindowsAvailable` confirmation.

The earlier lock-screen failure and `PreviewCompleted: false` fields in raw generation/Low/fixture records are historical observations. They remain unchanged; this later record supplies the actual GUI verdict. No merge, tag, NuGet publication or release was performed.
