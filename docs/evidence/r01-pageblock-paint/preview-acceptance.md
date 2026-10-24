# R01 p105 current-source macOS Preview record

Product source: `baeae53ddc0a24a35c7e909e973c8079ccc86bcc`; PR head before Preview: `aa4ff3c42a9f939d18ac8b6677d228064bbbd21e`.
Exact-file request: `/Users/wanghongyi/Projects/ofdrw.net/artifacts/roadmap-v12-20261021/resource-leases/gui-request-R01-20261023T1925Z.json`; request SHA-256: `240281f9032937483c1c63fe1db7b84d678720ba152e81cdeb85146b3eee9f24`.

Status: **macOS Preview inspection complete for all 39 exact PDFs and 42 pages**. The expected column is based on current-source PDF-to-PNG preparation; the actual column records inspection in Preview on 2026-10-24.

Each PDF was opened from its absolute path in the request. Preview showed the exact file URL and page count; each page was visually inspected. The 39 PDF hashes were checked against the request before GUI inspection. All task PDF windows were closed after inspection.

| # | PDF / page | Expected in Preview | Actual Preview observation |
| --- | --- | --- | --- |
| 01.1 | `numeric-font-ids.pdf` p1/1 | Two clear lines: pageblock font, then composite font; same black proportional face. | PASS — matched expected appearance in Preview. |
| 02.1 | `numeric-font-face-ids.pdf` p1/1 | Two bold black lines: direct face, then nested face. | PASS — matched expected appearance in Preview. |
| 03.1 | `numeric-media-ids.pdf` p1/1 | Solid red square near upper left; no missing-image icon. | PASS — matched expected appearance in Preview. |
| 04.1 | `numeric-media-ids-template.pdf` p1/1 | Solid red square near upper left on template page. | PASS — matched expected appearance in Preview. |
| 05.1 | `media-type-omitted.pdf` p1/1 | Solid red square near upper left. | PASS — matched expected appearance in Preview. |
| 06.1 | `numeric-composite-ids.pdf` p1/1 | Black numeric composite text near upper left. | PASS — matched expected appearance in Preview. |
| 07.1 | `composite-paint-resources.pdf` p1/1 | Red square left and blue square right, both complete. | PASS — matched expected appearance in Preview. |
| 08.1 | `numeric-paint-ids.pdf` p1/1 | Solid red square near upper left. | PASS — matched expected appearance in Preview. |
| 09.1 | `pageblock-arc.pdf` p1/1 | Open blue upper semicircle stroke with clean endpoints. | PASS — matched expected appearance in Preview. |
| 10.1 | `pageblock-clip-arc.pdf` p1/1 | Blue arch with rounded top and flat base; clipped contour is clean. | PASS — matched expected appearance in Preview. |
| 11.1 | `pageblock-clip-arc-default.pdf` p1/1 | Blue arch with rounded top and flat base; matches the explicit clip case. | PASS — matched expected appearance in Preview. |
| 12.1 | `pageblock-comment-whitespace.pdf` p1/1 | Blue right-side triangle with clean diagonal and edges. | PASS — matched expected appearance in Preview. |
| 13.1 | `clip-whitespace-flags.pdf` p1/1 | Blue left-side triangle with clean diagonal and edges. | PASS — matched expected appearance in Preview. |
| 14.1 | `layer09-block-unqualified.pdf` p1/1 | One black mixed block text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 15.1 | `layer16-block-unqualified.pdf` p1/1 | One black mixed block text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 16.1 | `layer09-block16.pdf` p1/1 | One black mixed block text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 17.1 | `layer16-block09.pdf` p1/1 | One black mixed block text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 18.1 | `block16-composite09.pdf` p1/1 | One black mixed composite text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 19.1 | `block16-composite-unqualified.pdf` p1/1 | One black mixed composite text line; no duplicate. | PASS — matched expected appearance in Preview. |
| 20.1 | `template-own-font.pdf` p1/2 | Black shared template text near upper left and black triangular shape lower right. | PASS — matched expected appearance in Preview. |
| 20.2 | `template-own-font.pdf` p2/2 | Same shared template text and triangle on second page; no template text loss. | PASS — matched expected appearance in Preview. |
| 21.1 | `textcode-inherited-start.pdf` p1/1 | Three intentionally co-started words overlap near upper left; fourth is separate below. Check against sample, not as an accidental new defect. | PASS — matched expected appearance in Preview. |
| 22.1 | `textcode-inherited-escape.pdf` p1/1 | Lead at left; escaped text to its right; next line below; no literal escape code shown. | PASS — matched expected appearance in Preview. |
| 23.1 | `unqualified-pageblock-clip.pdf` p1/1 | Blue left-side clipped triangle. | PASS — matched expected appearance in Preview. |
| 24.1 | `jpeg-media-pageblock.pdf` p1/1 | Solid red square from JPEG near upper left. | PASS — matched expected appearance in Preview. |
| 25.1 | `template-local-composite.pdf` p1/1 | Blue left-side triangle from template-local composite. | PASS — matched expected appearance in Preview. |
| 26.1 | `page-local-composite.pdf` p1/1 | Blue left-side triangle from page-local composite. | PASS — matched expected appearance in Preview. |
| 27.1 | `omitted-format-jpeg-pageblock.pdf` p1/1 | Solid red square, matching explicit JPEG case. | PASS — matched expected appearance in Preview. |
| 28.1 | `shared-res-drawparam.pdf` p1/1 | Solid red square from shared Res DrawParam. | PASS — matched expected appearance in Preview. |
| 29.1 | `alias-res-colorspace.pdf` p1/1 | Solid red square from normalized Res alias. | PASS — matched expected appearance in Preview. |
| 30.1 | `native-v1.pdf` p1/2 | DOCX -> explicit native OFD -> PDF: Chinese blue title and section heading, gray italic English subtitle, complete mixed Chinese/English paragraph, light-blue table header, blue borders, Alpha/Beta rows. | PASS — matched expected appearance in Preview. |
| 30.2 | `native-v1.pdf` p2/2 | Second page: blue section heading, complete Chinese sentence with local red bold phrase, date 2026-01-01; no extra page or clipping. | PASS — matched expected appearance in Preview. |
| 31.1 | `default-v4.pdf` p1/2 | DOCX -> default OFD -> PDF: same two-page layout; Chinese blue title and section heading, gray italic subtitle, mixed paragraph, light-blue table header and blue borders, Alpha/Beta rows. | PASS — matched expected appearance in Preview. |
| 31.2 | `default-v4.pdf` p2/2 | Second page: blue section heading, local red bold phrase and date 2026-01-01; no extra page or clipping. | PASS — matched expected appearance in Preview. |
| 32.1 | `composite-single-clips.pdf` p1/1 | Blue clipped rectangle left half and readable black clip sentinel across its top; no extra blue area. | PASS — matched expected appearance in Preview. |
| 33.1 | `document-composite-2016-res.pdf` p1/1 | One black resource namespace text line near upper left. | PASS — matched expected appearance in Preview. |
| 34.1 | `public-composite-unqualified-res.pdf` p1/1 | One black resource namespace text line near upper left. | PASS — matched expected appearance in Preview. |
| 35.1 | `nested-composite-pageblock.pdf` p1/1 | One black nested resource text line near upper left. | PASS — matched expected appearance in Preview. |
| 36.1 | `unreferenced-resource-pageblock-image.pdf` p1/1 | One black visible text line; malformed unreferenced image remains absent. | PASS — matched expected appearance in Preview. |
| 37.1 | `hidden-large-pageblock-image.pdf` p1/1 | One black visible text line; hidden large image remains absent. | PASS — matched expected appearance in Preview. |
| 38.1 | `ordinary-resource-pageblock.pdf` p1/1 | One black ordinary resource text line near upper left. | PASS — matched expected appearance in Preview. |
| 39.1 | `typed-source-clear.pdf` p1/1 | One blue typed edit text line; old source text absent. | PASS — matched expected appearance in Preview. |

## Closeout fields

- Preview application and actual viewing chain: macOS Preview; PageBlock OFD -> PDF -> Preview and generated-layout.docx -> explicit native/default OFD -> PDF -> Preview. Actual PDFs are the p105 current-source outputs in the GUI request.
- First and last inspection time (Asia/Shanghai): after 2026-10-24 23:07 (GUI grant) through 2026-10-24 23:21.
- PDF files/pages inspected: 39/39, 42/42. All exact PDF URLs and page counts were checked in Preview.
- Unexpected missing glyphs, color/style spread, table errors, clipping, overlap, duplicate text, or blank pages: none observed in the inspected pages. The three co-started words in textcode-inherited-start are intentional fixture overlap.
- Transient Preview/UI capture behavior: pageblock-arc first painted white, then displayed its blue open arc after reopening; public-composite-unqualified-res first capture was white, then displayed its text after reopening; nested-composite-pageblock first captures were thumbnail-sized, then a full-size capture showed its text. alias-res-colorspace first screenshot failed with zero-size capture and the next screenshot showed its red square. No content defect remained after refresh. Page PNGs and PDF hashes are preserved under ignored artifacts/roadmap-v12/R01/.
- Task Preview windows closed; unrelated windows left untouched: yes. The Preview Window menu listed only numeric-font-ids, numeric-font-face-ids, and numeric-media-ids after the other task windows were closed. Those three were closed individually; the Window menu then listed no document window. Preview’s pre-existing Open panel remains.
- Final visual acceptance result: PASS for these 39 PDFs/42 pages in macOS Preview, including the explicit native and default DOCX conversion chains. This does not establish arbitrary DOCX fidelity.
