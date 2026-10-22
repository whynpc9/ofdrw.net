# R07 PDF navigation and attachment acceptance — 2026-10-22

## Source and scope

- Base: `62887ad6af62605411e1968da662be48ce1c091c` (`origin/main`, R06 merged).
- Final source implementation: `7dc4e5170f4422ffe9a5a13af4ef18ba0a95d3b2` (full SHA in the retained manifest). The following evidence commit changes only the report, manifest and archive.
- Both dual-layer and vector importers were exercised with `CopyBookmarks` and `CopyAttachments` enabled. They remain off by default. See [PDF import: navigation and embedded attachments](../pdf-import-navigation-attachments.md) for the supported profile and fail-closed limits.
- The standard deterministic two-page PDF contains Chinese/English text, local styles, a table, affine graphics, a nested outline, a catalog embedded file and a printable visible page `FileAttachment` (`/F 4`). The selected source page order is `[2, 1]`. Separate cropped sources at 90° and 270° exercise annotation/bookmark coordinates with external `pdftoppm`, default PDFium and native vector rendering.

## Functional and package checks

| Check | Result |
| --- | --- |
| 59 focused PDF metadata regression cases and one local/external ID contract case (Release) | Passed |
| Full Release solution tests | 3,395 passed; 13 skipped across 11 assemblies |
| Source E2E `--r07-only` | Passed; OFD readback checked hierarchy/targets, attachment names and payload hashes, `GotoA` binding, both rotations and both rasterizers |
| Default converter package consumer | Passed with 12 local `0.1.0-r07.9` packages |
| PDF vector package consumer | Passed with 14 local `0.1.0-r07.9` packages from a fresh cache and package references only; R07 readback included |
| Bookmark-only CropBox source probe | Passed for external dual and vector raster fallback with `CopyBookmarks=true`, `CopyAttachments=false`; first-page image 1100 × 800 px, child target `(480,20)` pt in displayed coordinates, no attachment or annotation output |
| Rejected annotation flag probes | `/F 0` (nonprintable) and `/F 6` (hidden) each failed before OFD output; focused tests also reject absent Print and NoView |

Additional regressions cover local GoTo action chains, malformed outline roots, grouping nodes, unknown and artwork-bounding-box destinations, null and omitted destination operands, XYZ zoom, long colliding attachment names, decoded-size limits, dangling/ambiguous/external `GotoA` targets and 90°/270° displayed CropBox coordinates.

## Rendering and visual checks

- Chain: generated source PDF → dual-layer OFD → exported PDF → macOS Preview; and generated source PDF → vector OFD → exported PDF → macOS Preview. The 11 package PDFs came from the final `0.1.0-r07.9` consumer. The bookmark-only probe added source → external dual/vector-fallback OFD → exported PDF paths.
- Automated rendering: all 16 final PDFs rendered at 144 DPI with displayed CropBox, producing 24 page PNGs. The 14 package pages and six accepted bookmark-only pages match previously inspected pages byte for byte. The four new rejected-source pages were individually inspected: hidden `/F 6` omits the source icon; nonprintable `/F 0` still shows it in Poppler page rasterization, confirming why copying it would risk a visual/print mismatch. Both second pages retain legible `B中文`.
- PNG visual review: 14 package and six accepted probe pages match the prior individually accepted images byte for byte. On the cropped 90° probe, `A中文`, page bounds and the bookmark location align in source and both output chains. Four flagged source pages were inspected individually. No missing glyph, overlap, clipping, ghost text or blank page was observed in these checked PNGs.
- macOS Preview: passed for the 16 exact current PDFs / 24 pages under GUI lease `20261022T0519Z`. Each current file was reopened by absolute path, its URL and page count checked, and every page viewed. Source and exported dual/vector pages retained the expected text, geometry, styles and icons without a conversion-specific visual defect. Preview was quit after review. See the ignored desktop receipt `artifacts/roadmap-v12/R07/visual-acceptance-r9.md` for per-file hashes and pages.
- Preview may suppress a source PDF `FileAttachment` icon that Poppler paints. External-dual output retains a rasterized source icon; default PDFium-dual and native vector output draw a generic PushPin. This viewer difference is recorded separately from output acceptance.

Custom PDF annotation appearances, unsupported flag combinations and FitB/FitBH/FitBV navigation are outside this profile. An independent OFD reader's `GotoA` click behavior and a real-world PDF corpus remain unverified; local OFD readback verifies target and payload binding. Conclusions apply only to the listed sample pages and modes.

## Retained evidence

[R07 solid-compressed LFS archive](../evidence/pdf-import-navigation-attachments/r07-pdf-navigation-attachments-evidence.tar.zst) contains the source PDFs, output OFDs, exported PDFs, 24 page PNGs, readback report and an internal size/SHA-256 manifest. The [manifest](../evidence/pdf-import-navigation-attachments/manifest.json) is also available without extracting the archive. Archive size: 21,904,067 bytes; SHA-256: `569136fbaf0bb2317e1a8ab4891bc47aa687e5ad6d94ed37bc85bb5a0268cf69`. All 51 files and the internal manifest passed a decompression and hash round trip.

The solid archive compresses repeated fixture font data. It excludes package caches, build output, full test logs, historical evidence and source checkout. Local detailed logs and generated files remain in ignored `artifacts/roadmap-v12/R07/`. The completed desktop inspection receipt is `artifacts/roadmap-v12/R07/visual-acceptance-r9.md`.
