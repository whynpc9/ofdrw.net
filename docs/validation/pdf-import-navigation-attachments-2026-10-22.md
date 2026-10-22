# R07 PDF navigation and attachment acceptance — 2026-10-22

## Source and scope

- Base: `62887ad6af62605411e1968da662be48ce1c091c` (`origin/main`, R06 merged).
- Final source implementation: `eddc0a6576f3b08ca07d7e41123fa5c2bda38c54` (full SHA in the retained manifest). The following evidence commit changes only the report, manifest and archive.
- Both dual-layer and vector importers were exercised with `CopyBookmarks` and `CopyAttachments` enabled. They remain off by default. See [PDF import: navigation and embedded attachments](../pdf-import-navigation-attachments.md) for the supported profile and fail-closed limits.
- The standard deterministic two-page PDF contains Chinese/English text, local styles, a table, affine graphics, a nested outline, a catalog embedded file and a printable visible page `FileAttachment` (`/F 4`). The selected source page order is `[2, 1]`. Separate cropped sources at 90° and 270° exercise annotation/bookmark coordinates with external `pdftoppm`, default PDFium and native vector rendering.

## Functional and package checks

| Check | Result |
| --- | --- |
| 79 focused PDF metadata regression cases and one local/external ID contract case (Release) | Passed |
| Full Release solution tests | 3,415 passed; 13 skipped across 11 assemblies |
| Source E2E `--r07-only` | Passed; OFD readback checked hierarchy/targets, attachment names and payload hashes, `GotoA` binding, both rotations and both rasterizers |
| Default converter package consumer | Passed with 12 local `0.1.0-r07.11` packages |
| PDF vector package consumer | Passed with 14 local `0.1.0-r07.11` packages from a fresh cache and package references only; R07 readback included |
| Bookmark-only CropBox source probe | Passed for external dual and vector raster fallback with `CopyBookmarks=true`, `CopyAttachments=false`; first-page image 1100 × 800 px, child target `(480,20)` pt in displayed coordinates, no attachment or annotation output |
| Indirect-coordinate and reversed-rectangle probe | Passed in dual and vector modes: indirect `/XYZ` target `(480,20)` pt on a rotated CropBox page; `/Rect [40 50 20 30]` pin `(10,10,20,20)` pt; payload and page binding read back |
| Rejected annotation flag probes | `/F 0` (nonprintable) and `/F 6` (hidden) each failed before OFD output; focused tests also reject absent Print and NoView |

Additional regressions cover local GoTo action chains, malformed outline roots and unreachable /Last sibling-chain terminators, grouping nodes, unknown and artwork-bounding-box destinations, null, omitted and indirect numeric destination operands, XYZ zoom, reversed annotation rectangles, long colliding attachment names, decoded-size limits, dangling/ambiguous/external `GotoA` targets and 90°/270° displayed CropBox coordinates.

## Rendering and visual checks

- Chain: generated source PDF → dual-layer OFD → exported PDF → macOS Preview; and generated source PDF → vector OFD → exported PDF → macOS Preview. The 11 package PDFs came from the final `0.1.0-r07.11` consumer. The bookmark-only probe added source → external dual/vector-fallback OFD → exported PDF paths. The new edge probe added source with indirect bookmark coordinates and reversed attachment rectangle → dual/vector OFD → exported PDF paths.
- Automated rendering: all 19 final PDFs rendered at 144 DPI with displayed CropBox, producing 30 page PNGs. The 24 existing-profile page PNGs match the previously inspected r9/r10 pages byte for byte. The six new edge-probe pages were visually reviewed individually. The four new rejected-source pages were individually inspected: hidden `/F 6` omits the source icon; nonprintable `/F 0` still shows it in Poppler page rasterization, confirming why copying it would risk a visual/print mismatch. Both second pages retain legible `B中文`.
- PNG visual review: 14 package and six accepted probe pages match the prior individually accepted images byte for byte. On the cropped 90° probe, `A中文`, page bounds and the bookmark location align in source and both output chains. Four flagged source pages were inspected individually. The six edge pages show `A中文` on the rotated first page and `B中文` on the second; source, dual and vector align except for the expected generic vector pin shape. No missing glyph, overlap, clipping, ghost text or blank page was observed in these checked PNGs.
- macOS Preview: the r9 artifacts passed for 16 PDFs / 24 pages under GUI lease `20261022T0519Z` (ignored receipt `artifacts/roadmap-v12/R07/visual-acceptance-r9.md`). The regenerated r11 PDF containers, including the new edge probe, passed individual Preview inspection for 19 PDFs / 30 pages under GUI lease `20261022T0619Z` (ignored receipt `artifacts/roadmap-v12/R07/visual-acceptance-r11.md`). Every file was reopened by absolute path and closed; Preview was quit after the final page. The 24 existing-profile rendered page PNGs are byte-identical to those already inspected.
- Preview may suppress a source PDF `FileAttachment` icon that Poppler paints. External-dual output retains a rasterized source icon; default PDFium-dual and native vector output draw a generic PushPin. This viewer difference is recorded separately from output acceptance.

Custom PDF annotation appearances, unsupported flag combinations and FitB/FitBH/FitBV navigation are outside this profile. An independent OFD reader's `GotoA` click behavior and a real-world PDF corpus remain unverified; local OFD readback verifies target and payload binding. Conclusions apply only to the listed sample pages and modes.

## Retained evidence

[R07 solid-compressed LFS archive](../evidence/pdf-import-navigation-attachments/r07-pdf-navigation-attachments-evidence.tar.zst) contains the source PDFs, output OFDs, exported PDFs, 30 page PNGs, readback report and an internal size/SHA-256 manifest. The [manifest](../evidence/pdf-import-navigation-attachments/manifest.json) is also available without extracting the archive. Archive size: 22,709,820 bytes; SHA-256: `7f126b2930b6dec9ccb5a5a998bfa16e2acc846ff9be8a966367057708ccc0ed`. All 62 files and the internal manifest passed a decompression and hash round trip.

The solid archive compresses repeated fixture font data. It excludes package caches, build output, full test logs, historical evidence and source checkout. Local detailed logs and generated files remain in ignored `artifacts/roadmap-v12/R07/`. The completed r9 desktop receipt is `artifacts/roadmap-v12/R07/visual-acceptance-r9.md`; the completed r11 receipt is `artifacts/roadmap-v12/R07/visual-acceptance-r11.md`.
