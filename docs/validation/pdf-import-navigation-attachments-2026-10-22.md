# R07 PDF navigation and attachment acceptance — 2026-10-22

## Source and scope

- Base: `62887ad6af62605411e1968da662be48ce1c091c` (`origin/main`, R06 merged).
- Final implementation and review fixes: `8343bae489fcf220f84d9e774ef4f54f33b46008`. The following evidence commit changes only the retained report and archive.
- Both dual-layer and vector importers were exercised with `CopyBookmarks` and `CopyAttachments` enabled. They remain off by default. See [PDF import: navigation and embedded attachments](../pdf-import-navigation-attachments.md) for the supported profile and fail-closed limits.
- The standard deterministic two-page PDF contains Chinese/English text, local styles, a table, affine graphics, a nested outline, a catalog embedded file and a page `FileAttachment`. The selected source page order is `[2, 1]`. Separate cropped single-page PDFs at 90° and 270° exercise annotation and bookmark coordinates with external `pdftoppm`, default PDFium, and native vector rendering.

## Functional and package checks

| Check | Result |
| --- | --- |
| 19 focused R07 regression tests (Release) | Passed |
| Full Release solution tests | 3,353 passed; 13 skipped across 11 assemblies |
| Source E2E `--r07-only` | Passed; OFD readback checked outline hierarchy/targets, attachment names and payload hashes, and `GotoA` binding; both rotations and both rasterizers covered |
| Default converter package consumer | Passed with 12 local `0.1.0-r07.4` packages |
| PDF vector package consumer | Passed with 14 local `0.1.0-r07.4` packages from a fresh cache and package references only; R07 probe/readback included |

The additional regressions cover outline grouping nodes, attachment count and decoded-size limits, dangling/duplicate/external `GotoA` targets, and 90°/270° displayed CropBox coordinates. The package consumer checks the same behavior through built packages.

## Rendering and visual checks

- Chain: generated source PDF → dual-layer OFD → exported PDF → macOS Preview; and generated source PDF → vector OFD → exported PDF → macOS Preview. The inspected PDFs came from the final `0.1.0-r07.4` package consumer.
- Automated rendering: all 11 PDFs rendered at 144 DPI. The standard source, dual and vector PDFs each have two pages; each rotated variant has one. The final package's 14 PNGs matched, byte for byte, individually inspected current-source E2E PNGs. Both default PDFium rotation pages were also inspected directly after the final fix.
- PNG visual review: all 14 pages. The standard outputs follow `[2, 1]`; Chinese/English text, local style, table fills/borders, affine graphics, footers and page boundaries are legible. On both cropped rotations, the blue source annotation rectangle, `PIN`/`TARGET` text and one output icon align. No new missing glyph, overlap, clipping, ghost text or blank page was observed.
- macOS Preview: all 14 pages across the 11 final package PDFs were opened and checked individually on 2026-10-22. Each PDF window was closed before the next, and Preview was quit afterward. Standard source/dual/vector pages 1–2 and 90°/270° source/external-dual/default-PDFium-dual/vector page 1 were inspected. The output icon appears once within the blue rectangle: upper-left at 90°, lower-right at 270°. No new output visual defect was observed for these pages.
- Preview does not paint the *source PDF* `FileAttachment` icon, leaving its blue rectangle empty. The source `pdftoppm -cropbox` PNG paints it. External-dual output retains that rasterized source icon; default PDFium-dual and native vector output draw a generic PushPin. This viewer difference was recorded separately from output acceptance.

Custom PDF annotation appearances are outside this profile. An independent OFD reader's `GotoA` click behavior and a real-world PDF corpus remain unverified; local OFD readback verifies the target and payload binding. These results apply only to the listed sample pages and conversion modes.

## Retained evidence

[R07 solid-compressed LFS archive](../evidence/pdf-import-navigation-attachments/r07-pdf-navigation-attachments-evidence.tar.zst) contains the three source PDFs, eight output OFDs, eleven PDFs, fourteen page PNGs, the readback report and an internal size/SHA-256 manifest. The [manifest](../evidence/pdf-import-navigation-attachments/manifest.json) is also available without extracting the archive. Archive size: 21,686,361 bytes; SHA-256: `811334eaaa9527e7bae9b1084894af22dd5c9d84f337e81342ea4e3d663892c3`.

The solid archive compresses repeated embedded fixture font data across the necessary output files. It excludes package caches, build output, full test logs, historical evidence and source checkout. Local detailed logs and generated files remain in ignored `artifacts/roadmap-v12/R07/`. The desktop inspection receipt is `artifacts/roadmap-v12/R07/visual-acceptance-r4.md` and GUI release generation `20261022T0318Z`.
