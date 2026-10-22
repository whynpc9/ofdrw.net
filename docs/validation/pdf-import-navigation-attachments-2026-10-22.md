# R07 PDF navigation and attachment acceptance — 2026-10-22

## Source and scope

- Base: `62887ad6af62605411e1968da662be48ce1c091c` (`origin/main`, R06 merged).
- Implementation: `29433b54716c92eb417a3f90c9365db095311738`.
- The opt-in PDF bookmark and attachment migration is exercised through both dual-layer and vector importers. The switches remain off by default. The supported profile and fail-closed boundaries are in [PDF import: navigation and embedded attachments](../pdf-import-navigation-attachments.md).
- Sample: a deterministic two-page PDF with Chinese and English, styled text, a table, affine graphics, a nested outline, one catalog embedded file, and a page `FileAttachment`. The selected source page order is `[2, 1]`.

## Functional and package checks

| Check | Result |
| --- | --- |
| 9 focused R07 regression tests (Debug) | Passed |
| Full Release solution tests | 3,343 passed; 13 skipped |
| Source E2E `--r07-only` | Passed; both importers' OFD readback checked outline targets, attachment names/payload hashes and `GotoA` attachment binding |
| Default converter package consumer | Passed with 12 local packages |
| PDF vector package consumer | Passed with 14 local packages from a fresh package cache; R07 probe and readback included |

No real-world PDF corpus or independent OFD reader interaction was part of this acceptance.

## Rendering and visual checks

- Chain: generated source PDF → dual-layer OFD → exported PDF → macOS Preview; and generated source PDF → vector OFD → exported PDF → macOS Preview. These are the package consumer's generated artifacts.
- Automated rendering: source PDF and both exported PDFs have two pages. The package output PDFs were rendered at 144 DPI. The four package output PNGs match the individually reviewed current-source E2E PNGs byte for byte.
- PNG visual review: source pages 1–2, dual output pages 1–2, vector output pages 1–2. The outputs follow `[2, 1]`; Chinese and English text, local style, table fills/borders, affine graphics, footer and page boundaries are legible. No new missing glyph, overlap, clipping, ghost text or blank page was observed.
- macOS Preview: all six pages were opened and inspected on 2026-10-22, with each file closed before opening the next and Preview quit afterward. The gray generic PushPin appears at the selected page attachment position on both outputs. No new visual defect was observed. This conclusion applies to this sample and these six pages.
- The PDF annotation's custom appearance is outside the profile; the output deliberately uses the generic PushPin. An independent OFD reader's `GotoA` click behavior remains unverified; local OFD readback verifies the target and payload binding.

## Retained evidence

[R07 evidence archive](../evidence/pdf-import-navigation-attachments/r07-pdf-navigation-attachments-evidence.zip) contains the source PDF, two OFDs, their exported PDFs, six page PNGs, the R07 readback report and an internal size/SHA-256 manifest. The [manifest](../evidence/pdf-import-navigation-attachments/manifest.json) is also available without extracting the archive. Archive size: 23,957,574 bytes; SHA-256: `420c2eefc048c532310eebeededa693622353cb5fe8e000d558cfb2e01a684ba`.

The two output OFDs contain the embedded fixture font; the archive does not duplicate the package cache, build output, complete test logs, historical evidence or source checkout. Local detailed logs and generation files remain in ignored `artifacts/roadmap-v12/R07/`.
