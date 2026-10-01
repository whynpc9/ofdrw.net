# Issue 03 verification evidence

Implementation: all four API/CLI paths are present. Complete acceptance remains **pending** until macOS Preview and current-head PR review close. This record separates functional and visual results; green tests and PNG review do not complete Preview acceptance.

| Layer | Current result |
| --- | --- |
| Functional regression | 149/149 passed locally; Astra design/source review and Sol Low independent verification performed; later Astra findings fixed with dedicated regressions |
| 11-package local consumption | Passed on implementation snapshot; final committed candidate readback pending |
| API/CLI automatic matrix | 15 public synthetic samples, 29 PDF pages, four rendered SVG watermark pages; body-byte preservation, selected text/page order, watermark text counts and file hashes checked |
| PNG visual review | Final Noto render review pending; initial system-font pages were inspected, superseded by redistributable Noto evidence |
| macOS Preview | **Not completed**: Computer Use returned “The Mac is locked and automatic unlock could not unlock it.” User unlock requested; no Preview document window opened |
| GitHub CI / Codex / Cursor | Pending PR creation and current-head readback |

The sample is `generated-layout.docx` with only font declarations replaced by pinned Noto Sans CJK SC. Content includes proportional Latin, Chinese, bold/italic/color, table fills/borders, alignment and explicit page break. The source copy and Noto/OFL license are retained. The rich OFD adds a template, annotation, public attachment, shared font and deliberately noncryptographic signature appearance.

The actual path is **DOCX → explicit Native/default OFD → each tool result OFD → PDF → PNG**. Preview will open those same result PDFs after unlock. Native/default outputs have two pages. Watermark is visible on page 1; merge retains it and adds a third page. Split selects original pages 2,1. Mix uses source page 1's physical size and overlays a green TOP LAYER box over the blue UNDER LAYER template box. Clean removes the red signature square at (175,240), retains body/template/annotation/attachment bytes, and `verify-signatures` reports no declarations. Ordinary red watermark image at (175,215) is separate content.

The committed `evidence.tar.zst` contains real OFD/PDF/SVG files, every rendered page PNG, source DOCX/font/license and validation logs. `manifest.json` provides per-file bytes and SHA-256; no private source data is used. This keeps the review evidence accessible without relying on an ignored local artifacts directory. SVG contains full embedded fonts, so the archive uses long-window compression to deduplicate repeated payloads.

```bash
python3 scripts/unpack-document-tools-evidence.py docs/evidence/document-tools/files
python3 scripts/verify-document-tools-evidence.py docs/evidence/document-tools/files
# Reproduce with the exact bundled font:
OFDRW_DOCUMENT_TOOLS_FONTS=docs/evidence/document-tools/files/fonts ./scripts/run-document-tools-e2e.sh
```

Source baseline, environment, bundle hash, actual checked page list, sizes and PR head are recorded in `acceptance.json` after the final generation. Conclusions apply to these synthetic pages and the tested structural matrix. Unsupported raw objects/actions/resource references fail explicitly; this is not a claim of arbitrary complex OFD or Word fidelity, production signing, or target-reader interoperability.
