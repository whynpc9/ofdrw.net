# Issue 03 verification evidence

Implementation: all four API/CLI paths are present. Complete acceptance remains **pending** until macOS Preview and current-head PR review close. This record separates functional and visual results; green tests and PNG review do not complete Preview acceptance.

| Layer | Current result |
| --- | --- |
| Functional regression | 181/181 passed locally; Astra design/source review and Sol Low independent verification performed; later Astra findings fixed with dedicated regressions |
| 11-package local consumption | Passed: 11 packages at `0.1.0-issue03.6`, built from runtime commit `a4425bb`; not published |
| API/CLI automatic matrix | 20 public synthetic samples, 34 PDF pages, nine rendered SVG pages; body-byte preservation, selected text/page order, watermark text counts and file hashes checked |
| PNG visual review | Passed for the 34 PDF pages and 9 SVG renders listed in `acceptance.json`; contact sheets plus standalone checks of changed overlay pages/SVG |
| macOS Preview | **Not completed**: Computer Use returned “The Mac is locked and automatic unlock could not unlock it.” User unlock requested; no Preview document window opened |
| GitHub CI / Codex / Cursor | First `28113dd` Codex/Cursor Automation reviews complete; functional findings fixed and re-review pending. Preview thread remains unresolved |

The sample is `generated-layout.docx` with only font declarations replaced by pinned Noto Sans CJK SC. Content includes proportional Latin, Chinese, bold/italic/color, table fills/borders, alignment and explicit page break. The source copy and Noto/OFL license are retained. The rich OFD adds a template, annotation, public attachment, shared font and deliberately noncryptographic signature appearance.

The actual path is **DOCX → explicit Native/default OFD → each tool result OFD → PDF → PNG**. Preview will open those same result PDFs after unlock. Native/default outputs have two pages. Watermark is visible on page 1; merge retains it and adds a third page. Split selects original pages 2,1. Mix uses source page 1's physical size and overlays a green TOP LAYER box over the blue UNDER LAYER template box. Clean removes the red signature square at (175,240), retains body/template/annotation/attachment bytes, and `verify-signatures` reports no declarations. Ordinary red watermark image at (175,215) is separate content.

The committed `evidence.tar.zst` contains real OFD/PDF/SVG files, every rendered page PNG, source DOCX/font/license and validation logs. `manifest.json` provides per-file bytes and SHA-256; no private source data is used. This keeps the review evidence accessible without relying on an ignored local artifacts directory. SVG contains full embedded fonts, so the archive uses long-window compression to deduplicate repeated payloads.

```bash
python3 scripts/unpack-document-tools-evidence.py docs/evidence/document-tools/files
python3 scripts/verify-document-tools-evidence.py docs/evidence/document-tools/files
# Reproduce with the exact bundled font:
OFDRW_DOCUMENT_TOOLS_FONTS=docs/evidence/document-tools/files/fonts ./scripts/run-document-tools-e2e.sh
```

Runtime source `a4425bb` passed 181/181 under independent Sol Low verification; example source is `a4425bb`. Source baseline, environment, bundle hash, actual checked page list and sizes are recorded in [acceptance.json](acceptance.json). The 34.2 MiB bundle was extracted into an independent temp directory and all 194 file sizes/hashes matched. Extraction needs `zstd` and approximately 1 GiB of memory; the expanded SVG/font evidence occupies about 1,062 MiB. The example embeds the same licensed face used by Native measurement through the public font model, retaining self-contained outputs without changing viewer-local CJK defaults. Conclusions apply to these synthetic pages and the tested structural matrix. Unsupported raw objects/actions/resource references fail explicitly; this is not a claim of arbitrary complex OFD or Word fidelity, production signing, or target-reader interoperability.

Review-fix samples `annotation-clipped` and `annotation-clipped-mix` retain the exact sheared clip, 90-degree ROTATE glyphs, 1.5x SCALE glyphs and nested PageBlock content. Raster regressions additionally check original image-boundary clipping when image CTM exceeds its box. The native Preview gate remains uncompleted for all samples.

Second-round fixes limit all mutating XML scans to owned declarations, qualified direct nodes and typed implicit resources. The three italic samples distinguish generated single-slant factor compensation from explicit unmarked user CTM. Parallel-ticket integration boundaries are recorded in [Issue 02 integration notes](../../document-tools-integration-notes.md).

Third-round namespace regressions check complete XName document/page allowlists and all annotation index/container/primitive boundaries. Vendor content remains preserved, is not exposed by extract-text, and is rejected by Mix. Current 43 re-rendered PNG files are byte-identical to the R3 pages actually inspected.
