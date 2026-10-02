# Issue 03 verification evidence

The four API/CLI tools are implemented. Acceptance remains **pending** until macOS Preview, the new SVG font check and current-head review close. Functional tests, PNG inspection and Preview are separate gates.

| Layer | Recorded result |
| --- | --- |
| Functional regression | Independent Sol Low: 274/274, zero failures/skips; bounded Astra High source review found no remaining actionable issue |
| Local package consumption | All 11 packages at `0.1.0-issue03.23`, runtime source `8b84794`; not published |
| API/CLI matrix | 25 synthetic samples, 42 PDF pages and 14 SVG renders; exact Native/default text, selected-page order, clean body bytes and watermark counts checked |
| PNG visual review | All 56 current renders retained in 15 contact sheets; 55 pass listed expectations; new fixed-anchor SVG font fidelity remains unverified |
| macOS Preview | Historical 10/37 pages inspected at `c10f002`; Mac relocked. Latest regenerated 42 pages require reopening. Historical reviewed PDF bytes are retained |
| CI / automated review | Previous `e08e5ae` all six checks passed; vendor annotation/template declaration ordering and resource-only ID reservation repaired; fresh-head Codex/Cursor/CI pending; Preview thread remains open |

The actual chain is **DOCX → explicit Native/default OFD → tool OFD → PDF/SVG → PNG**. `generated-layout.docx` changes only font declarations to pinned OFL Noto Sans CJK SC. Chinese, proportional English, local bold/italic/color, table fills/borders, alignment and explicit page break remain. Native/default TextCode is compared exactly against all 202 source characters. The example embeds the measurement face through the public model, preserving the converter's viewer-local CJK defaults.

Watermarks remain after merge. Split selects pages 2,1. Mix keeps the first physical box and paints green TOP LAYER over blue UNDER LAYER. Clean removes the synthetic signature square at (175,240), preserves every non-signature entry byte, and leaves the ordinary watermark at (175,215). Synthetic signing does not establish cryptographic validity.

Fixtures also cover exact sheared appearance/image clips, nested PageBlock, rotated/scaled glyphs, generated versus unmarked italic matrices, and missing annotation assets. Hidden annotation text/path never paint or enter visible extraction. Known artwork survives unknown outer metadata during ordinary export; Mix rejects metadata it cannot preserve. Unsupported drawing descendants, resource references, malformed text/path leaves and empty/foreign clipping geometry fail before partial export. Unknown annotation records stay page-local when their IDs resolve; ambiguous records remain shared opaque metadata.

Retained XML of any suffix participates in page/template/resource/signature closure. Nested extension `BaseLoc` resolves the `state.dat` template dependency; split keeps the template while selecting only page 2. Invalid XML retains uncertain payloads, while BOM-prefixed plain text does not disable pruning. Mix rejects unmodeled DocBody extensions. Vendor annotation literals are classified before path resolution/deduplication, so they remain preserved and cannot suppress the standard NOTE. Same-ID vendor templates stay outside the standard template table. Content-based ID reservation prevents resource-only IDs from colliding with new objects even with a stale document hint; equivalent XML and dat/bin cases pass. Owned resource descriptors with arbitrary suffixes retain font/image payloads; the actual `.dat`/`.bin` watermark sample passes PDF/SVG inspection, and unowned or malformed tables remain protected. Original MIT rectangle glyphs reproduce the Issue 02 fixed-baseline shear in `italic-fixed-anchor`: PDF shape/slope and regular/italic regression pairs pass. Librsvg substitutes the new SVG's font; Ego confirms both embedded faces loaded, but browser screenshot timed out while the desktop was locked. That SVG font-fidelity check is explicitly incomplete.

`evidence.tar.zst` contains source/font/licenses, actual artifacts, current/historical contact sheets, validation logs, and five historical Preview PDFs. `manifest.json` records size/SHA-256 for all **319 files**. The **39.4 MiB** archive expands to about 1.39 GiB. Extraction uses `zstd` with long-window support and about 1 GiB of memory.

```bash
python3 scripts/unpack-document-tools-evidence.py docs/evidence/document-tools/files
python3 scripts/verify-document-tools-evidence.py docs/evidence/document-tools/files
OFDRW_DOCUMENT_TOOLS_FONTS=docs/evidence/document-tools/files/fonts ./scripts/run-document-tools-e2e.sh
```

Runtime/example source is `8b84794`. Independent tests ran at `1f1e61d`; source/test/build inputs are identical, with only five validator assertions added afterward and exercised by the complete matrix. [acceptance.json](acceptance.json) records exact scope, hashes, sizes and limitations; [preview-progress.json](preview-progress.json) preserves the historical partial inspection. [Issue 02 integration notes](../../document-tools-integration-notes.md) describe shared contracts and the read-only italic comparison. The branches have not been integrated. These results apply only to listed fixtures/pages; PNG review does not replace Preview or prove arbitrary Word/OFD fidelity, production signing or target-reader interoperability.
