# Issue 03 verification evidence

All four API/CLI paths are implemented. Acceptance remains **pending** until macOS Preview and current-head review close. Functional tests, rendered PNG review and Preview are separate gates.

| Layer | Recorded result |
| --- | --- |
| Functional regression | Independent Sol Low: 222/222, zero failures/skips; Astra High source review findings repaired |
| Local package consumption | All 11 packages at `0.1.0-issue03.18`, from runtime source `13e0d6c`; not published |
| API/CLI matrix | 21 public synthetic samples, 36 PDF pages and 10 SVG renders; exact Native/default text, selected-page order, clean body bytes and watermark counts checked |
| PNG visual review | Every current page inspected in 12 retained contact sheets with targeted crops; final 46 page hashes match those inspected bytes |
| macOS Preview | **Not completed**: Mac locked and automatic unlock failed; no Preview document window opened |
| CI / automated review | Previous head `e137c99` all six checks succeeded; per-Area and annotation asset/index findings repaired; new-head Codex/Cursor/CI re-review pending; Preview thread open |

The sample is `generated-layout.docx` with only font declarations replaced by pinned, licensed Noto Sans CJK SC. It retains Chinese, proportional English, local bold/italic/color, table fills/borders, alignment and explicit page break. Source DOCX and Noto/OFL are bundled. Native/default text is compared exactly against the 202 source characters. The example explicitly embeds the same measurement font through the public model without changing the converter's viewer-local CJK defaults.

The actual chain is **DOCX → explicit Native/default OFD → tool OFD → PDF/SVG → PNG**. Those same PDFs require Preview after unlock. Watermark appears on page 1 and remains after merge. Split selects pages 2,1. Mix keeps the first physical box and places green TOP LAYER over blue UNDER LAYER at identical coordinates. Clean removes the synthetic signature square at (175,240) while retaining all non-signature entry bytes; the ordinary watermark image at (175,215) remains independent. Synthetic signing does not claim cryptographic validity.

Additional fixtures cover exact sheared appearance/image clips, nested PageBlock, rotated/scaled glyphs, and generated versus unmarked explicit italic CTM. Outer metadata keeps known NOTE artwork in ordinary export while preventing Mix. Hidden annotation text/path retain XML but never paint or enter visible text extraction. Unsupported children, same-namespace leaf extensions, unimplemented drawing references and clips with any unsupported/foreign Area or without supported literal paths fail explicitly. Page/template data uses direct literal text; edited path literals preserve extension subtrees. Qualified resource/annotation/template ownership and XML-named signature payload shared-reference closure have dedicated regressions. Text watermarks remain usable under a policy forbidding images.

`evidence.tar.zst` contains actual OFD/PDF/SVG/PNG, all 12 current-page contact sheets, source/font/license, environment and validation logs. `manifest.json` records size/SHA-256 for all **222 files**. The **35.7 MiB** archive was independently extracted and verified. Extraction requires `zstd` with long-window support and about 1 GiB of memory; expanded evidence occupies about 1.1 GiB because SVG embeds full fonts.

```bash
python3 scripts/unpack-document-tools-evidence.py docs/evidence/document-tools/files
python3 scripts/verify-document-tools-evidence.py docs/evidence/document-tools/files
OFDRW_DOCUMENT_TOOLS_FONTS=docs/evidence/document-tools/files/fonts ./scripts/run-document-tools-e2e.sh
```

Runtime/example source: `13e0d6c`. See [acceptance.json](acceptance.json) for exact baseline, viewed files, environment, sizes and bundle hashes. [Issue 02 integration notes](../../document-tools-integration-notes.md) describe shared contracts; the branches have not been integrated. Findings apply only to these fixtures/pages and structural tests, not arbitrary Word/OFD fidelity, production signing or target-reader interoperability. PNG inspection does not replace Preview.

Annotation assets must resolve: unknown media IDs, missing payloads and missing declared page files reject the whole group without partial export. Unknown records with PageID affect only that page; ambiguous records use a shared aggregate rather than a pages-times-records allocation. Bundled `annotation-unknown-media` and `annotation-missing-payload` inputs include rejection logs. Actual API `mix` uses `scoped-annotations-input`: unmodeled page2 records block that page while page1 plus overlay retains the inspected output.
