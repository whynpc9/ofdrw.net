# Native embedded-font bold compatibility — 2026-10-17

Baseline: `77825d3107f767e334cd80d6db5e651e958b23bf`. This is the appended repair commit containing this report. Worktree: `xplat-font-identity/ofdrw.net`. The ignored `artifacts/font-bold-round2/source-binding.json` records the full final SHA and the source/artifact identities. No merge, tag, publication, golden or tolerance change.

## Evidence and cause

The independent first-round retest preserved source text, original embedded font identity, cross-platform structure, and PDF appearance. Direct Suwell 5.0.22.0831 OFD rendering nevertheless lost bold emphasis on both pages. These are separate acceptance results.

Controlled probes changed only resource flags, then resource names, then explicit text painting. The source conversation's `artifacts/xplat/20261017-bold-fix-retest/probe-review.json` records four files/eight pages, matching input hashes, and 100%/200% comparisons. Clearing resource Bold/Italic or removing style-name suffixes did not recover bold. Standard text stroke did recover it, but the initial legacy name-only width (4.5% em, minimum 0.08 mm) was visibly heavier than the paired PDF at 200%. That diagnostic probe is **not** an accepted product output.

The supported inference is a missing explicit painting contract for this regular embedded face/reader combination. We have no Suwell implementation evidence proving its internal handling of Weight. [GB/T 33190-2016, tables 44–45, printed pages 61–64](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/GBT_33190-2016_%E7%94%B5%E5%AD%90%E6%96%87%E4%BB%B6%E5%AD%98%E5%82%A8%E4%B8%8E%E4%BA%A4%E6%8D%A2%E6%A0%BC%E5%BC%8F%E7%89%88%E5%BC%8F%E6%96%87%E6%A1%A3.pdf) distinguishes physical font attributes from per-text Weight/Italic and permits text Fill/Stroke/StrokeColor. The corresponding scanned pages were inspected locally; no font bytes or names are manufactured to claim a real bold face.

## Final candidate contract

- Resource Bold/Italic now describe the selected embedded font's actual OpenType style. TextObject Weight/Italic retain the requested run style.
- Native/default DOCX text requesting bold with an embedded regular face emits a single semantic TextObject with Fill and Stroke. Stroke color/alpha match the text fill. The original TextCode, placement and font payload remain intact.
- Embedded synthetic bold uses the **existing PDF strategy, 0.025 em**, through shared `FontEmphasisMetrics`. PDF drawing still uses exactly its prior width. This is synthetic emphasis, not the designer's real bold face and not a claim of identical rendering across arbitrary readers/fonts.
- `NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.TextStroke` is the default. `DOCX_EMBEDDED_BOLD_SIMULATED` explicitly reports the selected family's simulation once per conversion.
- For strict embedded bold, set `NativeEmbeddedBoldBehavior = DocxEmbeddedBoldBehavior.Throw` and supply actual bold font files via `FontDirectories`. Missing selected bold face raises `InvalidDataException` before writing output. This option governs embedded native text only; it is not a global missing-font policy. The CLI uses the default; the strict option is exposed by the .NET API.
- Actual physical bold faces receive no synthetic stroke. Imported embedded SourceXml owns painting, even if a stale internal simulation hint exists. Other embedded typed text acquires no hint implicitly. Clone preserves the hint for freshly generated native text and preserves source-owned painting. Existing name-only width/shear behavior stays unchanged; embedded italic remains on its existing route.

## Local validation

All commands used writable task-local DOTNET_CLI_HOME, no telemetry/first-time experience, single-node builds, disabled build servers/shared compilation and the existing NuGet cache.

- Functional tests: **116/116** native/generated DOCX and font/bold policy tests; **85/85** PDF identity/isolation/style tests; **1/1** packaging name-only style test. No skips. Exact native/default source text, strict no-partial-output failure, physical-face selection, no extra stroke for bold faces, SourceXml rewrite and clone boundaries are covered. The physical-style selection unit test reuses the existing tiny synthetic style fixture; it is not visual acceptance of a commercial bold font.
- Local package consumer: **PASS**, eleven packages at unpublished `0.1.0-xplat-bold-fix.1`, installed into an isolated consumer/cache. Existing macOS visual corpus full-page checks remain RMSE 0 under unchanged 0.035 gates. See `artifacts/package-e2e/0.1.0-xplat-bold-fix.1/` and `package-e2e.log`. Existing ImageSharp NU1902/NU1903 warnings remain outside this font repair.
- Trusted CJK fixture: SHA-256 `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa`. PDF identity tests validate the original glyph bytes/metrics with host resolution rejected, both subset and full embedding. No new binary fixture.
- Actual converter generation: `candidate-v2-macos/outputs/{original,noto}-{native,default}`. Four OFDs/eight pages; all exact OFD text checks and whitespace-normalized PDF text-once checks pass. All four canonical OFD structures match. OFD is about 290.7 KB, essentially unchanged from the first repair (~291 KB).
- Automatic rendering: all eight PDF pages at 144 DPI. Their PNG hashes equal the first repair's paired PDFs, confirming the shared width extraction did not change PDF rendering.
- PNG visual review: both unique final page images inspected; all four cases have identical per-page hashes. Chinese/English, proportional spacing, local bold/italic/color, table fill/borders, alignment and explicit page break pass for the represented content.
- macOS Preview: all four final PDFs freshly opened via **DOCX → native/default OFD → PDF → Preview**, pages 1–2 inspected and captured individually; all document windows closed. No missing glyphs, style spread, clipping, overlap, ghosting or abnormal blank pages seen. This fixture does not cover images, headers/footers or automatic page numbers.

Logs/TRX, generation manifests, final OFD/PDF/PNGs, eight Preview screenshots and `visual-acceptance.json` are under ignored `artifacts/font-bold-round2/`. Original acceptance directories remain unchanged.

## Required fixed-commit retest

| Gate | Scope | Status at local handoff |
| --- | --- | --- |
| Three platforms | macOS/Windows/WSL; same pinned Noto and two DOCXs × native/default | Source conversation to run on final SHA |
| Direct Suwell | Each final generated OFD, both pages, title/table/red emphasis at 100% and 200%, paired PDF comparison | **NOT COMPLETED**; probe04 and local Preview do not satisfy this gate |
| Font identity/text | Original glyphs/metrics through OFD/PDF; no host fallback, duplicates or text loss | Local affected tests PASS; target rerun required |
| Windows visual golden | Existing independently reviewed Windows baseline required | **BLOCKED**, unchanged; do not promote current output to golden |
| Windows Audit023 | Two unchanged symbolic-link cases on a suitably privileged isolated runner | **BLOCKED by existing environment prerequisite**, unchanged |

This is a reviewable local candidate with explicit simulation and strict-failure contracts. Overall direct-reader/cross-platform acceptance remains incomplete until the source conversation records the final commit's reader results.
