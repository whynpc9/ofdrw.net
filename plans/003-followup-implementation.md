# 003 follow-up: direct BuiltIn PDF italic

Base: `3247c76f95c9a3312cc26144e24c4984767af26f`. Worktree: `audit-followup-003`.
The later implementation authorization supersedes the original plan's read-only/dependency scope limitation. Original transaction production code and ten regressions remain untouched.

## Chosen implementation

Keep PdfSharpCore/MigraDoc 1.3.67 unchanged. Use the public `IXGraphicsRenderer` seam to intercept MigraDoc's precise DrawString operation before CID/GID encoding. This supplies the original text, actual XFont, brush and baseline; there is no text-string search, PDF-operation guessing, overlay, font-metadata style relabeling or private reflection.

Resolve italic faces once per conversion, snapshot the selected physical sfnt through the existing content-addressed registry, and inspect its actual OS/2/head style. Record the physical XFont family identity. The renderer shears only an italic request using a nonitalic physical face, around the exact baseline; Save/Restore contains the transform. A genuine italic face passes through once. All nontext drawing is forwarded. Unsupported italic identities/anchors fail the candidate before output commit. No-italic documents retain the original RenderPages path.

The adapter uses the pinned PDFsharp simulation angle (`sin(20 degrees)`). It does not implement bold simulation. Existing OFD rendering/font resolution stays unchanged.

## Small experiments and rejected options

- New public DOCX conversion regression failed on the frozen baseline: middle glyph 3 skew=0; adjacent controls and baseline checks ran. TRX: `artifacts/plan-003-followup/baseline-tests/baseline-italic.trx`.
- The same actual-PDF assertion passed through the graphics seam; repeated identical text is distinguished by the exact drawing call.
- A resolver-flags-only change is insufficient, confirmed by historical investigation and pinned upstream source `d6ac8b092129a4f797365bbbf3eea2723d6c3ebd`, XGlyphTypeface.GetOrCreateFrom.
- Upstream repair would be small but requires a maintained fork package/source distribution and changes every PDF consumer. The graphics seam experiment proved an in-repository route without that dependency/packaging cost. No dependency was patched or upgraded.
- Post-encoding style replacement was rejected: PDFsharp font-table coalescing means physical font names do not prove the individual styled occurrence.

## Acceptance

Initial full regression 2817/2817, trusted fonts 65/65, Python 20/20, API docs and 11-package consumption passed at 643e06a. Final helper extraction/breadth controls and fresh visual generation are being verified. Current focused tests: 17/17 (seven new style/drawing checks and ten original transaction checks). Mixed-page controls verify 300x420 portrait and 420x300 landscape PageInfo against original MigraDoc RenderPages, including exact regular and italic anchors; another control verifies headers, PAGE/NUMPAGES fields, table fill/borders, images and different page sizes. The true italic deterministic control uses the repository's authored synthetic italic face, whose outlines are actually slanted; additionally inspect the system's genuine Arial italic in the visual boundary fixture. Never infer arbitrary complex DOCX fidelity from these samples.

Full results will be recorded in `artifacts/plan-003-followup/acceptance.md` and task coordination `003.json`. This plan is not itself completion evidence.

## Pre-existing source page-setup finding

The original BuiltIn ApplyPageSetup assigns source landscape dimensions and also sets Landscape; MigraDoc treats its page width/height before orientation and swaps them. A source model 420x300 therefore produces 300x420 through both original RenderPages and this adapter. The failed exploratory model control and its original no-italic output are retained in `artifacts/plan-003-followup/legacy-orientation-evidence`. This source page-setup behavior predates and is unchanged by the italic repair. It is not counted as passed landscape DOCX fidelity. The new direct prepared-page control independently proves the adapter preserves actual 420x300 landscape pages. No page-setup fix was included in this scoped follow-up.
