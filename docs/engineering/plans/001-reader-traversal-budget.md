# Plan 001: Let OfdReader read normal-size documents (stop rejecting ~24+ page files)

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Core/Models/OfdModelLimits.cs src/Ofdrw.Net.Core/Models/OfdModelMaterialization.cs src/Ofdrw.Net.Reader/Readers/OfdReader.cs docs/content/reference/conversion-contracts.md docs/content/troubleshooting/known-issues.md`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: MED
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

`OfdReader.ReadAsync` throws `InvalidDataException: OFD traversal exceeds its reference budget.` for ordinary documents. The advisor reproduced it with the SDK's own output: a `FlowDocument` with 1,200 short paragraphs (24 pages, a 40 KB OFD) written by `OfdPackageWriter` cannot be read back. 1,000 paragraphs (20 pages) still worked. Every feature that reads an OFD goes through this path: OFD→PDF, SVG, HTML and image export, text extraction, and every CLI edit command. So a 25-page DOCX→OFD output cannot be exported. The cause is a fixed internal budget that callers cannot change and that is charged several times per XML part. After this plan, documents of several hundred pages read and convert. Resource protection stays in place through the byte and materialization budgets.

## Current state

- `src/Ofdrw.Net.Core/Models/OfdModelLimits.cs` defines the internal default limits (lines 7-14):
  ```csharp
  // Internal until a caller-facing read/write option needs this contract.
  internal sealed class OfdModelLimits
  {
      internal int MaxDepth { get; set; } = 64;
      internal int MaxNodeCount { get; set; } = 100_000;
      internal int MaxReferenceCount { get; set; } = 200_000;
      internal int MaxExpandedOccurrenceCount { get; set; } = 100_000;
      internal long MaxSourceXmlBytes { get; set; } = 128L * 1024 * 1024;
      internal long MaxMaterializedBytes { get; set; } = 512L * 1024 * 1024;
  ```
  The same file's `OfdTraversalBudget` (lines 33-107) counts `Nodes` and `References` across the **whole document read**. `Fail(kind)` throws `OFD traversal exceeds its {kind} budget.`
- `src/Ofdrw.Net.Core/Models/OfdModelMaterialization.cs`:
  - Lines 195-252: `GuardEntry` pre-scans each XML entry once. It charges one node per `XmlReader` node (whitespace and end tags included), one reference per attribute, and one reference per whitespace-separated token in `CTM`/`Boundary`/`Value` attributes and in all text nodes.
  - Lines 272-280: `ReadXml` charges the same node and reference totals **again** on every DOM materialization:
    ```csharp
    // Every DOM materialization is work, including aliases of the same entry.
    ChargeCopy(2L * bytes.LongLength, token);
    Structure.NodeWork(work.Nodes, work.Depth);
    for (long index = 0; index < work.References; index++) { token.ThrowIfCancellationRequested(); Structure.Reference(); }
    ```
  - Memory is bounded separately by `MaxMaterializedBytes` (line 385, `ChargeCopy`/`ReserveReadModel`) and `MaxSourceXmlBytes`. Nothing allocates in proportion to `MaxNodeCount` or `MaxReferenceCount`. They are counters only (`grep -rn "\.MaxNodeCount\|\.MaxReferenceCount" src` matches only `OfdModelLimits.cs`).
- `src/Ofdrw.Net.Reader/Readers/OfdReader.cs:75` creates the budget per read: `return ReadArchive(archive, snapshot, new OfdModelMaterialization(new OfdModelLimits()), cancellationToken);`. `ListDocumentsAsync` (line 35) does the same.
- About 25 other call sites also do `new OfdModelLimits()` (writer, pruner, text extractor, editor, OFD→PDF). They all pick up the defaults, which is why this plan changes the defaults rather than one call site.
- The keyword locator passes its own explicit limits (`src/Ofdrw.Net.Reader/Keywords/Internal/OfdKeywordAnalysis.cs:133`), so it is **not** affected by the defaults. Do not change it.
- Measured cost: about 180 references per simple text object across all charges. So 200,000 references ≈ 1,100 text objects.
- No existing test depends on the default node/reference values. `grep -rn "reference budget\|node budget\|traversal exceeds" tests` returns nothing. Tests that exercise these budgets set explicit limits (`tests/Ofdrw.Net.Core.Tests/OfdModelTraversalTests.cs:163-192, 235, 318-320`).

Conventions: xUnit 2 tests, `public sealed class XxxTests`, async `Task` facts, file-scoped namespaces matching the project (`namespace Ofdrw.Net.Packaging.Tests;`). Model new tests on `tests/Ofdrw.Net.Packaging.Tests/PackageRoundTripTests.cs`. Code style: terse, `var`, invariant culture for every number written to output.

The generation API, from the public quickstart:
```csharp
var document = new FlowDocument();                       // Ofdrw.Net.Layout
document.Blocks.Add(new Paragraph("text"));
await new OfdPackageWriter().WriteAsync(document.Render(), output);   // Ofdrw.Net.Packaging
var read = await new OfdReader().ReadAsync(stream);       // Ofdrw.Net.Reader.Readers
```

## Commands you will need

Run from the repo root. Once per shell:
```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected on success |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 (existing warnings are OK) |
| New tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~LargeDocumentReadTests"` | all pass |
| PDF test | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F --filter "FullyQualifiedName~LargeDocumentPdfTests"` | all pass |
| Regression | `dotnet test tests/Ofdrw.Net.Core.Tests -c Release $F` and the same for `Ofdrw.Net.Packaging.Tests`, `Ofdrw.Net.Layout.Tests`, `Ofdrw.Net.Converter.Pdf.Tests`, `Ofdrw.Net.Cli.Tests` | all pass (tests that need LibreOffice/Poppler may fail if tools are missing; record which) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |
| Docs site | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | exit 0 (needs Node ≥ 22.12; if Node is missing, record it as not run) |

## Scope

**In scope**:
- `src/Ofdrw.Net.Core/Models/OfdModelLimits.cs`: default values and a comment only
- `tests/Ofdrw.Net.Packaging.Tests/LargeDocumentReadTests.cs` (create)
- `tests/Ofdrw.Net.Converter.Pdf.Tests/LargeDocumentPdfTests.cs` (create)
- `docs/content/reference/conversion-contracts.md`: one bullet in "## 资源预算与失败行为"
- `docs/content/troubleshooting/known-issues.md`: one table row

**Out of scope**:
- `OfdModelMaterialization.cs` accounting logic. The "charge every materialization" rule is a deliberate design comment, and changing it touches many audited allocation tests (`AllocationBudgetDecisionTests`, `Audit025RegressionTests`).
- Making limits public (`OfdPackageLoadOptions`). That is a public API decision; see Maintenance notes.
- `OfdKeywordOptions` / keyword-locator limits.
- `MaxDepth`, `MaxExpandedOccurrenceCount`, `MaxSourceXmlBytes`, `MaxMaterializedBytes`. Do not change them unless a STOP condition tells you to report.

## Git workflow

- Branch: `advisor/001-reader-traversal-budget`
- Commit style matches `git log` (conventional prefix), e.g. `fix(reader): raise default traversal budgets so large documents read`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Add failing regression tests

Create `tests/Ofdrw.Net.Packaging.Tests/LargeDocumentReadTests.cs`:
- Helper `static async Task<MemoryStream> WriteFlow(int paragraphs)`: build a `FlowDocument`, add `paragraphs` paragraphs of the form `$"Paragraph {i}: The quick brown fox jumps over the lazy dog, 中文段落测试内容第{i}段。"`, render, write with `OfdPackageWriter`, rewind, return. Also return or record `document.Render().Pages.Count` for comparison (render once and reuse the package).
- `[Theory] [InlineData(1200)] [InlineData(4000)] [InlineData(20000)] Reader_ShouldReadLargeFlowDocument(int paragraphs)`: read with `new OfdReader().ReadAsync(stream)`. Assert the page count equals the rendered package's page count, and that the total `OfdTextElement` count across pages equals the rendered total.
- `[Fact] TextExtractor_ShouldExtractLargeDocument`: 4,000 paragraphs. `new OfdTextExtractor().ExtractPages(read, includeTemplates: true)` (namespace `Ofdrw.Net.Reader.Extraction`) returns one string per page, and the last page contains `"Paragraph 3999:"`.
- `[Fact] ListDocuments_ShouldListLargeDocument`: 4,000 paragraphs. `ListDocumentsAsync` returns 1 descriptor. This is a **guard regression, expected to pass both before and after the fix**. `ListDocumentsAsync` (`OfdReader.cs:31-40`) materializes only `OFD.xml` and checks that the document entries exist; it never reads the page XML that exhausts the counters. It protects against a future change that makes listing materialize pages.

Create `tests/Ofdrw.Net.Converter.Pdf.Tests/LargeDocumentPdfTests.cs`:
- `[Fact] OfdToPdf_ShouldConvertLargeFlowDocument`: 2,000 paragraphs → OFD → `new OfdToPdfConverter().ConvertAsync(ofd, pdf)` (namespace `Ofdrw.Net.Converter.Pdf.Converters`). Open the PDF with PdfPig (`UglyToad.PdfPig.PdfDocument.Open(bytes)`, already referenced by this test project) and assert `NumberOfPages` equals the OFD page count. If the test project's fonts make this slow (> 60 s), reduce to 1,500 paragraphs, which is still above the old limit.

**Verify**: `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~LargeDocumentReadTests"` → the 1200/4000/20000 cases and the extractor fact **fail** with `OFD traversal exceeds its reference budget.` (or `node budget`). These are the tests that read page content. `ListDocuments_ShouldListLargeDocument` **passes**, as expected (see above); do not treat that as a STOP signal. If the page-reading cases pass already, STOP (the bug was fixed independently). If the listing fact fails, STOP and report: listing would then be materializing page content, contrary to the excerpt.

### Step 2: Raise the two counter defaults

In `OfdModelLimits.cs` change only:
```csharp
internal int MaxNodeCount { get; set; } = 10_000_000;
internal int MaxReferenceCount { get; set; } = 20_000_000;
```
Add a comment above them, in the same terse style, saying:
- These are whole-read work counters, and every DOM materialization re-charges them (`ReadXml`), so they must scale with real documents.
- Memory remains bounded by `MaxSourceXmlBytes` and `MaxMaterializedBytes`.
- About 180 references per simple text object were measured, so 20M ≈ 100k text objects.

**Verify**: build → exit 0. Re-run the Step 1 filters → all pass.

### Step 3: Full regression run

Run the "Regression" commands from the table.

**Verify**: every suite passes, except tests that fail only because an external tool (LibreOffice/Poppler/fonts) is missing. Re-run those on the base commit (`git stash` is shared across worktrees; instead use `git worktree add ../base 09a2e7e` or compare with CI) to prove they fail there too.

### Step 4: Visual sanity of the large PDF (AGENTS.md conversion rule)

Write the 2,000-paragraph OFD and its PDF to `artifacts/visual/001/large.ofd` and `artifacts/visual/001/large.pdf`. A temporary console app outside the repo, or an env-gated line in the test, is fine; do not commit generated files. Render **every** page:
```sh
pdftoppm -r 144 -png artifacts/visual/001/large.pdf artifacts/visual/001/page
ls artifacts/visual/001/page-*.png | wc -l      # must equal the OFD page count
sha256sum artifacts/visual/001/large.ofd artifacts/visual/001/large.pdf artifacts/visual/001/page-*.png > artifacts/visual/001/SHA256SUMS
```
Open **each** page PNG with an image viewer or your image-reading tool. For every page, check that text is present, nothing overlaps or is clipped, there are no blank or ghosted pages, and paragraph numbering continues from the previous page (no lost or duplicated paragraphs at page breaks). Chinese characters must not render as boxes **if** a CJK font is installed. Missing CJK glyphs with no CJK font installed is an environment limitation; record it, it is not a failure of this plan.

Record the review coverage as an explicit list: pages reviewed (e.g. `1-40 of 40`) and observations per page or page range. Any page not opened must be listed as **unreviewed**; it keeps the PNG review item incomplete. Do not describe a partial review as complete.

**Verify**: PNG count equals the page count; every PNG is non-empty (`for f in artifacts/visual/001/page-*.png; do test -s "$f" || echo EMPTY $f; done` prints nothing); the coverage list is in the report.

### Step 5: Document the behavior change

- In `docs/content/reference/conversion-contracts.md`, section `## 资源预算与失败行为`, add one bullet in the surrounding style: OFD reading uses internal whole-document XML node/reference work counters (defaults 10,000,000 / 20,000,000), with memory bounded by the source-XML and materialization byte budgets. preview.10 used 100,000 / 200,000, which rejected documents with roughly more than 1,000 text objects. Callers cannot configure these yet.
- In `docs/content/troubleshooting/known-issues.md`, add a row to the version-difference table: `| 较大 OFD 读取（约 1,000+ 文本对象） | 读取/转换报 "OFD traversal exceeds its reference budget" | 默认工作计数上限已提高，尚未发布 |`.
- Run the docs-site commands if Node is available.

**Verify**: `git diff --stat` shows only in-scope files. Docs commands exit 0, or are recorded as not run with the reason.

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/001/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/001/trusted-font-tests
rm -rf artifacts/001/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/001/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/001/packages --output-dir artifacts/001/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/001/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Test plan

- New: `LargeDocumentReadTests` (3 theory cases + 2 facts) and `LargeDocumentPdfTests` (1 fact), all passing after Step 2. All except `ListDocuments_ShouldListLargeDocument` fail before Step 2; the listing fact is a guard regression that passes both before and after.
- Existing explicit-limit tests in `OfdModelTraversalTests` must still pass unchanged, because they set their own limits.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] `dotnet build Ofdrw.Net.sln -c Release $F` exits 0
- [ ] New tests exist and pass; the page-reading tests failed before Step 2 (record the failure message) and the listing guard passed before and after
- [ ] Core, Packaging, Layout, Converter.Pdf, Cli test suites pass (environment-only failures listed and shown to also fail on base)
- [ ] `grep -n "MaxNodeCount\|MaxReferenceCount" src/Ofdrw.Net.Core/Models/OfdModelLimits.cs` shows the new defaults
- [ ] Docs bullet and known-issues row added
- [ ] Every page rendered to PNG; per-page review coverage recorded (unreviewed pages listed as incomplete); macOS Preview acceptance marked **not done (Linux/CLI environment)** in the report
- [ ] `python3 scripts/check-repository-storage.py` exits 0 after staging
- [ ] `docs/engineering/plans/README.md` status row updated

## STOP conditions

- After Step 2, any Step 1 test fails with a *different* budget message (e.g. `materialization budget`, `occurrence/diagnostic budget`, `source XML byte budget`). Report the exact message and paragraph count. Do not raise other limits on your own.
- Any existing test fails that asserts a rejection based on the old defaults.
- The 20,000-paragraph case takes more than 120 s. Report the timing; do not delete the case without approval.
- The excerpts above don't match the live code.

## Maintenance notes

- This is the minimal fix. The better long-term design is either to charge re-materializations by bytes rather than by node count again, or to expose these limits through `OfdPackageLoadOptions`, alongside `MaxEntryCount`/`MaxPageCount`. Either needs a public-API decision plus `scripts/check-public-api-docs.py`.
- Reviewers should confirm that no code allocates arrays sized by `MaxNodeCount`/`MaxReferenceCount`. Re-run `grep -rn "\.MaxNodeCount\|\.MaxReferenceCount" src` in review.
- Plan 009 edits `OfdReader.cs` too; land this first.
