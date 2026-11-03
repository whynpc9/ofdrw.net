# Plan 005: Add an OFD→PDF result/diagnostics channel and stop silently dropping seals

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs src/Ofdrw.Net.Converter.Pdf/OfdToPdfOptions.cs src/Ofdrw.Net.Cli/Program.cs`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW (additive public API; rendering of successful seals unchanged)
- **Depends on**: docs/engineering/plans/001 (large-document reads) is recommended first but not required
- **Category**: bug + direction (the maintainer explicitly requested the result channel)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

In the default OFD→PDF path, an official seal (签章) that fails to parse or draw is **silently dropped**. The PDF looks complete, and `ConvertAsync` returns `Task` with no way to report it. For e-invoices and official documents, a missing seal is the worst kind of silent loss. Two more problems:
- In the default path, nested seal OFDs are read with library-default budgets (`new OfdPackageLoadOptions()`) instead of the caller's `OfdToPdfOptions.PackageLoadOptions`. That contradicts the documented contract "`OfdToPdfOptions` 控制 OFD 加载和图片解码预算".
- A seal stamped on N pages is unzipped and parsed N times.

The maintainer asked for an OFD→PDF result/diagnostics channel, mirroring `DocxToOfdConverter.ConvertWithResultAsync`. This plan adds it, reports every skipped appearance through it, applies caller budgets, and caches nested seal packages.

## Current state

- `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
  - Lines 54-69, the public entry point:
    ```csharp
    public async Task ConvertAsync(Stream ofdInput, Stream pdfOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
    {
        ...null checks...
        var reader = new OfdReader();
        var package = await reader.ReadAsync(ofdInput, _options.PackageLoadOptions, cancellationToken).ConfigureAwait(false);
        await ConvertPackageAsync(package, pdfOutput, pages, cancellationToken).ConfigureAwait(false);
    }
    ```
  - Lines 72-73: `internal async Task ConvertPackageAsync(OfdDocumentPackage package, Stream pdfOutput, IReadOnlyList<int>? pages, CancellationToken cancellationToken, bool strictAppearanceBudgets = false, int? maximumSignatureAppearances = null)`. It is also used by image export with `strictAppearanceBudgets: true`; strict mode throws instead of skipping. **Do not change strict-mode behavior.**
  - Lines 180-226, `PrepareSignatureAppearancesAsync`: the nested-package cache and cumulative budget only apply `if (strictAppearanceBudgets)`:
    ```csharp
    if (!strictAppearanceBudgets || !nestedCache.TryGetValue(key, out var cached))
    {
        var remaining = strictAppearanceBudgets ? RemainingAppearanceBudget(usedBytes, usedEntries, usedPages) : null;
        var nested = await ReadAppearanceOfdPackageAsync(appearance.Data, remaining, annotationMaterial, cancellationToken)...
    ```
  - Lines 228-242: `RemainingAppearanceBudget` copies limit properties from `_options.PackageLoadOptions` (`MaxInputBytes`, `MaxEntryUncompressedBytes`, `MaxAnnotationObjectCount`, `MaxAnnotationXmlBytes`, `MaxCompressionRatio`, `MaxTotalUncompressedBytes`, `MaxEntryCount`, `MaxPageCount`). It does **not** copy `DocumentIndex`, which is correct because a nested seal is always read at index 0.
  - Line 254: `.ReadAsync(input, appearanceBudget ?? new OfdPackageLoadOptions(), cancellationToken)`. The default path ignores caller limits.
  - Lines 280-283: `catch { return null; }`. A nested seal that fails to parse is dropped silently in the default path.
  - Lines 361-368, `DrawSignatureAppearances`: `catch { /* Unsupported vendor payloads retain the existing tolerant PDF fallback. */ }`. Draw failures are dropped silently in the default path.
  - The same method also skips appearances **without** an exception, through three guard exits. Each throws only when `reusePayloads` (strict/image mode) is true, and otherwise `continue`s silently:
    ```csharp
    // lines 301-305: placement
    if (!PdfOperandGeometry.Box(appearance.XMillimeters - page.XMillimeters, ...))
    {   if (reusePayloads) throw new InvalidDataException("Selected signature stamp placement is not representable in PDF coordinates.");
        continue; }
    // lines 313-320: clip
    if (!PdfOperandGeometry.Box(clipX, clipY, clip.Width, clip.Height))
    {   if (reusePayloads) throw new InvalidDataException("Selected signature stamp Clip is not representable in PDF coordinates.");
        continue; }
    // lines 330-334: nested-form scale
    if (!PdfOperandGeometry.PositiveExtent(target.Width / formWidth) || !PdfOperandGeometry.PositiveExtent(target.Height / formHeight))
    {   if (reusePayloads) throw new InvalidDataException("Selected signature form scale is not representable by the PDF writer.");
        continue; }
    ```
    (The braces are condensed here.) The clip and scale exits sit inside the `try` whose `finally` calls `graphics.Restore(state)`, so `continue` still restores graphics state.
- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs:15-17`: `public static IReadOnlyList<OfdSignatureAppearance> Read(OfdDocumentPackage package, HashSet<string>? selectedPageIds = null, int? maximumAppearances = null, CancellationToken cancellationToken = default)`. Lines 65-70 return an empty list on any exception when `maximumAppearances` is null (the default path):
  ```csharp
  catch
  {
      // A malformed or unsupported signature must not prevent the document
      // body from being converted.
      return Array.Empty<OfdSignatureAppearance>();
  }
  ```
  The same file also drops stamps **without** an exception. Every strict check has the shape `if (maximumAppearances.HasValue) throw ...;` followed by a silent tolerant `return`/`continue`. The tolerant exits at `09a2e7e` are:

  | Line | Exit | Meaning |
  |---|---|---|
  | 22 | `return Array.Empty<...>()` | no `OFD.xml` metadata root |
  | 37 | `return Array.Empty<...>()` | selected DocBody missing from the metadata root |
  | 82 | `return;` | referenced signature list entry missing |
  | 96 | `continue;` | signature entry reference missing or empty |
  | 114 | `return;` | referenced signature entry missing |
  | 142 | `return;` | stamp has no supported appearance payload (`appearanceData.Length == 0`) |
  | 154 | `continue;` | stamp `Boundary` (or `PageRef`) not representable |
  | 169 | `continue;` | stamp `Clip` invalid or outside `Boundary` |

  Line 127 (`if (selectedPageIds is not null && !selectedPageIds.Contains(...)) continue;`) is **page selection**, not a skip. It must not produce a diagnostic. Re-list the exits with `grep -n "if (maximumAppearances.HasValue" -A2 src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs` and treat any difference from this table as drift.
- `IOfdToPdfConverter` (`src/Ofdrw.Net.Converter.Abstractions/Interfaces/IOfdToPdfConverter.cs`) only has `ConvertAsync`. **Do not change the interface** (breaking for implementers).
- Existing public result to mirror, for naming and doc style: `src/Ofdrw.Net.Converter.Docx/DocxConversionDiagnostic.cs`. It is a `DocxConversionDiagnosticSeverity` enum (`Information`, `Warning`, `Error`, each with `/// <summary>`) plus a sealed class with `Code`, `Message`, `Severity` and an XML-documented constructor. Note that `src/Ofdrw.Net.Converter.Pdf/PdfConversionResult.cs` already exists for **PDF→OFD**; do not reuse or rename it.
- CLI, `src/Ofdrw.Net.Cli/Program.cs:208-213`:
  ```csharp
  else
  {
      var ofdToPdf = new OfdToPdfConverter(new Ofdrw.Net.Converter.Pdf.OfdToPdfOptions
      { PackageLoadOptions = new OfdPackageLoadOptions { DocumentIndex = documentIndex } });
      await ofdToPdf.ConvertAsync(input, output.Stream, pages, cancellationToken).ConfigureAwait(false);
  }
  ```
  The DOCX diagnostics printer at lines 240-244 prints `$"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}"` to stderr. Mirror it.
- Public API documentation gate: `scripts/check-public-api-docs.py` (CI) fails on new public members without XML docs. Every new public type and member needs `/// <summary>`.
- Test exemplar: `tests/Ofdrw.Net.Converter.Pdf.Tests/StampClipRenderingTests.cs`. `SignedTwoPageSeal()` (lines 76-95) signs a 2-page document with a PNG seal via `OfdSignatureService` and a `PictureProvider` whose `SealData` is the picture bytes.

## Public API to add (exact shape)

In `src/Ofdrw.Net.Converter.Pdf/` (namespace `Ofdrw.Net.Converter.Pdf`), one new file `OfdToPdfConversionResult.cs`:
```csharp
/// <summary>Severity of an OFD-to-PDF conversion diagnostic.</summary>
public enum OfdToPdfDiagnosticSeverity { /// <summary>…</summary> Information, /// <summary>…</summary> Warning }

/// <summary>Describes content that OFD-to-PDF conversion skipped or degraded.</summary>
public sealed class OfdToPdfDiagnostic
{
    internal OfdToPdfDiagnostic(string code, string message, OfdToPdfDiagnosticSeverity severity, int? pageIndex = null, string? signatureId = null) { … }
    /// <summary>Stable machine-readable code, e.g. PDF_SIGNATURE_APPEARANCE_SKIPPED.</summary>
    public string Code { get; }
    public string Message { get; }               // documented
    public OfdToPdfDiagnosticSeverity Severity { get; }
    /// <summary>Zero-based source page index when the diagnostic concerns one page.</summary>
    public int? PageIndex { get; }
    /// <summary>OFD signature ID when known.</summary>
    public string? SignatureId { get; }
    /// <inheritdoc/> public override string ToString() => $"{Severity}: {Code}: {Message}";
}

/// <summary>Result of an OFD-to-PDF conversion.</summary>
public sealed class OfdToPdfConversionResult
{
    internal OfdToPdfConversionResult(IEnumerable<int> sourcePages, int renderedSignatureAppearances, IEnumerable<OfdToPdfDiagnostic> diagnostics) { … }
    /// <summary>Zero-based source page index for each output page, in output order.</summary>
    public IReadOnlyList<int> SourcePages { get; }
    /// <summary>Number of signature appearances drawn into the PDF.</summary>
    public int RenderedSignatureAppearances { get; }
    public IReadOnlyList<OfdToPdfDiagnostic> Diagnostics { get; }   // documented
}
```
On `OfdToPdfConverter`:
```csharp
/// <summary>Converts OFD to PDF and reports page mapping and skipped or degraded content.</summary>
public async Task<OfdToPdfConversionResult> ConvertWithResultAsync(Stream ofdInput, Stream pdfOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
```
`ConvertAsync` becomes `_ = await ConvertWithResultAsync(...)`, with the same exceptions as before.

Diagnostic codes, all `Warning`:
- `PDF_SIGNATURE_METADATA_UNREADABLE`: `OfdSignatureAppearanceReader.Read` fell into its tolerant catch, or took a metadata-level tolerant exit (table rows 22, 37, 82, 96, 114). The message includes the exception type name or a fixed reason (`metadata root`, `document body`, `signature list missing`, `signature reference`, `signature entry missing`). It never includes the exception message or raw XML (avoid echoing untrusted content verbatim).
- `PDF_SIGNATURE_APPEARANCE_UNREADABLE`: a nested OFD seal package could not be read (the tolerant `return null` path), **or** a stamp had no supported appearance payload (table row 142, reason `no supported payload`). Include `SignatureId` if available.
- `PDF_SIGNATURE_APPEARANCE_SKIPPED`: a stamp or appearance was not drawn in the default path. This covers:
  - the metadata reader's stamp exits (table rows 154 and 169, reasons `boundary` and `clip`), with `PageIndex` resolved from `PageRef` where possible, else null
  - the draw-time tolerant catch
  - the three draw-time guard exits (`placement`, `clip`, `form scale`)

  Include `PageIndex` and `SignatureId`, and state the reason in `Message`.

## Commands you will need

Once per shell:
```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F --filter "FullyQualifiedName~OfdToPdfResultTests"` | pass |
| PDF suite | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F` | pass |
| CLI suite | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Public API docs | `python3 scripts/check-public-api-docs.py` | exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**:
- `src/Ofdrw.Net.Converter.Pdf/OfdToPdfConversionResult.cs` (create)
- `src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs`: add the optional `onTolerantSkip` sink, thread it to the tolerant exits, and invoke it there only; keep parsing, selection logic and `ParseXml` unchanged (plan 007 changes `ParseXml`)
- `src/Ofdrw.Net.Cli/Program.cs`: the OFD→PDF branch only
- `tests/Ofdrw.Net.Converter.Pdf.Tests/OfdToPdfResultTests.cs` (create)
- `tests/Ofdrw.Net.Cli.Tests/`: one new test in an existing CLI test file or a new file
- Docs: `docs/content/reference/conversion-contracts.md`, `docs/content/guides/conversion.md` (one short example), and the `docs/content/troubleshooting/known-issues.md` table

**Out of scope**:
- `IOfdToPdfConverter` (do not change it)
- Strict-mode (`strictAppearanceBudgets == true`) behavior and messages
- An option to fail instead of skipping (deferred; see Maintenance notes)
- SVG/HTML/image converters' diagnostics
- Body-content error handling; it already throws (`OfdToPdfConverter.cs:453,521`)

## Git workflow

Branch `advisor/005-ofd-to-pdf-result`; commit e.g. `feat(pdf): report skipped signature appearances via ConvertWithResultAsync`. No push or PR unless instructed.

## Steps

### Step 1: Failing tests (`OfdToPdfResultTests.cs`)

Copy the `PictureProvider` and `SignedTwoPageSeal` pattern from `StampClipRenderingTests.cs` (lines 76-109) into private helpers. Parameterize the seal bytes.

Cases:
1. `ValidSeal_ReportsRenderedAppearancesAndNoDiagnostics`: valid PNG seal → `RenderedSignatureAppearances >= 2` (riding seal over 2 pages), `Diagnostics` empty, `SourcePages == [0, 1]`.
2. `UndrawableSeal_ReportsSkippedAppearance`: `SealData` = 64 bytes that are neither an image nor a ZIP (`Enumerable.Repeat((byte)0x42, 64)`). Conversion succeeds and `Diagnostics` contains `PDF_SIGNATURE_APPEARANCE_SKIPPED` with non-null `PageIndex`.
   - If signing rejects such bytes, use bytes that start with a valid PNG signature followed by garbage, so validation passes but decoding fails.
   - If signing still rejects them, rewrite the seal entry inside the signed ZIP afterwards. Find the `.png` entry under `Doc_0/Signs/` and replace its bytes. Signature digests becoming invalid is fine for this test.
3. `MalformedSignatureMetadata_ReportsMetadataUnreadable`: take a validly signed package and replace the bytes of the entry ending with `Signatures.xml` with `<Broken`. Expect `PDF_SIGNATURE_METADATA_UNREADABLE`.
   - If `OfdReader` itself now rejects that input with `InvalidDataException`, that is acceptable. Assert that instead, and pick a mutation the reader tolerates but the appearance reader does not (e.g. a `Signature` element whose `BaseLoc` points to a missing entry).
   - If no such mutation exists, delete this case and say so in the report.
4. `NestedSealBudget_UsesCallerOptions`: a seal whose `SealData` is a ZIP (an OFD written by `OfdPackageWriter`, 1 page). Convert with `new OfdToPdfOptions { PackageLoadOptions = new OfdPackageLoadOptions { MaxEntryCount = <main package entry count + 1> } }`.
   - Expected after the fix: the main package reads, but the nested seal OFD is rejected, because its entry count must fit within the caller's limit. Pick a nested package with more entries than that value, e.g. add 1 attachment and 2 pages.
   - The result then contains `PDF_SIGNATURE_APPEARANCE_UNREADABLE`. Before the fix it rendered silently with default budgets.
   - If building such a seal through `OfdSignatureService` is impossible, mark this case `Skip` with the reason and cover budget propagation with an internal-visible unit test instead (`InternalsVisibleTo` exists for this test project if `OfdToPdfConverter` internals are reachable; check `src/Ofdrw.Net.Converter.Pdf/*.csproj`).
2b. `UnrepresentableBoundary_ReportsSkippedAppearance`: take the valid signed package from case 1 and rewrite the stamp annotation's `Boundary` in the entry ending with `/Signature.xml` (find the `StampAnnot` element) to `1e300 1e300 40 20`. The metadata reader's boundary check (row 154) rejects it before drawing. The default conversion succeeds; `Diagnostics` contains `PDF_SIGNATURE_APPEARANCE_SKIPPED` whose message mentions `boundary`; `RenderedSignatureAppearances` is 0 for that stamp. Also assert that the strict image-export path (`OfdToImageConverter`, used in `StampClipRenderingTests`) still **throws** `InvalidDataException` for the same input.
2c. `InvalidClip_ReportsSkippedAppearance`: set `StampAnnot/@Clip` to a rectangle larger than its `Boundary` (e.g. `0 0 999 999`). Expect `PDF_SIGNATURE_APPEARANCE_SKIPPED` mentioning `clip` (row 169). Strict image export still throws.
2d. `MissingPayload_ReportsUnreadableAppearance`: delete the seal picture entry that `Signature.xml` references. Find it via the signed package's `Doc_0/Signs/` entries; it is the non-XML payload. Expect `PDF_SIGNATURE_APPEARANCE_UNREADABLE` mentioning `no supported payload` (row 142). If `OfdReader` itself rejects the package instead, assert that and note it.
2e. `DrawGuard_ReportsSkippedAppearance`: cover the draw-time guards (placement/clip/form scale, `OfdToPdfConverter.cs:301-334`). The metadata reader already validates geometry, so these guards are normally unreachable from public input. Use an internal-visible unit test that calls `DrawSignatureAppearances` directly with a crafted `PreparedSignatureAppearance` whose placement is unrepresentable relative to the page, e.g. page origin offset. Check `src/Ofdrw.Net.Converter.Pdf/*.csproj` for `InternalsVisibleTo`. If no such test is feasible, say so in the report.
5. `ConvertAsync_StillReturnsTask`: compile-level: `await new OfdToPdfConverter().ConvertAsync(...)` still works and produces bytes.

**Verify**: build fails (the API does not exist yet). After adding empty stubs, cases 2, 2b, 2c, 2d, 3 and 4 fail (2e may only compile once Step 3 adds the sink).

### Step 2: Add the public types (exact shape above) and `ConvertWithResultAsync`

Thread a `List<OfdToPdfDiagnostic>` and a rendered counter through `ConvertPackageAsync`. Add an optional parameter `ICollection<OfdToPdfDiagnostic>? diagnostics = null` and return the counter, or use a small internal `OfdToPdfRun` holder class. Image export passes nothing, so its behavior is unchanged. `SourcePages` = the normalized `selected` list already computed in `ConvertPackageAsync`.

**Verify**: build exits 0; `python3 scripts/check-public-api-docs.py` exits 0.

### Step 3: Report every tolerant skip

- `OfdSignatureAppearanceReader.Read`: add an optional `Action<string, string, string?, string?>? onTolerantSkip = null` parameter, with arguments `(code, reason, pageRef, signatureId)`. Thread it through the private helpers that contain the tolerant exits (`ReadSignatureList`, `ReadSignature`, …).
  - Invoke it from the bare `catch` (change it to `catch (Exception exception)` and pass `exception.GetType().Name` as the reason, code `PDF_SIGNATURE_METADATA_UNREADABLE`).
  - Invoke it immediately before **every** tolerant `return`/`continue` listed in the Current-state table, with the code and reason given under "Diagnostic codes". Include `PageRef` and the signature `ID` attribute when they are in scope.
  - Do **not** invoke it for the page-selection `continue` (line 127).
  - Leave every `if (maximumAppearances.HasValue) throw ...` line unchanged, so strict mode still throws and reports nothing.
  - The converter maps `pageRef` to a zero-based page index via the page IDs it already has (`orderedPages`), or null if unknown, and appends an `OfdToPdfDiagnostic`.
- `ReadAppearanceOfdPackageAsync`'s final `catch { return null; }` → `PDF_SIGNATURE_APPEARANCE_UNREADABLE`, recorded at the call site when the result is null. Also do this when the page-geometry check returns null (line 275).
- `DrawSignatureAppearances`' final `catch { }` → change to `catch (Exception)` and add `PDF_SIGNATURE_APPEARANCE_SKIPPED` with the page index and signature ID. Count successful draws for `RenderedSignatureAppearances`; increment after `DrawImage` returns.
- `DrawSignatureAppearances`' three guard exits (placement, lines 301-305; clip, 313-320; form scale, 330-334): immediately before each `continue`, record `PDF_SIGNATURE_APPEARANCE_SKIPPED` with the page index, signature ID, and reason `placement` / `clip` / `form scale`. Leave the `if (reusePayloads) throw ...` line before it unchanged, so strict mode still throws and records nothing.

Keep every existing `catch (X) { throw; }` clause and its order exactly as it is.

**Verify**: cases 1-3 and 2b-2e pass. `grep -n "if (maximumAppearances.HasValue" -A3 src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs` shows an `onTolerantSkip` call before every tolerant `return`/`continue` in the Current-state table. Every `continue` inside `DrawSignatureAppearances` is now preceded by a diagnostic: `grep -n -B1 "continue;" src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs` shows the diagnostic call on the line before each `continue` in that method.

### Step 4: Caller budgets and caching in the default path

- Add `private OfdPackageLoadOptions NestedAppearanceOptions()`. It returns a new `OfdPackageLoadOptions` copying the same eight limit properties `RemainingAppearanceBudget` copies, **without** subtracting and **without** `DocumentIndex`.
- At line 254, replace `appearanceBudget ?? new OfdPackageLoadOptions()` with `appearanceBudget ?? NestedAppearanceOptions()`. Leave the `when (appearanceBudget is not null)` filters as they are: strict mode is identified by a non-null budget, so keep passing `null` in the default path and resolve the options inside.
- Cache nested results in the default path too. Compute `key = BinaryIdentity.Hash(appearance.Data)` (already used at line 198) for both modes. In default mode store `(nested, fonts)` in `nestedCache`, including a null result, so a broken seal is parsed once and reported once per occurrence (each occurrence still produces its own `PDF_SIGNATURE_APPEARANCE_UNREADABLE` with its page). Strict-mode cumulative accounting is unchanged.

**Verify**: case 4 passes; the whole PDF suite passes (including `StampClipRenderingTests`, `MultiDocumentExportTests`, `Audit011RendererRegressionTests`).

### Step 5: CLI prints diagnostics

In `Program.cs` lines 208-213, call `ConvertWithResultAsync` and print each diagnostic to stderr as `$"{d.Severity}: {d.Code}: {d.Message}"` (same format as `PrintDocxDiagnostics`). The exit code stays 0. Add a CLI test: convert an OFD with an undrawable seal; assert exit 0 and that stderr contains `PDF_SIGNATURE_APPEARANCE_SKIPPED`. Capture stderr with `Console.SetError(new StringWriter())` and restore it in `finally`; check `tests/Ofdrw.Net.Cli.Tests` for an existing capture helper first and reuse it if present.

**Verify**: CLI suite passes.

### Step 6: Visual non-regression (AGENTS.md)

The output for valid seals must be unchanged.
- Generate the valid two-page seal OFD from case 1 into `artifacts/visual/005/seal.ofd`, using a throwaway console app or an env-gated test write.
- Convert it with the base commit's CLI (`git worktree add ../ofdrw-base 09a2e7e`, then build there) and with the new CLI.
- `pdftoppm -r 144 -png` both PDFs and compare page PNGs (`cmp` or `sha256sum`).
- Open both page PNGs and confirm the seal slices are visible.

**Verify**: PNG hashes are identical, or visually identical with any differences explained. The record is in the report. macOS Preview is reported as not done unless performed.

### Step 7: Docs

- `conversion-contracts.md`, section `## 资源预算与失败行为`, two bullets:
  - The default OFD→PDF path still tolerates unreadable or undrawable seals, but now reports them via `ConvertWithResultAsync().Diagnostics` (list the three codes) and the CLI's stderr.
  - Nested seal OFDs are read under the caller's `PackageLoadOptions` limits.
- `guides/conversion.md`: a 6-line C# example of `ConvertWithResultAsync` that prints diagnostics.
- `known-issues.md`: a row "OFD→PDF 签章外观失败 | 静默跳过 | 诊断报告，尚未发布".

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/005/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/005/trusted-font-tests
rm -rf artifacts/005/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/005/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/005/packages --output-dir artifacts/005/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/005/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] Build exits 0; `check-public-api-docs.py` exits 0
- [ ] `OfdToPdfResultTests` pass, with any skipped case explicitly justified; PDF and CLI suites pass
- [ ] `grep -n "new OfdPackageLoadOptions()" src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs` no longer matches the nested-read line
- [ ] `IOfdToPdfConverter.cs` unchanged (`git diff --quiet -- src/Ofdrw.Net.Converter.Abstractions`)
- [ ] Visual non-regression recorded
- [ ] Docs updated; storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- Strict-mode (image export) tests change behavior.
- Adding the diagnostics sink requires changing `OfdSignatureAppearanceReader`'s parsing or selection logic beyond invoking the callback at the listed exits, or the grep in Step 1 shows tolerant exits not in the Current-state table.
- Valid-seal PNGs differ between the base and new builds.
- The public API check reports undocumented symbols you cannot document (e.g. generated members).

## Maintenance notes

- A follow-up can add `OfdToPdfOptions.SignatureAppearanceFailure = Skip | Throw` for callers who want strictness. Image export's strict mode already contains the throwing logic.
- The SVG/HTML/image converters have their own tolerant paths; they can adopt `OfdToPdfDiagnostic`-style results later.
- Plan 007 hardens `OfdSignatureAppearanceReader.ParseXml`; land this plan first to avoid conflicts.
- Plan 014 (font-registry spike) should report registry pressure through this new channel.
