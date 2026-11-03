# Plan 005: Add an OFD→PDF result/diagnostics channel and stop silently dropping seals

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
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
- **Depends on**: plans/001 (large-document reads) is recommended first but not required
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
- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs:15-17`: `public static IReadOnlyList<OfdSignatureAppearance> Read(OfdDocumentPackage package, HashSet<string>? selectedPageIds = null, int? maximumAppearances = null, CancellationToken cancellationToken = default)`. Lines 65-70 return an empty list on any exception when `maximumAppearances` is null (the default path):
  ```csharp
  catch
  {
      // A malformed or unsupported signature must not prevent the document
      // body from being converted.
      return Array.Empty<OfdSignatureAppearance>();
  }
  ```
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
- `PDF_SIGNATURE_METADATA_UNREADABLE`: `OfdSignatureAppearanceReader.Read` fell into its tolerant catch. The message includes the exception type name, not the exception message (avoid echoing untrusted content verbatim).
- `PDF_SIGNATURE_APPEARANCE_UNREADABLE`: a nested OFD seal package could not be read (the tolerant `return null` path). Include `SignatureId` if available on the appearance.
- `PDF_SIGNATURE_APPEARANCE_SKIPPED`: drawing an appearance failed in the tolerant catch. Include `PageIndex` and `SignatureId`.

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
- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs`: add an optional diagnostics sink parameter only; keep `ParseXml` unchanged (plan 007 changes it)
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
5. `ConvertAsync_StillReturnsTask`: compile-level: `await new OfdToPdfConverter().ConvertAsync(...)` still works and produces bytes.

**Verify**: build fails (the API does not exist yet). After adding empty stubs, cases 2-4 fail.

### Step 2: Add the public types (exact shape above) and `ConvertWithResultAsync`

Thread a `List<OfdToPdfDiagnostic>` and a rendered counter through `ConvertPackageAsync`. Add an optional parameter `ICollection<OfdToPdfDiagnostic>? diagnostics = null` and return the counter, or use a small internal `OfdToPdfRun` holder class. Image export passes nothing, so its behavior is unchanged. `SourcePages` = the normalized `selected` list already computed in `ConvertPackageAsync`.

**Verify**: build exits 0; `python3 scripts/check-public-api-docs.py` exits 0.

### Step 3: Report the three tolerant paths

- `OfdSignatureAppearanceReader.Read`: add an optional `Action<string>? onTolerantFailure = null` parameter, called from the bare `catch` with the exception type name before returning empty. The converter turns it into `PDF_SIGNATURE_METADATA_UNREADABLE`. To capture the type name, change the bare `catch` to `catch (Exception exception)`.
- `ReadAppearanceOfdPackageAsync`'s final `catch { return null; }` → `PDF_SIGNATURE_APPEARANCE_UNREADABLE`, recorded at the call site when the result is null. Also do this when the page-geometry check returns null (line 275).
- `DrawSignatureAppearances`' final `catch { }` → change to `catch (Exception)` and add `PDF_SIGNATURE_APPEARANCE_SKIPPED` with the page index and signature ID. Count successful draws for `RenderedSignatureAppearances`; increment after `DrawImage` returns.

Keep every existing `catch (X) { throw; }` clause and its order exactly as it is.

**Verify**: cases 1-3 pass.

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

## Done criteria

- [ ] Build exits 0; `check-public-api-docs.py` exits 0
- [ ] `OfdToPdfResultTests` pass, with any skipped case explicitly justified; PDF and CLI suites pass
- [ ] `grep -n "new OfdPackageLoadOptions()" src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs` no longer matches the nested-read line
- [ ] `IOfdToPdfConverter.cs` unchanged (`git diff --quiet -- src/Ofdrw.Net.Converter.Abstractions`)
- [ ] Visual non-regression recorded
- [ ] Docs updated; storage gate exits 0; `plans/README.md` updated

## STOP conditions

- Strict-mode (image export) tests change behavior.
- Adding the diagnostics sink requires changing `OfdSignatureAppearanceReader`'s parsing or selection logic (beyond the catch and the callback).
- Valid-seal PNGs differ between the base and new builds.
- The public API check reports undocumented symbols you cannot document (e.g. generated members).

## Maintenance notes

- A follow-up can add `OfdToPdfOptions.SignatureAppearanceFailure = Skip | Throw` for callers who want strictness. Image export's strict mode already contains the throwing logic.
- The SVG/HTML/image converters have their own tolerant paths; they can adopt `OfdToPdfDiagnostic`-style results later.
- Plan 007 hardens `OfdSignatureAppearanceReader.ParseXml`; land this plan first to avoid conflicts.
- Plan 014 (font-registry spike) should report registry pressure through this new channel.
