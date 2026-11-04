# Plan 009: Reject non-finite numbers and normalize reader exceptions for malformed OFDs

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Reader/Readers/OfdReader.cs`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: docs/engineering/plans/001-reader-traversal-budget.md (land first; that plan's tests share the reader)
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Three small input-hygiene defects in `OfdReader` make malformed files behave inconsistently.
1. `ParseDouble` accepts `NaN`, `Infinity` and `1e999`, which overflows to Infinity on .NET Core but not on .NET Framework. Those values flow into font size, text positions, alpha and line width. Downstream converters then emit NaN coordinates or throw deep inside PDFsharp.
2. Two `TemplatePage` entries with the same ID crash `ReadAsync` with a raw `ArgumentException` from `ToDictionary`, instead of the documented `InvalidDataException`.
3. `ListDocumentsAsync` lets raw `XmlException` and `KeyNotFoundException` escape, while its sibling `ReadAsync` converts both.

Callers that catch `InvalidDataException` for bad input miss cases 2 and 3.

## Current state

`src/Ofdrw.Net.Reader/Readers/OfdReader.cs`:
- Lines 31-41, `ListDocumentsAsync`, has no exception conversion:
  ```csharp
  var archive = await _loader.LoadAsync(ofdStream, options ?? new OfdPackageLoadOptions(), cancellationToken).ConfigureAwait(false);
  var material = new OfdModelMaterialization(new OfdModelLimits());
  var xml = material.ReadXml("OFD.xml", archive.GetBytes("OFD.xml"), cancellationToken);
  ...
  ```
- Lines 54-65, `ReadAsync(Stream, OfdPackageLoadOptions, CancellationToken)`, converts them:
  ```csharp
  try { return await ReadCoreAsync(ofdStream, options, cancellationToken).ConfigureAwait(false); }
  catch (XmlException error) { throw new InvalidDataException("OFD package contains malformed XML.", error); }
  catch (KeyNotFoundException error) { throw new InvalidDataException("OFD package references a missing entry.", error); }
  ```
- Lines 253-262, TemplatePage map:
  ```csharp
  .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.BaseLoc))
  .ToDictionary(x => x.Id!, x => (BaseLoc: x.BaseLoc!, ZOrder: x.ZOrder), StringComparer.OrdinalIgnoreCase)
  ?? new Dictionary<string, (string BaseLoc, string? ZOrder)>(StringComparer.OrdinalIgnoreCase);
  ```
- Lines 1681-1693:
  ```csharp
  private static double ParseDouble(string? value, double fallback)
  {
      return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : fallback;
  }
  private static double? ParseNullableDouble(string? value)
  {
      return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
  }
  ```
  12 call sites. Notable: `FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d)` and `Weight = (int)ParseDouble(node.Attribute("Weight")?.Value, OfdTextElement.DefaultWeight)` (lines 1358-1359); attachment `SizeKilobytes = ParseNullableDouble(...)` (line 338).
- Box/matrix parsing already rejects non-finite values (`OfdBoxParser.TryParse`, line 1677: "OFD box must start with four finite numbers."). Follow that precedent.

Test conventions: `tests/Ofdrw.Net.Packaging.Tests/BadPackageCorpusTests.cs`, a sealed class of small deterministic hostile inputs, each asserting `Assert.ThrowsAsync<InvalidDataException>`. Build valid packages with `OfdPackageWriter` as in `PackageRoundTripTests.cs:19-40`, then mutate entries with `ZipArchive` in `Update` mode.

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
| New tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~ReaderInputHygieneTests"` | pass |
| Suites | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Core.Tests -c Release $F` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Reader/Readers/OfdReader.cs` (`ListDocumentsAsync`, the TemplatePage map, `ParseDouble`, `ParseNullableDouble`, and the `Weight` conversion line); `tests/Ofdrw.Net.Packaging.Tests/ReaderInputHygieneTests.cs` (create).

**Out of scope**:
- Declared `<PageRes>` support and font ID-rule unification. Both are recorded as follow-up candidates in the index.
- ZIP entry-name normalization in `OfdPackageLoader`.
- Range validation of positive-only fields beyond what is listed here.

## Git workflow

Branch `advisor/009-reader-input-hygiene`; commit e.g. `fix(reader): reject non-finite numbers and normalize malformed-input exceptions`. No push or PR unless instructed.

## Steps

### Step 1: Failing tests (`ReaderInputHygieneTests.cs`)

Helper: write a one-page package with one `OfdTextElement` (copy the `PackageRoundTripTests` builder). Then a mutation helper that loads the ZIP, finds the entry ending in `/Content.xml`, applies a string replacement, and returns the new bytes.
1. `[Theory] NaN / Infinity / -Infinity / 1e999 as TextObject Size` → reading succeeds and `FontSizeMillimeters` equals the fallback `4` (and is finite).
2. `Weight="NaN"` and `Weight="1e999"` → `Weight == OfdTextElement.DefaultWeight`.
3. Duplicate TemplatePage: mutate `Document.xml` to contain two `<ofd:TemplatePage ID="9" BaseLoc="..."/>` entries. To make one exist, write a package with a template page, or add the second line by copying the first. Assert `InvalidDataException`.
   - If building a template is awkward, inject `<ofd:TemplatePage ID="9" BaseLoc="Tpls/Tpl_0/Content.xml"/>` twice inside `<ofd:CommonData>`; the dictionary is built before the files are read.
   - Confirm in the live code (line ~250) which element contains TemplatePage and match it.
4. `ListDocuments_MalformedOfdXml`: replace `OFD.xml` with `<ofd:OFD` (truncated) → `ListDocumentsAsync` throws `InvalidDataException`, the same type as `ReadAsync` on the same input. Assert both.

**Verify**: cases 1 (at least NaN/Infinity), 2, 3 and 4 fail before Step 2.

### Step 2: Implement

- `ParseDouble`: return `fallback` unless the parse succeeds **and** `!double.IsNaN(parsed) && !double.IsInfinity(parsed)`. Do the same in `ParseNullableDouble` (return `null`). netstandard2.0 lacks `double.IsFinite`, so use the two checks.
- Weight line: `var weight = ParseDouble(..., OfdTextElement.DefaultWeight); Weight = weight >= 1 && weight <= 1000 ? (int)weight : OfdTextElement.DefaultWeight,` (an inline helper or local function is fine).
- TemplatePage: replace `ToDictionary` with an explicit loop over the same filtered sequence that throws `new InvalidDataException($"OFD document declares duplicate TemplatePage ID '{id}'.")` on a duplicate key (case-insensitive, same comparer).
- `ListDocumentsAsync`: wrap the body in the same two catches as `ReadAsync`. Extract a private static helper `Translate(Exception)` or duplicate the two catch clauses; either is fine.

**Verify**: new tests pass; the three suites pass.

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/009/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/009/trusted-font-tests
rm -rf artifacts/009/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/009/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/009/packages --output-dir artifacts/009/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/009/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] Build exits 0; `ReaderInputHygieneTests` pass; the Packaging, Pdf and Core suites pass
- [ ] `grep -n "ToDictionary(x => x.Id!" src/Ofdrw.Net.Reader/Readers/OfdReader.cs` returns nothing
- [ ] Storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- An existing test relies on `Infinity`/`NaN` values reaching the model (e.g. preflight or keyword tests asserting rejection downstream). Report the test; do not change it.
- The TemplatePage map is built in more than one place.

## Maintenance notes

- If more numeric fields are added to the reader, use `ParseDouble`/`ParseNullableDouble`; never call `double.Parse` directly.
- Follow-ups (not in this plan): honoring declared `<PageRes>` locations, and keying fonts by (resource file, canonical ID).
