# Plan 007: Route remaining XML parses of package content through the hardened parser

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Core/Models/OfdXmlEnvelope.cs src/Ofdrw.Net.Core/Models/OfdGraphicXmlContract.cs src/Ofdrw.Net.Core/Models/OfdPathStyle.cs src/Ofdrw.Net.Core/Models/OfdTextEmphasis.cs src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs src/Ofdrw.Net.Core/Models/OfdAnnotationContract.cs src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs src/Ofdrw.Net.Signatures/Signing/OfdSignatureService.cs src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs`
> If any in-scope file changed since this plan was written (other than plan 005's callback in `OfdSignatureAppearanceReader.Read`), compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: docs/engineering/plans/005-ofd-to-pdf-result-and-seal-diagnostics.md (both edit `OfdSignatureAppearanceReader.cs`)
- **Category**: security (+ small perf)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

The project's XML policy is DTDs prohibited, null resolver, and size-capped. It is applied by the main reader, preflight, DOCX validator and stamp parsing. Several parsers of **untrusted package bytes** still call `XDocument.Parse`/`XElement.Parse` with framework defaults, which *parse* DTDs and allow up to 10 million characters of internal entity expansion per document. External entities are not resolved, so this is not XXE. It is a CPU and memory amplifier on files supplied by users, in code paths that run on every OFD→PDF conversion and every signature verification. One of them, `OfdGraphicXmlContract.IsVisible`, also re-parses an element's entire XML on every call, from 9 call sites per render.

## Current state

- The hardened helper already exists, `src/Ofdrw.Net.Core/Models/OfdXmlEnvelope.cs:89-97`:
  ```csharp
  internal static XElement Parse(string xml)
  {
      using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
      {
          DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
          MaxCharactersInDocument = new OfdModelLimits().MaxSourceXmlBytes
      });
      return XElement.Load(reader, LoadOptions.PreserveWhitespace);
  }
  ```
  Core grants `InternalsVisibleTo` to Packaging, Signatures, Converter.Pdf and others (`src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj:8-21`).
- Default-settings parse sites to convert:
  1. `src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs:35`: `ofdXml = XDocument.Parse(archive.ReadUtf8Text(OfdConstants.OfdRootFile));`
  2. Same file, `:92`: `documentXml = XDocument.Parse(archive.ReadUtf8Text(docRoot));`. Inside a loop over `DocRoot` values that are **not de-duplicated** (lines 47-51, 64-77), so one small `Document.xml` listed many times is parsed many times.
  3. `src/Ofdrw.Net.Signatures/Verification/OfdSignatureVerifier.cs:68`: `ofd = XDocument.Parse(archive.ReadUtf8Text("OFD.xml"), LoadOptions.PreserveWhitespace);`
  4. `src/Ofdrw.Net.Signatures/Signing/OfdSignatureService.cs:437-443`, `ParseXml(byte[] bytes, string path)`: `XDocument.Parse(Encoding.UTF8.GetString(bytes).TrimStart('﻿'), LoadOptions.PreserveWhitespace)`
  5. `src/Ofdrw.Net.Converter.Pdf/Internal/OfdSignatureAppearanceReader.cs:440-445`, `ParseXml(byte[] data)`: the same pattern
  6. `src/Ofdrw.Net.Core/Models/OfdGraphicXmlContract.cs:87`: `return XElement.Parse(xml!).Attribute("Visible")?.Value.Trim().ToLowerInvariant() is not ("false" or "0");`. Called per element, per pass.
  7. `src/Ofdrw.Net.Core/Models/OfdPathStyle.cs:26`: `XElement.Parse(styleXml!)`
  8. `src/Ofdrw.Net.Core/Models/OfdTextEmphasis.cs:32`: `XElement.Parse(text.SourceXml!)`
  9. `src/Ofdrw.Net.Core/Fonts/EmbeddedFontCoverage.cs:13`: `try { return XElement.Parse(sourceXml!)....` (inside try)
  10. `src/Ofdrw.Net.Core/Models/OfdAnnotationContract.cs:73`: `var annotation = XElement.Parse(sourceXml);`
- Exemplar of a hardened document-level parse: `src/Ofdrw.Net.Signatures/Signing/OfdSignatureService.Stamps.cs:200-211` (`XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, ... })` then `XDocument.Load(reader)`).
- Behavior differences to be aware of:
  - `OfdXmlEnvelope.Parse` preserves whitespace, while `XElement.Parse(string)` discards insignificant whitespace. Sites 6-10 only read attributes or element names, so whitespace text nodes don't matter.
  - With `DtdProcessing.Prohibit`, a document containing `<!DOCTYPE` throws `XmlException`. Every site above already catches `Exception`/`XmlException` or runs on reader-validated XML (the reader rejects DTDs at `OfdModelMaterialization.cs:212`). Keep each site's existing catch behavior.

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
| New tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~XmlHardeningTests"` | pass |
| Suites | `dotnet test tests/Ofdrw.Net.Core.Tests ...`, `Ofdrw.Net.Packaging.Tests`, `Ofdrw.Net.Signatures.Tests`, `Ofdrw.Net.Converter.Pdf.Tests` (each `-c Release $F`) | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: the 10 sites listed; `OfdXmlEnvelope.cs` (add `ParseDocument`); `tests/Ofdrw.Net.Packaging.Tests/XmlHardeningTests.cs` (create; Packaging.Tests references Core, Packaging, Reader, Signatures and Layout).

**Out of scope**:
- `XElement.Parse` calls in `OfdPackageWriter.cs`, `OfdPackageWriter.Owned.cs`, `OfdDocumentMerger.cs`, `PdfVectorToOfdConverter.cs:181`. They parse model strings already validated by the reader or produced by the SDK. Changing them risks byte-exact writer output.
- `OfdPackageStructureChecker` path-resolution semantics (`Resolve`, lines 150-178). Aligning them with `OfdPackagePath` changes validation results and deserves its own decision.
- CLI `DocumentTools.cs` `XDocument.Load(root)`, which runs on reader-validated `OFD.xml`.

## Git workflow

Branch `advisor/007-xml-parse-hardening`; commit e.g. `fix(security): parse package XML with DTDs prohibited everywhere`. No push or PR unless instructed.

## Steps

### Step 1: Add `ParseDocument`

In `OfdXmlEnvelope`, add `internal static XDocument ParseDocument(string xml, LoadOptions options = LoadOptions.None)`. Use the same `XmlReaderSettings` as `Parse`; then `XDocument.Load(reader, options)`.

**Verify**: build exits 0.

### Step 2: Failing tests (`XmlHardeningTests.cs`)

Build a valid OFD with `OfdPackageWriter`, then rewrite the ZIP with `OFD.xml` replaced by a copy whose content starts with `<!DOCTYPE ofd:OFD [ <!ENTITY a "aaaaaaaaaa"> ]>` (a small, harmless internal entity; **no** recursive expansion payload is needed), followed by the original root.
- `StructureChecker_RejectsDtd`: `OfdPackageStructureChecker` reports an issue for `OFD.xml` (it currently parses it successfully). Find the public entry point in `src/Ofdrw.Net.Packaging/Validation/OfdPackageStructureChecker.cs` and assert on its issue list. Before the fix, it reports no OFD.xml parse issue.
- `SignatureVerifier_RejectsDtd`: `OfdSignatureVerifier` returns a report with an issue, and not a successful parse. Inspect `OfdSignatureVerifier.cs` around line 70 for the catch behavior, and assert on what it records.
- `StructureChecker_ParsesRepeatedDocRootOnce`: OFD.xml with 50 identical `DocBody/DocRoot` entries. The checker still completes and reports the same issues as with one. This is a behavioral assertion; there is no timing assertion.
- `IsVisible_UsesHardenedParser`: `OfdGraphicXmlContract` is internal, and Packaging.Tests has Core internals. A text element whose `SourceXml` contains a DOCTYPE → `IsVisible` throws `XmlException`. Before the fix it parsed successfully.

**Verify**: these 4 tests fail before Step 3.

### Step 3: Convert sites 1-5 (documents)

Replace each `XDocument.Parse(text[, options])` with `OfdXmlEnvelope.ParseDocument(text[, options])`. Keep the `TrimStart('﻿')` and the existing catches exactly. In `OfdPackageStructureChecker`, de-duplicate DocRoot paths, normalized as that file already normalizes them, before the parse loop. Use an ordinal-ignore-case `HashSet<string>`, matching the archive's case-insensitive lookup.

### Step 4: Convert sites 6-10 (fragments)

Replace `XElement.Parse(x)` with `OfdXmlEnvelope.Parse(x)`. For site 6 (`IsVisible`), also cache the result per string instance: `private static readonly ConditionalWeakTable<string, object> VisibleCache` (namespace `System.Runtime.CompilerServices`, available on netstandard2.0). Store boxed `true`/`false`, and look up with `TryGetValue` before parsing. `SourceXml` strings are immutable, so caching by instance is safe even though the property is settable: a new value is a new string instance.

**Verify**: build exits 0; Step 2 tests pass; all four suites pass.

### Step 5: Confirm no default-settings parse remains in scope

`grep -n "XDocument.Parse\|XElement.Parse" <the 10 files>` → only matches inside `OfdXmlEnvelope.cs`, or none.

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/007/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/007/trusted-font-tests
rm -rf artifacts/007/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/007/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/007/packages --output-dir artifacts/007/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/007/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] Build exits 0; `XmlHardeningTests` pass; Core/Packaging/Signatures/Pdf suites pass
- [ ] The Step 5 grep returns no default-settings parse in the 10 files
- [ ] `git diff --stat` touches only in-scope files
- [ ] Storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- Any existing test feeds a DOCTYPE into one of these paths and expects success. Report it; it may document intended tolerance.
- A signature-verification or rendering test changes its result. Whitespace preservation is the likely cause; report rather than switching `LoadOptions` blindly.

## Maintenance notes

- New code parsing package or model XML should call `OfdXmlEnvelope.Parse`/`ParseDocument`, never `XDocument.Parse`. Reviewers should grep for `XDocument.Parse(` and `XElement.Parse(` in new diffs.
- The structure checker's path resolution still differs from `OfdPackagePath` (a `..` that climbs above the root is silently dropped). That is recorded in the plans index as a follow-up candidate.
