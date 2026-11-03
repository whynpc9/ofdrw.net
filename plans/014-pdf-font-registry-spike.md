# Plan 014 (spike): Decide how long-running services avoid exhausting the process-wide PDF font registry

> **Executor instructions**: This is an **investigation/design spike**, not a feature build. Produce the measurements and the design note described below. Do not change product behavior unless a step explicitly says so. If anything in the "STOP conditions" section occurs, stop and report. When done, update the status row for this plan in `plans/README.md` — unless a reviewer dispatched you and told you they maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Converter.Pdf/Internal/OfdEmbeddedFontResolver.cs src/Ofdrw.Net.Converter.Pdf/PdfFontRegistry.cs src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs`

## Status

- **Priority**: P2
- **Effort**: M (spike); the follow-up implementation is likely L
- **Risk**: LOW for the spike (no product change)
- **Depends on**: plans/005-ofd-to-pdf-result-and-seal-diagnostics.md (Option C below uses its diagnostics channel)
- **Category**: perf / direction
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

OFD→PDF registers every embedded font face in a **process-wide, never-evicted** registry: 256 MiB / 4,096 faces by default. Subset fonts differ per document, so nothing is shared between documents. A long-running server that converts many OFDs with embedded fonts will, after roughly 1–2k documents, fail **every** further conversion with "The configured process-wide font face budget has been reached." until it restarts. The behavior is documented (`docs/content/reference/conversion-contracts.md:71`: "达到限额会明确失败，不驱逐仍可能被 PDFsharp 使用的字体"), so it is a known trade-off. But it is an operational cliff for the most obvious deployment (a conversion web service), and users currently get no warning before it happens. The spike decides whether safe eviction is possible, or what the supported mitigation is.

## Current state

- `src/Ofdrw.Net.Converter.Pdf/Internal/OfdEmbeddedFontResolver.cs`:
  - Lines 11-21: a singleton `IFontResolver` with `_faceByFamily` / `_fontByFace` dictionaries, `_maximumBytes = 256 MiB`, `_maximumFaces = 4096`.
  - Lines 72-104, `RegisterFace`: hashes the bytes (`BinaryIdentity.Hash`) and builds family `ofd-font-{identity}-{r|b}{n|i}`. It throws `InvalidOperationException` when either budget is reached. Identical bytes are stored once.
  - There is no removal API.
  - `GetFont(faceName)` (lines 123-131) returns bytes to PdfSharpCore on demand.
  - `EnsureInstalled` (lines 141-152): PDFsharp's `GlobalFontSettings.FontResolver` can only be set before PDFsharp's first font use.
- `src/Ofdrw.Net.Converter.Pdf/PdfFontRegistry.cs`: public `MaximumRegisteredFontBytes`, `MaximumRegisteredFaces`, `RegisteredFontBytes`, `RegisterFontFace`, `EnsureInstalled`. Doc comment at line 24: "existing PDFsharp cache entries are never evicted while in use."
- Callers: `src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs` lines 81, 119, 310, 351, 375, 396, 434 and 542 register per document. Nested seal OFDs register their own fonts too (`OfdToPdfConverter.cs:206`).
- Dependency: `PdfSharpCore` 1.3.67 (`src/Ofdrw.Net.Converter.Pdf/Ofdrw.Net.Converter.Pdf.csproj`). Whether PdfSharpCore keeps its own static caches keyed by face name (e.g. `FontFactory`, `XFontSource`, `OpenTypeFontface` caches) is the crux. Removing a face from our dictionaries is unsafe if PdfSharpCore still holds the name and later calls `GetFont`.

## Questions the spike must answer (with evidence)

1. **Exhaustion curve**: how many distinct documents with embedded subset fonts can one process convert before failure, and at what memory? Measure with a console app outside `src/`, under `artifacts/spike-014/` (not committed). Loop:
   - generate an OFD with 1–3 *distinct* embedded fonts per document (for example, subset the test font `tests/Ofdrw.Net.Converter.Pdf.Vector.Tests/fonts/regular.ttf` or the trusted-font fixtures, with a varying glyph set via the existing font-subset API, or append a varying dummy table so the bytes differ)
   - convert it to PDF
   - record `PdfFontRegistry.RegisteredFontBytes`, the registered face count (via reflection or an internal accessor in a scratch build; do not commit), the process working set, and the iteration at failure.
2. **PdfSharpCore caching**: read the PdfSharpCore 1.3.67 sources (NuGet package or the GitHub tag `v1.3.67`). List every static cache that retains font bytes or face names (class, field, key). Can they be cleared through public API (e.g. `FontFactory.Reset()` or similar)? Is clearing safe while another thread is drawing?
3. **Eviction feasibility**: given (2), is any of these safe?
   - **A. Per-conversion scope**: reference-count faces per conversion, and after `PdfDocument.Save` release faces whose count is 0, also clearing the matching PdfSharpCore cache entries.
   - **B. Generation reset**: when the registry is full and no conversion is in flight (global reader/writer lock), clear our registry and PdfSharpCore's caches together.
   - **C. No eviction, better signals**: report `PDF_FONT_REGISTRY_PRESSURE` (Warning) via plan 005's `OfdToPdfConversionResult.Diagnostics` when either budget passes 80%, expose a public face count, and document process recycling (e.g. worker restart after N conversions) as the supported mitigation.
4. For A/B, prototype in a scratch branch that is **not** part of the deliverable. Run the existing font tests: `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false --filter "FullyQualifiedName~Font"`. Also run a 2-thread concurrent conversion loop for 10 minutes. Record any wrong-glyph or `GetFont` miss.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
dotnet build Ofdrw.Net.sln -c Release $F
```

## Deliverables

1. `docs/architecture/pdf-font-registry-lifetime.md` (new design note; `docs/architecture/` is the repo's location for design decisions). Contents:
   - the measured exhaustion curve (table)
   - the PdfSharpCore cache inventory with source links (file and line at tag v1.3.67)
   - options A/B/C with evidence-based safety verdicts
   - the recommendation
   - open questions
   Keep it factual; mark every unmeasured claim as such.
2. A one-paragraph addition to `docs/content/reference/conversion-contracts.md` (after the line about the 256 MiB / 4096 limit) giving the measured approximate document count and the currently recommended mitigation (process recycling). Mark it as guidance for main.
3. A follow-up implementation plan **only if** the maintainer picks an option. Do not implement A/B/C in this spike.

## Scope

**In scope**: the two documentation files above; throwaway code under `artifacts/spike-014/` (never committed).
**Out of scope**: any change under `src/`; upgrading PdfSharpCore.

## Git workflow

Branch `advisor/014-pdf-font-registry-spike`; commit e.g. `docs: record PDF font registry lifetime investigation`. No push or PR unless instructed.

## Done criteria

- [ ] The design note exists with the four sections answered, or explicitly marked unanswerable with the reason
- [ ] `git diff --stat -- src` is empty
- [ ] Docs paragraph added; `python3 scripts/check-repository-storage.py` exits 0 after staging
- [ ] `plans/README.md` updated; the report states the recommended option

## STOP conditions

- The measured exhaustion point is far higher than 4,096 faces, e.g. because PdfSharpCore dedupes something unexpected. Report it, since the premise changes.
- A prototype produces wrong glyphs. Record it in the note as evidence against that option; do not try to fix it.

## Maintenance notes

The roadmap already defers a PDF-backend decision (D10: ImageSharp/TFM). If the spike concludes that only a backend change fixes eviction properly, link the note from `docs/engineering/roadmap.md`'s "遗留问题与独立候选" table rather than starting that migration.
