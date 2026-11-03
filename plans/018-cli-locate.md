# Plan 018: Add a read-only `ofdrw locate` command (keyword hits with page and box, JSON)

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Cli src/Ofdrw.Net.Reader/Keywords src/Ofdrw.Net.Converter.Pdf/Geometry docs/content/reference/cli.md docs/content/guides/keyword-replacement.md`
> `Program.cs` edits from plans 005/008/015/016/017 are expected. Changes under `Keywords/` or `Geometry/` are a STOP condition.

## Status

- **Priority**: P3
- **Effort**: M
- **Risk**: LOW (read-only; no output OFD)
- **Depends on**: plans/017-cli-attachments.md (serialize `Program.cs` edits)
- **Category**: direction (requested by the maintainer)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

`ofdrw replace-text` rewrites a document in one shot and rejects the whole batch when geometry is unsupported. There is no way to **preview** what it would match. The SDK's read-only `OfdKeywordLocator` already returns hits with page index, matched text, glyph bounds and a `CanReplace` flag, plus diagnostics. A `locate` command exposes it for scripting (e.g. "find the 签名 field's position") and as a dry run before `replace-text`, using the same trusted-font identities.

## Current state

- `src/Ofdrw.Net.Reader/Keywords/OfdKeywordLocator.cs:18-33`:
  ```csharp
  public sealed class OfdKeywordLocator
  {
      public OfdKeywordLocator(IOfdGlyphGeometryProvider geometryProvider, OfdKeywordLocatorLimits? locatorLimits = null)
      public OfdKeywordSearchResult Find(OfdDocumentPackage source, string keyword, OfdKeywordSearchOptions? options = null, ...)
      public OfdKeywordSearchResult Find(OfdDocumentPackage source, IReadOnlyList<int> pageIndexes, string keyword, ...)
  ```
  Read the full parameter lists in the file; there may be a trailing `CancellationToken`.
- `src/Ofdrw.Net.Reader/Keywords/OfdKeywordSearchResult.cs`:
  - Lines 238-252: `OfdKeywordSearchResult { Hits, Diagnostics, IsComplete }`
  - Lines 107-138: `OfdKeywordHit { PageIndex, MatchedText, SourceVersion, CanReplace, Fragments, BoundsMillimeters (OfdGlyphBox?), BoundsKind }`
  - Lines 142-153: `OfdKeywordDiagnostic { Code, PageIndex, SourcePath, MakesScopeIncomplete, ... }`
- `src/Ofdrw.Net.Reader/Keywords/OfdGlyphGeometryContracts.cs:36-46`: `OfdGlyphBox(double left, double top, double right, double bottom)` with matching properties.
- Geometry provider: `src/Ofdrw.Net.Converter.Pdf/Geometry/OfdTrustedEmbeddedFontGeometryProvider.cs:45`, `public OfdTrustedEmbeddedFontGeometryProvider(IReadOnlyList<OfdTrustedEmbeddedFontIdentity> trustedFonts)` ("Empty lists trust nothing"). Identity: `OfdTrustedEmbeddedFontIdentity(string sha256, int byteCount, bool bold = false, bool italic = false)`.
- The `--trusted-font SHA256:BYTECOUNT` parsing already exists in `src/Ofdrw.Net.Cli/TextReplacementArguments.cs:31-38`:
  ```csharp
  case "--trusted-font":
      if (trusted.Count == OfdTrustedEmbeddedFontGeometryProvider.MaximumTrustedIdentities)
          throw new ArgumentException("Too many trusted original-font identities.");
      var colon = value.IndexOf(':');
      if (colon != 64 || value.Length > 76 || !int.TryParse(value.AsSpan(colon + 1), NumberStyles.None, CultureInfo.InvariantCulture, out var length))
          throw new ArgumentException("Trusted font must be SHA256:BYTECOUNT.");
      trusted.Add(new OfdTrustedEmbeddedFontIdentity(value[..colon], length)); break;
  ```
  Extract it into `internal static OfdTrustedEmbeddedFontIdentity ParseTrustedFont(string value, int currentCount)` in the same file, and call it from both parsers. This is a pure refactor; `TextReplacementArgumentsTests` must still pass.
- Guide: `docs/content/guides/keyword-replacement.md` (SDK usage and trusted-font concept). CLI reference: `docs/content/reference/cli.md`.
- Test exemplars:
  - `tests/Ofdrw.Net.Packaging.Tests/KeywordLocatorTests.cs` uses a fake `IOfdGlyphGeometryProvider` (`new Provider()`); the CLI can't inject one.
  - `tests/Ofdrw.Net.Converter.Pdf.Tests/KeywordLocatorRealFontTests.cs` uses real trusted fonts and shows how fixtures with embedded fonts are built and how their SHA-256 and byte count are obtained.
  - CLI tests capturing Console must use `[Collection("Owned CLI console")]`.

## Command contract

```
ofdrw locate <input.ofd> --find TEXT [--trusted-font SHA256:BYTECOUNT ...] [--pages 1,3-5] [--document N] [--json]
```
- Read the input with `OfdReader`, at the `--document` index (1-based; reuse `ParseOneBasedPage`). Pages are 1-based distinct positions; reuse `TextReplacementArguments.ParsePages` semantics.
- Provider: `new OfdTrustedEmbeddedFontGeometryProvider(identities)`. Zero identities are allowed: the locator then reports what it can, and hits may lack exact geometry or have `CanReplace=false`, as the SDK decides.
- Human output:
  - one line per hit: `page <1-based> "<MatchedText>" box L,T,R,B (mm, 3 decimals, invariant) replaceable=<true|false>`
  - then one line per diagnostic: `diagnostic <Code> page <n|-> <SourcePath|->`
  - then `Complete: <true|false>`
- `--json`: `{"schema":"ofdrw-cli-locate/1","complete":bool,"hits":[{"page":1,"text":"..","bounds":{"left":..,"top":..,"right":..,"bottom":..}|null,"canReplace":bool}],"diagnostics":[{"code":"..","page":1|null,"sourcePath":".."|null,"makesScopeIncomplete":bool}]}`, camelCase, 1-based pages, nothing else on stdout.
- Exit codes: `0` = complete, ≥ 1 hit; `2` = complete, 0 hits; `3` = incomplete (`IsComplete == false`); `1` = error; `130` = cancelled.
- No output file is ever written.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F --filter "FullyQualifiedName~LocateCliTests"` | pass (the real-font case may be Skipped with a reason) |
| Real-font run | `python3 scripts/prepare-trusted-fonts.py --output-dir artifacts/trusted-fonts`, then rerun with `OFDRW_TEST_TRUSTED_CJK_FONT=$PWD/artifacts/trusted-fonts/fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf` | the real-font case runs and passes |
| CLI suite | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Cli/LocateCommand.cs` (create); `src/Ofdrw.Net.Cli/TextReplacementArguments.cs` (extract `ParseTrustedFont` only); `src/Ofdrw.Net.Cli/Program.cs` (allowlist, early dispatch, help); `tests/Ofdrw.Net.Cli.Tests/LocateCliTests.cs` (create); `docs/content/reference/cli.md`; `docs/content/guides/keyword-replacement.md`.

**Out of scope**: any SDK change in `Keywords/` or `Geometry/`; changing `replace-text` behavior; a `--dry-run` flag on `replace-text`. `locate` is the preview.

## Git workflow

Branch `advisor/018-cli-locate`; commit e.g. `feat(cli): add read-only locate command`. No push or PR unless instructed.

## Steps

### Step 1: Refactor `ParseTrustedFont`

Extract it as described. **Verify**: `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F --filter "FullyQualifiedName~TextReplacementArguments"` passes unchanged.

### Step 2: Tests (`LocateCliTests.cs`, `[Collection("Owned CLI console")]`)

1. `Locate_RequiresFind`: exit 1.
2. `Locate_BadTrustedFont`: `--trusted-font abc` → exit 1 with the same message as replace-text.
3. `Locate_Json_Shape`: a simple OFD (`OfdPackageWriter`, one `OfdTextElement` "hello world", name-only font) with `--find world --json` → stdout parses and `schema == "ofdrw-cli-locate/1"`. Exit code ∈ {0, 2, 3}, and it is consistent with `complete` and the hit count (assert the mapping rule, not a specific outcome; name-only fonts may be incomplete).
4. `Locate_NoMatch`: `--find zzz` on the same file → `hits` is empty; exit 2 if `complete`, else 3.
5. `Locate_RealFont_FindsHitWithBounds`: use a `[Fact]` subclass that skips when `OFDRW_TEST_TRUSTED_CJK_FONT` is unset (pattern: `tests/Ofdrw.Net.Core.Tests/Audit024ProtocolTests.cs:128-131`). Build an OFD with that font embedded, following `KeywordLocatorRealFontTests`, and pass `--trusted-font <sha256>:<bytes>` computed from the file. Expect exit 0, page 1, non-null bounds, `canReplace` true.
6. `Locate_WritesNothing`: the directory listing is unchanged after a run.

**Verify**: tests fail (unknown command).

### Step 3: Implement `LocateCommand.cs`

Use a dedicated argument parser (pairs, literal values, reject unknown or duplicate options except repeated `--trusted-font`), `OfdReader` → `OfdKeywordLocator.Find(...)` → the renderers. Dispose any `IDisposable` results or providers, if the SDK types implement it (check). Add `"locate"` to the allowlist, the early dispatch and help (`locate       Read-only keyword search with page and glyph bounds; preview for replace-text.`).

**Verify**: new tests pass (case 5 skipped or passing); the CLI suite passes; the real-font run passes.

### Step 4: Docs

`cli.md`: command row, usage, exit-code sub-table, marked "main, not yet released". `keyword-replacement.md`: a short "CLI preview" paragraph showing `ofdrw locate ... --json` before `ofdrw replace-text`, with the same trusted-font explanation and a link instead of duplicated text.

## Done criteria

- [ ] Build exits 0; `LocateCliTests` pass (case 5 passes in the real-font run); CLI suite passes
- [ ] `git diff --stat -- src/Ofdrw.Net.Reader src/Ofdrw.Net.Converter.Pdf` is empty
- [ ] Docs updated; storage gate exits 0; `plans/README.md` updated

## STOP conditions

- `OfdKeywordLocator.Find` requires a cancellation token or session lifecycle that differs from the excerpt in a way that affects the design (e.g. mandatory disposal of a session the CLI can't own). Report the actual signature.
- The real-font case returns `canReplace=false` for a plain Regular embedded font. Report it: it would contradict `replace-text` acceptance.

## Maintenance notes

- `schema: ofdrw-cli-locate/1` is a public contract.
- If `replace-text` later gains new admission rules, `locate`'s `canReplace` follows automatically, because both use the same SDK.
