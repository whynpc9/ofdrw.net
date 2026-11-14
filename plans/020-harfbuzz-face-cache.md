# Plan 020: Shape text without copying the font on every call

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Text.HarfBuzz src/Ofdrw.Net.Layout/Text/TextFontFace.cs tests/Ofdrw.Net.Layout.Tests`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW–MED
- **Depends on**: none
- **Category**: perf
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

`HarfBuzzTextShaper` is the only shaper behind the paragraph layout engine, the
prelayout queries, the Flow "final" text engine and the DOCX native paragraph
engine. Every single shape call clones the complete font in managed memory and
then copies it again into native memory before building a fresh HarfBuzz face.
The paragraph engine shapes thousands of runs per document, so the cost is
(font size × 2) per run.

Measured at commit `427433a` on Linux x64, .NET 10.0.401, Release build, with
`e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/LXGWWenKai-Regular.ttf` (25.6 MB),
frame width 80 mm, font size 4 mm, through
`HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(...)`:

| Input (ordinary breakable CJK sentences) | Lines | Time | Managed bytes allocated |
|---|---:|---:|---:|
| 2,000 characters | 100 | 407 ms | 2.4 GB |
| 8,000 characters | 400 | 1,038 ms | 9.8 GB |
| 32,000 characters | 1,600 | 4,256 ms | 39.3 GB |

That is about 1.2 MB of allocation per character. With a 0.57 MB Latin font the
same engine lays out 64,000 characters in 409 ms, so nearly all the cost is the
font copy. After this plan, the font is copied into native memory once per
`TextFontFace` instance and every later shape call reuses it.

## Current state

Files:

- `src/Ofdrw.Net.Text.HarfBuzz/HarfBuzzTextShaper.cs` — the whole adapter (176
  lines). `ShapeCore` (lines 43–137) does all the work. `CreateOwnedBlob` (lines
  139–151) makes the native copy.
- `src/Ofdrw.Net.Layout/Text/TextFontFace.cs` — immutable, public, sealed font
  face. Holds the selected face bytes privately; the only public accessor clones.
- `tests/Ofdrw.Net.Layout.Tests/` — xUnit tests of the shaper and engines:
  `TextPrelayoutTests.cs`, `TextContextShapingTests.cs`, `ParagraphLayoutTests.cs`,
  `ParagraphLayoutF08Tests.cs`, `FlowTextEngineTests.cs`,
  `TextSizedFontSpanShapingTests.cs`, `TextPositionedRunTests.cs`,
  `TextPositionedConsumerTests.cs`.

The hot path today (`HarfBuzzTextShaper.cs:54-62`):

```csharp
            if (request.Text.Length == 0)
                return new TextShapedRun(request, Array.Empty<TextShapedGlyph>(), engineIdentity);
            using var blob = CreateOwnedBlob(request.Face.GetFaceBytes());
            using var face = new Face(blob, 0);
            if (face.UnitsPerEm != request.Face.UnitsPerEm)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "HarfBuzz and the managed reader disagree about font units.");
            using var font = new Font(face);
            font.SetFunctionsOpenType();
            font.SetScale(request.Face.UnitsPerEm, request.Face.UnitsPerEm);
```

`CreateOwnedBlob` (`HarfBuzzTextShaper.cs:139-151`):

```csharp
    private static Blob CreateOwnedBlob(byte[] faceBytes)
    {
        var pinnedBytes = GCHandle.Alloc(faceBytes, GCHandleType.Pinned);
        try
        {
            // Duplicate copies the bytes into native-owned storage before unpinning the array.
            return new Blob(pinnedBytes.AddrOfPinnedObject(), faceBytes.Length, MemoryMode.Duplicate);
        }
        finally
        {
            pinnedBytes.Free();
        }
    }
```

The managed clone (`TextFontFace.cs:153-154`):

```csharp
    /// <summary>Returns an independent copy; callers cannot mutate this face.</summary>
    public byte[] GetFaceBytes() => (byte[])_faceBytes.Clone();
```

Facts you need:

- `TextFontFace` is immutable after construction (its bytes never change) and
  admits faces up to 32 MiB (`TextFontFace.cs:20`).
- `HarfBuzzTextShaper` has no instance state. `HarfBuzzTextPrelayout` (same
  file, lines 155–176) creates a new shaper per engine, and callers may use
  engines from several threads. Any cache must therefore be static and
  thread-safe.
- The project targets `netstandard2.0;netstandard2.1`
  (`src/Ofdrw.Net.Text.HarfBuzz/Ofdrw.Net.Text.HarfBuzz.csproj`) and references
  `HarfBuzzSharp` 14.2.1.301. `System.Runtime.CompilerServices.ConditionalWeakTable<TKey,TValue>`
  exists on both targets.
- `Ofdrw.Net.Layout` does NOT grant `InternalsVisibleTo` to
  `Ofdrw.Net.Text.HarfBuzz`. Do not add one: keep the cache inside the HarfBuzz
  project and use only the public `TextFontFace` API.
- Exception contract of `ShapeCore` (lines 126–136): `TextLayoutException`
  passes through; native-loading failures become
  `TextLayoutError.NativeUnavailable`; `InvalidOperationException`,
  `ArgumentException`, `IOException` become `TextLayoutError.InvalidFont`. Keep
  this contract exactly, including for failures that now happen while creating
  the cached face.

Repository conventions:

- Code style: file-scoped namespaces, `var`, XML doc comments on public members,
  short `//` comments that explain a constraint. See the excerpts above.
- Tests: xUnit `[Fact]`/`[Theory]`. Test fonts are loaded with the helpers in
  `tests/Ofdrw.Net.Layout.Tests/ParagraphLayoutTests.cs:343-355`:
  `ParagraphLayoutTests.Face("LXGW WenKai", "LXGWWenKai-Regular.ttf")` returns a
  `TextFontFace` from `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/`. Those fonts
  are Git LFS files.

## Commands you will need

Set the environment once per shell (required by the repository's `AGENTS.md`):

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch test fonts | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` (about 41 pre-existing warnings) |
| Layout tests | `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| DOCX native tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~DocxNativeParagraphLayoutTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Fixed-font harness | `dotnet run --project e2e/Ofdrw.Net.Text.HarfBuzz.E2E/Ofdrw.Net.Text.HarfBuzz.E2E.csproj -c Release --no-build --no-restore -- <output-dir>` | exit 0, writes `r21-ascii.ofd`, `r21-han.ofd`, `text-paragraph-justify.ofd`, `text-paragraph-rtl-center.ofd` and PDFs |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Text.HarfBuzz/HarfBuzzTextShaper.cs`
- `tests/Ofdrw.Net.Layout.Tests/HarfBuzzFaceCacheTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch, even though they look related):

- `src/Ofdrw.Net.Layout/Text/TextFontFace.cs` — public API; no new members, no
  `InternalsVisibleTo`.
- `src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs` — the number
  of shape calls it makes is a separate plan (026).
- The pinned HarfBuzzSharp version and the `ExpectedNativeVersion` check.
- Any change to the glyphs, advances, offsets, clusters or ink bounds returned.
  Output must be identical to today.

## Git workflow

- Branch: `codex/harfbuzz-face-cache` (the repository's merged branches use the
  `codex/<slug>` form, e.g. `codex/release-0.3.1`).
- Commit messages: one imperative sentence, capitalized, no prefix. Example from
  `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Record the baseline output

Before changing code, build and run the fixed-font harness into a scratch
directory outside the repository, then record hashes:

```sh
BASE="$(mktemp -d)"
dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project e2e/Ofdrw.Net.Text.HarfBuzz.E2E/Ofdrw.Net.Text.HarfBuzz.E2E.csproj -c Release --no-build --no-restore -- "$BASE"
(cd "$BASE" && sha256sum *.ofd text-paragraph-receipt.json | sort -k2 > "$BASE/../baseline-020.sha256")
```

**Verify**: `wc -l < "$BASE/../baseline-020.sha256"` → at least 5.

If two baseline runs produce different hashes for a file (run the harness twice
to check), that file contains run-dependent data: exclude it from the later
comparison and note it in your report. Compare the `.ofd` files that are stable.

### Step 2: Add a per-face native cache

In `HarfBuzzTextShaper.cs`:

1. Add a private nested sealed class, for example `NativeFace`, that owns one
   `Blob` and one `Face` and exposes the `Face`. Give it a `Dispose` that
   disposes both, and a finalizer that calls it, so native memory is released
   when the owning `TextFontFace` is collected.
2. Add `private static readonly ConditionalWeakTable<TextFontFace, NativeFace> Faces = new();`.
   Key by the `TextFontFace` instance (reference identity). Do not key by
   `FaceSha256`: an instance key needs no hashing and lets the entry die with
   the face.
3. Replace the three lines that create `blob` and `face` per call with a lookup
   `Faces.GetValue(request.Face, CreateNativeFace)`. `CreateNativeFace` calls
   `request.Face.GetFaceBytes()` exactly once, calls the existing
   `CreateOwnedBlob`, creates `new Face(blob, 0)`, performs the existing
   `UnitsPerEm` comparison and throws the same `TextLayoutException(TextLayoutError.InvalidFont, "HarfBuzz and the managed reader disagree about font units.")`
   on mismatch (dispose the blob and face before throwing).
4. `ConditionalWeakTable.GetValue` may run the factory on two threads at once
   and keep only one result. Avoid leaking the loser: guard creation with a
   static lock object and re-check with `TryGetValue` inside the lock, then
   `Add`.
5. Keep `using var font = new Font(face);` per call, with `SetFunctionsOpenType`
   and `SetScale` exactly as today. A `Font` is cheap and is not shared between
   threads.
6. Do NOT dispose the cached `Face`/`Blob` at the end of `ShapeCore`.
7. Keep the lookup inside the existing `try` so creation failures are still
   translated by the existing `catch` clauses.

**Verify**: build command → `0 Error(s)`.

### Step 3: Add tests

Create `tests/Ofdrw.Net.Layout.Tests/HarfBuzzFaceCacheTests.cs` (namespace
`Ofdrw.Net.Layout.Tests`, `public sealed class HarfBuzzFaceCacheTests`), modeled
on the top of `ParagraphLayoutTests.cs`. Use
`ParagraphLayoutTests.Face(name, file)` for fonts.

1. `RepeatedShapingDoesNotCopyTheFacePerCall` — load
   `LXGWWenKai-Regular.ttf` (about 25.6 MB) as one `TextFontFace`. Shape the
   run `"敏捷的棕色狐狸"` once to warm the cache
   (`new HarfBuzzTextShaper().Shape(new TextShapeRequest("敏捷的棕色狐狸", 0, face, TextDirection.LeftToRight, "Hani", "zh"))`).
   Then record `GC.GetAllocatedBytesForCurrentThread()`, shape the same request
   200 more times, and assert the delta is below `face.GetFaceBytes().Length`
   (one font size; before this plan the delta is about 200 font sizes — do not
   count the `GetFaceBytes()` call used for the assertion itself inside the
   measured region).
2. `CachedFaceGivesIdenticalRuns` — shape the same request with two different
   `TextFontFace` instances built from the same bytes and assert the glyph
   sequences are equal element by element (glyph id, source start, source
   length, advances, offsets). Read the property names from
   `TextShapedGlyph` in `src/Ofdrw.Net.Layout/Text/TextShapeContracts.cs`.
3. `ConcurrentShapingOfOneFaceIsConsistent` — with one `TextFontFace`, run
   `Parallel.For(0, 64, ...)` shaping the same Latin request with
   `NotoSans-Regular.ttf`; assert every result has the same glyph ids and
   advances as a single-threaded reference result.
4. `UnitsMismatchStillFailsAsInvalidFont` is NOT required (it cannot be
   triggered through the public API); skip it.

**Verify**: `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --filter "FullyQualifiedName~HarfBuzzFaceCacheTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` → `Passed: 3`, `Failed: 0`.

### Step 4: Prove output is unchanged

1. Run the full Layout test project and the DOCX native tests (commands above).
2. Rebuild, rerun the fixed-font harness into a second scratch directory and
   compare with the baseline:

```sh
AFTER="$(mktemp -d)"
dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
dotnet run --project e2e/Ofdrw.Net.Text.HarfBuzz.E2E/Ofdrw.Net.Text.HarfBuzz.E2E.csproj -c Release --no-build --no-restore -- "$AFTER"
(cd "$AFTER" && sha256sum *.ofd text-paragraph-receipt.json | sort -k2) | diff - "$BASE/../baseline-020.sha256"
```

**Verify**: both test commands report `Failed: 0`; the `diff` prints nothing for
the files that were stable in Step 1.

### Step 5: Render and look at the pages

The repository requires a visual check whenever font or layout code changes,
even when hashes match. Render the harness PDFs from `$AFTER` to PNG and view
each page:

```sh
for name in r21-ascii r21-han text-paragraph-justify text-paragraph-rtl-center; do
  pdftoppm -png -r 144 "$AFTER/$name.pdf" "$AFTER/$name-page"
done
```

Open every PNG and confirm: text present, no missing glyphs (boxes), no
overlapping or shifted glyphs, Chinese and Latin both legible. Record which
pages you viewed in your report. If you cannot view images, say so in the
report and mark the PNG review as not done.

macOS Preview acceptance cannot be performed on Linux. Report it as "not
performed" — never as passed.

**Verify**: `ls "$AFTER"/*-page-*.png | wc -l` → at least 4.

## Test plan

- New file `tests/Ofdrw.Net.Layout.Tests/HarfBuzzFaceCacheTests.cs` with the
  three tests of Step 3 (allocation bound, identical runs across instances,
  concurrent consistency).
- Structural pattern: `tests/Ofdrw.Net.Layout.Tests/ParagraphLayoutTests.cs`.
- Verification: full Layout test project → `Failed: 0`, including 3 new tests.

## Done criteria

- [ ] Build command prints `0 Error(s)`
- [ ] `dotnet test tests/Ofdrw.Net.Layout.Tests/...` → `Failed: 0`; `HarfBuzzFaceCacheTests` has 3 passing tests
- [ ] DOCX native test filter → `Failed: 0`
- [ ] `grep -n "GetFaceBytes" src/Ofdrw.Net.Text.HarfBuzz/HarfBuzzTextShaper.cs` shows exactly one call, inside the face-creation helper
- [ ] Harness `.ofd` hashes equal the Step 1 baseline (for stable files)
- [ ] `git status --short` lists only the in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated; report states PNG review result and "macOS Preview: not performed"

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- Any existing Layout or DOCX native test fails after Step 2 — that means the
  shaped output changed, which this plan forbids.
- A harness `.ofd` that was stable in Step 1 differs after the change.
- `HarfBuzzSharp.Face` or `Blob` turns out not to be safe to share between
  threads in this version (the concurrency test fails or crashes). Do not add a
  global lock around shaping to make it pass; report instead.
- You find you need to modify `TextFontFace.cs` or add `InternalsVisibleTo`.

## Maintenance notes

- The cache holds one native copy of each font for as long as its
  `TextFontFace` instance is alive. A service that creates a new `TextFontFace`
  per request for the same bytes gets no benefit; the user docs for the text
  engine (`docs/content/reference/text-paragraph-layout.md`) may later advise
  reusing face instances. That doc change is deliberately not part of this plan.
- Reviewers should check: the finalizer path of the cached native objects, the
  creation lock, and that the `catch` clauses still wrap face creation.
- Plan 026 (emergency line breaking) builds on this: its timings are only
  meaningful once this plan has landed.
