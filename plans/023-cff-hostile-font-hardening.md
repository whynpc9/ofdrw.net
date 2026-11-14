# Plan 023: A hostile CFF font can only make its own font "unsupported"

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Core/Fonts tests/Ofdrw.Net.Layout.Tests`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

OFD documents from third parties embed fonts. When such a document is exported
to PDF, every embedded OpenType/CFF font is interpreted by the SDK's managed CFF
CharString interpreter and converted glyph by glyph to a TrueType rendition.
The intended contract is: a font that cannot be handled becomes "unsupported"
for that export (the converter catches `NotSupportedException` and
`InvalidDataException`) and the rest of the document still converts.

Three gaps break that contract:

1. **No whole-font budget.** Budgets are per glyph (200,000 interpreter
   operations, 32,768 points). A font may declare 65,535 glyphs that all reuse
   one shared subroutine, so a font of a few hundred KB can demand about 10^10
   interpreter operations and gigabytes of converted output. The export then
   stalls or dies with `IOException`/`OutOfMemoryException`, which the converter
   does not catch.
2. **Wrong exception type.** A CFF INDEX whose last offset is close to
   `int.MaxValue` makes a `checked` addition throw `OverflowException` before the
   bounds check runs. That escapes the converter's filter and aborts the whole
   export.
3. **NaN passes the coordinate guard.** CharString arithmetic can produce
   Infinity and NaN. The point guard uses `Math.Abs(x) > 1_000_000`, which is
   false for NaN, so NaN coordinates reach SVG path data and the PDF rendition.

After this plan each of these ends in `InvalidDataException` from the CFF code,
so the existing handlers mark that one font unsupported. Output for valid fonts
does not change by a single byte.

## Current state

Files:

- `src/Ofdrw.Net.Core/Fonts/CffCharStringOutlines.cs` (723 lines) — contains
  `GlyphOutlineSources`, `CffFontProgram`, `CffIndex`, `CffDict` and
  `CffCharStringOutlines` with its private nested `Interpreter`.
- `src/Ofdrw.Net.Core/Fonts/CffTrueTypeConversion.cs` — `Convert(byte[] faceBytes, CancellationToken)`.
- Callers (do not edit; listed so you know the exception contract):
  - `src/Ofdrw.Net.Converter.Pdf/Internal/DocumentFontContext.cs:70-76`
    ```csharp
                    try { data = CffTrueTypeConversion.Convert(data, _cancellationToken); }
                    catch (Exception exception) when (exception is NotSupportedException or System.IO.InvalidDataException)
                    {
                        _unsupportedCoverage[font] = new NotSupportedException(
                            "PDF export could not build a TrueType rendition of the CFF font program: " + exception.Message, exception);
                        continue;
                    }
    ```
  - `src/Ofdrw.Net.Packaging/OfdFontEmbeddingPlan.cs:221` catches only
    `InvalidDataException` around `CffSubset.Create`.
  - `src/Ofdrw.Net.Layout/Text/TextFontFace.cs:115` already translates
    `InvalidDataException || OverflowException || IndexOutOfRangeException` —
    this shows the intended contract.
- Tests: `tests/Ofdrw.Net.Layout.Tests/ParagraphLayoutF08Tests.cs` (lines
  339–423 are the CFF tests). Test font:
  `tests/Ofdrw.Net.Layout.Tests/testdata/cff-han/OfdrwCIHanCFF-Regular.otf`,
  located with the helper `ParagraphLayoutF08Tests.CffPath(file)` (internal
  static, line 24).

The per-glyph budgets (`CffCharStringOutlines.cs:291-295`):

```csharp
    private const int MaximumFaceBytes = 32 * 1024 * 1024;
    private const int MaximumSourceBytes = 64 * 1024 * 1024;
    private const int MaximumDepth = 10;
    private const int MaximumOperations = 200_000;
    private const int MaximumPoints = 32768;
```

The interpreter loop head (`CffCharStringOutlines.cs:420-427`):

```csharp
        private void Execute(int start, int end, int depth)
        {
            if (depth > MaximumDepth) throw new InvalidDataException("CFF subroutine nesting exceeded its budget.");
            var cursor = start;
            while (cursor < end)
            {
                if (++_operations > MaximumOperations) throw new InvalidDataException("CFF CharString exceeded its operation budget.");
                if ((_operations & 1023) == 0) _token.ThrowIfCancellationRequested();
```

Operand push and the point guard (`CffCharStringOutlines.cs:398-402`, `704-710`):

```csharp
        private void Push(double value)
        {
            if (_sp >= _stack.Length) throw new InvalidDataException("CFF CharString operand stack overflow.");
            _stack[_sp++] = value;
        }
...
        private TrueTypeOutlinePoint Point(double x, double y, bool onCurve)
        {
            if (++_points > MaximumPoints) throw new InvalidDataException("CFF outline point budget exceeded.");
            if (Math.Abs(x) > 1_000_000 || Math.Abs(y) > 1_000_000)
                throw new InvalidDataException("CFF point exceeds bounded design coordinates.");
            return new TrueTypeOutlinePoint(x, y, onCurve);
        }
```

The INDEX reader (`CffCharStringOutlines.cs:196-206`):

```csharp
        for (var i = 0; i <= count; i++)
        {
            uint value = 0;
            for (var j = 0; j < offSize; j++) value = (value << 8) | cff[offsetsStart + i * offSize + j];
            if (value == 0 || value > int.MaxValue || i == 0 && value != 1 || i > 0 && value < offsets[i - 1])
                throw new InvalidDataException("Invalid CFF INDEX offsets.");
            offsets[i] = (int)value;
        }
        var end = checked(dataStart + offsets[count] - 1);
        OpenTypeFace.Require(cff, dataStart, end - dataStart);
        return new CffIndex(start, dataStart, end, offsets);
```

The conversion loop (`CffTrueTypeConversion.cs:19-41`):

```csharp
    internal static byte[] Convert(byte[] faceBytes, CancellationToken cancellationToken = default)
    {
        var face = new OpenTypeFace(faceBytes);
        var outlines = new CffCharStringOutlines(faceBytes);
        var count = outlines.GlyphCount;
        using var glyf = new MemoryStream();
        var loca = new byte[(count + 1) * 4];
        int maxPoints = 0, maxContours = 0;
        for (var gid = 0; gid < count; gid++)
        {
            if ((gid & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            OpenTypeFace.Put32(loca, gid * 4, (uint)glyf.Length);
            if (gid == 0) continue; // .notdef stays empty; it is never drawn.
            var outline = outlines.Read(gid, cancellationToken);
            var contours = Quadratic(outline);
            if (contours.Count == 0) continue;
            var points = contours.Sum(contour => contour.Count);
            if (points > 65535 || contours.Count > 32767) throw new NotSupportedException("A CFF glyph is too complex for TrueType.");
            maxPoints = Math.Max(maxPoints, points); maxContours = Math.Max(maxContours, contours.Count);
            var bytes = Encode(contours);
            glyf.Write(bytes, 0, bytes.Length);
            while (glyf.Length % 4 != 0) glyf.WriteByte(0);
        }
```

`Read` creates one `Interpreter` per glyph (`CffCharStringOutlines.cs:349-360`):

```csharp
    public TrueTypeGlyphOutline Read(int glyphId, CancellationToken cancellationToken = default)
    {
        if (glyphId <= 0 || glyphId >= GlyphCount)
            throw new NotSupportedException("The requested original GID is outside the admitted font.");
        var interpreter = new Interpreter(_program, _program.LocalSubrs(glyphId), cancellationToken);
```

Conventions: C# with file-scoped namespaces, `internal` types, budget failures
are `InvalidDataException` with a short English sentence. Tests are xUnit in
`tests/Ofdrw.Net.Layout.Tests` (that project can see `Ofdrw.Net.Core` internals).

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Layout tests | `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Packaging tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~FontSubset" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Skia CFF tests | `dotnet test tests/Ofdrw.Net.Graphics.SkiaSharp.Tests/Ofdrw.Net.Graphics.SkiaSharp.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Core/Fonts/CffCharStringOutlines.cs`
- `src/Ofdrw.Net.Core/Fonts/CffTrueTypeConversion.cs`
- `tests/Ofdrw.Net.Layout.Tests/CffHardeningTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- `DocumentFontContext.cs`, `OfdFontEmbeddingPlan.cs`, `OfdExplicitGlyphFont.cs`
  — the fix is to throw the type they already handle, not to widen their filters.
- `FontParseBudget` — it is the cmap-parsing budget with its own message; do not
  reuse it for CharString work.
- `src/Ofdrw.Net.Graphics.SkiaSharp` — a separate, Skia-based CFF reader.
- The per-glyph limits and any behavior for valid fonts.

## Git workflow

- Branch: `codex/cff-hostile-font-hardening`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Pin today's output for a valid font (before any source change)

Create `tests/Ofdrw.Net.Layout.Tests/CffHardeningTests.cs` (namespace
`Ofdrw.Net.Layout.Tests`, `public sealed class CffHardeningTests`). Add:

- `ValidFontRenditionIsUnchanged` — read
  `ParagraphLayoutF08Tests.CffPath("OfdrwCIHanCFF-Regular.otf")`, call
  `CffTrueTypeConversion.Convert(bytes)`, compute the lowercase hex SHA-256 of
  the result and `Assert.Equal` it to a constant.

Run the test once with a placeholder constant, copy the actual hash from the
failure message into the constant, and run again.

**Verify**: `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --filter "FullyQualifiedName~CffHardeningTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`
→ `Passed: 1`. This hash is your baseline; it must never change in later steps.

### Step 2: INDEX offsets fail as `InvalidDataException`

In `CffIndex.Read`, replace the `checked` end computation with a `long`
comparison that throws the existing message type:

```csharp
        var endLong = (long)dataStart + offsets[count] - 1;
        if (endLong > cff.Length) throw new InvalidDataException("CFF INDEX data exceeds the font program.");
        var end = (int)endLong;
```

Keep the following `OpenTypeFace.Require(...)` line.

Add test `IndexWithOverflowingLastOffsetIsInvalidData`: build a 13-byte array
by hand — count `00 01`, offSize `04`, offsets `00 00 00 01` and
`7F FF FF FF`, then one data byte — and assert
`Assert.Throws<InvalidDataException>(() => CffIndex.Read(bytes, 0))`.

**Verify**: filter `CffHardeningTests` → `Passed: 2`.

### Step 3: Non-finite operands are rejected

In `Interpreter.Push`, reject NaN and Infinity before storing:

```csharp
            if (double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidDataException("CFF CharString produced a non-finite operand.");
```

In `Interpreter.Point`, make the guard NaN-safe by testing the accepted range:

```csharp
            if (!(Math.Abs(x) <= 1_000_000) || !(Math.Abs(y) <= 1_000_000))
                throw new InvalidDataException("CFF point exceeds bounded design coordinates.");
```

Add test `ArithmeticOverflowInCharStringIsInvalidData`. Build the input from the
valid Han font so no font has to be authored from scratch:

1. `var face = new OpenTypeFace(bytes); var program = CffFontProgram.Parse(face);`
2. Choose a glyph id `gid >= 1` with `program.CharStrings.ItemLength(gid) >= 48`.
3. Copy the `"CFF "` table bytes (`face.Tables["CFF "]`), and at
   `program.CharStrings.ItemStart(gid)` overwrite the CharString with a short
   Type 2 program that pushes a large number, squares it repeatedly with the
   two-byte `dup` (`12 27`) and `mul` (`12 24`) operators until the value
   overflows to Infinity, and ends with `endchar` (`14`). Read
   `Interpreter.Execute` for the number and operator encodings it accepts.
4. Put the modified table back (`face.Tables["CFF "] = modified;`), rebuild with
   `face.Build(0x4F54544F)` (check `OpenTypeFace.Build`'s signature) and assert
   `Assert.Throws<InvalidDataException>(() => new CffCharStringOutlines(rebuilt).Read(gid))`.

**Verify**: filter `CffHardeningTests` → `Passed: 3`, and
`ValidFontRenditionIsUnchanged` still passes with the same hash.

### Step 4: A whole-font budget for the TrueType rendition

1. In `CffCharStringOutlines.cs` add an internal sealed class
   `CffWorkBudget` with a `long Remaining` set from the constructor and a method
   `void Spend(long work)` that throws
   `InvalidDataException("CFF font exceeded its total interpretation budget.")`
   when `work > Remaining`, otherwise subtracts.
2. Add an overload
   `internal TrueTypeGlyphOutline Read(int glyphId, CffWorkBudget? budget, CancellationToken cancellationToken)`
   that passes the budget to the `Interpreter` (add an optional constructor
   parameter and field). Make the existing public `Read(int, CancellationToken)`
   call it with `null`. Do not change the `IGlyphOutlineSource` interface.
3. In `Interpreter.Execute`, next to the per-glyph check
   (`if (++_operations > MaximumOperations) ...`), add `_budget?.Spend(1);`.
4. In `CffTrueTypeConversion`:
   - add constants `DefaultMaximumOperations = 256_000_000L` and
     `DefaultMaximumGlyfBytes = 64L * 1024 * 1024`;
   - add an internal overload
     `Convert(byte[] faceBytes, CancellationToken cancellationToken, long maximumOperations, long maximumGlyfBytes)`
     containing the current body; the existing two-parameter `Convert` calls it
     with the defaults;
   - create one `CffWorkBudget(maximumOperations)` before the loop and call
     `outlines.Read(gid, budget, cancellationToken)`;
   - after computing `bytes = Encode(contours)`, before writing, throw
     `new InvalidDataException("CFF TrueType rendition exceeds its size budget.")`
     when `glyf.Length + bytes.Length + 3 > maximumGlyfBytes`.

Add tests (use the valid Han font and small budgets — no hostile font needed):

- `TotalOperationBudgetStopsTheRendition` —
  `Assert.Throws<InvalidDataException>(() => CffTrueTypeConversion.Convert(bytes, default, maximumOperations: 1000, maximumGlyfBytes: long.MaxValue))`.
- `TotalSizeBudgetStopsTheRendition` — same with
  `maximumOperations: long.MaxValue, maximumGlyfBytes: 1024`.
- `DefaultBudgetsAdmitTheValidFont` — `Convert(bytes)` succeeds (this is also
  covered by the hash test; keep it as an explicit statement).

**Verify**: filter `CffHardeningTests` → `Passed: 6`; the hash test still passes.

### Step 5: Regression and visual check

1. Run the full Layout tests, the Packaging `FontSubset` filter and the Skia
   test project (commands above).
2. Generate the CFF visual sample and render it:

```sh
OUT="$(mktemp -d)"
OFDRW_PARAGRAPH_ARTIFACTS="$OUT" dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~EveryF08OutputFeatureRoundTripsAndExports" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
pdftoppm -png -r 144 "$OUT/f08-visual-sample.pdf" "$OUT/f08-page"
```

Open each PNG. The last line of the sample reads
`Arabic مرحبا بالعالم and CFF 思源黑体字体测试.` — confirm the CFF Chinese
characters are drawn (no boxes, no missing glyphs, no misplaced outlines) and
the other lines are legible. Record the pages viewed. If you cannot view
images, say so and mark the PNG review as not done. macOS Preview acceptance is
not possible on Linux: report it as "not performed".

**Verify**: all three test commands `Failed: 0`; `ls "$OUT"/f08-page-*.png` lists
at least one file.

## Test plan

- New file `tests/Ofdrw.Net.Layout.Tests/CffHardeningTests.cs` with 6 tests:
  rendition hash unchanged, INDEX overflow, arithmetic overflow in a
  CharString, operation budget, size budget, defaults admit the valid font.
- Pattern: the CFF section of `ParagraphLayoutF08Tests.cs` (lines 339–423).

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `CffHardeningTests` → 6 passed; the hash constant from Step 1 is unchanged in the final diff
- [ ] `grep -n "checked(dataStart + offsets\[count\]" src/Ofdrw.Net.Core/Fonts/CffCharStringOutlines.cs` → no matches
- [ ] Full Layout tests, Packaging `FontSubset` filter and Skia tests → `Failed: 0`
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated; report states PNG review result and "macOS Preview: not performed"

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- `ValidFontRenditionIsUnchanged` fails after any step (output for a valid font
  changed).
- An existing oracle test (`ManagedCharStringOutlinesMatchTheIndependentOracle`,
  `PlainCffFaceMatchesTheExistingSourceCodeProOracle`) fails.
- The valid Han font needs more than the default 256,000,000 operations or
  64 MiB of glyf data. Report the measured numbers instead of raising the
  defaults yourself.
- No glyph in the test font has a CharString of at least 48 bytes (Step 3).

## Maintenance notes

- `CffSubset.Create` (font embedding) also interprets glyphs through
  `CollectSubroutines`, bounded by
  `options.MaximumGlyphClosureOperations`; it was not changed here. If a future
  change makes it interpret every glyph of a font, give it a `CffWorkBudget` too.
- The default budgets are generous on purpose (a full CJK CFF font needs tens
  of millions of operations). If a legitimate font is rejected in the field, the
  diagnostic will say "exceeded its total interpretation budget"; raise the
  constant with a recorded measurement rather than removing the check.
- Reviewers: confirm no caller's `catch` filter was widened, and that the
  public `Read` path (used by SVG export and explicit-glyph drawing) still has
  no whole-font budget — it reads one glyph per request by design.
