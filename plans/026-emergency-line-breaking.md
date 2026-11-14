# Plan 026: Long unbreakable text lays out in linear time and within its stated budgets

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs tests/Ofdrw.Net.Layout.Tests`
> If the engine file changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition. (New test files from plans 020 and 023 are expected.)

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED (line choices must stay identical)
- **Depends on**: plans/020-harfbuzz-face-cache.md (land it first; this plan's timings assume it)
- **Category**: perf, bug
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The paragraph layout engine breaks a segment that has no legal line break
(a long URL, a base64 or log line, Thai text) at grapheme boundaries when the
caller chooses `TextOverflowPolicy.BreakGraphemes`. The Flow, plain-text and
DOCX "final" text paths always choose that policy. Three defects:

1. **Each emergency line costs the whole remaining segment.** The search for
   the cut point starts by building a trial line of half the remaining segment,
   then a quarter, and so on. Measured at `427433a` (Linux x64, Release,
   `NotoSans-Regular.ttf`, width 80 mm, size 4 mm, one string of the letter `a`):
   16,000 characters → 2.2 s; 32,000 → 5.2 s. With the 25 MB CJK face, 8,000
   characters take 9.6 s.
2. **The "line budget" counts trial builds, not lines.** With default options
   (`MaximumLines` = 10,000) a 90,000-character unbreakable string — about 2,570
   lines — fails after 17 s with `Paragraph layout exceeded its line budget.`
   Measured at `427433a`.
3. **Letter spacing in right-to-left runs does a linear scan per grapheme**
   (quadratic per run).

After this plan the cut-point search looks only at a window about one line
long, `MaximumLines` limits accepted lines as documented, and the RTL lookup is
constant time. The lines produced must be exactly the same as today.

## Current state

File: `src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs` (1,222
lines). Everything below is inside the private nested class `Session`.

Constants (lines 23–27):

```csharp
    private const double Epsilon = 1e-9;
    private const int MaximumItemUtf16 = 4000;
    private const int MaximumFitAttempts = 32;
    private const int MaximumEmergencyMisses = 4;
    private const int MaximumNearMisses = 3;
```

Per-paragraph data used here (lines 128–129): `paragraph.Prefix` (measured
advance prefix sums in mm, one entry per code point index) and
`paragraph.MaximumEm` (largest font size in the paragraph, mm).
`Estimate(paragraph, start, end, indent)` returns the estimated content width of
`[start, end)` in mm, including tab handling; the main loop trusts it with a
slack of `paragraph.MaximumEm / 2`.

Where the emergency path is entered (`BreakLines`, lines 717–728):

```csharp
                if (line is null)
                {
                    var limit = firstOver >= 0 ? firstOver : fitting.Count > 0 ? fitting[0] : count;
                    ...
                    line = EmergencyBreak(paragraph, start, limit, available, first);
                }
```

`EmergencyBreak` (lines 735–768):

```csharp
        private DraftLine EmergencyBreak(Paragraph paragraph, int start, int limit, double available, bool first)
        {
            var boundaries = new List<int>();
            for (var index = start + 1; index <= limit; index++)
                if (paragraph.GraphemeStart[index]) boundaries.Add(index);
            // Largest grapheme boundary whose rebuilt (reshaped) line fits. The measured prefix is only a
            // heuristic (a ligature's whole advance sits on its first code point), and rebuilt widths are
            // not monotonic either: an Arabic lam-alef ligature makes "لا" narrower than "ل". So a binary
            // search on rebuilt widths only proposes a start; a bounded upward scan then tests later cuts
            // and stops after MaximumEmergencyMisses consecutive overflows (or the build budget).
            var low = 0;
            var high = boundaries.Count - 1;
            DraftLine? best = null;
            var bestIndex = -1;
            var builds = 0;
            while (low <= high && builds < MaximumFitAttempts)
            {
                builds++;
                var middle = (low + high) / 2;
                var draft = BuildLine(paragraph, start, boundaries[middle], first, boundaries[middle] == paragraph.Length);
                if (draft.ContentWidth <= available + 1e-7) { best = draft; bestIndex = middle; low = middle + 1; }
                else high = middle - 1;
            }
            var misses = 0;
            for (var index = bestIndex + 1; index < boundaries.Count && misses < MaximumEmergencyMisses &&
                     builds < 2 * MaximumFitAttempts; index++)
            {
                builds++;
                var draft = BuildLine(paragraph, start, boundaries[index], first, boundaries[index] == paragraph.Length);
                if (draft.ContentWidth <= available + 1e-7) { best = draft; misses = 0; }
                else misses++;
            }
            if (best is not null) return best;
            throw new TextLayoutException(TextLayoutError.LayoutOverflow,
                "A single extended grapheme cluster exceeds the line width.",
                sourceStart: Offset(paragraph, start),
                sourceLength: Offset(paragraph, boundaries.Count > 0 ? boundaries[0] : limit) - Offset(paragraph, start));
        }
```

`limit` is the end of the unbreakable segment, so `boundaries` covers the whole
remaining segment and the first trial line is half of it. `BuildLine` costs time
proportional to the length it builds.

The build counter (`BuildLine`, lines 776–779):

```csharp
            _token.ThrowIfCancellationRequested();
            if (++_lineCount > _options.MaximumLines * 4L + 64)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded its line budget.");
```

Accepted lines are already limited separately when pages are assembled (lines
1081–1083):

```csharp
                if (++lines > _options.MaximumLines)
                    throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded MaximumLines.");
```

Per accepted line the engine performs at most `MaximumNearMisses` (3) +
`MaximumFitAttempts` (32) + `2 * MaximumFitAttempts` (64) = 99 builds.

The RTL letter-spacing lookup (`ApplyLetterSpacing`, lines 853–879): a local
function scans the whole `clusters` dictionary for every grapheme:

```csharp
                var clusters = new Dictionary<int, (int End, int LastGlyph)>();
                ...
                int? ClusterOf(int codePoint)
                {
                    foreach (var pair in clusters)
                        if (codePoint >= pair.Key && codePoint < pair.Value.End) return pair.Key;
                    return null;
                }
```

`clusters` maps a cluster's first code point index (paragraph-relative) to its
end index and last glyph. All keys and ends lie inside `[run.Start, run.End]`.

Existing tests that protect the tricky cases
(`tests/Ofdrw.Net.Layout.Tests/ParagraphLayoutTests.cs`):
`OverflowIsRejectedOrBrokenOnlyAtGraphemeBoundaries` (line 128),
`EmergencyBreakRebuildsCutsInsideALigature` (line 268, covers `"ffi"` and the
Arabic `"لالالا"` case where a longer line is narrower than a shorter one),
`LetterSpacingCountsEveryGraphemeOfALigatureAndWidensRtlGaps` (line 291).
The test helper there:

```csharp
    private static ParagraphLayoutResult Layout(string text, double width = 80, TextAlignment alignment = TextAlignment.Start,
        TextParagraphDirection direction = TextParagraphDirection.Auto, double? pageHeight = null,
        TextOverflowPolicy overflow = TextOverflowPolicy.Reject, TextStyle? style = null) =>
        HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, style ?? Style(),
            new ParagraphLayoutOptions(width, pageHeight, direction, alignment, overflow: overflow)));
```

Each result line exposes `SourceStart` and `SourceLength`
(`ParagraphLayoutResult.cs:211-213`). Fonts come from
`ParagraphLayoutTests.Face(name, file)` (internal static), files under
`e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/` (Git LFS): `NotoSans-Regular.ttf`,
`DejaVuSans-TextFixture.ttf`, `LXGWWenKai-Regular.ttf`, `NotoSansThai-Regular.ttf`.

User docs state the budgets (`docs/content/reference/text-paragraph-layout.md:91`):
`MaximumLines`（10,000）, `MaximumShapeCalls`（100,000）. No doc change is
needed: after this plan `MaximumLines` means what that line says.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fonts | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Layout tests | `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| DOCX native tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~DocxNativeParagraphLayoutTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs` (only
  `EmergencyBreak`, the budget line in `BuildLine`, the `_lineCount` field name
  if you rename it, and `ApplyLetterSpacing`)
- `tests/Ofdrw.Net.Layout.Tests/EmergencyBreakTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- The normal (non-emergency) line-fitting loop in `BreakLines`, `Estimate`,
  `Measure`, `BuildLine`'s shaping logic, bidi, justification.
- The values of `MaximumFitAttempts`, `MaximumEmergencyMisses`,
  `MaximumNearMisses` and the public option defaults.
- `src/Ofdrw.Net.Text.HarfBuzz` (plan 020) and
  `src/Ofdrw.Net.Layout/Internal/Flow/FlowFinalTextLayout.cs` (plan 027).

## Git workflow

- Branch: `codex/emergency-line-breaking`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Characterize today's line choices (before any source change)

Create `tests/Ofdrw.Net.Layout.Tests/EmergencyBreakTests.cs` (namespace
`Ofdrw.Net.Layout.Tests`, `public sealed class EmergencyBreakTests`). Add a
private helper that lays text out with `TextOverflowPolicy.BreakGraphemes` and
returns the string `string.Join(",", result.Lines.Select(l => l.SourceStart + ":" + l.SourceLength))`.

Add one `[Theory]` `LineChoicesAreUnchanged(string caseName, string expected)`
whose cases are built inside the test from `caseName` (so no long literals are
needed in attributes). Cases:

| caseName | Text | Font (single-family style, size 4) | Width mm | Extra |
|---|---|---|---:|---|
| `latin` | `"a"` repeated 600 times | Noto Sans | 30 | — |
| `mixed-widths` | `"WiMl"` repeated 150 times | Noto Sans | 30 | — |
| `ligatures` | `"ffi"` repeated 120 times | DejaVu Sans | 20 | — |
| `arabic` | `"لا"` repeated 120 times | DejaVu Sans | 20 | — |
| `spaced` | `"a"` repeated 400 times | Noto Sans, `.WithLetterSpacing(0.5)` | 30 | — |
| `indent` | `"a"` repeated 400 times | Noto Sans | 30 | first-line indent 10 mm |
| `thai` | `"ภาษาไทย"` repeated 60 times | Noto Sans Thai | 30 | — |
| `after-words` | `"word word " + "a"×300 + " end"` | Noto Sans | 30 | — |

Styles are built like this (the constructors are the ones used in
`ParagraphLayoutTests.cs:15-21` and `:269`; `Face` is `internal static` there):

```csharp
    private static TextStyle One(string family, string file, double size = 4) =>
        new(new TextFontFamily(family, ParagraphLayoutTests.Face(family, file)), size);
    // One("Noto Sans", "NotoSans-Regular.ttf"), One("DejaVu Sans", "DejaVuSans-TextFixture.ttf"),
    // One("Noto Sans Thai", "NotoSansThai-Regular.ttf")
```

Cache each face in a `static readonly Lazy<TextStyle>` so the font file is read
once. If the face constructor rejects the family name for a file, use the name
the existing tests pass for that file (see `ParagraphLayoutTests.cs:10-13`).

For the expected value store the lowercase hex SHA-256 of the helper's string
(keeps the file small). Run once with placeholder constants, copy each actual
hash from the failure output into the test data, and run again until all pass.
These constants are the baseline: they must not change in later steps.

For the `indent` case construct `ParagraphLayoutOptions` with
`firstLineIndentMillimeters: 10` (see the constructor at
`ParagraphLayoutRequest.cs:360-363`).

**Verify**: `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --filter "FullyQualifiedName~EmergencyBreakTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`
→ `Passed: 8`, `Failed: 0`. Commit this step on its own before changing the engine.

### Step 2: Search a window about one line long

In `EmergencyBreak`, stop collecting `boundaries` once the estimate is clearly
past the line. Replace the first loop with:

```csharp
            var boundaries = new List<int>();
            var indent = first ? _options.FirstLineIndentMillimeters : 0;
            // The cut can only lie where the estimate is near the frame: rebuilt widths differ from
            // the estimate only by shaping changes local to the cut. Keep MaximumEmergencyMisses
            // boundaries past that point so the upward scan below sees the same overflows as before.
            var beyond = 0;
            for (var index = start + 1; index <= limit && beyond <= MaximumEmergencyMisses; index++)
            {
                if (!paragraph.GraphemeStart[index]) continue;
                boundaries.Add(index);
                if (Estimate(paragraph, start, index, indent) > available + 4 * paragraph.MaximumEm) beyond++;
            }
```

Leave the binary search, the upward scan and the final `throw` exactly as they
are. Check how `BreakLines` passes the indent to `Estimate` (it uses
`first ? _options.FirstLineIndentMillimeters : 0`) and use the same expression.

Why this keeps results identical: the binary search returns the largest fitting
boundary when "fits" is monotonic, and that boundary has an estimate within the
slack, so it is inside the window; the upward scan stops after
`MaximumEmergencyMisses` consecutive overflows, all of which are inside the
window too.

**Verify**: build → `0 Error(s)`; `EmergencyBreakTests` → `Passed: 8` with the
Step 1 constants untouched; the full Layout test project → `Failed: 0`.

### Step 3: Make `MaximumLines` count lines, not trial builds

In `BuildLine` replace the budget check with a build budget that an accepted
line can never exhaust (99 builds per line, see "Current state"):

```csharp
            if (++_lineBuilds > _options.MaximumLines * 100L + 64)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded its line-building budget.");
```

Rename the field `_lineCount` to `_lineBuilds` (it is used only here; confirm
with `grep -n "_lineCount" src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs`).
The accepted-line check at line 1082 stays unchanged.

Add tests to `EmergencyBreakTests`:

- `DefaultBudgetsAdmitALongUnbreakableString` — 90,000 × `"a"`, Noto Sans,
  width 80, default options except `overflow: TextOverflowPolicy.BreakGraphemes`.
  Assert no exception, `result.Lines.Count` is between 2,000 and 3,500, and the
  line ranges tile the input (each line starts where the previous one ended;
  the last ends at 90,000).
- `MaximumLinesStillLimitsAcceptedLines` — 2,000 × `"a"`, width 30,
  `maximumLines: 5`; assert `TextLayoutException` with
  `Error == TextLayoutError.BudgetExceeded` and message containing
  `MaximumLines`.

**Verify**: `EmergencyBreakTests` → `Passed: 10`. Before this plan the first new
test fails after about 17 s; note the measured duration after the change in your
report (expected: a few seconds or less).

### Step 4: Constant-time cluster lookup for RTL letter spacing

In `ApplyLetterSpacing`, replace the scanning `ClusterOf` with a lookup table
built once per run, right after the `clusters` dictionary is filled:

```csharp
                var clusterStart = new int[run.End - run.Start];
                for (var index = 0; index < clusterStart.Length; index++) clusterStart[index] = -1;
                foreach (var pair in clusters)
                    for (var index = Math.Max(pair.Key, run.Start); index < pair.Value.End && index < run.End; index++)
                        clusterStart[index - run.Start] = pair.Key;
                int? ClusterOf(int codePoint) =>
                    codePoint >= run.Start && codePoint < run.End && clusterStart[codePoint - run.Start] >= 0
                        ? clusterStart[codePoint - run.Start] : (int?)null;
```

`DraftRun` has `Start` and `End` (paragraph-relative code point indices; see
where `logical.Add(new DraftRun { Start = index, End = runEnd, ... })` is built
in `BuildLine`). If two clusters overlapped, the old scan returned the first in
dictionary order; clusters of one shaped run do not overlap, so the table is
equivalent.

**Verify**: the full Layout test project → `Failed: 0` (includes
`LetterSpacingCountsEveryGraphemeOfALigatureAndWidensRtlGaps`).

### Step 5: Regression and visual check

1. Run the full Layout tests and the DOCX native filter (commands above).
2. Generate and render the paragraph visual sample (it exercises letter spacing,
   Arabic, tabs and CJK):

```sh
OUT="$(mktemp -d)"
OFDRW_PARAGRAPH_ARTIFACTS="$OUT" dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~EveryF08OutputFeatureRoundTripsAndExports" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
pdftoppm -png -r 144 "$OUT/f08-visual-sample.pdf" "$OUT/f08-page"
```

Open each PNG and confirm all five sample lines are present and legible, the
letter-spaced text has even gaps, and Arabic reads right to left with no
overlapping glyphs. Record the pages viewed; if you cannot view images, say so.
macOS Preview acceptance is not possible on Linux: report "not performed".

**Verify**: both test commands `Failed: 0`; at least one `f08-page-*.png` exists.

## Test plan

- New `tests/Ofdrw.Net.Layout.Tests/EmergencyBreakTests.cs`: 8 characterization
  cases (Step 1) and 2 budget tests (Step 3).
- Pattern: `tests/Ofdrw.Net.Layout.Tests/ParagraphLayoutTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `EmergencyBreakTests` → 10 passed; the 8 hash constants are identical to those committed in Step 1 (`git diff <step-1-commit> -- tests/Ofdrw.Net.Layout.Tests/EmergencyBreakTests.cs` shows only added tests)
- [ ] Full Layout tests and DOCX native filter → `Failed: 0`
- [ ] `grep -n "MaximumLines \* 4L" src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutEngine.cs` → no matches
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated; report gives the measured time of `DefaultBudgetsAdmitALongUnbreakableString`, the PNG review result and "macOS Preview: not performed"

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- Any Step 1 constant changes after Step 2 or Step 4. Report the case name and
  both line lists. Do not update the constant and do not widen the slack to make
  it pass without reporting.
- `EmergencyBreakRebuildsCutsInsideALigature` fails.
- `DefaultBudgetsAdmitALongUnbreakableString` now fails with
  `Paragraph layout exceeded MaximumShapeCalls.` — report the shape-call count
  (`result.ShapeCalls` on a smaller input) instead of changing that default.
- The fix seems to need changes to `BuildLine`'s shaping or to `Estimate`.

## Maintenance notes

- The window relies on `Estimate` being within `4 × MaximumEm` of the rebuilt
  width at the true cut. Anything that makes rebuilt lines much narrower than
  measured (a new kind of long ligature or contextual shrinking) needs this
  slack revisited; the characterization cases will show it.
- Build cost per emergency line is now proportional to one line, but each trial
  still reshapes the cut run. If profiles later show this path hot, the next
  lever is fewer trial builds, not a wider window.
- Reviewers: look at the window loop's use of `indent`, and confirm the accepted
  line check at the page-assembly loop is still the only user-visible
  `MaximumLines` failure.
