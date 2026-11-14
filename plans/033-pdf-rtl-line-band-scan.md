# Plan 033: The PDF right-to-left import checks neighbouring lines in linear time

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Converter.Pdf/Internal/PdfRtlVisualLineImport.cs tests/Ofdrw.Net.Converter.Pdf.Tests/PdfRtlVisualLineImportTests.cs`
> If either file changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: perf
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

PDF → OFD has an opt-in mode for right-to-left text
(`PdfToOfdOptions.RightToLeftTextImport = VisualLineProfile`; the default is
`Reject`). It groups the page's glyphs into baseline "clusters" (one per text
line) and, for every cluster containing right-to-left text, compares it against
**every** other cluster on the page to make sure no other line intrudes into its
band.

That is (right-to-left lines) × (all lines) comparisons per page. The number of
accepted right-to-left lines is capped by `MaxTextObjectsPerPage` (default
50,000), but the number of clusters is not capped at this point, so a crafted
page with very many tiny lines costs billions of comparisons before any limit
applies. Ordinary documents (tens of lines per page) are unaffected. This was
derived from the code; it was **not measured** during planning.

After this plan the same decision is made in constant time per line from two
precomputed arrays.

## Current state

`src/Ofdrw.Net.Converter.Pdf/Internal/PdfRtlVisualLineImport.cs`
(`internal static class PdfRtlVisualLineImport`, 391 lines). Constants (lines
29–30):

```csharp
    private const double BaselineTolerance = 0.05;
    private const double NeighbourLineDistance = 0.6;
```

Clusters are built top to bottom (lines 108–130). `clusterY[k]` is the baseline
Y of cluster `k`'s first glyph and `clusterEm[k]` its largest glyph size:

```csharp
        // Baseline clusters, top to bottom (PDF Y grows upwards).
        var sorted = horizontal.OrderByDescending(item => item.Letter.StartBaseLine.Y).ThenBy(item => item.Order).ToList();
        var clusters = new List<List<(Letter Letter, int Order)>>();
        var clusterY = new List<double>();
        var clusterEm = new List<double>();
        foreach (var item in sorted)
        {
            ...
            if (last >= 0 && Math.Abs(clusterY[last] - y) <= BaselineTolerance * Math.Max(clusterEm[last], em))
            {
                clusters[last].Add(item);
                clusterEm[last] = Math.Max(clusterEm[last], em);
            }
            else
            {
                clusters.Add(new List<(Letter, int)> { item });
                clusterY.Add(y);
                clusterEm.Add(em);
            }
        }
```

Because the glyphs are sorted by descending Y, `clusterY` is non-increasing in
the cluster index. `Em(...)` (line 383) always returns a finite value greater
than zero.

The scan to replace (lines 136–147):

```csharp
        for (var index = 0; index < clusters.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!clusters[index].Any(item => PdfImportReadingOrder.ContainsRightToLeft(item.Letter.Value ?? string.Empty)))
                continue;
            for (var other = 0; other < clusters.Count; other++)
            {
                if (other == index) continue;
                if (Math.Abs(clusterY[other] - clusterY[index]) < NeighbourLineDistance * Math.Max(clusterEm[other], clusterEm[index]))
                    throw PdfImportReadingOrder.ProfileRejection(number,
                        "right-to-left text shares its line band with glyphs on a different baseline (superscript, overlap or ruby).");
            }
            // Only directly painted, unclipped glyphs are geometry evidence.
            painted ??= PdfPaintedTextEvidence.Scan(document, page, cancellationToken);
```

The order of effects in this loop matters and must be preserved: lines above
the current one have already been imported when its band check runs.

Why two arrays are enough. Write `N` for `NeighbourLineDistance`. Cluster `i`
conflicts with cluster `j` when `|Y[j] − Y[i]| < N × max(Em[j], Em[i])`, that is
when `j` is inside `i`'s own band **or** `i` is inside `j`'s band:

- *Own band*: the closest clusters to `i` are `i − 1` and `i + 1` (Y is
  sorted). If neither is inside `i`'s band, no cluster is.
- *Someone else's band, from above* (`j < i`, so `Y[j] ≥ Y[i]`):
  `Y[j] − N × Em[j] < Y[i]`. It holds for some `j < i` exactly when the
  **minimum** of `Y[j] − N × Em[j]` over all `j < i` is below `Y[i]` — a running
  minimum.
- *From below* (`j > i`): `Y[j] + N × Em[j] > Y[i]` for some `j > i` exactly when
  the running **maximum** from the end exceeds `Y[i]`.

The rearranged comparisons can differ from the original only when a distance
equals its threshold to within floating-point rounding.

Existing tests: `tests/Ofdrw.Net.Converter.Pdf.Tests/PdfRtlVisualLineImportTests.cs`
— `EachGuardRejectsItsOwnUnverifiedEvidence` has the row
`[InlineData("superscript", "different baseline")]` that exercises this check,
and `ProfileImportsDeclaredLogicalTextFromIndependentGenerators` (8 fixture
rows) proves accepted pages still import the same text. The test project can
call `internal` members of the converter (the tests call
`PdfRtlVisualLineImport.Import` directly).

Design background: `docs/architecture/f06-pdf-rtl-import.md`. No documentation
change is needed; behaviour is unchanged.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| RTL tests | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~PdfRtl" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| All PDF converter tests | same command without `--filter` | `Failed: 0` |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Converter.Pdf/Internal/PdfRtlVisualLineImport.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/PdfRtlLineBandTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- The clustering loop, `ImportLine`, the painted-evidence checks, the constants,
  the rejection message text.
- `PdfToOfdOptions` limits and defaults.

## Git workflow

- Branch: `codex/pdf-rtl-line-band-scan`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Add the band index

In `PdfRtlVisualLineImport.cs` add a nested type:

```csharp
    /// <summary>Answers "does another baseline cluster intrude into this cluster's line band, or
    /// this cluster into another's" in constant time. Y must be non-increasing (top to bottom).</summary>
    internal sealed class LineBands
    {
        private readonly IReadOnlyList<double> _y, _em;
        private readonly double[] _lowestReachAbove, _highestReachBelow;

        internal LineBands(IReadOnlyList<double> y, IReadOnlyList<double> em)
        {
            _y = y; _em = em;
            _lowestReachAbove = new double[y.Count];
            _highestReachBelow = new double[y.Count];
            var lowest = double.PositiveInfinity;
            for (var index = 0; index < y.Count; index++)
            {
                _lowestReachAbove[index] = lowest;
                var reach = y[index] - NeighbourLineDistance * em[index];
                if (reach < lowest) lowest = reach;
            }
            var highest = double.NegativeInfinity;
            for (var index = y.Count - 1; index >= 0; index--)
            {
                _highestReachBelow[index] = highest;
                var reach = y[index] + NeighbourLineDistance * em[index];
                if (reach > highest) highest = reach;
            }
        }

        internal bool Conflicts(int index)
        {
            var own = NeighbourLineDistance * _em[index];
            if (index > 0 && (Math.Abs(_y[index - 1] - _y[index]) < own || _lowestReachAbove[index] < _y[index])) return true;
            return index + 1 < _y.Count &&
                   (Math.Abs(_y[index + 1] - _y[index]) < own || _highestReachBelow[index] > _y[index]);
        }
    }
```

Use `if (reach < lowest)` / `if (reach > highest)` exactly as written (not
`Math.Min`/`Math.Max`), so a NaN coordinate is ignored instead of poisoning the
running value.

**Verify**: build → `0 Error(s)`.

### Step 2: Prove equivalence with the old scan, then scale

Create `tests/Ofdrw.Net.Converter.Pdf.Tests/PdfRtlLineBandTests.cs` (namespace
`Ofdrw.Net.Converter.Pdf.Tests`, `using Ofdrw.Net.Converter.Pdf.Internal;`).

- `MatchesTheExhaustiveScan`: with `new Random(20261113)`, run 2,000 trials. Each
  trial: pick a cluster count 1–40; generate Y values by starting at 800 and
  subtracting a random step in (0, 30) per cluster (so Y is strictly
  decreasing); pick each Em in (0.5, 50). For every index compare
  `new PdfRtlVisualLineImport.LineBands(y, em).Conflicts(index)` with the
  reference written in the test, copied from the old loop:

```csharp
    private static bool Reference(IReadOnlyList<double> y, IReadOnlyList<double> em, int index)
    {
        for (var other = 0; other < y.Count; other++)
        {
            if (other == index) continue;
            if (Math.Abs(y[other] - y[index]) < 0.6 * Math.Max(em[other], em[index])) return true;
        }
        return false;
    }
```

  On a mismatch fail with a message that prints the trial's Y and Em arrays and
  the index. Also assert that across all trials both `true` and `false` results
  occurred at least 1,000 times each (so the test is not vacuous).
- `HandCases` (`[Theory]` or several asserts): single cluster → false; two
  clusters 100 apart with Em 10 → false for both; two clusters 5 apart with Em
  10 → true for both; three clusters at Y 300, 200, 100 with Em 10, 10, 400 →
  true for index 0 and 1 (the large bottom cluster's band reaches them) and
  true for index 2; same with Em 10, 10, 10 → false for all.
- `ScalesLinearly`: 200,000 clusters, Y = 1,000,000 − index × 2, Em = 1; build
  `LineBands` and call `Conflicts` for every index inside a
  `System.Diagnostics.Stopwatch`; assert all results are false and elapsed time
  is under 5 seconds (the exhaustive scan would need 4 × 10^10 comparisons).

**Verify**: RTL tests command → `PdfRtlLineBandTests` all pass. Note the
measured time of `ScalesLinearly` in your report.

### Step 3: Use it in `Import`

Declare `LineBands? bands = null;` next to `PdfPaintedTextEvidence? painted = null;`
and replace the inner `for (var other ...)` loop with:

```csharp
            bands ??= new LineBands(clusterY, clusterEm);
            if (bands.Conflicts(index))
                throw PdfImportReadingOrder.ProfileRejection(number,
                    "right-to-left text shares its line band with glyphs on a different baseline (superscript, overlap or ruby).");
```

Everything else in the loop stays in the same order.

**Verify**: RTL tests command → `Failed: 0` (includes the `superscript` row and
the 8 fixture rows). Then all PDF converter tests → `Failed: 0`.

## Test plan

- New `tests/Ofdrw.Net.Converter.Pdf.Tests/PdfRtlLineBandTests.cs`: seeded
  equivalence test against the exhaustive scan, hand cases, and a scale test.
- Existing `PdfRtlVisualLineImportTests` unchanged and passing.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `PdfRtlLineBandTests` pass; all PDF converter tests → `Failed: 0`
- [ ] `grep -n "for (var other = 0" src/Ofdrw.Net.Converter.Pdf/Internal/PdfRtlVisualLineImport.cs` → no matches
- [ ] `git status --short` lists only in-scope files; storage gate exits 0
- [ ] `plans/plan-index.md` status row updated; report states the `ScalesLinearly` time and that no visual review was performed because accepted pages produce identical output (only the rejection decision's cost changes)

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- `MatchesTheExhaustiveScan` reports a mismatch. Include the printed arrays in
  your report; do not change the seed, the ranges or the reference to make it
  pass.
- Any existing RTL test changes its result.

## Maintenance notes

- `LineBands` requires Y sorted top to bottom. If the clustering order ever
  changes, this class must change with it; `MatchesTheExhaustiveScan` only
  generates sorted input.
- The total cluster count per page is still unbounded before
  `MaxTextObjectsPerPage` is applied to the whole page. With this plan every
  remaining step is linear in the number of glyphs, so no new limit was added.
