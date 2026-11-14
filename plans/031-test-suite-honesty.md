# Plan 031: Tests that cannot run say "skipped", and CLI tests stop sharing the console

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- tests/Ofdrw.Net.Packaging.Tests/ArchivalCopyTests.cs tests/Ofdrw.Net.Cli.Tests`
> New test files from plans 021, 028 and 030 are expected. If
> `ArchivalCopyTests.cs` or the five classes named in Step 2 changed, compare
> the "Current state" excerpts against the live code; on a mismatch, treat it
> as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: tests
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

Two small defects make the test suite report more than it proves.

1. `CaseDistinctDestinationWorksOnCaseSensitiveVolume` returns early — and is
   reported **Passed** — on case-insensitive volumes (the default on Windows and
   macOS, both of which are CI platforms). The behaviour it names is therefore
   only tested on Linux, and nothing in the results says so.
2. Five CLI test classes call the in-process CLI entry point, which writes to
   the process-wide `Console`, without joining the collection that serialises
   console users. One of them replaces `Console.Error` with a writer that counts
   `Warning:` lines and cancels on the first one; a `Warning:` line from a test
   running in parallel would be counted and would trigger the cancellation. No
   failure from this has been observed; it is a latent flake.

## Current state

### The silent pass

`tests/Ofdrw.Net.Packaging.Tests/ArchivalCopyTests.cs:450-466`:

```csharp
    [Fact]
    public async Task CaseDistinctDestinationWorksOnCaseSensitiveVolume()
    {
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "archive.ofd"); var destination = Path.Combine(folder, "Archive.ofd");
            File.WriteAllBytes(source, Zip(Seed("<Permissions><Edit>false</Edit></Permissions>")));
            if (File.Exists(destination)) return; // This volume treats the two names as one file.
            var inputHash = Hash(File.ReadAllBytes(source));
            ...
```

The repository's pattern for "this environment cannot run the test" is a small
`FactAttribute` subclass that sets `Skip`
(`tests/Ofdrw.Net.Cli.Tests/Audit004RegressionTests.cs:149-180`):

```csharp
    public sealed class NativeCaseFactAttribute : FactAttribute
    {
        public NativeCaseFactAttribute()
        {
            if (!OperatingSystem.IsWindows() && !OperatingSystem.IsMacOS())
                Skip = "Requires the actual Windows/macOS case comparison branch.";
        }
    }
```

The same file has `SymbolicLinkFactAttribute`, which probes the file system in
its constructor and cleans up in `finally`. The test projects use xUnit 2.9.3.

### The console collection

`tests/Ofdrw.Net.Cli.Tests/TextReplacementCommitNotificationTests.cs:12-13`:

```csharp
[CollectionDefinition("Owned CLI console", DisableParallelization = true)]
public sealed class OwnedCliConsoleCollection { }
```

Most CLI test classes carry `[Collection("Owned CLI console")]`. These call
`global::Cli.RunAsync(...)` without it (count of calls at `427433a`):

| Class (file of the same name) | `Cli.RunAsync` calls | Replaces `Console.*` itself |
|---|---:|---|
| `DocumentToolsCliTests` | 15 | yes (`Console.SetError`) |
| `ImageCommandTests` | 12 | no |
| `MultiDocumentCliTests` | 5 | no |
| `Audit004RegressionTests` | 3 | no |
| `TextReplacementArgumentsTests` | 1 | no |

A collection with `DisableParallelization = true` never runs alongside other
collections, but classes without any collection attribute each form their own
collection and run in parallel with one another — so these five can overlap.

The writer in `DocumentToolsCliTests.cs:53-65`:

```csharp
    private sealed class CancelOnMergeWarning(CancellationTokenSource cancellation) : StringWriter
    {
        public int Warnings { get; private set; }
        public override void WriteLine(string? value)
        {
            base.WriteLine(value);
            if (value?.StartsWith("Warning:", StringComparison.Ordinal) == true)
            {
                Warnings++;
                cancellation.Cancel();
            }
        }
    }
```

and its use (lines 39–47) asserts `error.Warnings` equals exactly 0 or 2.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Packaging tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests/Ofdrw.Net.Packaging.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Unguarded classes | `grep -L "Owned CLI console" $(grep -l "Cli.RunAsync" tests/Ofdrw.Net.Cli.Tests/*.cs)` | (empty after Step 2) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `tests/Ofdrw.Net.Packaging.Tests/ArchivalCopyTests.cs`
- In `tests/Ofdrw.Net.Cli.Tests/`: `DocumentToolsCliTests.cs`,
  `ImageCommandTests.cs`, `MultiDocumentCliTests.cs`,
  `Audit004RegressionTests.cs`, `TextReplacementArgumentsTests.cs`, and any
  newer file the "Unguarded classes" command lists (attribute line only)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- Product code under `src/`.
- Test bodies and assertions (other than the one early `return`).
- xUnit configuration files or parallelisation settings.

## Git workflow

- Branch: `codex/test-suite-honesty`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Report the case-sensitivity test as skipped where it cannot run

In `ArchivalCopyTests.cs` add a nested attribute to the test class, modelled on
`SymbolicLinkFactAttribute`:

```csharp
    public sealed class CaseSensitiveVolumeFactAttribute : FactAttribute
    {
        public CaseSensitiveVolumeFactAttribute()
        {
            var directory = Path.Combine(Path.GetTempPath(), "ofdrw-case-probe-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                File.WriteAllText(Path.Combine(directory, "probe"), "probe");
                if (File.Exists(Path.Combine(directory, "PROBE")))
                    Skip = "The temporary volume is case-insensitive; case-distinct file names are one file here.";
            }
            finally { Directory.Delete(directory, recursive: true); }
        }
    }
```

Change the test's `[Fact]` to `[CaseSensitiveVolumeFact]` and replace the early
return with an assertion that keeps the precondition explicit:

```csharp
            Assert.False(File.Exists(destination));
```

**Verify**: build → `0 Error(s)`. On Linux run the Packaging tests with
`--filter "FullyQualifiedName~CaseDistinctDestinationWorksOnCaseSensitiveVolume"`
→ `Passed: 1`, `Skipped: 0`. (On Windows/macOS CI the same test must now show
as skipped with the reason above; you cannot check that locally — say so in the
report.)

### Step 2: Put every in-process CLI test class in the console collection

Add `[Collection("Owned CLI console")]` on the line above the class
declaration of each of the five classes in the table, exactly as
`PageEditCliTests.cs:7-8` does:

```csharp
[Collection("Owned CLI console")]
public sealed class PageEditCliTests : IDisposable
```

Then run the "Unguarded classes" command and add the attribute to any further
class it lists.

**Verify**: the "Unguarded classes" command prints nothing; build →
`0 Error(s)`; the full CLI test project → `Failed: 0`. Record the wall-clock
duration reported by `dotnet test` before and after (serialising five classes
makes the run somewhat longer; report both numbers).

### Step 3: Full check

Run the Packaging and CLI test projects in full.

**Verify**: both `Failed: 0`; the total test counts are unchanged from before
the plan (no test added or removed).

## Test plan

No new tests. One existing test changes its reporting on case-insensitive
volumes; five classes change their collection.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `grep -n "return; // This volume" tests/Ofdrw.Net.Packaging.Tests/ArchivalCopyTests.cs` → no matches
- [ ] "Unguarded classes" command prints nothing
- [ ] Packaging and CLI test projects → `Failed: 0`, test counts unchanged
- [ ] `git status --short` lists only in-scope files; storage gate exits 0
- [ ] `plans/plan-index.md` status row updated; report states the CLI test duration before/after and that the skip path was not exercised locally

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- A CLI test starts failing after joining the collection (it may depend on
  running in parallel, or on console state another test leaves behind). Report
  the test and its output.
- The CLI test run becomes more than twice as long. Report the numbers and
  leave the decision to the maintainer.

## Maintenance notes

- New test classes that call `global::Cli.RunAsync` need the collection
  attribute; the "Unguarded classes" command is a quick review check.
- A search for other tests that return early instead of skipping found only
  this one (`grep -rn "if (.*) return; //" tests --include="*.cs"` at
  `427433a`); other early-return shapes were not searched exhaustively.
