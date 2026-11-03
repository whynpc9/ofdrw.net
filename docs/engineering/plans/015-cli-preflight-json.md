# Plan 015: Add `ofdrw preflight` with human and `--json` output

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Cli src/Ofdrw.Net.Packaging/Preflight docs/content/reference/cli.md docs/content/guides/archival-preflight.md`
> Plans 005 and 008 also edit `src/Ofdrw.Net.Cli/Program.cs`. If they landed, the command list and help text will differ slightly from the excerpts; that is expected. Any other drift is a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW (additive command; JSON becomes a public contract)
- **Depends on**: docs/engineering/plans/008-verify-signatures-exit-codes.md (serialize `Program.cs` edits)
- **Category**: direction (requested by the maintainer)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

`ArchivalPreflightChecker` (six local archival rules, bounded budgets, a rich report) is SDK-only. Its guide says so explicitly: "The library introduces no serialization or network dependency and no CLI." Archive and records users are exactly the people who check OFD files in batch from shell pipelines and CI. A `preflight` command with stable JSON output and meaningful exit codes makes the feature usable without writing C#, on top of an API that already takes a stream and enforces its own budgets.

## Current state

- SDK API, `src/Ofdrw.Net.Packaging/Preflight/ArchivalPreflightChecker.cs:25`:
  ```csharp
  public async Task<ArchivalPreflightReport> CheckAsync(Stream input, ArchivalPreflightOptions options, CancellationToken cancellationToken = default)
  ```
  The seekable input must be at position 0. The method never writes to it.
- Options (`ArchivalPreflightModels.cs:32-51`): `RequireSingleDocument = true`, `RequireNoEncryption = true`, `LoadOptions`, plus XML budgets.
- Report (`ArchivalPreflightModels.cs:131-180`): `ProfileVersion`, `RunState` (`Complete|Incomplete`), `Scope`, `Certification`, `RequireSingleDocument`, `RequireNoEncryption`, `BudgetDescription`, `EntryCount`, `DocumentCount`, `RuleResults` (`PreflightRuleResult`: `ResultId`, `RuleId`, `Outcome`, `ReasonCode`, `Evidence`, …), `CoverageGaps`, `SummaryOutcome`, and the method `GetRuleOutcome(ruleId)`. Outcomes: `PreflightOutcome { Pass, Fail, NotApplicable, Unknown }`.
- Guide: `docs/content/guides/archival-preflight.md`. It documents rules APF-001..006, that reports "can be serialized by the caller's serializer (for example System.Text.Json with JsonStringEnumConverter)", and that "a zero-object result never proves success".
- CLI structure (`src/Ofdrw.Net.Cli/Program.cs`):
  - Lines 33-35: allowlist `if (command is not ("convert" or ... or "replace-text"))` → "Unknown command".
  - Lines 44-56: early dispatch for commands with their own argument parsers, e.g. `if (command == "replace-text") return await ReplaceTextAsync(args.Skip(1).ToArray(), cancellationToken)...`. Follow this pattern with a dedicated parser in a new partial-class file.
  - Lines 141-150: errors → `Console.Error.WriteLine($"Error: {ex.Message}"); return 1;`; cancellation → `130`.
  - `PrintHelp()` at line 672+ has usage lines and a `Commands:` section.
  - Commands live in partial-class files such as `src/Ofdrw.Net.Cli/DocumentTools.cs` (`internal static partial class Cli { private static async Task<int> DocumentToolsAsync(string command, string[] args, CancellationToken token) ... }`).
  - Output files go through `AtomicOutput` (`src/Ofdrw.Net.Cli/AtomicOutput.cs`: `await using var output = new AtomicOutput(path); ... await output.CommitAsync(token);`).
- The CLI project targets net10.0 and is not trimmed or AOT, so `System.Text.Json` reflection serialization is available.
- CLI tests (`tests/Ofdrw.Net.Cli.Tests`): call `await global::Cli.RunAsync([...])` directly. Tests that redirect `Console` must carry `[Collection("Owned CLI console")]` (defined in `TextReplacementCommitNotificationTests.cs:12`). Exemplar of stdout capture: `NativeEmbeddedBoldArgumentsTests.cs:80-90` (`var saved = Console.Out; Console.SetOut(output); try {...} finally { Console.SetOut(saved); }`). Temp-dir pattern: `DocumentToolsCliTests.cs:12-15`.
- The preflight SDK tests are in `tests/Ofdrw.Net.Packaging.Tests/ArchivalPreflightTests.cs`; look there for how to build passing and failing packages (e.g. a missing FontFile → APF-003 Fail).

## Command contract (implement exactly)

```
ofdrw preflight <input.ofd> [--json] [--output <report.json>] [--allow-multiple-documents] [--allow-encryption]
```
- `--allow-multiple-documents` → `RequireSingleDocument = false`; `--allow-encryption` → `RequireNoEncryption = false`. All other options keep SDK defaults.
- `--output` requires `--json` and writes the JSON via `AtomicOutput`. With `--json` and no `--output`, JSON goes to stdout and **nothing else** goes to stdout.
- Human output (no `--json`), to stdout:
  - `Preflight <SummaryOutcome> (run <RunState>, profile <ProfileVersion>)`
  - then one line per rule ID in ascending order: `APF-00N <aggregate outcome>`
  - then one line per result whose outcome is not `Pass`: `  <RuleId> <Outcome> <ReasonCode> <Evidence.Entry> <Evidence.Selector>`
  - then `Coverage gaps: N`
- JSON (UTF-8, indented, camelCase property names, enums as strings via `JsonStringEnumConverter`):
  ```json
  { "schema": "ofdrw-cli-preflight/1",
    "input": { "fileName": "<file name only>", "bytes": 1234, "sha256": "<lowercase hex>" },
    "ruleOutcomes": { "APF-001": "Pass", ... },
    "report": { ...System.Text.Json serialization of ArchivalPreflightReport... } }
  ```
  `ruleOutcomes` is built from the distinct `RuleId`s in `RuleResults` via `report.GetRuleOutcome`. Use the file **name**, not the full path: the report must not leak directory structure.
- Exit codes (this command): `0` = `RunState == Complete` and `SummaryOutcome` is `Pass` or `NotApplicable`; `2` = `SummaryOutcome == Fail`; `3` = `Unknown`, or `RunState == Incomplete`; `1` = argument/IO error; `130` = cancelled.
- Input: open with `File.OpenRead` (seekable, position 0). If the file is missing, throw `FileNotFoundException` ("Input file does not exist.", as other commands do).

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F --filter "FullyQualifiedName~PreflightCliTests"` | pass |
| CLI suite | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Manual | `dotnet run --project src/Ofdrw.Net.Cli -c Release -- preflight some.ofd --json` | JSON on stdout |
| Docs site | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | exit 0 (or recorded as not run if Node is unavailable) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Cli/PreflightCommand.cs` (create, `internal static partial class Cli`); `src/Ofdrw.Net.Cli/Program.cs` (allowlist entry, early dispatch line, help text); `tests/Ofdrw.Net.Cli.Tests/PreflightCliTests.cs` (create); `docs/content/reference/cli.md`; `docs/content/guides/archival-preflight.md`.

**Out of scope**: any change in `src/Ofdrw.Net.Packaging/Preflight/` (the SDK contract is fixed); `ArchivalRuleService` / `ArchivalCopyConverter` commands (a possible follow-up, `archive-copy`); exposing XML budget options as flags.

## Git workflow

Branch `advisor/015-cli-preflight`; commit e.g. `feat(cli): add preflight command with JSON report`. No push or PR unless instructed.

## Steps

### Step 1: Tests first (`PreflightCliTests.cs`, `[Collection("Owned CLI console")]`)

Build fixtures with the same approach as `ArchivalPreflightTests`: a minimal valid package via `OfdPackageWriter` (expect Pass or NotApplicable), and a failing one (expect Fail).
1. `Preflight_Human_PassExitsZero`: exit 0; stdout starts with `Preflight `.
2. `Preflight_Json_IsParseableAndStable`: with `--json`, stdout parses with `JsonDocument.Parse`. Assert:
   - `schema == "ofdrw-cli-preflight/1"`
   - `input.sha256` equals the file's SHA-256 and `input.fileName` equals the name only
   - `report.ruleResults` is an array
   - each `ruleOutcomes` value is one of the four enum names
   - the stdout text contains no directory separator from the temp path (no path leak)
3. `Preflight_Fail_ExitsTwo`: the failing fixture → exit 2, and `ruleOutcomes` has a `Fail`.
4. `Preflight_Output_WritesAtomically`: with `--json --output r.json`, the file exists, stdout is empty, and there are no `.ofdrw-*.tmp` leftovers.
5. `Preflight_OutputWithoutJson_IsError`: exit 1.
6. `Preflight_MissingInput_IsError`: exit 1.
7. `Preflight_Garbage_IsIncomplete`: a non-ZIP input file → exit 3 (the SDK returns `RunState=Incomplete` rather than throwing, per the guide). If the SDK throws instead, assert exit 1 and note it.

**Verify**: tests fail (unknown command).

### Step 2: Implement `PreflightCommand.cs`

Contents: an argument parser for the contract, plus `CheckAsync`, the human/JSON renderers and the exit-code mapping. Unknown options throw `ArgumentException("Unknown preflight option: ...")`. Compute SHA-256 with `System.Security.Cryptography.SHA256.HashData(stream)`, then rewind before calling `CheckAsync` (it requires position 0). Add `"preflight"` to the allowlist and the early dispatch in `Program.cs`. Add the usage line and `Commands:` entry to `PrintHelp`: `preflight    Run local archival preflight rules APF-001..006 (not certification).`

**Verify**: new tests pass; the CLI suite passes.

### Step 3: Docs

- `docs/content/reference/cli.md`: add the command to the command table, the usage block, and an exit-code sub-table for `preflight` (0/2/3/1/130), with the same "main, not yet released" marking the file uses for unreleased features.
- `docs/content/guides/archival-preflight.md`: replace "and no CLI" with a short sentence that the CLI `ofdrw preflight` (main, unreleased) wraps the same API and emits `ofdrw-cli-preflight/1` JSON. Add a 3-line usage example. Keep every disclaimer about non-certification.

**Verify**: docs commands pass, or are recorded as not run.

## Done criteria

- [ ] Build exits 0; `PreflightCliTests` (7) pass; CLI suite passes
- [ ] `git diff --stat -- src/Ofdrw.Net.Packaging` is empty
- [ ] `dotnet run --project src/Ofdrw.Net.Cli -c Release -- --help` lists `preflight`
- [ ] Docs updated; storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- `System.Text.Json` cannot serialize `ArchivalPreflightReport` (e.g. cycles or unsupported types). Report it; do not add attributes to SDK types.
- `CheckAsync` throws for inputs the guide says produce `Incomplete`. Report it, with the input description.

## Maintenance notes

- `schema: ofdrw-cli-preflight/1` is now a public contract. Additive fields are fine; renames or removals need `/2`.
- Because the embedded `report` object mirrors SDK property names, renaming a public preflight property is now also a CLI JSON break. Note this in the preflight models' review checklist.
- Plans 016–018 add more commands to the same files; keep help output grouped.
