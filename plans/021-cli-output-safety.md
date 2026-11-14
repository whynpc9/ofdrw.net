# Plan 021: CLI output files never destroy the input or widen its permissions

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Cli tests/Ofdrw.Net.Cli.Tests docs/content/reference/cli.md`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug, security
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

Two defects in how the `ofdrw` CLI publishes output files:

1. `ofdrw attachments extract doc.ofd --id 7 --output doc.ofd` and
   `ofdrw preflight doc.ofd --json --output doc.ofd` replace the source document
   with the attachment bytes or the JSON report and exit 0. Every conversion
   command already rejects "output path equals input path"; these two commands
   were added later without the guard.
2. Every CLI write goes through `AtomicOutput`, which writes a new temporary
   file and renames it over the destination. The new file gets the process
   default mode (usually `0644` on Linux/macOS), so an in-place edit such as
   `ofdrw reorder secret.ofd secret.ofd --pages 2,1` turns a `0600` document
   into a world-readable one. In-place editing is a documented feature.

After this plan, the two commands refuse an output path that is the input, and
replacing an existing file on Unix keeps that file's permission bits.

## Current state

Files:

- `src/Ofdrw.Net.Cli/AtomicOutput.cs` — stages output beside the destination and
  renames on commit. Used by 14 call sites (every file-writing command).
- `src/Ofdrw.Net.Cli/ConversionPathIdentity.cs` — `RequireDifferentPaths(input, output)`;
  resolves symlinks and compares case-insensitively on Windows/macOS. Throws
  `ArgumentException("Conversion input and output must use different paths.")`.
- `src/Ofdrw.Net.Cli/AttachmentsCommand.cs` — `attachments list|extract|remove`.
- `src/Ofdrw.Net.Cli/PreflightCommand.cs` — `preflight [--json] [--output FILE]`.
- `tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`,
  `tests/Ofdrw.Net.Cli.Tests/PreflightCliTests.cs`,
  `tests/Ofdrw.Net.Cli.Tests/PageEditCliTests.cs` — existing CLI tests.
- `docs/content/reference/cli.md` — user-facing CLI reference (Chinese).

`AtomicOutput.cs` in full (39 lines):

```csharp
/// <summary>Stages output beside its destination and publishes only a completed file.</summary>
internal sealed class AtomicOutput : IAsyncDisposable
{
    private readonly string _destination;
    private readonly string _temporaryPath;
    private bool _committed;

    internal AtomicOutput(string destination)
    {
        _destination = Path.GetFullPath(destination);
        var directory = Path.GetDirectoryName(_destination)!;
        Directory.CreateDirectory(directory);
        _temporaryPath = Path.Combine(directory, $".ofdrw-{Guid.NewGuid():N}.tmp");
        Stream = new FileStream(_temporaryPath, FileMode.CreateNew, FileAccess.Write,
            FileShare.None, 81920, FileOptions.Asynchronous);
    }

    internal FileStream Stream { get; }

    internal async Task CommitAsync(CancellationToken cancellationToken)
    {
        await Stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        await Stream.DisposeAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        // Both paths are on the same filesystem. A failed conversion never opens
        // the destination for writing, and replacement exposes a complete file.
        File.Move(_temporaryPath, _destination, overwrite: true);
        _committed = true;
    }
    ...
```

`AttachmentsCommand.cs:62-74` (the extract branch; `input` is the validated
input path from line 48):

```csharp
        if (subcommand == "extract")
        {
            // Never derive the destination from the stored attachment name.
            var outputPath = values.TryGetValue("--output", out var output) && output.Length > 0
                ? output : throw new ArgumentException("--output is required for attachments extract.");
            var package = await ReadToolInputAsync(input, token, documentIndex, allowMultiple: true).ConfigureAwait(false);
            var attachment = SelectAttachment(package, id);
            if (attachment.IsExternal) throw new InvalidOperationException("External attachments have no embedded payload.");
            await using var file = new AtomicOutput(outputPath);
```

`PreflightCommand.cs:60-62` (after argument parsing):

```csharp
        if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("An input OFD path is required.");
        if (outputPath is not null && !json) throw new ArgumentException("--output requires --json.");
        if (!File.Exists(input)) throw new FileNotFoundException("Input file does not exist.", input);
```

How the existing guard is used (`Program.cs:452-453`, called at `Program.cs:190`,
`283`, `303`):

```csharp
    private static void EnsureDifferentPaths(string inputPath, string outputPath)
        => ConversionPathIdentity.RequireDifferentPaths(inputPath, outputPath);
```

Facts you need:

- The CLI targets `net10.0`, so `File.GetUnixFileMode`, `File.SetUnixFileMode`
  and `FileStreamOptions.UnixCreateMode` are available. They throw
  `PlatformNotSupportedException` on Windows: guard with
  `!OperatingSystem.IsWindows()`.
- The repo already uses this pattern:
  `src/Ofdrw.Net.Crypto.Password/OfdPasswordEnvelope.cs:481` —
  `if (!OperatingSystem.IsWindows()) fileOptions.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;`
- Any `ArgumentException` thrown by a command is printed as `Error: <message>`
  on stderr and becomes exit code 1 (top-level handler in `Program.cs`).
- In-place editing must keep working: `reorder`, `delete-pages`, `crop`,
  `watermark`, `split`, `clean-signatures`, `attachments remove` are tested with
  input == output (see `PageEditCliTests.DeletePages_InPlace`). Do NOT add the
  path-identity guard to them.

Test conventions (`tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`):

- Classes carry `[Collection("Owned CLI console")]`, create a temp directory in
  the constructor and delete it in `Dispose`.
- Commands run in-process through a helper that captures the console:
  `var (code, stdout, stderr) = await Run("attachments", "extract", await Seeded(), "--id", "7", "--output", output);`
  `Seeded()` returns the path of a small OFD with attachment ID `7`
  (5 payload bytes) and an external attachment ID `8`.
- Platform-specific tests use a small custom attribute that sets `Skip` so the
  result shows as Skipped, never as a silent pass. Exemplar:
  `SymbolicLinkFactAttribute` in `tests/Ofdrw.Net.Cli.Tests/Audit004RegressionTests.cs:149`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Cli/AtomicOutput.cs`
- `src/Ofdrw.Net.Cli/AttachmentsCommand.cs`
- `src/Ofdrw.Net.Cli/PreflightCommand.cs`
- `tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`
- `tests/Ofdrw.Net.Cli.Tests/PreflightCliTests.cs`
- `tests/Ofdrw.Net.Cli.Tests/OutputPermissionTests.cs` (create)
- `docs/content/reference/cli.md`
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- The in-place edit commands listed above — they must keep accepting
  input == output.
- `ConversionPathIdentity.cs` — reuse it unchanged.
- File ownership (`chown`) and Windows ACL preservation — see Maintenance notes.
- The mode of files that did not exist before (keep today's default).
- `src/Ofdrw.Net.Cli/Program.cs` option parsing — that is plan 028.

## Git workflow

- Branch: `codex/cli-output-safety`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Reject input == output in `attachments extract`

In `AttachmentsCommand.cs`, immediately after `outputPath` is computed in the
extract branch (before `ReadToolInputAsync`), add:

```csharp
            ConversionPathIdentity.RequireDifferentPaths(input, outputPath);
```

**Verify**: build → `0 Error(s)`.

### Step 2: Reject input == output in `preflight --output`

In `PreflightCommand.cs`, after the `File.Exists(input)` check at line 62, add:

```csharp
        if (outputPath is not null) ConversionPathIdentity.RequireDifferentPaths(input, outputPath);
```

**Verify**: build → `0 Error(s)`.

### Step 3: Keep the permission bits of a replaced file (Unix only)

In `AtomicOutput.cs`:

1. In the constructor, before creating the stream, capture the existing mode:
   if `!OperatingSystem.IsWindows()` and `File.Exists(_destination)`, store
   `File.GetUnixFileMode(_destination)` in a private `UnixFileMode?` field.
   Wrap the read in a `try`/`catch (IOException)` and
   `catch (UnauthorizedAccessException)` that leaves the field null (fall back
   to today's behavior).
2. When the field has a value, create the temporary file owner-only so the
   content is never readable more widely than the final file while it is being
   written: build the stream with `FileStreamOptions` (`Mode = FileMode.CreateNew`,
   `Access = FileAccess.Write`, `Share = FileShare.None`, `BufferSize = 81920`,
   `Options = FileOptions.Asynchronous`,
   `UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite`). When the
   field is null, keep the existing constructor call exactly as it is.
3. In `CommitAsync`, after the stream is disposed and before `File.Move`, when
   the field has a value call `File.SetUnixFileMode(_temporaryPath, mode.Value)`.
4. Update the class comment to mention that replacing an existing file keeps its
   Unix permission bits.

**Verify**: build → `0 Error(s)`; full CLI tests → `Failed: 0` (no existing test
may change behavior).

### Step 4: Tests

Add to `AttachmentsCliTests.cs`:

- `Extract_OntoInput_FailsAndKeepsInput` — `var input = await Seeded();` record
  `File.ReadAllBytes(input)`; run
  `attachments extract <input> --id 7 --output <input>`; assert exit code 1,
  stderr contains `different paths`, the input bytes are unchanged, and no
  `.ofdrw-*.tmp` file remains in the directory.

Add to `PreflightCliTests.cs` (use its existing `Valid()` and `Run` helpers):

- `Output_OntoInput_FailsAndKeepsInput` — run
  `preflight <input> --json --output <input>`; assert exit code 1 and input
  bytes unchanged.

Create `tests/Ofdrw.Net.Cli.Tests/OutputPermissionTests.cs`
(`[Collection("Owned CLI console")]`, temp directory, same `Run` helper shape
as `AttachmentsCliTests`). Add a private `UnixFactAttribute : FactAttribute`
whose constructor sets `Skip = "Unix permission bits are not available on Windows."`
when `OperatingSystem.IsWindows()`. Build the input with the `Seeded()` content
from `AttachmentsCliTests` (copy the small entry dictionary and ZIP helper; do
not make the original helper public). Tests:

- `InPlaceEdit_KeepsOwnerOnlyMode` (`[UnixFact]`) — set the seeded file to
  `UserRead | UserWrite` with `File.SetUnixFileMode`; run
  `attachments remove <file> <file> --id 7`; assert exit 0 and
  `File.GetUnixFileMode(file) == (UnixFileMode.UserRead | UnixFileMode.UserWrite)`.
- `Replace_KeepsExistingWiderMode` (`[UnixFact]`) — pre-create the output file
  with mode `UserRead | UserWrite | GroupRead`; run `attachments extract ... --output <that file>`;
  assert the mode is unchanged and the content is the 5 payload bytes.
- `NewOutput_IsCreatedAndReadable` (`[Fact]`) — extract to a path that does not
  exist; assert exit 0 and the file can be read (guards the fallback path).

**Verify**: `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --filter "FullyQualifiedName~OutputPermissionTests|FullyQualifiedName~AttachmentsCliTests|FullyQualifiedName~PreflightCliTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`
→ `Failed: 0`, and the 5 new tests appear as Passed on Linux/macOS.

### Step 5: Document the behavior

In `docs/content/reference/cli.md`, find the paragraph that starts
`文件型命令先验证参数，再在目标目录写临时文件` (near line 111). Append two
sentences in Chinese, marked as not yet released (the repository's rule is that
unreleased `main` behavior is labelled as such):

- `attachments extract` 与 `preflight --output` 的输出路径不得与输入相同。
- 在 Linux/macOS 上替换已存在的文件时保留其权限位；新建文件沿用进程默认权限。
- Add the marker `（main，尚未发布）` to these sentences.

**Verify**: `grep -n "尚未发布" docs/content/reference/cli.md` → at least one
match. Then run the docs check command. If `npm ci` cannot reach the registry,
record "docs site checks not run: no network" in your report instead of
claiming they passed.

## Test plan

- 1 new test in `AttachmentsCliTests.cs`, 1 in `PreflightCliTests.cs`, 3 in the
  new `OutputPermissionTests.cs` (listed in Step 4).
- Pattern: `tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`.
- Verification: full CLI test project → `Failed: 0`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] Full `Ofdrw.Net.Cli.Tests` run → `Failed: 0`, 5 new tests present
- [ ] `grep -c "RequireDifferentPaths" src/Ofdrw.Net.Cli/AttachmentsCommand.cs src/Ofdrw.Net.Cli/PreflightCommand.cs` → `1` for each file
- [ ] `grep -n "SetUnixFileMode" src/Ofdrw.Net.Cli/AtomicOutput.cs` → one match
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- Any existing in-place edit test (`PageEditCliTests`, `MultiDocumentCliTests`,
  `DocumentToolsCliTests`) fails after Step 3.
- `File.SetUnixFileMode` fails on the CI filesystem for a reason other than
  Windows (report the exception text).
- You find that fixing this requires changing a command other than
  `attachments extract` and `preflight`.

## Maintenance notes

- Not covered: when a different user (for example root) replaces a file, the new
  file is owned by that user; .NET has no managed `chown`. Not covered on
  Windows: the replaced file's explicit ACL is lost because the new file
  inherits the directory ACL. `File.Replace` would preserve it; that is a
  follow-up needing a Windows host to verify.
- Any new CLI command that writes a file whose path the user supplies should
  call `ConversionPathIdentity.RequireDifferentPaths` unless in-place editing is
  an intended feature of that command.
- Reviewers: check that the temporary file is created owner-only only when a
  mode is being preserved, so brand-new outputs keep their current default mode.
