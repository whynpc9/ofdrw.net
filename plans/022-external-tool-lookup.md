# Plan 022: External tools are never picked up from the working directory

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Core/Processes tests/Ofdrw.Net.Core.Tests`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S–M
- **Risk**: LOW–MED (the Windows half can only be verified on a Windows host or CI)
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The SDK starts two external tools by bare name: `pdftoppm` (PDF rasterizing
fallback) and `soffice` (LibreOffice, when no installation is found in the known
locations). Its own launcher resolves a bare name by looking in the **current
working directory before `PATH`**. People run `ofdrw pdf-to-ofd` or a service
inside the directory that holds the documents they received. If that directory
contains a file named `pdftoppm` (for example from an extracted archive), the
SDK executes it. This is an untrusted search path (CWE-426).

After this plan, a bare tool name is resolved only from the application's own
directory and the absolute directories listed in `PATH`. A name that contains a
directory separator (an explicit relative or absolute path chosen by the caller)
keeps working as today.

## Current state

Files:

- `src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs` — launcher with a Unix
  implementation (`UnixGroup`, starts the target through a native supervisor)
  and a Windows implementation (`WindowsGroup`, `CreateProcessW` + job object).
  722 lines of intricate, audited process-ownership code. You change two small
  spots only.
- `src/Ofdrw.Net.Core/Processes/ExternalProcessRunner.cs` — the only caller of
  `OwnedProcessGroup.Start`.
- `src/Ofdrw.Net.Converter.Pdf/Internal/PdfToPpmRasterizer.cs:36` — starts
  `new ProcessStartInfo { FileName = "pdftoppm", Arguments = args }`.
- `src/Ofdrw.Net.Converter.Docx/Internal/LibreOfficeExecutableResolver.cs:31-33`
  — returns the bare name `"soffice.com"` (Windows) or `"soffice"` when no
  known installation path exists.
- `tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs` — existing launcher tests.

Unix resolution today (`OwnedProcessGroup.cs:485-496`, inside `UnixGroup`):

```csharp
        private static string ResolveExecutable(string name)
        {
            if (Path.IsPathRooted(name)) return name;
            if (name.IndexOf(Path.DirectorySeparatorChar) >= 0) return Path.GetFullPath(name);
            foreach (var root in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() }
                .Concat((Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator)))
            {
                var candidate = Path.GetFullPath(Path.Combine(root, name));
                if (File.Exists(candidate) && access(candidate, 1) == 0) return candidate;
            }
            throw new Win32Exception(2, "Executable not found: " + name);
        }
```

It is called once (`OwnedProcessGroup.cs:197`): `var executable = ResolveExecutable(info.FileName);`

Windows launch today (`OwnedProcessGroup.cs:572-581`, inside the `WindowsGroup`
constructor). `CreateProcessW` receives a null application name and a command
line that starts with the quoted bare name, so Windows applies its default
search order, which includes the current directory before `PATH`:

```csharp
                var command = new StringBuilder(Quote(info.FileName));
                // Windows Process.Start sends Arguments verbatim to the target.
                if (!string.IsNullOrEmpty(info.Arguments)) command.Append(' ').Append(info.Arguments);
                else foreach (var argument in arguments) command.Append(' ').Append(Quote(argument));
                ...
                if (!CreateProcessW(null, command, IntPtr.Zero, IntPtr.Zero, true, 0x80000 | 0x400 | 0x4 | 0x08000000,
                    environment, string.IsNullOrEmpty(info.WorkingDirectory) ? null : info.WorkingDirectory, ref startup, out var process))
                    throw NativeError("CreateProcess " + info.FileName);
```

Facts you need:

- Callers already treat a missing tool as `System.ComponentModel.Win32Exception`:
  `PdfToPpmRasterizer.cs:39` catches it and returns null;
  `DocxToPdfConverter.cs:569` catches it and reports "LibreOffice was not
  found...". Keep throwing `Win32Exception` with native error code 2 for "not
  found".
- `Ofdrw.Net.Core` targets `netstandard2.0;netstandard2.1`. Use
  `RuntimeInformation.IsOSPlatform(OSPlatform.Windows)`, not
  `OperatingSystem.IsWindows()`.
- `Ofdrw.Net.Core` grants `InternalsVisibleTo` to `Ofdrw.Net.Core.Tests`, so
  internal types are testable from `tests/Ofdrw.Net.Core.Tests`.
- Existing tests launch either an absolute path or the bare name `dotnet`
  resolved through `PATH` (`ExternalProcessTests.cs:126`). None relies on the
  working directory.
- Repository rule for tests: never change the test process's current directory.

Code style: file-scoped namespace `Ofdrw.Net.Core.Processes`, `internal static`
helpers, short `//` comments explaining constraints. Tests: xUnit, see
`ExternalProcessTests.cs`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Core tests | `dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| PDF import tests | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` (needs `pdftoppm` and fonts; tests that lack a tool show as Skipped) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Core/Processes/ExecutableSearch.cs` (create)
- `src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs` (two spots only: the body
  of `UnixGroup.ResolveExecutable`, and the first line that builds `command` in
  the `WindowsGroup` constructor)
- `tests/Ofdrw.Net.Core.Tests/ExecutableSearchTests.cs` (create)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- Everything else in `OwnedProcessGroup.cs`: supervisor protocol, job objects,
  pipes, timeouts, the blocking waits. They were reviewed earlier and are
  deliberately left alone.
- `PdfToPpmRasterizer.cs`, `LibreOfficeExecutableResolver.cs`,
  `DocxToPdfConverter.cs` — callers keep passing bare names; the fix is central.
- Microsoft Word automation paths.

## Git workflow

- Branch: `codex/external-tool-lookup`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Add a pure, testable search helper

Create `src/Ofdrw.Net.Core/Processes/ExecutableSearch.cs`:

```csharp
namespace Ofdrw.Net.Core.Processes;

/// <summary>Resolves a bare tool name without consulting the working directory.</summary>
internal static class ExecutableSearch
{
    // Returns an absolute path, or null when the name is bare and nothing matches.
    internal static string? Find(string name, string baseDirectory, string? pathVariable,
        bool windows, Func<string, bool> isExecutableFile)
}
```

Rules, in order:

1. `Path.IsPathRooted(name)` → return `name` unchanged.
2. `name` contains `Path.DirectorySeparatorChar` or
   `Path.AltDirectorySeparatorChar` → return `Path.GetFullPath(name)` (an
   explicit relative path is the caller's choice).
3. Otherwise build the directory list: `baseDirectory` first, then each entry of
   `pathVariable` split on `Path.PathSeparator`. Skip entries that are empty,
   whitespace, or not rooted (`!Path.IsPathRooted(entry)`): an empty or relative
   `PATH` entry means "current directory" and must not be searched. On Windows
   also strip one pair of surrounding double quotes from an entry.
4. Candidate file names: on non-Windows just `name`. On Windows: `name` if
   `Path.HasExtension(name)`, otherwise `name + ".exe"` then `name + ".com"`.
5. Return the first `Path.Combine(directory, candidate)` for which
   `isExecutableFile` returns true; otherwise null. Wrap `Path.Combine` /
   `Path.GetFullPath` on a `PATH` entry in a `try`/`catch (ArgumentException)`
   that skips the malformed entry.

The helper must not call `Directory.GetCurrentDirectory()` or
`Environment.CurrentDirectory`.

**Verify**: build → `0 Error(s)`.

### Step 2: Use it on Unix

Replace the body of `UnixGroup.ResolveExecutable` with a call to the helper:

```csharp
        private static string ResolveExecutable(string name) =>
            ExecutableSearch.Find(name, AppContext.BaseDirectory, Environment.GetEnvironmentVariable("PATH"),
                windows: false, candidate => File.Exists(candidate) && access(candidate, 1) == 0)
            ?? throw new Win32Exception(2, "Executable not found: " + name);
```

**Verify**: `grep -n "GetCurrentDirectory" src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs`
→ no matches. Build → `0 Error(s)`. Core tests → `Failed: 0`.

### Step 3: Use it on Windows

In the `WindowsGroup` constructor, resolve before building the command line and
quote the resolved absolute path instead of `info.FileName`:

```csharp
                var executable = ExecutableSearch.Find(info.FileName, AppContext.BaseDirectory,
                    Environment.GetEnvironmentVariable("PATH"), windows: true, File.Exists)
                    ?? throw new Win32Exception(2, "Executable not found: " + info.FileName);
                var command = new StringBuilder(Quote(executable));
```

Leave the `CreateProcessW` call and everything after it unchanged. Place the
resolution so that a thrown `Win32Exception` goes through the constructor's
existing cleanup (`finally`) — read the surrounding `try` block and keep the
statement inside it.

**Verify**: build → `0 Error(s)`. (Behavior on Windows is verified by the unit
tests of Step 4 plus the Windows CI job; say so in your report if you are not on
Windows.)

### Step 4: Tests

Create `tests/Ofdrw.Net.Core.Tests/ExecutableSearchTests.cs` (namespace
`Ofdrw.Net.Core.Tests`). The helper takes a predicate, so no real files are
needed: pass a `HashSet<string>` of "existing" paths and
`candidate => existing.Contains(candidate)`. Build directory names with
`Path.Combine(Path.GetTempPath(), "...")` so they are rooted on every OS, and
build the `PATH` string with `Path.PathSeparator`.

Cases (each a `[Fact]`):

1. `BareName_IsFoundInPathDirectory` — tool exists in the second `PATH`
   directory → that absolute path.
2. `BareName_InWorkingDirectoryOnly_IsNotFound` — "existing" contains
   `Path.Combine(Directory.GetCurrentDirectory(), "tool")` only → result null.
3. `EmptyAndRelativePathEntries_AreSkipped` — `PATH` is
   `"" + sep + "." + sep + "relative/dir" + sep + realDir`; "existing" contains
   the tool under the current directory and under `realDir` → the `realDir`
   path is returned.
4. `BaseDirectory_IsSearchedFirst` — tool exists in base directory and in a
   `PATH` directory → base directory path.
5. `ExplicitRelativePath_IsResolvedAgainstWorkingDirectory` — name
   `"sub" + Path.DirectorySeparatorChar + "tool"` → `Path.GetFullPath` of it,
   even though the predicate returns false.
6. `RootedName_IsReturnedUnchanged`.
7. `WindowsMode_ProbesExeThenCom` — `windows: true`, name `"pdftoppm"`,
   "existing" contains only `<dir>/pdftoppm.com` → that path; when both `.exe`
   and `.com` exist → `.exe`.
8. `WindowsMode_NameWithExtension_IsNotExtended` — name `"soffice.com"`,
   "existing" contains `<dir>/soffice.com.exe` only → null.

Also add one launcher-level regression to `ExternalProcessTests.cs`:
`Runner_ShouldNotResolveBareNameFromWorkingDirectory` — create an executable
file with a unique random name (for example `"ofdrw-planted-" + Guid.NewGuid().ToString("N")`)
inside the test's temp directory `_directory`, set
`ProcessStartInfo.WorkingDirectory = _directory`, start it by bare name and
assert `Win32Exception`. This does not change the test process's working
directory; it proves the child working directory is not searched either. Use a
tiny shell script (`#!/bin/sh\nexit 0\n`, mode `UserRead|UserWrite|UserExecute`
via `File.SetUnixFileMode`) on Unix and skip visibly on Windows using a small
`FactAttribute` subclass that sets `Skip` (pattern:
`UnixSupervisorFactAttribute` in `tests/Ofdrw.Net.Core.Tests/UnixSupervisorLaunchTests.cs:135`).

**Verify**: `dotnet test tests/Ofdrw.Net.Core.Tests/Ofdrw.Net.Core.Tests.csproj -c Release --filter "FullyQualifiedName~ExecutableSearchTests|FullyQualifiedName~ExternalProcessTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`
→ `Failed: 0`; 8 `ExecutableSearchTests` pass.

### Step 5: Regression run

Run the full Core test project and the PDF converter test project (commands
above).

**Verify**: both report `Failed: 0`.

## Test plan

- New `ExecutableSearchTests.cs` with the 8 cases above; one new test in
  `ExternalProcessTests.cs`.
- Pattern: `tests/Ofdrw.Net.Core.Tests/ExternalProcessTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `grep -rn "GetCurrentDirectory\|Environment.CurrentDirectory" src/Ofdrw.Net.Core/Processes/` → no matches
- [ ] `grep -n "Quote(info.FileName)" src/Ofdrw.Net.Core/Processes/OwnedProcessGroup.cs` → no matches
- [ ] Core tests → `Failed: 0`, including 9 new tests (1 may be Skipped on Windows)
- [ ] Converter.Pdf tests → `Failed: 0`
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated; report states whether Windows behavior was exercised

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- Any existing test in `ExternalProcessTests`, `Audit024*` or
  `UnixSupervisorLaunchTests` fails after Step 2 — those protect the process
  ownership protocol.
- You find another place in `src/` that starts a process without going through
  `ExternalProcessRunner` (`grep -rn "Process.Start" src --include="*.cs"`).
  Report it; do not fix it here.
- The Windows change appears to need edits to the `CreateProcessW` call itself.

## Maintenance notes

- Behavior change for users: a tool placed only in the working directory is no
  longer found. The next release note under `docs/content/releases/` should say
  so and point to `PATH`, `OFDRW_LIBREOFFICE_PATH` or
  `DocxConversionOptions.LibreOfficePath`.
- `LibreOfficeExecutableResolver.ResolveWorkingDirectory` still returns
  `Environment.CurrentDirectory` for a bare name; that only sets the child's
  working directory and is unaffected.
- Reviewers: confirm the Windows resolution throws inside the constructor's
  `try` so handles are released, and that `.bat`/`.cmd` are intentionally not
  probed.
