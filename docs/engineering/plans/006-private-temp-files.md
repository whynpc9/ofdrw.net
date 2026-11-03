# Plan 006: Stage document temp files in owner-only directories

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Core/IO src/Ofdrw.Net.Layout/Editing/Internal/OfdTextReplacementFileSession.cs src/Ofdrw.Net.Packaging/OfdAccountedTemporaryFile.cs src/Ofdrw.Net.Converter.Pdf/Converters/PdfToOfdConverter.cs src/Ofdrw.Net.Converter.Pdf/Internal/PdfToPpmRasterizer.cs src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageIoBudget.cs src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdConverter.cs src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs src/Ofdrw.Net.Converter.Docx/Converters/DocxToOfdConverter.cs`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW
- **Depends on**: none
- **Category**: security
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Several operations copy the caller's **plaintext document**, or rendered page images of it, into the shared system temp directory using default file permissions. On Linux, `/tmp` is shared and default mode is typically 0644/0755, so other local users can read the documents while the operation runs. The raster path also uses a fixed, predictable parent directory (`/tmp/ofdrw-net-raster/`) that another user can create first. The repo already does this correctly in two places (owner-only `UnixCreateMode` in the password-crypto package, `mkdtemp` for the process supervisor). This plan applies the same protection everywhere through one shared helper.

## Current state

Sites that create temp files or directories (verified with `grep -rn "GetTempPath()" src`):

| File:line | What is staged |
|---|---|
| `src/Ofdrw.Net.Layout/Editing/Internal/OfdTextReplacementFileSession.cs:113-114` | full input OFD snapshot: `snapshot = Path.Combine(Path.GetTempPath(), "ofdrw-input-" + Guid.NewGuid().ToString("N") + ".tmp"); stream = new FileStream(snapshot, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);`. Line 102 also charges `8L * Path.GetTempPath().Length` to a budget; keep a similar charge for the new path length. |
| `src/Ofdrw.Net.Packaging/OfdAccountedTemporaryFile.cs:45-49` | replacement output spool, same pattern (`"ofdrw-replacement-"`). Has `Fail()` / dispose cleanup that deletes `path`. |
| `src/Ofdrw.Net.Converter.Pdf/Converters/PdfToOfdConverter.cs:67-80` | input PDF: `var tempPdfPath = Path.Combine(Path.GetTempPath(), $"ofdrw-net-{Guid.NewGuid():N}.pdf");` then `File.Create`; `finally { try { File.Delete(tempPdfPath); } catch (IOException) { } catch (UnauthorizedAccessException) { } }` |
| `src/Ofdrw.Net.Converter.Pdf/Internal/PdfToPpmRasterizer.cs:19-27` | rendered PNGs: `var tempDir = Path.Combine(Path.GetTempPath(), "ofdrw-net-raster", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(tempDir);`. Arguments built by interpolation: `$"-f {pageNumber} -l {pageNumber} -r {...} -png -singlefile " + (cropBox ? "-cropbox " : "") + $"\"{pdfPath}\" \"{outputPrefix}\""`. Cleanup in `finally` at lines 58-70 deletes `tempDir` recursively. |
| `src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageIoBudget.cs:66-67` | image staging stream (`"ofdrw-image-"`) |
| `src/Ofdrw.Net.Converter.Pdf.Vector/PdfVectorToOfdConverter.cs:38-41` | input PDF (`"ofdrw-vector-"`) |
| `src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:113-122` | staging directory with `input.docx` |
| `src/Ofdrw.Net.Converter.Docx/Converters/DocxToPdfConverter.cs:403-409` | `CreateWorkDirectory`: LibreOffice work dir under `Path.GetTempPath()`. The Microsoft Word engine uses `GetMicrosoftWordTempRoot()` (sandbox container); **leave the Word root unchanged**. |
| `src/Ofdrw.Net.Converter.Docx/Converters/DocxToOfdConverter.cs:92-93` | DualLayer staging directory |

Exemplar of the safe pattern, `src/Ofdrw.Net.Core/Processes/NativeSupervisorAssets.cs:39-41, 78`:
```csharp
var template = Encoding.UTF8.GetBytes(Path.Combine(Path.GetTempPath(), "ofdrw-supervisor-XXXXXX") + "\0");
if (mkdtemp(template) == IntPtr.Zero) throw Failure("Cannot create a private owned-process supervisor directory.");
var directory = Encoding.UTF8.GetString(template, 0, Array.IndexOf(template, (byte)0));
...
[DllImport("libc", SetLastError = true)] private static extern IntPtr mkdtemp([In, Out] byte[] template);
```
`mkdtemp` creates the directory with mode 0700 atomically and fails if the name exists. On Windows, `Path.GetTempPath()` is the per-user `%TEMP%`, which is already private to the user, so a GUID-named subdirectory is enough there.

Facts:
- `Ofdrw.Net.Core` grants `InternalsVisibleTo` to Packaging, Layout, Reader, Converter.Docx, Converter.Pdf, Converter.Pdf.Vector and the test projects (`src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj:8-21`). A helper in `src/Ofdrw.Net.Core/IO/` is reachable from every site.
- Libraries target netstandard2.0/2.1. `UnixCreateMode` and `ProcessStartInfo.ArgumentList` are **not** available, so use `mkdtemp` and a quoting helper.
- Quoting helper to reuse: `src/Ofdrw.Net.Converter.Docx/Internal/ProcessArguments.cs`, `internal static string Join(params string[] arguments)`, used at `DocxToPdfConverter.cs:533`. It lives in Converter.Docx, which Converter.Pdf does not reference, so **move it** to `src/Ofdrw.Net.Core/Processes/ProcessArguments.cs` (namespace `Ofdrw.Net.Core.Processes`) and update the Docx usage.
- No test depends on the temp-file name patterns (`grep -rn "ofdrw-replacement\|ofdrw-net-raster\|ofdrw-vector-" tests` is empty).

## Commands you will need

Once per shell:
```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| Helper tests | `dotnet test tests/Ofdrw.Net.Core.Tests -c Release $F --filter "FullyQualifiedName~PrivateTemporaryDirectoryTests"` | pass |
| Suites | `dotnet test Ofdrw.Net.sln -c Release --no-build $F` | pass (record tool-dependent failures and show they also fail on base) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**:
- `src/Ofdrw.Net.Core/IO/PrivateTemporaryDirectory.cs` (create)
- `src/Ofdrw.Net.Core/Processes/ProcessArguments.cs` (moved from Converter.Docx; delete the old file)
- The files in the table above, at their temp-creation and cleanup lines only
- `tests/Ofdrw.Net.Core.Tests/PrivateTemporaryDirectoryTests.cs` (create)

**Out of scope**:
- `NativeSupervisorAssets.cs`: already safe. You may switch it to the helper only if trivial; leaving it alone is preferred.
- `src/Ofdrw.Net.Crypto.Password`: already uses `UnixCreateMode`.
- CLI `AtomicOutput.cs`: writes beside the destination by design.
- The Microsoft Word temp root.
- Any test file that creates its own temp dirs.

## Git workflow

Branch `advisor/006-private-temp-files`; commit e.g. `fix(security): stage temporary document copies in owner-only directories`. No push or PR unless instructed.

## Steps

### Step 1: Helper and tests

Create `PrivateTemporaryDirectory` (namespace `Ofdrw.Net.Core.IO`, `internal static class`):
```csharp
/// Creates a new directory under the system temp path that only the current user can access.
internal static string Create(string prefix)
// Unix (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) == false): mkdtemp(prefix + "XXXXXX"); throw IOException with Win32Exception(Marshal.GetLastWin32Error()) on failure.
// Windows: Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), prefix + Guid.NewGuid().ToString("N"))).
internal static void TryDelete(string? directory)
// Recursive delete; swallow IOException and UnauthorizedAccessException only (same policy as PdfToOfdConverter's finally).
```
`prefix` must not contain path separators; throw `ArgumentException` otherwise.

Tests in `PrivateTemporaryDirectoryTests.cs`:
- Create → directory exists. On non-Windows, `File.GetUnixFileMode(dir)` (net10 test project) equals `UnixFileMode.UserRead | UserWrite | UserExecute`.
- Two creates return different paths.
- `TryDelete` removes a directory containing a file; `TryDelete(null)` and deleting a missing path do not throw.
- A prefix with `/` throws `ArgumentException`.

**Verify**: helper tests pass.

### Step 2: Move `ProcessArguments` to Core

`git mv src/Ofdrw.Net.Converter.Docx/Internal/ProcessArguments.cs src/Ofdrw.Net.Core/Processes/ProcessArguments.cs`, fix the namespace, and add `using Ofdrw.Net.Core.Processes;` where Docx uses it.

**Verify**: build exits 0; `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F` passes.

### Step 3: Convert each site

For each row in the table:
- **Single-file sites** (snapshot, spool, staged PDF, image stage, vector PDF): `var dir = PrivateTemporaryDirectory.Create("ofdrw-<same-name>-");` then put the file at `Path.Combine(dir, "<fixed name>")`, e.g. `input.ofd.tmp`, `spool.tmp`, `input.pdf`. Extend the existing cleanup so it also deletes `dir`, using `PrivateTemporaryDirectory.TryDelete(dir)` after the file delete, on every path where the file was deleted (including `Fail()` / dispose).
- `OfdTextReplacementFileSession`: keep a budget charge of `8L * <new full path>.Length` in place of `8L * Path.GetTempPath().Length`, at the same point, before the allocation.
- **Directory sites** (Docx staging, DualLayer staging, LibreOffice work dir, raster): replace `Path.Combine(Path.GetTempPath(), "x-" + guid)` + `Directory.CreateDirectory` with `PrivateTemporaryDirectory.Create("x-")`. In `CreateWorkDirectory`, only the non-Word branch changes. If callers later call `Directory.CreateDirectory` on the returned path, that is harmless; leave it.
- **Raster**: remove the shared `ofdrw-net-raster` parent entirely (`PrivateTemporaryDirectory.Create("ofdrw-net-raster-")`). Build the arguments with `ProcessArguments.Join("-f", n, "-l", n, "-r", dpi, "-png", "-singlefile", [optional "-cropbox"], pdfPath, outputPrefix)`. Do **not** add `--`: pdftoppm does not document it, and both paths are absolute.

**Verify**: build exits 0; `grep -rn "GetTempPath()" src --include=*.cs` matches only `PrivateTemporaryDirectory.cs`, `NativeSupervisorAssets.cs` and the Word-root branch, or explain any other remaining match.

### Step 4: Regression

Run the full test suite. Tests that need LibreOffice or Poppler must still pass where those tools are installed (`which soffice pdftoppm`). Specifically run:
- `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F --filter "FullyQualifiedName~PdfToOfd"`
- `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~TextReplacement"`

**Verify**: pass. After the run, `ls /tmp | grep -c "^ofdrw-"` is 0, or no more than before the run: there are no leaked dirs.

### Step 5: Conversion sanity (AGENTS.md)

The rendering paths are touched (raster temp dir and argument quoting), so run one PDF→OFD and one DOCX→OFD (default and `--docx-ofd-mode native`) conversion via the CLI into `artifacts/visual/006/`. Then OFD→PDF → `pdftoppm -r 144 -png`. Confirm the PNG pages match the base-commit output (`sha256sum`), or are visually identical. Use `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` as the DOCX input.

**Verify**: record hashes and observations.

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/006/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/006/trusted-font-tests
rm -rf artifacts/006/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/006/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/006/packages --output-dir artifacts/006/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/006/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] Build exits 0; helper tests pass; full suite passes (environment-only failures listed)
- [ ] The `grep -rn "GetTempPath()" src` result is as described in Step 3
- [ ] `test ! -e src/Ofdrw.Net.Converter.Docx/Internal/ProcessArguments.cs`
- [ ] No leaked temp dirs after the test run
- [ ] Conversion sanity recorded (Preview not done unless performed)
- [ ] Storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- `mkdtemp` is unavailable on a target in CI (e.g. a musl or other unusual runtime) — report rather than adding fallbacks.
- LibreOffice fails to write into a 0700 work directory. That is unexpected, since it runs as the same user; report the error output.
- Any audited budget/allocation test (`AllocationBudgetDecisionTests`, `Audit025RegressionTests`, `TextReplacementBudgetTests`) fails because the charged path length changed. Report the expected vs actual numbers; do not edit the tests.

## Maintenance notes

- New code that stages user content must use `PrivateTemporaryDirectory.Create`. Reviewers should grep new diffs for `GetTempPath()`.
