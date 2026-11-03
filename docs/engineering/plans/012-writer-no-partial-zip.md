# Plan 012: Never leave a valid-looking partial ZIP when an OFD write fails midway

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Packaging/OfdPackageWriter.cs src/Ofdrw.Net.Packaging/OfdPackageSignatureCleaner.cs`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: MED (touches the writer's output stream; output bytes must stay identical on success)
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

`OfdPackageWriter.WriteAsync` builds all entries in memory, then streams them into `new ZipArchive(destination, Create, leaveOpen: true)` inside a `using`. If a write throws partway, the `using` still disposes the archive, and `ZipArchive.Dispose` writes a **valid central directory for the entries written so far**. A write can throw on cancellation, disk full or network failure. A caller writing directly to a file or HTTP response then holds a well-formed ZIP that is missing `Doc_0/...` parts, which downstream tools may accept as an OFD. The CLI is safe because it stages through `AtomicOutput`; API callers are not. The fix makes a failed write leave a stream that is **not** a valid ZIP, and leaves successful output byte-for-byte unchanged.

## Current state

- `src/Ofdrw.Net.Packaging/OfdPackageWriter.cs:58-65`, the public write path:
  ```csharp
  using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
  foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
  {
      var zipEntry = zip.CreateEntry(entry.Key, plan.Compression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
      using var stream = zipEntry.Open();
      await stream.WriteAsync(entry.Value, 0, entry.Value.Length, cancellationToken).ConfigureAwait(false);
  }
  return result;
  ```
- `OfdPackageWriter.cs:258-266`, the prepared/owned write path, has the same shape inside `using (var zip = ...) { foreach ... }`. Its comment at lines 268-269 acknowledges: "Failure here cannot roll back an arbitrary caller stream".
- `src/Ofdrw.Net.Packaging/OfdPackageSignatureCleaner.cs:45-51`: same shape (`using var zip = ...; foreach ... await output.WriteAsync(...)`).
- **Do not touch** `WriteZipTerminalAsync` (`OfdPackageWriter.cs:160+`). It is the controlled-edit terminal path with its own audited failure handling.
- Facts about `System.IO.Compression.ZipArchive` (Create mode): the end-of-central-directory record is written only in `Dispose`, through the underlying stream's `Write`/`Flush`. If those calls are discarded, the output lacks an EOCD, and `new ZipArchive(stream, ZipArchiveMode.Read)` throws `InvalidDataException`.

## Design

Add `internal sealed class FailClosedOutputStream : Stream` in `src/Ofdrw.Net.Packaging/FailClosedOutputStream.cs`. It is a transparent pass-through wrapper over the destination:
- It forwards `CanRead/CanSeek/CanWrite/Length/Position/Seek/SetLength/Read/Write/WriteAsync/Flush/FlushAsync` to the inner stream, so the ZIP bytes are identical.
- `internal void Poison()`: afterwards, `Write`, `WriteAsync` and `Flush*` silently do nothing. `Seek`/`Position` setters keep forwarding, so ZipArchive bookkeeping doesn't throw.
- `Dispose(bool)` does **not** dispose the inner stream (leave-open semantics).

Usage at each of the three sites:
```csharp
var guarded = new FailClosedOutputStream(destination);
var zip = new ZipArchive(guarded, ZipArchiveMode.Create, leaveOpen: true);
try
{
    foreach (...) { ... same body ... }
}
catch
{
    guarded.Poison();          // the central directory is never written
    try { zip.Dispose(); } catch { /* the primary exception wins */ }
    throw;
}
zip.Dispose();                 // success: writes the central directory normally
return result;
```

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F --filter "FullyQualifiedName~WriterFailClosedTests"` | pass |
| Suites | `dotnet test tests/Ofdrw.Net.Packaging.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Packaging/FailClosedOutputStream.cs` (create); `OfdPackageWriter.cs`, at the two sites above only; `OfdPackageSignatureCleaner.cs`, at one site; `tests/Ofdrw.Net.Packaging.Tests/WriterFailClosedTests.cs` (create); one bullet in `docs/content/reference/conversion-contracts.md` "## 资源预算与失败行为".

**Out of scope**: `WriteZipTerminalAsync`; buffering the whole ZIP in memory (rejected: it would double peak memory and bypass the writer's output reservation accounting); CLI `AtomicOutput`.

## Git workflow

Branch `advisor/012-writer-no-partial-zip`; commit e.g. `fix(packaging): fail closed when an OFD write is interrupted`. No push or PR unless instructed.

## Steps

### Step 1: Failing tests (`WriterFailClosedTests.cs`)

Test helper `ThrowAfterBytesStream : MemoryStream` that throws `IOException` from `Write`/`WriteAsync` once `Length` exceeds N bytes.
1. `Write_FailureMidway_LeavesNoReadableZip`: a package with 3 pages, each with a 200 KB `OfdImageElement`, or anything that yields several entries. With N = 50% of a successful write's size, `WriteAsync` throws `IOException`. Then `new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read)` must throw `InvalidDataException`. Before the fix it opens with fewer entries.
2. `Write_Cancelled_LeavesNoReadableZip`: a destination that cancels a `CancellationTokenSource` on its first `WriteAsync`. `WriteAsync` throws `OperationCanceledException`, and the bytes are not a readable ZIP.
3. `Write_Success_IsByteIdentical`: build the same package twice; write once through the normal API. Compute the expected bytes by writing through a plain `MemoryStream` **before** making the change, saving the SHA-256 in the test as a constant, or regenerate with the base commit (`git worktree add ../base 09a2e7e`). Assert the hash is unchanged after the change. If the writer output contains timestamps, compare entry names and contents instead of raw hashes, and say so.
4. `SignatureCleaner_FailureMidway_LeavesNoReadableZip`: same as 1, through `OfdSignatureCleaner.CleanAsync`. The CLI exemplar is in `src/Ofdrw.Net.Cli/DocumentTools.cs:39-46`.

**Verify**: tests 1, 2 and 4 fail before Step 2; test 3 passes.

### Step 2: Implement the design above at the three sites

**Verify**: all `WriterFailClosedTests` pass; all three suites pass.

### Step 3: Docs

Add a bullet to `conversion-contracts.md` "## 资源预算与失败行为": when `OfdPackageWriter`/`OfdSignatureCleaner` fail or are cancelled while writing to a caller stream, the bytes already written are deliberately not a valid ZIP. Callers must still discard the partial output. Use a temp-file-then-rename pattern for files.

## Done criteria

- [ ] Build exits 0; new tests pass; Packaging/Cli/Pdf suites pass
- [ ] `grep -n "new ZipArchive(destination, ZipArchiveMode.Create" src/Ofdrw.Net.Packaging/OfdPackageWriter.cs` shows only the `WriteZipTerminalAsync` occurrence
- [ ] Storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- Test 3 shows a byte difference on success: the wrapper is not transparent (e.g. `CanSeek` changes the ZIP layout). Report it; do not accept output changes.
- `ZipArchive.Dispose` throws on a poisoned stream in a way that masks the original exception despite the try/catch. Report the stack.
- Allocation/budget regression tests (`AllocationBudgetDecisionTests`, `Audit025RegressionTests`) fail because of the extra wrapper allocation. Report the numbers; do not edit the tests.

## Maintenance notes

Any new ZIP-writing path to a caller-owned stream should use `FailClosedOutputStream`. `WriteZipTerminalAsync` could adopt it later after review by whoever owns controlled-edit auditing.
