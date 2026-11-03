# Plan 016: Add `ofdrw delete-pages` and `ofdrw crop` CLI commands

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Cli src/Ofdrw.Net.Layout/Editing/OfdDocumentEditor.cs docs/content/reference/cli.md docs/content/guides/document-tools.md`
> Edits to `Program.cs` from plans 005/008/015 are expected. Any change to `OfdDocumentEditor.cs` is a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S–M
- **Risk**: LOW (additive commands)
- **Depends on**: plans/015-cli-preflight-json.md (serialize `Program.cs` edits)
- **Category**: direction (requested by the maintainer)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

The SDK has `OfdDocumentEditor.RemovePages` and `CropPage`, and `docs/content/guides/document-tools.md` promises document operations "via CLI or SDK". But the CLI has no way to delete pages or crop. Ops users who only install the `ofdrw` dotnet tool cannot do the most common edit, removing a page, without writing C#. Both commands are thin wrappers over existing, validated SDK calls, using the CLI's existing read → edit → atomic-write plumbing.

## Current state

- `src/Ofdrw.Net.Layout/Editing/OfdDocumentEditor.cs`:
  - Lines 55-95: `public static void RemovePages(OfdDocumentPackage package, IEnumerable<int> pageIndexes)`. Zero-based. Throws `ArgumentOutOfRangeException` for out-of-range indexes, and validates annotations, navigation and the source envelope. It does **not** reject removing every page.
  - Lines 97-125: `public static void CropPage(OfdPage page, double xMillimeters, double yMillimeters, double widthMillimeters, double heightMillimeters)`. Sets the page box; width and height must be positive; content coordinates are unchanged.
- The closest CLI exemplar is `reorder`. `src/Ofdrw.Net.Cli/Program.cs:100-110` dispatches:
  ```csharp
  if (command == "reorder")
  {
      var pageOrder = ParsePages(options.Pages)
          ?? throw new ArgumentException("--pages is required for reorder.");
      await ReorderAsync(options.InputPath, options.OutputPath, pageOrder, options.DocumentIndex ?? 0, cancellationToken).ConfigureAwait(false);
      WriteSuccessNotification($"Reordered {options.InputPath} -> {options.OutputPath}");
      return 0;
  }
  ```
  `ReorderAsync` (lines 356-378):
  ```csharp
  if (!File.Exists(inputPath)) throw new FileNotFoundException("Input file does not exist.", inputPath);
  Ofdrw.Net.Core.Models.OfdDocumentPackage package;
  await using (var input = File.OpenRead(inputPath))
  {
      package = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = documentIndex }, cancellationToken).ConfigureAwait(false);
  }
  OfdDocumentEditor.ReorderPages(package, pageOrder);
  await using var output = new AtomicOutput(outputPath);
  var result = await new OfdPackageWriter().WriteWithResultAsync(package, output.Stream, cancellationToken).ConfigureAwait(false);
  foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
  await output.CommitAsync(cancellationToken).ConfigureAwait(false);
  ```
  The input is closed before the atomic commit, so in-place updates are supported, as for `reorder` per `docs/content/reference/cli.md`: "页重排支持原位更新".
- `ParsePages` (`Program.cs:564+`) parses 1-based `1,3-5` into zero-based indexes, preserving order and duplicates.
- `ParseOptions` (`Program.cs:474-549`) supports `--input/-i`, `--output/-o`, `--pages/-p`, `--document` and positionals. Unknown `-x` options throw `Unknown option`. A new `--box` option must be added there, as a `string? Box` in `CliOptions` (`Program.cs:750-772`).
- Command allowlist: `Program.cs:33-35`. Help: `PrintHelp()` line 672+.
- Tests: `tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs` (temp dir via `PathFor`, packages written with `OfdPackageWriter`, read back with `OfdReader`). The reorder tests are the closest pattern; `grep -rn "\"reorder\"" tests/Ofdrw.Net.Cli.Tests`.

## Command contract

```
ofdrw delete-pages <input.ofd> <output.ofd> --pages 2,4-5 [--document N]
ofdrw crop <input.ofd> <output.ofd> --box X,Y,W,H [--pages 1,3] [--document N]
```
- `delete-pages`:
  - `--pages` is required (1-based).
  - Duplicates are de-duplicated.
  - Deleting **every** page → `ArgumentException("delete-pages cannot remove every page.")` → exit 1, and the output is not written.
  - Success message: `Deleted N page(s) {input} -> {output}`.
- `crop`:
  - `--box` is required: four invariant-culture numbers in millimetres, separated by commas (`double.Parse(..., NumberStyles.Float, CultureInfo.InvariantCulture)`). Reject non-finite values and anything other than 4 parts.
  - `--pages` defaults to all pages.
  - Every selected page gets the same box.
  - Success message: `Cropped N page(s) {input} -> {output}`.
- Both support in-place updates (input == output), like `reorder`, and use `AtomicOutput`.
- `--document` is 1-based like other commands; check how `options.DocumentIndex` is populated (zero-based after parsing) and match `reorder`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F --filter "FullyQualifiedName~PageEditCliTests"` | pass |
| CLI suite | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Visual | `dotnet run --project src/Ofdrw.Net.Cli -c Release -- ofd-to-pdf -i X.ofd -o X.pdf && pdftoppm -r 144 -png X.pdf X` | PNGs |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Cli/PageEditCommands.cs` (create, partial `Cli`); `src/Ofdrw.Net.Cli/Program.cs` (allowlist, dispatch, `--box` option, help); `tests/Ofdrw.Net.Cli.Tests/PageEditCliTests.cs` (create); `docs/content/reference/cli.md`; `docs/content/guides/document-tools.md`.

**Out of scope**: `OfdDocumentEditor` (no SDK changes); multi-DocBody editing beyond the `--document` selection that `reorder` already supports; crop semantics beyond `CropPage` (no content clipping).

## Git workflow

Branch `advisor/016-cli-delete-pages-crop`; commit e.g. `feat(cli): add delete-pages and crop commands`. No push or PR unless instructed.

## Steps

### Step 1: Tests (`PageEditCliTests.cs`)

Fixture: 4 pages, each with one `OfdTextElement` `"P1"`..`"P4"`.
1. `DeletePages_RemovesSelected`: `--pages 2,4` → output has 2 pages, texts `P1`, `P3`; exit 0.
2. `DeletePages_InPlace`: input == output → succeeds, and there are no `.ofdrw-*.tmp` leftovers.
3. `DeletePages_All_Fails`: `--pages 1-4` → exit 1; a pre-existing output file still contains `"unchanged"`.
4. `DeletePages_OutOfRange_Fails`: `--pages 9` → exit 1.
5. `DeletePages_RequiresPages`: exit 1.
6. `Crop_SetsBox`: `--box 10,20,80,60 --pages 1` → page 1 has X=10, Y=20, W=80, H=60 (±0.001), and the other pages are unchanged.
7. `Crop_AllPagesByDefault`.
8. `Crop_InvalidBox`: `--box 1,2,3`, `--box 1,2,0,5`, `--box 1,2,NaN,5` → exit 1 each.

**Verify**: tests fail (unknown command).

### Step 2: Implement

Put the code in `PageEditCommands.cs`. Reuse the `ReorderAsync` read/write shape exactly. Add `"delete-pages" or "crop"` to the allowlist. Dispatch them after `ParseOptions`, next to `reorder`, because they share its option set. Add `--box` to `ParseOptions`/`CliOptions`. Reject `--box` on other commands with `ArgumentException("--box applies only to crop.")`. Add help usage lines and `Commands:` entries.

**Verify**: new tests and the CLI suite pass.

### Step 3: Visual check of crop (AGENTS.md: page output changes)

Create a 1-page OFD with text in each corner (a `FlowDocument` or a builder), and crop it to the central region with the CLI. Then `ofd-to-pdf` → `pdftoppm -r 144` into `artifacts/visual/016/`. Open the PNG: it should show only the central region, with the page size equal to the box. If the CropPage semantics produce something else (e.g. the content offset is not applied), **STOP** and report with the PNG: the SDK semantics would need clarification, which is not this plan's job.

**Verify**: PNG reviewed; observations recorded; macOS Preview reported as not done unless performed.

### Step 4: Docs

- `cli.md`: two rows in the command table and usage examples, marked "main, not yet released" in the file's existing style.
- `document-tools.md`: add the CLI equivalents next to the SDK `RemovePages`/`CropPage` description, if it is described; otherwise add a short "CLI" subsection.

## Done criteria

- [ ] Build exits 0; 8 new tests pass; CLI suite passes
- [ ] `git diff --stat -- src/Ofdrw.Net.Layout` is empty
- [ ] Crop PNG reviewed and recorded
- [ ] Docs updated; storage gate exits 0; `plans/README.md` updated

## STOP conditions

- `RemovePages` rejects common documents because of annotation or navigation validation errors that the CLI cannot explain. Report an example.
- The crop visual result contradicts the SDK doc comment ("Sets a page physical crop box").

## Maintenance notes

Plans 017/018 add more commands; keep `PrintHelp` ordering consistent (document-tool commands together).
