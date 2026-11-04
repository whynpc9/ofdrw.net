# Plan 017: Add `ofdrw attachments list | extract | remove`

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Cli src/Ofdrw.Net.Core/Models/OfdAttachment.cs src/Ofdrw.Net.Core/Models/OfdDocumentPackage.cs docs/content/reference/cli.md`
> `Program.cs` edits from plans 005/008/015/016 are expected. Changes to the two Core model files are a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW–MED (writes user-selected output paths; removal touches attachment cleanup)
- **Depends on**: docs/engineering/plans/016-cli-delete-pages-and-crop.md (serialize `Program.cs` edits)
- **Category**: direction (requested by the maintainer)
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

OFD attachments (invoice XML, source files) are a core use case. The SDK exposes `OfdDocumentPackage.Attachments` (with `Data`) and `RemoveAttachment(attachment, Logical|Physical)`, the R18 physical cleanup. The CLI cannot list, extract or remove attachments, so users have to write C# to see what is inside an OFD or to strip an attachment before sending it.

## Current state

- `src/Ofdrw.Net.Core/Models/OfdAttachment.cs`: public `Id`, `Name`, `MediaType` (default `application/octet-stream`), `Format`, `CreationDate`, `ModificationDate`, `SizeKilobytes`, `Visible`, `Usage`, `IsExternal`, `ExternalPath`, `Data` (`byte[]`), plus nested `Attachments` and `Pages`.
- `src/Ofdrw.Net.Core/Models/OfdDocumentPackage.cs`:
  - Line 104: `public List<OfdAttachment> Attachments { get; } = [];`
  - Lines 108-130: `public bool RemoveAttachment(OfdAttachment attachment, OfdAttachmentRemovalMode mode = OfdAttachmentRemovalMode.Logical)`. Physical removal throws `NotSupportedException("Physical attachment removal requires a reader-owned internal payload in this document.")` for external or not-reader-owned attachments. An unsafe physical request makes the later **write** fail before touching the destination.
  - Lines 7-14: `OfdAttachmentRemovalMode { Logical, Physical }`. Logical retains the historical ZIP entries; Physical deletes the payload.
- The reader populates `package.Attachments` (`src/Ofdrw.Net.Reader/Readers/OfdReader.cs:328-350`).
- CLI plumbing to reuse:
  - `src/Ofdrw.Net.Cli/DocumentTools.cs:132-148`: `ReadToolInputAsync(path, token, documentIndex, allowMultiple)` reads with `OfdReader` and rejects multi-DocBody unless allowed. `WriteToolOutputAsync(package, path, token)` writes via `AtomicOutput` and prints writer diagnostics as warnings.
  - `AtomicOutput` (`src/Ofdrw.Net.Cli/AtomicOutput.cs`) stages beside the destination and commits atomically.
  - Dispatch pattern for subcommand groups: `Program.cs:44-56` (early `if (command == "...") return await XAsync(args.Skip(1).ToArray(), token)`), with each group parsing its own arguments as in `DocumentToolsAsync` (`DocumentTools.cs:11-62`).
- Test fixture to copy: `tests/Ofdrw.Net.Packaging.Tests/R18AttachmentPhysicalCleanupTests.cs:1125-1136`, `Seed(bool external)`. It builds a ZIP with `Doc_0/Metadata/List.xml` declaring an internal attachment (`ID='7' Name='original.bin'`, payload `Doc_0/Payload/original.bin`) and optionally an external one (`External='true'`). Copy it, with its `Zip`/`Utf8` helpers, into the CLI test as a private helper; the CLI test project does not reference the Packaging test project.
- CLI tests that capture Console must use `[Collection("Owned CLI console")]`.

## Command contract

```
ofdrw attachments list <input.ofd> [--document N] [--json]
ofdrw attachments extract <input.ofd> --id ID --output <file> [--document N]
ofdrw attachments remove <input.ofd> <output.ofd> --id ID [--physical] [--document N]
```
- **list**, human form: one line per top-level attachment, `ID<TAB>Name<TAB>MediaType<TAB>Bytes<TAB>internal|external`.
  - `--json`: `{"schema":"ofdrw-cli-attachments/1","attachments":[{"id":..,"name":..,"mediaType":..,"bytes":..,"external":..,"externalPath":..}]}`, camelCase, nothing else on stdout.
  - Exit 0 even when there are no attachments.
- **extract**:
  - Selects by `Id`, using exact ordinal comparison.
  - Errors (exit 1, no file written): no match; more than one match ("ambiguous"); external attachment ("external attachments have no embedded payload").
  - Writes `Data` to `--output` via `AtomicOutput`. **Never** derive the output path from the attachment `Name` (path-traversal safety). `--output` is mandatory.
- **remove**:
  - Same selection rules as extract.
  - Removal mode: `--physical` → `Physical`, otherwise `Logical`.
  - Writes the package via `WriteToolOutputAsync`, so in-place updates are safe.
  - A `NotSupportedException` from physical removal → exit 1 with its message.
  - After writing, print `Removed attachment {id} ({mode})`.
- `--document` is 1-based (reuse `ParseOneBasedPage`, as `DocumentToolsAsync` does at `DocumentTools.cs:50`).

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F --filter "FullyQualifiedName~AttachmentsCliTests"` | pass |
| CLI suite | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Cli/AttachmentsCommand.cs` (create); `src/Ofdrw.Net.Cli/Program.cs` (allowlist, early dispatch, help); `tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs` (create); `docs/content/reference/cli.md`; `docs/content/reference/core-model.md` (one CLI cross-reference line in its attachment section, if it has one).

**Out of scope**: adding attachments (`attachments add`; the writer contract for new attachments needs separate review); nested attachments (top-level only; list prints a `nested: N` count only if trivial, otherwise ignore); any Core/Packaging change.

## Git workflow

Branch `advisor/017-cli-attachments`; commit e.g. `feat(cli): list, extract and remove OFD attachments`. No push or PR unless instructed.

## Steps

### Step 1: Tests (`AttachmentsCliTests.cs`, `[Collection("Owned CLI console")]`)

Using the copied `Seed(external: true)`:
1. `List_Human`: two lines; contains `7\toriginal.bin` and `external`.
2. `List_Json`: parses; 2 entries; the internal one has `bytes == 5`.
3. `Extract_WritesPayload`: `--id 7 --output x.bin` → bytes equal `{0,1,2,3,255}`.
4. `Extract_External_Fails`: `--id 8` → exit 1 and no file is created.
5. `Extract_Unknown_Fails`: `--id 99` → exit 1.
6. `Remove_Logical`: output re-read has 1 attachment, and the ZIP still contains `Doc_0/Payload/original.bin`.
7. `Remove_Physical`: the ZIP no longer contains `Doc_0/Payload/original.bin`, and 1 attachment remains (the external one).
8. `Remove_PhysicalExternal_Fails`: `--id 8 --physical` → exit 1; a pre-existing output file is unchanged.
9. `Remove_InPlace`: input == output works, with no `.ofdrw-*.tmp` leftovers.
10. `UnknownSubcommand_Fails`: `attachments frob` → exit 1.
11. `MultiDocument_ListAndExtractSecondDocument`: start from `MultiDocumentFixture.CreateEntriesAsync()` (`tests/TestSupport/MultiDocumentFixture.cs`, already compiled into the CLI test project and used by `MultiDocumentCliTests.cs:19-20`). It builds two DocBodies, `Doc_0` and `Doc_1`.
    - Add an attachment to **`Doc_1` only**: append `<Attachments>Attachs/List.xml</Attachments>` (in the namespace of `Doc_1/Document.xml`'s root, after its `Pages` element) to `Doc_1/Document.xml`. Add `Doc_1/Attachs/List.xml`, declaring `<Attachment ID="21" Name="second.bin"><FileLoc>second.bin</FileLoc></Attachment>` in the same namespace, and the payload entry `Doc_1/Attachs/second.bin` = `{9,8,7}`. Zip it with `MultiDocumentFixture.Zip`.
    - `attachments list <file> --document 2 --json` → exit 0, exactly one attachment with id `21`.
    - `attachments list <file> --document 1 --json` → exit 0, zero attachments.
    - `attachments extract <file> --document 2 --id 21 --output x.bin` → bytes `{9,8,7}`.
    - `attachments remove <file> out.ofd --document 2 --id 21` → exit 1 (multi-DocBody write is rejected), and `out.ofd` is not created.
    - If the reader does not resolve `FileLoc` relative to the list file as assumed, adjust the fixture paths. Do not touch the reader. Confirm against `OfdReader.cs:328-350`.

**Verify**: tests fail (unknown command).

### Step 2: Implement `AttachmentsCommand.cs`

Use `internal static partial class Cli`. Parse the subcommand from `args[0]`, then flags in pairs as `DocumentToolsAsync` does (literal values; reject unknown or duplicate options).
- For the **read-only** `list` and `extract`, call `ReadToolInputAsync(path, token, documentIndex, allowMultiple: true)`. The default `allowMultiple: false` rejects any OFD with more than one DocBody, even after `--document` selects one, so `--document N` would be unusable.
- For `remove`, call `ReadToolInputAsync(path, token, documentIndex)` (single-DocBody only, the same write-safety restriction as the other page-editing tools) and `WriteToolOutputAsync`. On multi-DocBody input it fails with the helper's existing message; document that in Step 3. Use `System.Text.Json` for `--json`. Add `"attachments"` to the allowlist and the early dispatch; add help usage lines and a `Commands:` entry.

**Verify**: new tests and the CLI suite pass.

### Step 3: Docs

In `cli.md`, add the three usages to the command table and usage block, marked "main, not yet released" in the file's existing style. Note that extraction never uses the stored attachment name as a path, and explain the difference between logical and physical removal in one sentence each, consistent with the `OfdAttachmentRemovalMode` doc comments. State that `list` and `extract` accept multi-DocBody input with `--document N`, while `remove` requires a single-DocBody document.

## Done criteria

- [ ] Build exits 0; 11 new tests pass (including the multi-DocBody case); CLI suite passes
- [ ] `git diff --stat -- src/Ofdrw.Net.Core src/Ofdrw.Net.Packaging` is empty
- [ ] Docs updated; storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- The reader does not populate `Data` for internal attachments (test 3 gets empty bytes). Report it; do not modify the reader.
- Logical removal output fails to re-read, or changes page content.

## Maintenance notes

- `schema: ofdrw-cli-attachments/1` is a public contract.
- `attachments add` is the natural next verb; it needs a decision on how the writer allocates IDs and paths for new attachments.
