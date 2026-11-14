# Plan 035: `ofdrw pdf-to-ofd` can copy bookmarks and attachments; design note for outline export to PDF

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Cli/Program.cs src/Ofdrw.Net.Converter.Pdf/PdfToOfdOptions.cs docs/content/reference/cli.md docs/content/guides/pdf-import-navigation-attachments.md`
> Plans 028 and 030 change `Program.cs` and `cli.md` before this one, so line
> numbers below will have moved; locate code by the quoted text. If a quoted
> statement no longer exists in any form, treat it as a STOP condition.

## Status

- **Priority**: P3 (feature)
- **Effort**: S (Part A) + S (Part B, a written design note only)
- **Risk**: LOW
- **Depends on**: plans/028-cli-argument-strictness.md, plans/030-cli-terminal-output-escaping.md (same files; land them first)
- **Category**: direction
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The SDK can copy a PDF's outline (bookmarks) and embedded files into the OFD it
produces (`PdfToOfdOptions.CopyBookmarks` / `CopyAttachments`, documented in
`docs/content/guides/pdf-import-navigation-attachments.md`). The CLI cannot: it
always constructs the converter with default options, so anyone converting from
the command line loses bookmarks and attachments with no way to ask for them.
Part A adds two flags.

In the other direction, OFD → PDF drops the OFD outline entirely: the package
model has `Outlines`, the PDF exporter never reads it. That is a larger feature
with real mapping decisions, so Part B only produces a short design note for the
maintainer; it changes no product code.

Part A's premise was checked during planning: a minimal nine-object PDF with
one outline item and one embedded file, converted by `PdfToOfdConverter` at
`427433a`, gave `Outlines = null`, no attachments with the options off, and one
outline `Chapter A` plus attachment `note.txt` (`HELLO`) with them on.

## Current state

### CLI conversion path

`src/Ofdrw.Net.Cli/Program.cs`, `ConvertAsync` (about line 216 at `427433a`):

```csharp
        else if (mode == ConversionMode.PdfToOfd)
        {
            var converter = new PdfToOfdConverter();
            var result = await converter.ConvertWithResultAsync(input, output.Stream, pages, cancellationToken).ConfigureAwait(false);
            foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
        }
```

(After plan 030 the warning line wraps the diagnostic in `TerminalText.Escape`.)

`ConvertAsync` receives its settings as separate parameters from `RunAsync`:

```csharp
            await ConvertAsync(
                mode,
                options.InputPath,
                options.OutputPath,
                pages,
                options.DocxEngine,
                options.DocxOfdMode,
                options.NativeEmbeddedBoldBehavior,
                options.LibreOfficePath,
                options.FontDirectories, options.DocumentIndex ?? 0, cancellationToken).ConfigureAwait(false);
```

Flags are parsed in `ParseOptions` as `switch` cases, e.g.:

```csharp
                case "--include-templates":
                    options.IncludeTemplates = true;
                    break;
```

into `private sealed class CliOptions` (bool properties such as
`public bool IncludeTemplates { get; set; }`). Options that do not apply to the
chosen command are rejected in `RunAsync` right after parsing, in this style:

```csharp
            if (options.Box is not null && command != "crop")
                throw new ArgumentException("--box applies only to crop.");
```

(plan 028 adds more checks of the same kind, including a helper that rejects
DOCX-only options after `var mode = ResolveMode(...)`).

Help text is one raw string in `PrintHelp()`; the relevant lines:

```
          ofdrw pdf-to-ofd --input <input.pdf> --output <output.ofd> [--pages 1,3-5]
```

and an `Options:` list with one line per option.

### SDK options

`src/Ofdrw.Net.Converter.Pdf/PdfToOfdOptions.cs:79-83`:

```csharp
    /// <summary>Copy local PDF outline destinations into OFD Outlines. Off by default for compatibility.</summary>
    public bool CopyBookmarks { get; set; }

    /// <summary>Copy embedded document files and page FileAttachment annotations. Off by default for compatibility.</summary>
    public bool CopyAttachments { get; set; }
```

Unsupported outline targets make the conversion fail before output is written
(see the guide, section "书签映射"); the CLI already reports such failures as
`Error: ...` with exit code 1.

### OFD outline model (for Part B)

`src/Ofdrw.Net.Core/Models/OfdNavigation.cs`: `OfdDocumentPackage.Outlines` is
an `OfdOutlineCollection` (`Items`), each `OfdOutline` has `Title`, `Expanded`,
`Count`, `Actions` and `Children`; an `OfdGoToAction` carries either a
`Destination` (`Page`, `Type` ∈ `XYZ, Fit, FitH, FitV, FitR`, optional `Left`,
`Top`, `Right`, `Bottom`, zoom) or a named `Bookmark`. The PDF exporter
(`src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`) writes PDF with
PdfSharpCore 1.3.67 and contains no reference to `Outlines`
(`grep -n "\.Outlines" src/Ofdrw.Net.Converter.Pdf/Converters/OfdToPdfConverter.cs`
shows only glyph-outline code). The reverse mapping rules already implemented
for import are in `src/Ofdrw.Net.Converter.Pdf/Internal/PdfImportMetadataMigrator.cs`.

### Tests and docs

CLI tests: `tests/Ofdrw.Net.Cli.Tests/` — pattern
`tests/Ofdrw.Net.Cli.Tests/PageEditCliTests.cs` (class attribute
`[Collection("Owned CLI console")]`, temp directory, `Run` helper calling
`global::Cli.RunAsync(args)` with the console captured). Reading a result:
`await new OfdReader().ReadAsync(stream)` → `package.Outlines`,
`package.Attachments` (`Name`, `Data`).

A PDF fixture that was verified during planning (plain ASCII, built in code):

```csharp
    private static byte[] PdfWithOutlineAndAttachment()
    {
        var objects = new[]
        {
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /Names << /EmbeddedFiles 7 0 R >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\nendstream",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "<< /Title (Chapter A) /Parent 5 0 R /Dest [3 0 R /Fit] >>",
            "<< /Names [(note.txt) 8 0 R] >>",
            "<< /Type /Filespec /F (note.txt) /EF << /F 9 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 5 >>\nstream\nHELLO\nendstream",
        };
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var index = 0; index < objects.Length; index++)
        {
            offsets.Add(Encoding.ASCII.GetByteCount(pdf.ToString()));
            pdf.Append($"{index + 1} 0 obj\n{objects[index]}\nendobj\n");
        }
        var xref = Encoding.ASCII.GetByteCount(pdf.ToString());
        pdf.Append($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append($"{offset:0000000000} 00000 n \n");
        pdf.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
```

User docs: `docs/content/reference/cli.md` (command table under `## 常用命令`)
and `docs/content/guides/pdf-import-navigation-attachments.md` (section
`## 启用复制` shows only C#). Unreleased behaviour is marked
`（main，尚未发布）`. Design notes live in `docs/architecture/`.

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

- `src/Ofdrw.Net.Cli/Program.cs`
- `tests/Ofdrw.Net.Cli.Tests/PdfImportNavigationCliTests.cs` (create)
- `docs/content/reference/cli.md`
- `docs/content/guides/pdf-import-navigation-attachments.md`
- `docs/architecture/ofd-outline-pdf-export.md` (create — Part B)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- `PdfToOfdOptions`, the converters, the vector extension, budgets
  (`MaxBookmarkCount` etc. keep their defaults; no CLI options for them).
- `OfdToPdfConverter` — Part B writes a note, not code.
- Any other CLI command.

## Git workflow

- Branch: `codex/cli-pdf-import-navigation`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Part A — CLI flags

### Step A1: Tests (failing first)

Create `tests/Ofdrw.Net.Cli.Tests/PdfImportNavigationCliTests.cs` following
`PageEditCliTests.cs`, with the fixture method above and a `Run` helper
returning `(int Code, string Errors)`.

- `DefaultImportCopiesNothing`: `pdf-to-ofd -i in.pdf -o out.ofd` → 0;
  `package.Outlines` is null; `package.Attachments` is empty. (Passes today.)
- `CopyBookmarksCopiesTheOutline`: add `--copy-bookmarks` → 0; single outline
  item titled `Chapter A`; attachments empty.
- `CopyAttachmentsCopiesEmbeddedFiles`: add `--copy-attachments` → 0; outlines
  null; one attachment named `note.txt` whose data is the ASCII bytes `HELLO`.
- `BothFlagsWorkThroughConvert`: `convert in.pdf out.ofd --copy-bookmarks --copy-attachments`
  → 0; both present.
- `FlagsAreRejectedForOtherCommands` (`[Theory]` over the two flags): first make
  an OFD with `text-to-ofd`, then `ofd-to-pdf -i a.ofd -o a.pdf <flag>` → 1,
  error contains `applies only to PDF to OFD conversion.`, and `a.pdf` does not
  exist.

**Verify**: build; CLI tests filtered with
`--filter "FullyQualifiedName~PdfImportNavigationCliTests"` → the first test
passes, the others fail with `Unknown option: --copy-...`.

### Step A2: Implement

1. `CliOptions`: add `public bool CopyBookmarks { get; set; }` and
   `public bool CopyAttachments { get; set; }`.
2. `ParseOptions`: add cases `--copy-bookmarks` and `--copy-attachments` setting
   them (same shape as `--include-templates`).
3. `RunAsync`: (a) next to the existing "applies only to" checks, reject the
   flags for commands that are never a conversion —
   `if ((options.CopyBookmarks || options.CopyAttachments) && command is not ("convert" or "pdf-to-ofd")) throw new ArgumentException($"{(options.CopyBookmarks ? "--copy-bookmarks" : "--copy-attachments")} applies only to PDF to OFD conversion.");`
   and (b) after `var mode = ResolveMode(...)`, the same exception when
   `mode != ConversionMode.PdfToOfd`.
4. `ConvertAsync`: add two `bool` parameters (`copyBookmarks`,
   `copyAttachments`) at the end before `cancellationToken`, pass
   `options.CopyBookmarks, options.CopyAttachments` at the call site, and build
   the converter as

```csharp
            var converter = new PdfToOfdConverter(new Ofdrw.Net.Converter.Pdf.PdfToOfdOptions
            { CopyBookmarks = copyBookmarks, CopyAttachments = copyAttachments });
```

   Check with `grep -n "ConvertAsync(" src/Ofdrw.Net.Cli/*.cs` that there is no
   other caller to update.
5. `PrintHelp`: change the usage line to
   `ofdrw pdf-to-ofd --input <input.pdf> --output <output.ofd> [--pages 1,3-5] [--copy-bookmarks] [--copy-attachments]`
   and add to the `Options:` list:
   `--copy-bookmarks, --copy-attachments pdf-to-ofd: copy the PDF outline / embedded files; off by default.`

**Verify**: build → `0 Error(s)`; `PdfImportNavigationCliTests` → all pass; full
CLI test project → `Failed: 0`.

### Step A3: Documentation

- `docs/content/reference/cli.md`, table under `## 常用命令`: add a row after
  "PDF → OFD":
  `| PDF → OFD，复制书签和附件（main，尚未发布） | \`ofdrw pdf-to-ofd -i input.pdf -o output.ofd --copy-bookmarks --copy-attachments\` |`
- `docs/content/guides/pdf-import-navigation-attachments.md`, at the end of
  section `## 启用复制`, add a short paragraph and a `sh` code block showing the
  same command, stating that the CLI flags map to `CopyBookmarks` /
  `CopyAttachments`, that the count/size budgets use the SDK defaults, and
  `（main，尚未发布）`.

**Verify**: docs check commands all exit 0.

## Part B — Design note for OFD outline → PDF (no product code)

Create `docs/architecture/ofd-outline-pdf-export.md` (start with a `# ` heading
and no frontmatter, like the other files in that directory; write it in the
same language as `docs/architecture/r03-named-areas-design.md`). It must answer, from
the code, each of the following, citing `file:line` for every claim:

1. **What the model offers**: the outline tree, action types and destination
   fields (`src/Ofdrw.Net.Core/Models/OfdNavigation.cs`), including which
   actions an outline item can carry besides `OfdGoToAction`.
2. **What the PDF writer can express**: open the PdfSharpCore 1.3.67 sources in
   the NuGet cache (`~/.nuget/packages/pdfsharpcore/1.3.67/`; if only binaries
   are there, list the public members of `PdfSharpCore.Pdf.PdfOutline` and
   `PdfOutlineCollection` with a throwaway reflection snippet run outside the
   repository). List which of `XYZ, Fit, FitH, FitV, FitR` it can write and
   whether zoom and "leave unchanged" coordinates are expressible.
3. **A mapping table**: one row per OFD destination type → PDF destination,
   with the coordinate conversion (OFD millimetres, origin top-left → PDF
   points, origin bottom-left) and the cases that cannot be expressed. Use the
   import-side rules in `PdfImportMetadataMigrator.cs` as the mirror image and
   say where the round trip would not be exact.
4. **Page selection and multi-document packages**: what happens to an outline
   item whose target page is not exported (`OfdToPdfConverter` takes a page
   list and a `DocumentIndex`).
5. **Proposed surface**: an opt-in `OfdToPdfOptions` property (name, default
   `false`), a per-item policy for unsupported items (the import side fails the
   whole conversion — say whether export should match and why), and budgets
   (count, depth) with proposed defaults equal to the import side's.
6. **Test and acceptance sketch**: which fixtures, how the result is read back
   (PdfPig can read outlines; the tests already reference it), what visual
   check applies under the repository's rules.
7. **Open questions for the maintainer** — at most five, each with your
   recommended answer.

Keep it under about 150 lines. Do not write converter code, and do not add the
file to site navigation.

**Verify**: the file exists, every section above is present, and
`grep -c "\.cs:" docs/architecture/ofd-outline-pdf-export.md` is at least 8
(claims are cited). `git diff --stat` shows no change under `src/` from Part B.

## Test plan

- New `tests/Ofdrw.Net.Cli.Tests/PdfImportNavigationCliTests.cs` (5 tests, one
  a 2-row theory).

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `PdfImportNavigationCliTests` pass; full CLI test project → `Failed: 0`
- [ ] `dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll --help | grep -c "copy-bookmarks"` → at least `2`
- [ ] Docs check commands exit 0; storage gate exits 0
- [ ] `docs/architecture/ofd-outline-pdf-export.md` exists with the seven sections
- [ ] `plans/plan-index.md` status row updated; the report says Part B is a proposal awaiting a maintainer decision

## STOP conditions

Stop and report back (do not improvise) if:

- A quoted statement in "Current state" cannot be found in any form.
- Plan 028 has not landed (no "applies only to DOCX conversion" check exists in
  `Program.cs`): the applicability checks in Step A2 assume its structure.
- The fixture PDF fails to convert in the default test (`DefaultImportCopiesNothing`)
  — report the error text.
- Part B seems to require running or modifying the converter to answer a
  question. Record the question under "Open questions" instead.

## Maintenance notes

- The CLI exposes only the two switches. If users hit the count/size budgets,
  add options then; do not pre-empt.
- A conversion with `--copy-bookmarks` fails as a whole on an unsupported
  outline target (by SDK design). If that proves too strict for CLI users, the
  place to relax it is the SDK policy, not the CLI.
- Part B's note is an input to a future plan; it is not a commitment.
