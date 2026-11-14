# Plan 036: `ofdrw areas` lists and fills named areas, with a user guide

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Cli src/Ofdrw.Net.Layout/Editing/OfdNamedAreaEditor.cs docs/content/reference/cli.md`
> Plans 021, 028, 030 and 035 change the CLI and `cli.md` before this one;
> locate code by the quoted text, not by line number. If
> `OfdNamedAreaEditor.cs` changed, compare the signatures under "Current state"
> with the live file; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P3 (feature)
- **Effort**: M
- **Risk**: LOW (new command; no existing behaviour changes)
- **Depends on**: plans/030-cli-terminal-output-escaping.md (uses `TerminalText.Escape`), plans/035-cli-pdf-import-navigation.md (same files; land it first)
- **Category**: direction
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The SDK has a complete "template" workflow: a document is laid out with
reserved rectangles (`AreaHolderBlock`), written to OFD, and later — possibly by
another system — each rectangle is filled by name with text or an image,
without moving anything else (`OfdNamedAreaEditor`). It is implemented, tested
(`tests/Ofdrw.Net.Layout.Tests/NamedAreaTests.cs`) and has a design record
(`docs/architecture/r03-named-areas-design.md`).

It is nearly invisible to users: there is no CLI command, and the only mention
in the user guides is half a sentence in
`docs/content/guides/flow-layout.md`. Filling a form field in an existing OFD is
exactly the kind of step people script, so this plan adds an `areas` command
modelled on the existing `attachments` command, and a short guide.

## Current state

### The SDK surface

`src/Ofdrw.Net.Layout/Editing/OfdNamedAreaEditor.cs` (namespace
`Ofdrw.Net.Layout.Editing`):

```csharp
public sealed class OfdNamedArea
{
    public string Name { get; }
    public int PageIndex { get; }          // zero-based
    public string PageId { get; }
    public double XMillimeters { get; }
    public double YMillimeters { get; }
    public double WidthMillimeters { get; }
    public double HeightMillimeters { get; }
    public bool IsFilled { get; }
}

public sealed class NamedAreaTextOptions
{
    public string FontFamily { get; set; } = "SimSun";
    public double FontSizeMillimeters { get; set; } = 3.704166666666667;
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public OfdColor Color { get; set; } = OfdColor.Black;
    public int MaxCharacters { get; set; } = 10000;
}

public sealed class OfdNamedAreaEditor
{
    public IReadOnlyList<OfdNamedArea> GetAreas(OfdDocumentPackage package)
    public void FillText(OfdDocumentPackage package, string name, string text,
        NamedAreaTextOptions? options = null, CancellationToken cancellationToken = default)
    public void FillImage(OfdDocumentPackage package, string name, byte[] encodedImage, string mediaType,
        double sourceWidthPixels, double sourceHeightPixels, CancellationToken cancellationToken = default)
}
```

Behaviour the CLI relies on (all from that file):

- `GetAreas` returns an empty list for a document without an index.
- Unknown name → `KeyNotFoundException("Named area was not found.")`; already
  filled → `InvalidOperationException("Named area has already been filled.")`;
  signed document → `NotSupportedException("Named area fill cannot rewrite a signed document.")`;
  text that does not fit → `InvalidOperationException("Named area text exceeds its Boundary.")`.
  Nothing is modified when a fill is rejected.
- `FillImage` accepts PNG or JPEG up to 16 MiB, media type `image/png` or
  `image/jpeg`, and **checks the declared pixel size against the image header**
  ("Declared PNG dimensions do not match the encoded header."). The caller must
  pass the true width and height.

Creating a document with an area (from `NamedAreaTests.cs`; namespace
`Ofdrw.Net.Layout`):

```csharp
        var flow = new FlowDocument();
        flow.Blocks.Add(new Paragraph("HEADER"));
        flow.Blocks.Add(new AreaHolderBlock("name", 18));     // name, height in mm
        flow.Blocks.Add(new Paragraph("FOOTER"));
        var package = flow.Render();
        await new OfdPackageWriter().WriteAsync(package, stream);
```

A valid 1 × 1 PNG used by the existing tests:
`iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l3sAAAAASUVORK5CYII=`
(Base64).

### The command to model on

`src/Ofdrw.Net.Cli/AttachmentsCommand.cs` (`internal static partial class Cli`):
a subcommand word, then an option loop with an allow-list per subcommand,
duplicate detection, and positional counting:

```csharp
    private static async Task<int> AttachmentsAsync(string[] args, CancellationToken token)
    {
        if (args.Length == 0) throw new ArgumentException("attachments requires list, extract or remove.");
        var subcommand = args[0];
        if (subcommand is not ("list" or "extract" or "remove"))
            throw new ArgumentException($"Unknown attachments subcommand: {subcommand}");

        var positionals = new List<string>();
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var flags = new HashSet<string>(StringComparer.Ordinal);
        var allowed = subcommand switch
        {
            "list" => new[] { "--document", "--json" },
            ...
        };
        for (var index = 1; index < args.Length; index++)
        {
            var arg = args[index];
            if (!arg.StartsWith("-", StringComparison.Ordinal)) { positionals.Add(arg); continue; }
            if (Array.IndexOf(allowed, arg) < 0) throw new ArgumentException($"Unknown attachments {subcommand} option: {arg}");
            if (arg is "--json" or "--physical")
            {
                if (!flags.Add(arg)) throw new ArgumentException($"Duplicate option: {arg}");
                continue;
            }
            if (values.ContainsKey(arg)) throw new ArgumentException($"Duplicate option: {arg}");
            if (index + 1 >= args.Length) throw new ArgumentException($"{arg} requires a value.");
            values[arg] = args[++index];
        }
```

Reading and writing use shared helpers (`src/Ofdrw.Net.Cli/DocumentTools.cs`):

```csharp
    private static async Task<OfdDocumentPackage> ReadToolInputAsync(string path, CancellationToken token, int documentIndex = 0, bool allowMultiple = false)
    private static async Task WriteToolOutputAsync(OfdDocumentPackage package, string path, CancellationToken token)
```

`ReadToolInputAsync` rejects multi-document packages unless
`allowMultiple: true` (read-only listing passes `true`).
`WriteToolOutputAsync` writes through `AtomicOutput` (temporary file beside the
destination, committed on success), so writing over the input is allowed, as
the other editing commands do. `ParseOneBasedPage(text)` converts a 1-based
`--document` value to a zero-based index. `WriteSuccessNotification(message)`
prints the success line.

JSON output pattern (same file, `ListAttachments`):

```csharp
        var document = new JsonObject { ["schema"] = AttachmentsSchema, ["attachments"] = items };
        Console.Out.Write(document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine);
```

Dispatch in `Program.cs` `RunAsync`: the command name must be added to the
`command is not (... or "attachments" or "locate")` list, and a line
`if (command == "attachments") return await AttachmentsAsync(args.Skip(1).ToArray(), cancellationToken).ConfigureAwait(false);`
shows the call shape. Exceptions become `Error: <message>` with exit code 1 in
`RunAsync`'s `catch`. Help text is the raw string in `PrintHelp()` with a
`Usage:` list and a `Commands:` list.

### Tests and docs

CLI test pattern: `tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`
(`[Collection("Owned CLI console")]`, temp directory, `Run` helper returning
`(int Code, string Stdout, string Stderr)`). The CLI test project references the
CLI project, so `Ofdrw.Net.Layout` types are available to tests.

User docs: `docs/content/reference/cli.md` (command table under `## 常用命令`),
guides in `docs/content/guides/` with frontmatter like:

```
---
title: "PDF 书签与附件导入"
description: "…"
sidebar:
  order: 160
  label: "PDF 书签与附件"
---
```

The highest `sidebar.order` among guides at `427433a` is 200
(`ses-interop.md`). Unreleased behaviour is marked `（main，尚未发布）`.
Documentation rules: `docs/engineering/documentation.md`.

## Command design (fixed by this plan)

```
ofdrw areas list <input.ofd> [--document N] [--json]
ofdrw areas fill <input.ofd> <output.ofd> --name NAME (--text TEXT | --image FILE) [--font NAME] [--font-size MM] [--document N]
```

- `list` text output: one line per area, tab-separated:
  `name`, 1-based page number, `x`, `y`, `width`, `height` (millimetres,
  invariant culture, format `0.###`), `blank` or `filled`. The name goes through
  `TerminalText.Escape`. No areas → no output, exit 0.
- `list --json`: standard output is only
  `{ "schema": "ofdrw-cli-areas/1", "areas": [ { "name", "page", "xMillimeters", "yMillimeters", "widthMillimeters", "heightMillimeters", "filled" } ] }`
  with `page` 1-based.
- `fill`: exactly one of `--text` / `--image`. `--font` and `--font-size`
  (millimetres, finite, > 0, invariant culture) apply to `--text` only and are
  rejected with `--image`. The image media type comes from the file extension
  (`.png` → `image/png`, `.jpg`/`.jpeg` → `image/jpeg`, otherwise an error), as
  `watermark --image` does.
- Exit codes: 0 success, 1 any error, 130 cancelled (the general table).
- Success line for `fill`: `Filled area NAME -> OUTPUT`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Run the CLI | `dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll <args>` | — |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Cli/AreasCommand.cs` (create)
- `src/Ofdrw.Net.Cli/Program.cs` (command list, dispatch line, help text)
- `tests/Ofdrw.Net.Cli.Tests/AreasCliTests.cs` (create)
- `docs/content/guides/named-areas.md` (create)
- `docs/content/reference/cli.md`
- `docs/content/guides/flow-layout.md` (one link only, Step 4)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- `OfdNamedAreaEditor` and anything else under `src/` outside the CLI.
- Options for bold, italic, colour, or reading text from a file (see
  Maintenance notes).
- Creating areas from the CLI (areas are created by `AreaHolderBlock` at layout
  time).

## Git workflow

- Branch: `codex/cli-named-areas`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Tests (failing first)

Create `tests/Ofdrw.Net.Cli.Tests/AreasCliTests.cs` following
`AttachmentsCliTests.cs`. Fixture helper: build the three-block `FlowDocument`
shown above with two areas (`"name"`, height 18, and `"logo"`, height 15), write
it to `template.ofd` in the temp directory. Tests:

- `List_PrintsOneRowPerArea`: `areas list template.ofd` → 0; two lines; each
  splits into 7 tab-separated fields; first fields are `name` and `logo`; last
  field `blank`; page field `1`.
- `List_Json`: `--json` → stdout parses as JSON, `schema` is
  `ofdrw-cli-areas/1`, two entries, `filled` false, `page` 1, and
  `heightMillimeters` 18 for `name`.
- `List_DocumentWithoutAreas_PrintsNothing`: an OFD made with `text-to-ofd` →
  0 and empty stdout.
- `FillText_WritesTextInsideTheArea`: `areas fill template.ofd out.ofd --name name --text "Alice 中文"`
  → 0. Read `out.ofd` with `OfdReader`: a text element with that text exists,
  its `XMillimeters`/`YMillimeters` lie inside the area rectangle reported by
  `new OfdNamedAreaEditor().GetAreas(...)`, the `name` area is filled, `logo`
  is not, and the `HEADER`/`FOOTER` text elements have the same coordinates as
  in `template.ofd`.
- `FillText_InPlace`: same with output path equal to input path → 0; no
  `.ofdrw-*.tmp` files remain in the directory.
- `FillImage_FitsTheImage`: write the 1 × 1 PNG to `dot.png`;
  `areas fill template.ofd out.ofd --name logo --image dot.png` → 0; one image
  element on the page, inside the `logo` rectangle.
- `Fill_Twice_Fails`: fill `name`, then fill `name` again on the result → 1,
  stderr contains `already been filled`, and the second output file is not
  created (use a different output path for the second call).
- `Fill_UnknownName_Fails`: `--name nope` → 1, stderr contains `not found`,
  output not created.
- `Fill_TooMuchText_FailsWithoutOutput`: `--text` of 5,000 `W` characters into
  the 18 mm area → 1, stderr contains `exceeds its Boundary`, output not
  created.
- `Fill_ArgumentErrors` (`[Theory]`): neither `--text` nor `--image`; both;
  `--image dot.png --font X`; missing `--name`; `--font-size 0`;
  `--image file.gif`; unknown option `--bold`; a duplicated `--name` → each
  exit 1 and no output file.
- `UnknownSubcommand_Fails`: `areas frob template.ofd` → 1.

**Verify**: build; CLI tests filtered with
`--filter "FullyQualifiedName~AreasCliTests"` → every test fails with
`Unknown command: areas` behaviour (exit code 1 where 0 is expected; the tests
expecting 1 may pass for the wrong reason — that is acceptable at this step).

### Step 2: Implement `AreasCommand.cs`

Create `src/Ofdrw.Net.Cli/AreasCommand.cs` as another part of
`internal static partial class Cli`, with
`private static async Task<int> AreasAsync(string[] args, CancellationToken token)`
structured exactly like `AttachmentsAsync` (subcommand check, allow-list,
duplicate detection, positional count: `list` takes 1, `fill` takes 2). Allowed
options: `list` → `--document`, `--json`; `fill` → `--document`, `--name`,
`--text`, `--image`, `--font`, `--font-size`. Error texts use the same shapes:
`areas requires list or fill.`, `Unknown areas subcommand: X`,
`Unknown areas fill option: X`, `Duplicate option: X`, `X requires a value.`,
`areas fill takes an input and an output OFD path.`, `--name is required.`,
`areas fill requires exactly one of --text or --image.`,
`--font and --font-size apply only to --text.`, `Invalid --font-size value.`,
`Image must be PNG/JPEG.`

`list`:

```csharp
        var package = await ReadToolInputAsync(input, token, documentIndex, allowMultiple: true).ConfigureAwait(false);
        var areas = new Ofdrw.Net.Layout.Editing.OfdNamedAreaEditor().GetAreas(package);
```

then print per the "Command design" section. Format numbers with
`value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)`.

`fill`:

```csharp
        var package = await ReadToolInputAsync(input, token, documentIndex).ConfigureAwait(false);
        var editor = new Ofdrw.Net.Layout.Editing.OfdNamedAreaEditor();
        if (text is not null)
        {
            var options = new Ofdrw.Net.Layout.Editing.NamedAreaTextOptions();
            if (font is not null) options.FontFamily = font;
            if (fontSize is double size) options.FontSizeMillimeters = size;
            editor.FillText(package, name, text, options, token);
        }
        else
        {
            var bytes = await File.ReadAllBytesAsync(imagePath, token).ConfigureAwait(false);
            var info = SixLabors.ImageSharp.Image.Identify(bytes)
                ?? throw new InvalidDataException("Image must be PNG/JPEG.");
            editor.FillImage(package, name, bytes, mediaType, info.Width, info.Height, token);
        }
        await WriteToolOutputAsync(package, output, token).ConfigureAwait(false);
        WriteSuccessNotification($"Filled area {TerminalText.Escape(name)} -> {output}");
        return 0;
```

Before reading the image file, check its length
(`new FileInfo(imagePath).Length > 16 * 1024 * 1024` →
`ArgumentException("Image encoded byte budget exceeded.")`) so an oversized file
is not loaded. Parse `--font-size` with
`double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size)`
and require it finite and greater than zero.

In `Program.cs`: add `"areas"` to the accepted-command list, add the dispatch
line next to the `attachments` one, and add to `PrintHelp()`:

```
          ofdrw areas list <input.ofd> [--document 2] [--json]
          ofdrw areas fill <input.ofd> <output.ofd> --name NAME (--text TEXT | --image FILE) [--font NAME] [--font-size MM] [--document 2]
```

and under `Commands:`

```
          areas        List named areas, or fill one blank area with text or a PNG/JPEG; other content does not move.
```

**Verify**: build → `0 Error(s)`; `AreasCliTests` → all pass; full CLI test
project → `Failed: 0`.

### Step 3: Visual check

Produce a filled sample and render it:

```sh
OUT="$(mktemp -d)"
```

Add to `FillText_WritesTextInsideTheArea` (or a separate test) the usual
artifact hook: when the environment variable `OFDRW_TEST_ARTIFACTS` is set,
copy `template.ofd` and `out.ofd` into that directory. Then:

```sh
OFDRW_TEST_ARTIFACTS="$OUT" dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~AreasCliTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
for f in template out; do dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll ofd-to-image "$OUT/$f.ofd" "$OUT/$f.png" --pages 1; done
```

Open both PNGs. Expected: `HEADER` and `FOOTER` at identical positions in both;
in `out.png` the fill text sits between them inside the reserved band, Chinese
characters are present and legible, nothing overlaps. The viewing chain is
OFD → PDF → PNG (the image exporter). Fill text uses a font **name** without
embedding, so the glyph shapes depend on fonts installed on the machine; note
which font the render used if the output shows a substitution warning. macOS
Preview acceptance is not possible on Linux: report "not performed".

**Verify**: both PNG files exist and the review found no defect — or, if
`ofd-to-image` fails because no font is available for the fill text, record the
exact message and report the visual check as not completed.

### Step 4: Documentation

1. Create `docs/content/guides/named-areas.md` with frontmatter (`title`,
   `description`, `sidebar.order: 210`, `sidebar.label`), in Chinese like the
   other guides, with these sections and nothing invented beyond them:
   - what a named area is and when to use it (2–3 sentences);
   - creating one: the `AreaHolderBlock` snippet from "Current state";
   - filling from code: `GetAreas`, `FillText`, `FillImage` (short snippet);
   - filling from the command line: the two `ofdrw areas` commands, the list
     output columns, the JSON schema name;
   - limits, each taken from `OfdNamedAreaEditor.cs` or
     `docs/architecture/r03-named-areas-design.md`: one fill per area; text
     that does not fit is rejected and nothing is written; text is measured
     with approximate metrics and the font is referenced by name, not embedded;
     images are PNG/JPEG up to 16 MiB and keep their aspect ratio; signed
     documents are rejected; areas exist only in documents created with
     `AreaHolderBlock`.
   Mark the CLI part `（main，尚未发布）`. Check in the design document from
   which version the SDK API has shipped before writing a "since" statement; if
   you cannot establish it from `docs/content/releases/`, write no version.
2. `docs/content/reference/cli.md`: two rows in the `## 常用命令` table
   (`列出命名区域（main，尚未发布）` and `填充命名区域（main，尚未发布）`) with the
   commands, and a short `## 命名区域` section after `## 附件` describing the
   list columns, the JSON schema and the exit codes, linking to the new guide
   with the same URL style as the neighbouring links.
3. In `docs/content/guides/flow-layout.md`, where named areas are mentioned in
   the list near "段落按行、表格按行、固定块与命名区域按整体提供片段", add a link to the
   new guide. (This file is in scope for that one link only.)

**Verify**: docs check commands all exit 0; in the built site output the new
page exists and the two links to it resolve (the `check` script reports broken
links; state what it covered).

## Test plan

- New `tests/Ofdrw.Net.Cli.Tests/AreasCliTests.cs`: 11 tests (one an 8-row
  theory), modelled on `AttachmentsCliTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `AreasCliTests` pass; full CLI test project → `Failed: 0`
- [ ] `dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll --help | grep -c "ofdrw areas"` → `2`
- [ ] Docs check commands exit 0; `docs/content/guides/named-areas.md` has a `title`
- [ ] `git status --short` lists only in-scope files (plus the one-link change in `flow-layout.md`); storage gate exits 0
- [ ] `plans/plan-index.md` status row updated; report lists functional tests, PNG review and "macOS Preview: not performed" separately

## STOP conditions

Stop and report back (do not improvise) if:

- The SDK signatures under "Current state" differ from the live file.
- `SixLabors.ImageSharp.Image.Identify` does not resolve from the CLI project.
  Do not add a package reference; deliver `list` and `fill --text` only, leave
  `--image` rejected with `--image is not available in this build.`, and report.
- `TerminalText` does not exist (plan 030 has not landed).
- A fill changes the coordinates of text outside the area in
  `FillText_WritesTextInsideTheArea` — that would be an SDK defect; report it
  with both coordinate sets.
- The guide would need a claim you cannot source from the two files named in
  Step 4.

## Maintenance notes

- Deliberately left out: `--bold`, `--italic`, `--color`, reading text from a
  file, filling several areas in one invocation. Add them when someone asks;
  the SDK already supports the first three through `NamedAreaTextOptions`.
- Fill text is positioned with approximate metrics and a name-only font. If the
  SDK later lets the paragraph layout engine fill areas, the CLI should gain a
  font-directory option at the same time rather than before.
- Reviewers: check that `fill` never writes the output when the editor throws,
  and that `list` works on multi-document packages with `--document`.
