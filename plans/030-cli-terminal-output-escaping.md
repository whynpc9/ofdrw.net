# Plan 030: The CLI escapes control characters in text that comes from input files

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Cli`
> Plans 021 and 028 change files in this directory, so line numbers below will
> have moved; locate every site by the quoted code, not by number. If a quoted
> statement no longer exists in any form, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S–M
- **Risk**: LOW
- **Depends on**: plans/028-cli-argument-strictness.md (same file `Program.cs`; land it first)
- **Category**: security
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The `ofdrw` CLI is run on OFD files received from other parties. Several
status and error lines include strings taken from the file — ZIP entry names,
attachment names, signature diagnostics — and print them unchanged. Checked at
`427433a`: a package whose ZIP entry name contains an escape character (U+001B)
and a carriage return makes `ofdrw verify-signatures` print those bytes raw
inside `Error: OFD package entry has an unsafe path: ...`.

A terminal interprets such bytes. A crafted file can therefore recolour, erase
or overwrite what the operator sees — for example make a failure line look like
a success line — and a tab or line break in an attachment name forges extra
columns or rows in `attachments list`, which scripts parse. Exit codes are not
affected, but people read the text.

After this plan every string that originates in an input file is printed with
control and bidirectional-format characters replaced by visible `\uXXXX`
escapes. JSON output and the extracted text payload are unchanged.

## Current state

The CLI lives in `src/Ofdrw.Net.Cli/` (global namespace; `Program.cs` is
`internal static partial class Cli`; helper types are separate `internal` types
in their own files, e.g. `AtomicOutput.cs`). `Ofdrw.Net.Cli.csproj` grants
`InternalsVisibleTo` to `Ofdrw.Net.Cli.Tests`.

Sites that print file-derived text (all found with
`grep -n "Console\.\(Error\.\)\?Write" src/Ofdrw.Net.Cli/*.cs`):

| File | Statement (as of `427433a`) | File-derived part |
|---|---|---|
| `Program.cs` (`RunAsync` catch) | `Console.Error.WriteLine($"Error: {ex.Message}");` | exception messages quote entry names, IDs, paths |
| `Program.cs` (4 sites), `PageEditCommands.cs` (1), `DocumentTools.cs` (2) | `Console.Error.WriteLine($"Warning: {diagnostic}");` | diagnostic text |
| `Program.cs` (2 sites) | `Console.Error.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");` | `Message` |
| `Program.cs` (`VerifySignaturesAsync`) | `Console.WriteLine($"Issue: {issue}");` | `issue` |
| `Program.cs` (`VerifySignaturesAsync`) | `Console.WriteLine($"Signature {signature.Id}: references " + ... + $"method {signature.SignatureMethod}");` | `Id`, `SignatureMethod` |
| `Program.cs` (`VerifySignaturesAsync`) | `Console.WriteLine($"  {signature.CryptographicMessage}");` | `CryptographicMessage` |
| `PreflightCommand.cs` | `Console.WriteLine($"  {result.RuleId} {result.Outcome} {result.ReasonCode} {result.Evidence.Entry} {result.Evidence.Selector}");` | `Entry`, `Selector` |
| `AttachmentsCommand.cs` (`ListAttachments`) | `Console.WriteLine(string.Join("\t", attachment.Id ?? string.Empty, attachment.Name, attachment.MediaType, ...));` | `Id`, `Name`, `MediaType` |
| `LocateCommand.cs` | `Console.WriteLine($"page {hit.PageIndex + 1} \"{hit.MatchedText}\" box {box} replaceable=...");` | `MatchedText` |
| `LocateCommand.cs` | `Console.WriteLine($"diagnostic {diagnostic.Code} page ... {diagnostic.SourcePath ?? "-"}");` | `SourcePath` |

Sites that must stay as they are:

- `Program.cs` `ExtractTextAsync`: `Console.WriteLine(text);` — this is the
  extracted document text, the command's payload.
- Every `--json` branch (`System.Text.Json` already escapes control characters).
- Help text, `Conversion canceled.`, success notifications
  (`SuccessNotification.cs`, `TextReplacementCommand.cs`) and
  `Unknown command: {args[0]}` — these print only what the user typed.

One source of such strings, for orientation
(`src/Ofdrw.Net.Packaging/Archive/OfdPackageLoader.cs:109-118`):

```csharp
    private static string NormalizeAndValidate(string path)
    {
        var normalized = path.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal) ||
            normalized.Split('/').Any(x => x == ".."))
        {
            throw new InvalidDataException($"OFD package entry has an unsafe path: {path}");
        }
```

The library messages are not changed by this plan: a library cannot know where
its messages will be shown; the program that prints them escapes.

Test conventions (`tests/Ofdrw.Net.Cli.Tests/AttachmentsCliTests.cs`): class
attribute `[Collection("Owned CLI console")]`, a temp directory per test class,
a `Run` helper returning `(int Code, string Stdout, string Stderr)`, and
packages built from an entry dictionary:

```csharp
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'>...<Attachments>Metadata/List.xml</Attachments></Document>"),
            ...
            ["Doc_0/Metadata/List.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='7' Name='original.bin'><FileLoc>../Payload/original.bin</FileLoc></Attachment>...</Attachments>"),
        };
        using (var zip = MultiDocumentFixture.Zip(entries)) await File.WriteAllBytesAsync(path, zip.ToArray());
```

(see the `Seeded()` method there for the full dictionary and the `Ns` constant).

User docs: `docs/content/reference/cli.md`, section `## 附件` describes the
tab-separated `attachments list` format; section `## 退出码与输出安全` covers
output behaviour. Unreleased behaviour is marked `（main，尚未发布）`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| List write sites | `grep -n "Console\.\(Error\.\)\?Write" src/Ofdrw.Net.Cli/*.cs` | (inspection) |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Cli/TerminalText.cs` (create)
- `src/Ofdrw.Net.Cli/Program.cs`, `PageEditCommands.cs`, `DocumentTools.cs`,
  `PreflightCommand.cs`, `AttachmentsCommand.cs`, `LocateCommand.cs` — only the
  statements in the table above
- `tests/Ofdrw.Net.Cli.Tests/TerminalTextTests.cs` (create)
- `docs/content/reference/cli.md`
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- Exception and diagnostic messages in the libraries under `src/` other than
  the CLI.
- JSON output, the `extract-text` payload, file output.
- Wrapping or replacing `Console.Out` / `Console.Error` globally.

## Git workflow

- Branch: `codex/cli-terminal-output-escaping`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: The helper and its unit tests

Create `src/Ofdrw.Net.Cli/TerminalText.cs`:

```csharp
using System.Globalization;
using System.Text;

/// <summary>Makes text that originates in an input file safe to print on one terminal line.</summary>
internal static class TerminalText
{
    /// <summary>Replaces control and bidirectional-format characters with visible \uXXXX escapes.
    /// Other text, including CJK, is returned unchanged.</summary>
    internal static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        StringBuilder? builder = null;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (!NeedsEscape(character)) { builder?.Append(character); continue; }
            builder ??= new StringBuilder(value, 0, index, value.Length + 8);
            builder.Append("\\u").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
        }
        return builder?.ToString() ?? value;
    }

    private static bool NeedsEscape(char c) =>
        c < 0x20 || c == 0x7F || c is >= (char)0x80 and <= (char)0x9F ||   // C0 (incl. tab, CR, LF), DEL, C1
        c == 0x061C || c == 0x200E || c == 0x200F ||                        // bidi marks
        c is >= (char)0x202A and <= (char)0x202E ||                         // bidi embeddings and overrides
        c is >= (char)0x2066 and <= (char)0x2069 ||                         // bidi isolates
        c == 0x2028 || c == 0x2029;                                         // line and paragraph separators
}
```

Create `tests/Ofdrw.Net.Cli.Tests/TerminalTextTests.cs` (class
`TerminalTextTests`, `[Collection("Owned CLI console")]`, temp directory and
`Run` helper copied from `AttachmentsCliTests.cs`). Unit tests:

- `PlainAndCjkTextIsUnchanged`: `"abc 中文 é"` → same instance content; `null`
  and `""` → `""`.
- `ControlAndBidiCharactersBecomeVisible` (`[Theory]`): for each of U+0009,
  U+000A, U+000D, U+001B, U+007F, U+009B, U+200F, U+202E, U+2066, U+2028 —
  build `"a" + character + "b"` in the test body from an `int` code point
  (`[InlineData(0x1B)]` etc.) and assert the result equals
  `"a\\u" + code.ToString("X4") + "b"` and contains no character for which
  `char.IsControl` is true.

**Verify**: build → `0 Error(s)`; CLI tests filtered with
`--filter "FullyQualifiedName~TerminalTextTests"` → all pass.

### Step 2: Integration tests (failing first)

Add to `TerminalTextTests`:

- `ErrorMessagesEscapeEntryNames`: build a ZIP with `System.IO.Compression`
  containing `OFD.xml` (any content) and a second entry whose name is
  `"../a" + (char)0x1B + "b" + (char)0x0D + "c"`; run `verify-signatures <path>`.
  Assert `Code == 1`, `Stderr` contains `\u001B` and `\u000D` as literal
  six-character texts, `Stderr` contains no U+001B character, and
  `Stderr.TrimEnd('\r', '\n')` contains no `'\r'` or `'\n'`.
- `AttachmentRowsKeepFiveColumns`: copy the `Seeded()` dictionary from
  `AttachmentsCliTests.cs` but give the first attachment
  `Name='a&#9;b&#x202E;c'` (XML character references for tab and
  right-to-left override). Run `attachments list <path>`. Assert `Code == 0`,
  the first output line split on `'\t'` has exactly 5 fields, and its second
  field equals `a\u0009b‮c` (literal backslash text; in C# write
  `@"a\u0009b‮c"`).
- `JsonOutputIsUnchanged`: same package, `attachments list <path> --json` →
  parse with `JsonDocument`; the attachment's `name` property equals the real
  string `"a\tb‮c"` (the JSON layer does its own escaping; the value
  round-trips).

**Verify**: the first two tests FAIL (raw characters present / 6 fields), the
JSON test passes.

### Step 3: Apply the helper at every site in the table

Wrap each file-derived value, not the fixed text around it. Examples:

```csharp
            Console.Error.WriteLine($"Error: {TerminalText.Escape(ex.Message)}");
```

```csharp
        foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {TerminalText.Escape(diagnostic.ToString())}");
```

```csharp
                Console.WriteLine(string.Join("\t", TerminalText.Escape(attachment.Id), TerminalText.Escape(attachment.Name),
                    TerminalText.Escape(attachment.MediaType),
```

```csharp
                Console.WriteLine($"  {result.RuleId} {result.Outcome} {result.ReasonCode} {TerminalText.Escape(result.Evidence.Entry)} {TerminalText.Escape(result.Evidence.Selector)}");
```

For interpolated values whose type is an enum or a number, no wrapping is
needed. When unsure whether a property is a string, check its declaration; wrap
every `string` (and `object.ToString()`) value.

**Verify**: build → `0 Error(s)`; `TerminalTextTests` all pass; the full CLI
test project → `Failed: 0`.

### Step 4: Audit the remaining write sites

Run the "List write sites" grep. For every line, confirm it is either (a)
wrapped, (b) in the "must stay" list, or (c) prints only constants, numbers and
enums. Put the classification of each line in your report.

**Verify**: no line is unclassified.

### Step 5: Document

In `docs/content/reference/cli.md`, section `## 退出码与输出安全`, add one
paragraph in Chinese:

> 文本模式的状态、诊断和列表输出会把来自输入文件的控制字符和双向格式字符显示为 `\uXXXX`（例如制表符显示为 `\u0009`），因此 `attachments list` 的列数固定；`--json` 输出和 `extract-text` 提取的正文不做此替换，在终端直接查看不可信文档的正文时建议写入文件（main，尚未发布）。

**Verify**: docs check commands all exit 0.

## Test plan

- New `tests/Ofdrw.Net.Cli.Tests/TerminalTextTests.cs`: 2 unit tests (one a
  10-row theory) and 3 integration tests, modelled on `AttachmentsCliTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `TerminalTextTests` all pass; full CLI test project → `Failed: 0`
- [ ] `grep -c "TerminalText.Escape" src/Ofdrw.Net.Cli/*.cs` shows hits in `Program.cs`, `PageEditCommands.cs`, `DocumentTools.cs`, `PreflightCommand.cs`, `AttachmentsCommand.cs`, `LocateCommand.cs`
- [ ] Step 4 classification included in the report
- [ ] Docs check commands exit 0; storage gate exits 0
- [ ] `plans/plan-index.md` status row updated

## STOP conditions

Stop and report back (do not improvise) if:

- A statement from the table cannot be found in any form.
- An existing test asserts raw control characters or exact multi-line error
  text that the escaping changes. Report the test; do not edit it.
- The reader rejects the attachment name used in `AttachmentRowsKeepFiveColumns`
  (then the row format is already protected upstream — report which check
  rejects it and keep the other tests).

## Maintenance notes

- Any new CLI line that prints a string from an input file must go through
  `TerminalText.Escape`. Reviewers: ask this for every new `Console.Write*`.
- Multi-line exception messages now appear on one line with `\u000A`. That is
  intended; do not special-case line feeds.
- `locate` text output puts the matched text in quotes; a quote inside the
  match is not escaped. Scripts should use `--json`.
