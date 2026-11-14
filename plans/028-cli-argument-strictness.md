# Plan 028: The CLI rejects arguments it would otherwise silently ignore

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Cli/Program.cs docs/content/reference/cli.md`
> If `Program.cs` changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition. (A small change to `cli.md` from plan 021 is expected.)

## Status

- **Priority**: P2
- **Effort**: S–M
- **Risk**: MED (scripts that pass ignored arguments start failing — that is the point, but it is a behaviour change)
- **Depends on**: plans/021-cli-output-safety.md (touches the same doc page; land it first)
- **Category**: bug
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

The shared option parser of the `ofdrw` CLI accepts several inputs and then does
something other than what was typed, with exit code 0. All of these were run
against the Release build at `427433a`:

| Command line | What happened |
|---|---|
| `ofdrw merge -o a.ofd b.ofd a.ofd` | `Merged 3 OFD files -> -o` — wrote a file named `-o`, using the intended output as an input |
| `ofdrw ofd-to-pdf -i a.ofd x.pdf y.pdf` | `Converted a.ofd -> y.pdf` — `x.pdf` ignored |
| `ofdrw ofd-to-pdf -i a.ofd -o z.pdf stray.pdf` | converted; `stray.pdf` ignored |
| `ofdrw extract-text -i a.ofd -i b.ofd` | printed the text of `b.ofd` only |
| `ofdrw extract-text a.ofd --pages 99` | printed the whole document; `--pages` ignored |
| `ofdrw verify-signatures a.ofd out.txt --pages 7 --include-templates` | verified; three arguments ignored |
| `ofdrw ofd-to-pdf -i a.ofd -o q.pdf --docx-engine word --libreoffice /nonexistent --font-directory /nonexistent` | converted; DOCX options ignored |
| `ofdrw ofd-to-pdf -i a.ofd -o p.pdf --pages +1` | converted page 1 (sign accepted; parsing uses the current culture) |

A tool used in batch jobs must fail on a command line it does not understand.
After this plan each row above exits 1 with a message naming the offending
argument.

## Current state

All code is in `src/Ofdrw.Net.Cli/Program.cs` (`internal static partial class Cli`).

Dispatch (`RunAsync`, lines 46–67): `text-to-ofd`, `ofd-to-html`,
`ofd-to-image`, `image-to-ofd`, `replace-text`, `preflight`, `attachments`,
`locate`, `watermark`, `split`, `mix`, `clean-signatures` and `merge` have their
own parsers. Everything else — `convert`, `docx-to-pdf`, `docx-to-ofd`,
`pdf-to-ofd`, `ofd-to-pdf`, `ofd-to-svg`, `extract-text`, `verify-signatures`,
`reorder`, `delete-pages`, `crop` — goes through `ParseOptions` (line 68):

```csharp
            var options = ParseOptions(args.Skip(1).ToArray());
            if (IsHelpRequested(options))
            {
                PrintHelp();
                return 0;
            }

            if (options.RequireSignatures && command != "verify-signatures")
                throw new ArgumentException("--require-signatures applies only to verify-signatures.");
            if (options.Box is not null && command != "crop")
                throw new ArgumentException("--box applies only to crop.");
```

Those two checks are the existing pattern for "option does not apply". Later in
the same method (lines 139–141):

```csharp
            var mode = ResolveMode(command, options.InputPath, options.OutputPath);
            if (options.DocumentIndex.HasValue && mode != ConversionMode.OfdToPdf)
                throw new ArgumentException("--document applies only to OFD input commands.");
```

`ParseOptions` (lines 496–577), abridged:

```csharp
                case "-i":
                case "--input":
                    options.InputPath = ReadValue(args, ref i, arg);
                    break;
                case "-o":
                case "--output":
                    options.OutputPath = ReadValue(args, ref i, arg);
                    break;
                case "-p":
                case "--pages":
                    options.Pages = ReadValue(args, ref i, arg);
                    break;
                case "--box":
                    if (options.Box is not null) throw new ArgumentException("Duplicate --box.");
                    options.Box = ReadValue(args, ref i, arg);
                    break;
                ...
                case "--libreoffice":
                    options.LibreOfficePath = ReadValue(args, ref i, arg);
                    break;
                case "--docx-engine":
                    options.DocxEngine = ParseDocxEngine(ReadValue(args, ref i, arg));
                    break;
                case "--docx-ofd-mode":
                    options.DocxOfdMode = ParseDocxToOfdMode(ReadValue(args, ref i, arg));
                    break;
                case "--native-embedded-bold-behavior":
                    options.NativeEmbeddedBoldBehavior = ParseNativeEmbeddedBoldBehavior(ReadValue(args, ref i, arg));
                    break;
                case "--font-directory":
                    options.FontDirectories.Add(ReadValue(args, ref i, arg));
                    break;
                default:
                    if (arg.StartsWith("-", StringComparison.Ordinal))
                    {
                        throw new ArgumentException($"Unknown option: {arg}");
                    }

                    positionals.Add(arg);
                    break;
            }
        }

        if (options.InputPath is null && positionals.Count > 0)
        {
            options.InputPath = positionals[0];
        }

        if (options.OutputPath is null && positionals.Count > 1)
        {
            options.OutputPath = positionals[1];
        }

        if (positionals.Count > 2)
        {
            throw new ArgumentException("Too many positional arguments.");
        }
```

Page numbers (lines 645–653):

```csharp
    private static int ParseOneBasedPage(string value)
    {
        if (!int.TryParse(value, out var page) || page <= 0)
        {
            throw new ArgumentException($"Invalid page number: {value}");
        }

        return page - 1;
    }
```

`merge` (lines 408–421):

```csharp
        var skipUnsupported = false;
        var paths = new List<string>();
        foreach (var arg in args)
        {
            if (arg is "-h" or "--help") { PrintHelp(); return 0; }
            if (arg == "--skip-unsupported") skipUnsupported = true;
            else paths.Add(arg);
        }
        if (paths.Count < 3)
        {
            throw new ArgumentException(
                "Merge requires an output path followed by at least two input OFD files.");
        }
```

All `ArgumentException`s are caught at lines 160–164 and printed as
`Error: <message>` with exit code 1.

Which commands really use which option (read from `RunAsync` lines 79–153 and
`ConvertAsync` lines 167–232):

| Option | Used by |
|---|---|
| output path | every `ParseOptions` command except `verify-signatures` (optional for `extract-text`) |
| `--pages` | `convert`, `docx-to-ofd`, `pdf-to-ofd`, `ofd-to-pdf`, `ofd-to-svg`, `reorder`, `delete-pages`, `crop` (`docx-to-pdf` already rejects it) |
| `--include-templates` | `extract-text` |
| `--docx-engine`, `--docx-ofd-mode`, `--native-embedded-bold-behavior`, `--libreoffice`, `--font-directory` | only when the conversion mode is DOCX→PDF or DOCX→OFD |

Test conventions (`tests/Ofdrw.Net.Cli.Tests/PageEditCliTests.cs`): the class
has `[Collection("Owned CLI console")]`, a per-test temp directory, and

```csharp
    private static async Task<int> Run(params string[] args)
    {
        var saved = Console.Error; var savedOut = Console.Out;
        using var errors = new StringWriter(); using var output = new StringWriter();
        try { Console.SetError(errors); Console.SetOut(output); return await global::Cli.RunAsync(args); }
        finally { Console.SetError(saved); Console.SetOut(savedOut); }
    }
```

An OFD test input can be produced with the CLI itself:
`Run("text-to-ofd", textPath, ofdPath)` (a UTF-8 text file in, an OFD out).

User docs: `docs/content/reference/cli.md`, section `## 退出码与输出安全`
(exit code 1 = "参数、输入、转换或输出失败"). Behaviour on `main` that is not in
a published version is marked `（main，尚未发布）`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Callers of the CLI | `grep -rnE "ofdrw |Ofdrw.Net.Cli.dll|Cli.RunAsync" --include="*.sh" --include="*.py" --include="*.yml" --include="*.cs" --include="*.md" e2e scripts .github docs/content` | (inspection) |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Cli/Program.cs`
- `tests/Ofdrw.Net.Cli.Tests/ArgumentStrictnessTests.cs` (create)
- `docs/content/reference/cli.md`
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- The commands with their own parsers (listed under "Dispatch" above) except
  `merge`. Several already reject duplicates; do not unify them.
- Option names, exit codes, help text layout (add no new options).
- `--font-directory` stays repeatable.

## Git workflow

- Branch: `codex/cli-argument-strictness`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Write the failing tests

Create `tests/Ofdrw.Net.Cli.Tests/ArgumentStrictnessTests.cs` following
`PageEditCliTests.cs` (collection attribute, temp directory, `Dispose`). Use a
`Run` helper that returns `(int Code, string Errors)` (return
`errors.ToString()` too). In the constructor create `a.ofd` and `b.ofd` from two
small text files with `text-to-ofd`.

One `[Fact]` per row (each asserts exit code 1, the quoted fragment in the error
text, and — where an output path is named — that the file was not created):

| Test | Arguments | Error contains |
|---|---|---|
| `MergeRejectsUnknownOption` | `merge -o out.ofd a.ofd b.ofd` | `Unknown option: -o`; no file named `-o` in the temp directory or the current directory |
| `ExtraPositionalIsRejectedWithInputOption` | `ofd-to-pdf -i a.ofd x.pdf y.pdf` | `Too many positional arguments.` |
| `ExtraPositionalIsRejectedWithBothOptions` | `ofd-to-pdf -i a.ofd -o z.pdf stray.pdf` | `Too many positional arguments.` |
| `DuplicateInputIsRejected` | `extract-text -i a.ofd --input b.ofd` | `Duplicate --input.` |
| `DuplicateOutputIsRejected` | `ofd-to-pdf -i a.ofd -o z.pdf -o y.pdf` | `Duplicate --output.` |
| `DuplicatePagesIsRejected` | `ofd-to-pdf -i a.ofd -o z.pdf -p 1 --pages 1` | `Duplicate --pages.` |
| `PagesDoesNotApplyToExtractText` | `extract-text a.ofd --pages 1` | `--pages does not apply to extract-text.` |
| `VerifySignaturesTakesNoOutput` | `verify-signatures a.ofd out.txt` | `verify-signatures does not take an output path.` |
| `IncludeTemplatesAppliesOnlyToExtractText` | `ofd-to-pdf -i a.ofd -o z.pdf --include-templates` | `--include-templates applies only to extract-text.` |
| `DocxOptionsApplyOnlyToDocxInput` (`[Theory]`: `--docx-engine word`, `--docx-ofd-mode native`, `--native-embedded-bold-behavior throw`, `--libreoffice /x`, `--font-directory /x`) | `ofd-to-pdf -i a.ofd -o z.pdf <option> <value>` | `applies only to DOCX conversion.` and the option name |
| `SignedPageNumberIsRejected` | `ofd-to-pdf -i a.ofd -o z.pdf --pages +1` | `Invalid page number: +1` |

Plus two tests that must keep passing (they pass before and after):

- `OutputOptionWithPositionalInputStillWorks`: `ofd-to-pdf -o z.pdf a.ofd` → 0, `z.pdf` exists.
- `InputOptionWithPositionalOutputWorks`: `ofd-to-pdf -i a.ofd z.pdf` → 0, `z.pdf`
  exists. (Before this plan this one fails with `Input and output paths are
  required.`; it is listed with the failing tests.)

Never change the process working directory in a test; always pass absolute
paths built from the temp directory.

**Verify**: build, then run the CLI tests with
`--filter "FullyQualifiedName~ArgumentStrictnessTests"` → all new tests fail
except `OutputOptionWithPositionalInputStillWorks`.

### Step 2: Duplicates and positionals in `ParseOptions`

1. Add `var seen = new HashSet<string>(StringComparer.Ordinal);` and a local
   function `void Once(string name) { if (!seen.Add(name)) throw new ArgumentException($"Duplicate {name}."); }`.
   Call it with the long name at the start of these cases: `--input`,
   `--output`, `--pages`, `--libreoffice`, `--docx-engine`, `--docx-ofd-mode`,
   `--native-embedded-bold-behavior`. Leave the existing `--box` and
   `--document` checks as they are.
2. Record which DOCX-only options were typed: add
   `public List<string> DocxOnlyOptions { get; } = new();` to `CliOptions` and add
   the option's long name in the five DOCX cases (including `--font-directory`).
3. Replace the positional assignment at the end with:

```csharp
        var next = 0;
        if (options.InputPath is null && next < positionals.Count) options.InputPath = positionals[next++];
        if (options.OutputPath is null && next < positionals.Count) options.OutputPath = positionals[next++];
        if (next < positionals.Count) throw new ArgumentException("Too many positional arguments.");
```

**Verify**: build → `0 Error(s)`; the duplicate and positional tests pass.

### Step 3: Applicability checks in `RunAsync`

Directly after the existing `--box applies only to crop.` check add, in the same
style:

```csharp
            if (options.IncludeTemplates && command != "extract-text")
                throw new ArgumentException("--include-templates applies only to extract-text.");
            if (options.Pages is not null && command is "extract-text" or "verify-signatures")
                throw new ArgumentException($"--pages does not apply to {command}.");
            if (command == "verify-signatures" && options.OutputPath is not null)
                throw new ArgumentException("verify-signatures does not take an output path.");
```

Mind operator precedence: write the second condition as
`options.Pages is not null && command is ("extract-text" or "verify-signatures")`.

For the DOCX-only options, the check must cover every command that is not a
DOCX conversion. Add a helper and call it (a) before each early branch that is
not a conversion — `extract-text`, `verify-signatures`, `delete-pages`, `crop`,
`reorder`, `ofd-to-svg` — simplest is one call right after the checks above
guarded by `command is not ("convert" or "docx-to-pdf" or "docx-to-ofd" or "pdf-to-ofd" or "ofd-to-pdf")`,
and (b) after `var mode = ResolveMode(...)` guarded by
`mode is not (ConversionMode.DocxToPdf or ConversionMode.DocxToOfd)`:

```csharp
    private static void RejectDocxOnlyOptions(CliOptions options)
    {
        if (options.DocxOnlyOptions.Count > 0)
            throw new ArgumentException($"{options.DocxOnlyOptions[0]} applies only to DOCX conversion.");
    }
```

**Verify**: the applicability tests pass; the full CLI test project →
`Failed: 0` (in particular `NativeEmbeddedBoldArgumentsTests`, which passes
`--font-directory` to `docx-to-ofd` and `convert` with a `.docx` input).

### Step 4: Page numbers and `merge`

1. `ParseOneBasedPage`: parse digits only, culture-independent:

```csharp
        if (!int.TryParse(value, System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out var page) || page <= 0)
```

2. `MergeAsync`: in the loop, before `else paths.Add(arg);` add

```csharp
            else if (arg.StartsWith("-", StringComparison.Ordinal)) throw new ArgumentException($"Unknown option: {arg}");
```

`MergeAsync` is called inside `RunAsync`'s `try`, so the exception becomes
`Error: ...` with exit 1 like the others.

**Verify**: `ArgumentStrictnessTests` → all pass; full CLI test project →
`Failed: 0`.

### Step 5: Check the repository's own callers

Run the "Callers of the CLI" grep. For every command line found in `e2e/`,
`scripts/`, `.github/` and `docs/content/`, check it against the new rules
(duplicates, extra positionals, DOCX options on non-DOCX commands, `--pages` on
`extract-text`). You are looking for a caller this plan would break.

**Verify**: none found. If one is found, STOP (see below).

### Step 6: Document

In `docs/content/reference/cli.md`, section `## 退出码与输出安全`, add one
paragraph in Chinese after the paragraph beginning `文件型命令先验证参数`:

> 通用转换与页面命令拒绝无法执行的参数并返回 `1`：重复的 `--input`/`--output`/`--pages`、多余的位置参数、对当前命令无效的选项（例如 `extract-text` 的 `--pages`、非 DOCX 转换中的 DOCX 选项），以及 `merge` 中未知的 `-` 开头参数（main，尚未发布）。

**Verify**: docs check commands all exit 0.

## Test plan

- New `tests/Ofdrw.Net.Cli.Tests/ArgumentStrictnessTests.cs` with the cases in
  Step 1 (11 rejecting cases incl. a 5-row theory, 2 accepting cases).
- Pattern: `tests/Ofdrw.Net.Cli.Tests/PageEditCliTests.cs`.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `ArgumentStrictnessTests` all pass; full CLI test project → `Failed: 0`
- [ ] No existing test was edited (`git diff --stat` shows only the files in scope)
- [ ] Docs check commands exit 0; storage gate exits 0
- [ ] `plans/plan-index.md` status row updated; the report lists the callers inspected in Step 5

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- An existing test, e2e harness, script, workflow or documented example uses a
  command line the new rules reject. Report it — the maintainer decides whether
  the caller or the rule changes. Do not weaken a rule to make it pass.
- A check seems to need a change in one of the out-of-scope command parsers.

## Maintenance notes

- New options added to `ParseOptions` need a decision at the same time: which
  commands use them, and a matching "applies only to" check.
- `--help` handling is unchanged: `ReadValue` still lets `-h` be a value.
- This is a behaviour change for callers; it must be listed under
  compatibility in the next release note (`docs/content/releases/`).
