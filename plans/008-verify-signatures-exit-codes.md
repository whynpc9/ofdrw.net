# Plan 008: Make `verify-signatures` distinguish "unsigned" and document every exit code

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Cli/Program.cs docs/content/reference/cli.md docs/content/guides/signatures.md tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs`
> If any in-scope file changed since this plan was written (plan 005 changes only the OFD→PDF branch of `Program.cs`), compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW (opt-in flag; default exit codes unchanged)
- **Depends on**: plans/005 (both edit `Program.cs`; land 005 first to avoid conflicts)
- **Category**: dx / security
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

`ofdrw verify-signatures` returns exit code **0** both for "fully valid" and for "the document has no signatures at all". A script like `ofdrw verify-signatures -i x.ofd && accept` therefore accepts a document whose signatures were stripped. Exit codes 2 ("references intact, signed value needs a registered verifier") and 3 ("invalid or incomplete") exist and are tested, but the CLI reference documents only 0/1/130. An existing test asserts 0 for unsigned documents (`DocumentToolsCliTests.cs:93`), and scripts may rely on it. So this plan keeps the default and adds an explicit `--require-signatures` flag that makes "unsigned" a failure (exit 4).

## Current state

- `src/Ofdrw.Net.Cli/Program.cs:85-95` dispatches the command:
  ```csharp
  if (command == "verify-signatures")
  {
      if (options.DocumentIndex.HasValue) throw new ArgumentException("--document does not apply to whole-package signature verification.");
      if (string.IsNullOrWhiteSpace(options.InputPath))
      {
          throw new ArgumentException("An input OFD path is required.");
      }
      return await VerifySignaturesAsync(options.InputPath, cancellationToken)
          .ConfigureAwait(false);
  }
  ```
- `Program.cs:296-353`, `VerifySignaturesAsync(string inputPath, CancellationToken cancellationToken)`. The exit-code tail:
  ```csharp
  if (!report.HasSignatureDeclarations && report.Issues.All(issue => issue == "The document does not declare a signatures list."))
  {
      Console.WriteLine("Signature verification: no signature declarations.");
      return 0;
  }
  if (report.FullyValid) { Console.WriteLine("Signature verification: fully valid."); return 0; }
  if (report.ReferenceIntegrityValid && report.Signatures.All(signature =>
          signature.CryptographicStatus == OfdCryptographicVerificationStatus.Unsupported))
  { Console.WriteLine("...signed-value algorithm requires a registered verifier."); return 2; }
  Console.WriteLine("Signature verification: invalid or incomplete.");
  return 3;
  ```
- Option parsing is a `switch` in `ParseOptions` (around `Program.cs:489-520`); boolean flag exemplar at lines 505-506:
  ```csharp
  case "--include-templates":
      options.IncludeTemplates = true;
  ```
  Options class `CliOptions` is at `Program.cs:750-772` (`public bool IncludeTemplates { get; set; }`).
- Help text is in `PrintHelp()` at `Program.cs:672+`. Line 688 reads `ofdrw verify-signatures --input <input.ofd>`; the options list is near line 738.
- Docs: `docs/content/reference/cli.md:67-73` has the exit-code table:
  ```
  | 退出码 | 含义 |
  | `0` | 成功 |
  | `1` | 参数、输入、转换或输出失败 |
  | `130` | 用户取消，例如 Ctrl-C |
  ```
  `docs/content/guides/signatures.md:26-29` shows `ofdrw verify-signatures -i signed.ofd`.
- Tests: `tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs`. Line 93 asserts `0` for a cleaned (unsigned) document; lines 125-139 assert `3` for a missing declared list; `Audit001_CliIncompleteReport` (lines 141-182) asserts `2`/`3`. Tests call `global::Cli.RunAsync([...])` directly. Helper `PathFor(name)` builds a path in the test's temp directory.

## Commands you will need

Once per shell:
```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests -c Release $F` | pass |
| Docs site | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | exit 0 (record as not run if there is no Node ≥ 22.12) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Cli/Program.cs` (verify-signatures dispatch, `VerifySignaturesAsync`, `ParseOptions`, `CliOptions`, `PrintHelp`); `tests/Ofdrw.Net.Cli.Tests/DocumentToolsCliTests.cs` (add tests); `docs/content/reference/cli.md`; `docs/content/guides/signatures.md`.

**Out of scope**: `OfdSignatureVerifier` and the report semantics; the existing default exit codes 0/2/3; other commands' exit codes.

## Git workflow

Branch `advisor/008-verify-signatures-exit-codes`; commit e.g. `feat(cli): add --require-signatures and document verify exit codes`. No push or PR unless instructed.

## Steps

### Step 1: Failing tests

In `DocumentToolsCliTests.cs` add:
- `Verify_RequireSignatures_UnsignedReturns4`: write `new OfdDocumentPackage()` via `OfdPackageWriter` to `PathFor("unsigned.ofd")`. `RunAsync(["verify-signatures", path, "--require-signatures"])` → `4`. Without the flag → `0`, which is the current behavior and stays.
- `Verify_RequireSignatures_DoesNotChangeOtherCodes`: reuse the `Audit001_CliIncompleteReport` fixture-building code for the `"complete"` kind: extract it into a private helper inside the same test class without changing the existing test's assertions. With `--require-signatures` → still `2`.

**Verify**: CLI tests → the new tests fail. Unknown options are rejected with exit 1; confirm that's the failure.

### Step 2: Implement

- Add `public bool RequireSignatures { get; set; }` to `CliOptions` and a `case "--require-signatures": options.RequireSignatures = true; break;` in `ParseOptions`, matching the `--include-templates` exemplar.
- Change the signature to `VerifySignaturesAsync(string inputPath, bool requireSignatures, CancellationToken cancellationToken)` and pass `options.RequireSignatures`.
- In the "no signature declarations" branch: if `requireSignatures`, print `Signature verification: no signature declarations (required).` and `return 4;`, else keep `return 0;`.
- In the dispatch, reject `--require-signatures` for other commands, mirroring the existing `--document` check style: `if (options.RequireSignatures && command != "verify-signatures") throw new ArgumentException("--require-signatures applies only to verify-signatures.");`. Put it right after `ParseOptions` returns, before command branching.
- `PrintHelp`: change the usage line to `ofdrw verify-signatures --input <input.ofd> [--require-signatures]` and add an options line: `--require-signatures Treat a document without signature declarations as a failure (exit 4).`

**Verify**: build exits 0; CLI tests pass, including line 93's existing `0` expectation.

### Step 3: Docs

- `cli.md`: below the generic table, add a sub-table for `verify-signatures`:
  - `0`: fully valid, or no signature declarations when `--require-signatures` is not given
  - `2`: references intact, but no registered verifier for the signed-value algorithm
  - `3`: invalid or incomplete
  - `4`: no signature declarations with `--require-signatures`
  - `1` / `130`: as above

  Add a note that automation accepting documents should use `--require-signatures` and treat only `0` as success. Mark the flag and code 4 as "main, not yet released", following the file's existing style for unreleased flags (see the `--native-embedded-bold-behavior` sentence in the same file).
- `signatures.md`: add `ofdrw verify-signatures -i signed.ofd --require-signatures` next to the existing example, with one sentence explaining it.

**Verify**: docs build/check commands exit 0, or are recorded as not run.

## Done criteria

- [ ] Build exits 0; CLI tests pass (2 new)
- [ ] `grep -n "require-signatures" src/Ofdrw.Net.Cli/Program.cs docs/content/reference/cli.md` matches in both
- [ ] The existing `Assert.Equal(0, ... "verify-signatures", output)` at `DocumentToolsCliTests.cs:93` is unchanged and passes
- [ ] Storage gate exits 0; `plans/README.md` updated

## STOP conditions

- `ParseOptions` treats unknown flags positionally rather than rejecting them, so the "unknown option" premise is false. Report it and adapt only after confirmation.
- The maintainer prefers changing the default (unsigned → non-zero). That is a breaking change and outside this plan.

## Maintenance notes

- A future breaking release could make `--require-signatures` the default. If so, update `DocumentToolsCliTests.cs:93` and the release notes together.
- Plans 015–018 also add commands and options to `Program.cs`; keep help text alphabetical/grouped as existing.
