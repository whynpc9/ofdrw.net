# Plan 011: Make font- and platform-dependent DOCX tests skip visibly (or fail) instead of passing empty

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- tests/Ofdrw.Net.Converter.Docx.Tests/DocxRevisionTests.cs tests/Ofdrw.Net.Converter.Docx.Tests/LibreOfficeFontStagingTests.cs .github/workflows/ci.yml`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: tests
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Two CJK font tests and one Windows-fonts test `return;` before any assertion when their font or directory is missing. xUnit 2 reports them as **passed**, so a developer machine, or a CI runner where font installation silently failed, shows green while testing nothing. The helper's own comment says "A required gate must not silently return before its assertions run". The repo already has the correct pattern: attributes that set `Skip` at discovery time, so the skip and its reason are visible in TRX.

## Current state

- `tests/Ofdrw.Net.Converter.Docx.Tests/DocxRevisionTests.cs`
  - Lines 52-57:
    ```csharp
    [Fact]
    public async Task Native_ShouldDeclareViewerLocalCjkNamesWithoutEmbeddingASubstituteFace()
    {
        var metricFont = TryFindCjkMetricFont();
        if (metricFont is null)
            return;
    ```
  - Lines 101-106: same pattern in `Native_ShouldKeepDistinctViewerLocalCjkStyleResources`.
  - Lines 605-620+, helper:
    ```csharp
    private static string? TryFindCjkMetricFont()
    {
        // Use the same offline fixture supplied to the mandatory PDF font gate.
        // A required gate must not silently return before its assertions run.
        var trusted = Environment.GetEnvironmentVariable("OFDRW_TEST_TRUSTED_CJK_FONT");
        if (!string.IsNullOrEmpty(trusted))
        {
            if (!File.Exists(trusted)) throw new FileNotFoundException("Required CJK fixture is missing.", trusted);
            return trusted;
        }
        if (Environment.GetEnvironmentVariable("OFDRW_REQUIRE_TRUSTED_FONTS") == "1")
            throw new InvalidOperationException("Required DOCX font gate needs OFDRW_TEST_TRUSTED_CJK_FONT.");
        foreach (var directory in DocxFontCatalog.DefaultPlatformFontDirectories()) { ... }
    ```
- `tests/Ofdrw.Net.Converter.Docx.Tests/LibreOfficeFontStagingTests.cs:82-95`:
  ```csharp
  [Fact]
  public void ShouldRestrictPreferredCjkOnlyForLargePlatformFontDirectories()
  {
      var custom = Path.Combine(Path.GetTempPath(), "ofdrw-corporate-fonts");
      Assert.False(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(custom, 80));
      var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
      if (string.IsNullOrEmpty(windows))
          return;
      var fonts = Path.Combine(windows, "Fonts");
      if (!Directory.Exists(fonts))
          return;
      Assert.False(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(fonts, 32));
      Assert.True(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(fonts, 33));
  }
  ```
  The first assertion runs everywhere; only the Windows half silently returns.
- Exemplar skip attribute, `tests/Ofdrw.Net.Core.Tests/Audit024ProtocolTests.cs:128-131`:
  ```csharp
  private sealed class UnixProtocolFactAttribute : FactAttribute
  {
      public UnixProtocolFactAttribute() { if (OperatingSystem.IsWindows()) Skip = "Unix socket protocol; Windows retains the Job backend."; }
  }
  ```
  Another, `tests/Ofdrw.Net.Converter.Pdf.Tests/TrustedFontGeometryTests.cs:304-310` (`RealTrustedFontsTheoryAttribute`), mentions `OFDRW_REQUIRE_TRUSTED_FONTS=1` in its skip message.
- CI:
  - The main DOCX step "Test BuiltIn without office applications" (`.github/workflows/ci.yml:278-280`) filters `BuiltIn|Auto|Native|DualLayer`, so it **includes** these `Native_*` tests without any trusted-font env.
  - The step "Verify original font identity through native OFD and PDF" (`ci.yml:282-288`) runs `python scripts/prepare-trusted-fonts.py --output-dir artifacts/trusted-fonts` with `OFDRW_REQUIRE_TRUSTED_FONTS: "1"`. `scripts/prepare-trusted-fonts.py:198-200` reports env `OFDRW_REQUIRE_TRUSTED_FONTS=1` and `OFDRW_TEST_TRUSTED_CJK_FONT=<destination>/fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf`.
- The trusted-font bundle is an LFS fixture under `tests/fixtures/trusted-fonts/`, which needs `git lfs pull --include="e2e/**,tests/**" --exclude=""`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F` | exit 0 |
| Without font | `env -u OFDRW_TEST_TRUSTED_CJK_FONT dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F --filter "FullyQualifiedName~DocxRevisionTests" --logger "console;verbosity=normal"` | the two CJK tests show as **Skipped** with a reason, unless a platform CJK font is installed, in which case they run and pass |
| With font | `python3 scripts/prepare-trusted-fonts.py --output-dir artifacts/trusted-fonts && OFDRW_REQUIRE_TRUSTED_FONTS=1 OFDRW_TEST_TRUSTED_CJK_FONT=$PWD/artifacts/trusted-fonts/fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F --filter "FullyQualifiedName~DocxRevisionTests"` | all pass, none skipped |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `tests/Ofdrw.Net.Converter.Docx.Tests/DocxRevisionTests.cs`, `tests/Ofdrw.Net.Converter.Docx.Tests/LibreOfficeFontStagingTests.cs`, and `.github/workflows/ci.yml` (one new step in the ubuntu branch of the `builtin-docx` job).

**Out of scope**: product code; `scripts/run-trusted-font-tests.py` (it targets the PDF test project only); other test projects.

## Git workflow

Branch `advisor/011-no-silent-test-returns`; commit e.g. `test: skip CJK and Windows-font DOCX tests visibly instead of passing empty`. No push or PR unless instructed.

## Steps

### Step 1: CJK attribute

In `DocxRevisionTests.cs`, add a nested `private sealed class CjkMetricFontFactAttribute : FactAttribute`. Its constructor sets `Skip = "No CJK metric font found; set OFDRW_TEST_TRUSTED_CJK_FONT (scripts/prepare-trusted-fonts.py) or OFDRW_REQUIRE_TRUSTED_FONTS=1 to fail instead."` when **both**:
- `OFDRW_REQUIRE_TRUSTED_FONTS` is not `"1"`, and
- `TryFindCjkMetricFont()` returns null. Make the helper `internal static` or keep it `private static`; nested classes can call it.

Wrap the call in try/catch so that discovery never throws. If it throws, do **not** skip: let the test run, and the helper throws again with its message.

Replace `[Fact]` with `[CjkMetricFontFact]` on the two tests. Replace `var metricFont = TryFindCjkMetricFont(); if (metricFont is null) return;` with `var metricFont = TryFindCjkMetricFont() ?? throw new InvalidOperationException("CJK metric font disappeared after discovery.");`.

**Verify**: "Without font" and "With font" commands behave as in the table.

### Step 2: Windows-fonts split

Split `ShouldRestrictPreferredCjkOnlyForLargePlatformFontDirectories` into:
- the original `[Fact]` keeping only the `custom` assertion;
- `[WindowsFontsFact] ShouldRestrictPreferredCjkForWindowsFontsDirectory` with the Windows assertions and no early returns. The nested `WindowsFontsFactAttribute` sets `Skip` when `!OperatingSystem.IsWindows()` or the `Windows\Fonts` directory does not exist.

**Verify**: on Linux, the new test shows as Skipped, and the original passes.

### Step 3: Run the CJK tests in CI with the trusted font

In `ci.yml`, job `builtin-docx`, after the step "Verify original font identity through native OFD and PDF" (which already prepares `artifacts/trusted-fonts`), add:
```yaml
- name: Test DOCX CJK viewer-local font declarations with the trusted fixture
  if: matrix.os == 'ubuntu-latest'
  env:
    OFDRW_REQUIRE_TRUSTED_FONTS: "1"
    OFDRW_TEST_TRUSTED_CJK_FONT: ${{ github.workspace }}/artifacts/trusted-fonts/fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf
  run: dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --no-build --no-restore --filter "FullyQualifiedName~DocxRevisionTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
```
Confirm the font path by running "With font" locally and checking the file exists.

**Verify**: YAML parses (`python3 -c "import yaml; yaml.safe_load(open('.github/workflows/ci.yml'))"`).

### Step 4: Sweep for other silent returns

`grep -rn -B2 "^\s*return;\s*$" tests --include=*.cs | grep -B1 "if (.*\(null\|Exists\|IsNullOrEmpty\)"`. List any other test methods that return before asserting because of missing tools, fonts or platform. **Do not fix them in this plan**; include the list in your report.

## Done criteria

- [ ] `grep -n "if (metricFont is null)" tests/Ofdrw.Net.Converter.Docx.Tests/DocxRevisionTests.cs` returns nothing
- [ ] `LibreOfficeFontStagingTests.cs` has no `return;` in the split test
- [ ] Both local runs behave as in the table; the CI step is added and YAML parses
- [ ] Report lists any other silent-return tests found in Step 4
- [ ] Storage gate exits 0; `plans/README.md` updated

## STOP conditions

- `scripts/prepare-trusted-fonts.py` writes the font somewhere other than `<output-dir>/fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf`. Report the actual path and adjust only the env value.
- The two CJK tests fail when run with the trusted font. That reveals a real bug previously hidden by the silent return; report it and do not change assertions.

## Maintenance notes

New tests that depend on optional fonts, tools or platforms must use a skip attribute (pattern above), never an early `return`. Reviewers should flag `return;` before the first `Assert` in test bodies.
