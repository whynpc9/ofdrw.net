# Plan 003: Draw DOCX underlines in Native DOCX→OFD output

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs src/Ofdrw.Net.Layout/FlowDocument.cs src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S
- **Risk**: MED (changes page output; needs visual re-acceptance)
- **Depends on**: none. Run after plan 002 if both are scheduled, because they share the DOCX test project.
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Native mode is the default for DOCX→OFD. The DOCX reader already parses `<w:u>` into `BuiltInTextFormat.Underline`, and the BuiltIn **PDF** renderer draws it (`BuiltInPdfRenderer.cs:303`). The **OFD** renderer never reads the flag, so underlined text (hyperlinks, fill-in blanks such as `姓名：________`, legal emphasis) silently loses its underline, with no diagnostic. The public `FlowDocument` already knows how to draw underlines as OFD paths; this plan reuses that exact geometry for Native DOCX.

## Current state

- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs:666-669` parses underline:
  ```csharp
  if (properties.Underline?.Val?.Value is UnderlineValues underline)
  {
      target.Underline = underline != UnderlineValues.None;
  }
  ```
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs:539-544`: the `Style` local function does **not** copy underline:
  ```csharp
  FlowTextStyle Style(BuiltInTextFormat format) => new()
  {
      FontFamily = NormalizeFontFamily(format.FontFamily),
      FontSizeMillimeters = PointsToMillimeters(format.FontSizePoints ?? DefaultFontSizePoints),
      Bold = format.Bold, Italic = format.Italic, Source = format
  };
  ```
- `BuiltInOfdRenderer.cs:584-625`, `DrawLine`: groups glyphs by `ReferenceEquals(line.Glyphs[i].Style.Source, format)`, emits one `OfdTextElement` per group, then `x += groupWidth;`. Key locals: `x`, `top`, `baseline` (= `line.BaselineMillimeters`), `groupWidth`, `format`, and `ParseColor(format.ColorHex)` (static helper at line 735, returns `OfdColor`).
- The exemplar to copy is `src/Ofdrw.Net.Layout/FlowDocument.cs:487-498`:
  ```csharp
  if (style.Underline && groupWidth > 0)
  {
      var underlineY = line.BaselineMillimeters + style.UnderlineOffsetMillimeters;
      _page.Elements.Add(new OfdPathElement
      {
          LayerType = "Body", XMillimeters = x, YMillimeters = top ?? _y,
          WidthMillimeters = groupWidth, HeightMillimeters = line.Height,
          Stroke = true, Fill = false, LineWidthMillimeters = style.UnderlineWidthMillimeters,
          StrokeColor = (OfdColor)style.Source!,
          AbbreviatedData = FormattableString.Invariant($"M 0 {underlineY:0.######} L {groupWidth:0.######} {underlineY:0.######}")
      });
  }
  ```
  The default geometry from `FlowDocument.cs:438-439` is `underlineOffset = size * 0.1d`, `underlineWidth = size * 0.06d` (size = font size in mm).
- `FlowTextStyle` (`src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs:13-25`) already has `Underline`, `UnderlineOffsetMillimeters`, `UnderlineWidthMillimeters`. The DOCX project can see these internal types; the renderer already uses them.
- `DrawLine` is used for body (line 278), table cells (335) and headers/footers (844, 846), so the change covers all of them.
- Test conventions: see `tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs:367-393` (`Native_WrappedSeparatorSpaces_ShouldNotShiftVisibleAlignment`), which builds a DOCX via `CreateMinimalDocx(body)`, converts with explicit Native and the default converter, and reads back with `OfdReader`.

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
| New tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F --filter "FullyQualifiedName~DocxUnderlineTests"` | pass |
| DOCX suite | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F` | pass (record environment-only failures) |
| PDF suite | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F` | pass |
| CLI | `dotnet run --project src/Ofdrw.Net.Cli -c Release -- docx-to-ofd -i IN -o OUT [--docx-ofd-mode native]` and `... ofd-to-pdf -i OUT.ofd -o OUT.pdf` | exit 0 |
| PNG | `pdftoppm -r 144 -png OUT.pdf PREFIX` | PNGs |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests` | exit 0 |

## Scope

**In scope**:
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs`: `Style` local function and `DrawLine`
- `tests/Ofdrw.Net.Converter.Docx.Tests/DocxUnderlineTests.cs` (create)
- `scripts/validation/docx-underline/generate-sample.py` (create)
- `docs/content/reference/conversion-contracts.md`: one bullet

**Out of scope**:
- Justified alignment (plan 004).
- Double/dotted/wavy underline styles (render all non-`none` values as a single solid line, as the PDF renderer does), strike-through, underline colour (`w:u w:color`).
- `FlowDocument.cs`. Copy its geometry; do not refactor it into a shared helper in this plan (see Maintenance notes).
- `BuiltInPdfRenderer.cs`.

## Git workflow

Branch `advisor/003-native-docx-underline`; commit e.g. `fix(docx): draw underlines in Native OFD output`. No push or PR unless instructed.

## Steps

### Step 1: Failing test

`DocxUnderlineTests.cs`. Body:
```xml
<w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Under</w:t></w:r><w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>NoLine</w:t></w:r></w:p>
```
For explicit Native and default: convert, read back with `OfdReader`, and on page 1:
- the text element whose `Text == "Under"` exists
- exactly one `OfdPathElement` has `Stroke == true`, `XMillimeters` equal (±0.001) to that text's `XMillimeters`, and `WidthMillimeters` equal (±0.001) to its `WidthMillimeters`
- no stroked path has an X matching the "Plain " or "NoLine" elements

A second `[Fact]`: an underlined run that wraps across two lines in a narrow page (copy the `sectPr` from `DocxConversionTests.cs:377`) → two underline paths, one per line.

**Verify**: the new tests fail (no path elements found).

### Step 2: Implement

1. In `Style(...)`, set `Underline = format.Underline`, `UnderlineOffsetMillimeters = size * 0.1d`, `UnderlineWidthMillimeters = size * 0.06d`, where `size` is the same mm font size already computed. Introduce a local variable instead of computing it twice.
2. In `DrawLine`, immediately after `page.Elements.Add(text);` and before `x += groupWidth;`, add the underline path. Use the FlowDocument block above with these substitutions:
   - `top` for `top ?? _y`
   - `ParseColor(format.ColorHex)` for the stroke colour
   - offset/width from `line.Glyphs[start].Style`
   - the condition `format.Underline && groupWidth > 0`
   Use `FormattableString.Invariant` exactly as shown.
3. Do not underline the zero-width trailing separators at automatic wraps. They have width 0, so `groupWidth` already excludes them; nothing to do. Mention it in the code comment.

**Verify**: new tests pass; the full DOCX and PDF suites pass.

### Step 3: Visual acceptance (AGENTS.md)

Write `scripts/validation/docx-underline/generate-sample.py`: standard library `zipfile`, fixed timestamps `(2026,1,1,0,0,0)`, output path `argv[1]`. Content:
- mixed Chinese/English paragraphs with underlined words mid-line
- an underlined **bold** run and an underlined red (`w:color 'C00000'`) run
- an underlined run that wraps across a line
- a fill-in line `姓名：` + underlined spaces (`xml:space="preserve"`)
- a 2×2 table with an underlined cell word, and a footer containing underlined text
- enough content for 2 pages

Then run the same pipeline as other DOCX plans:
```sh
mkdir -p artifacts/visual/003
python3 scripts/validation/docx-underline/generate-sample.py artifacts/visual/003/underline.docx
CLI="dotnet run --project src/Ofdrw.Net.Cli -c Release --"
$CLI docx-to-ofd -i artifacts/visual/003/underline.docx -o artifacts/visual/003/native.ofd --docx-ofd-mode native
$CLI docx-to-ofd -i artifacts/visual/003/underline.docx -o artifacts/visual/003/default.ofd
for m in native default; do $CLI ofd-to-pdf -i artifacts/visual/003/$m.ofd -o artifacts/visual/003/$m.pdf; pdftoppm -r 144 -png artifacts/visual/003/$m.pdf artifacts/visual/003/$m; done
$CLI extract-text artifacts/visual/003/native.ofd artifacts/visual/003/native.txt
sha256sum artifacts/visual/003/* > artifacts/visual/003/SHA256SUMS
```
Review each PNG page, zooming on underlined words. Check:
- the line sits just below the baseline, does not cut descenders badly, and matches the text colour
- it does not extend past the run or spread to neighbouring runs
- there is no duplicate text
- the table and footer are underlined
- the fill-in blank shows a line

Check `native.txt` has the original text, and that the PDF has no duplicated text layer (`pdftotext artifacts/visual/003/native.pdf - | head` shows each word once).

**Verify**: PNGs exist; the review checklist is in the report.

### Step 4: Docs

Add one bullet in `docs/content/reference/conversion-contracts.md` under "## 原始 DOCX 文本与分页": Native draws a single solid underline path for runs whose `w:u` is not `none` (style variants and underline colour are not distinguished; main only, not yet released).

### Final step: Full regression and local package-consumption E2E (AGENTS.md)

This plan changes the shared read/layout/font/conversion chain. AGENTS.md therefore requires, in addition to the targeted suites above, the **full regression** and a **local package-consumption E2E**. Converting through the source CLI does not exercise the packaged libraries. Run from the repo root, after all other steps:
```sh
dotnet build Ofdrw.Net.sln -c Release $F
dotnet test Ofdrw.Net.sln -c Release --no-build $F --logger trx --results-directory artifacts/003/full-tests
python3 scripts/run-trusted-font-tests.py --no-build --results-directory artifacts/003/trusted-font-tests
rm -rf artifacts/003/packages   # pack-only requires an empty candidate directory
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only --no-build --packages-dir artifacts/003/packages
./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --consume-only --packages-dir artifacts/003/packages --output-dir artifacts/003/package-e2e
```

**Verify**: every command exits 0. A failure caused only by a missing external tool (LibreOffice, Poppler, fonts) must be shown to fail identically on the base commit (`git worktree add ../ofdrw-base 09a2e7e`), and recorded in the report as an environment limitation; it does not count as a pass. Open the rendered pages under `artifacts/003/package-e2e` for the samples this plan affects, and record which pages you reviewed. Report the TRX totals (total/passed/skipped/failed) for both test runs.

## Test plan

Two facts × two modes as in Step 1; existing suites as regression.

## Done criteria

- [ ] Full `dotnet test Ofdrw.Net.sln`, `run-trusted-font-tests.py` and the local package-consumption E2E (pack + consume) were run per the Final step; TRX totals and environment-only failures are recorded
- [ ] Build exits 0; `DocxUnderlineTests` pass; DOCX and PDF suites pass
- [ ] `grep -n "Underline" src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs` shows the new code
- [ ] `artifacts/visual/003/` contains PNGs for both modes and the review notes are in the report. macOS Preview reported as **not done** unless performed.
- [ ] Storage gate exits 0; `docs/engineering/plans/README.md` updated

## STOP conditions

- `DrawLine`/`Style` differ from the excerpts.
- An existing test asserts an exact count of `OfdPathElement`s on a page that contains underlined text and now fails. Report the test name; do not change its expectation unilaterally.
- Underlines visibly overlap the next line or descenders in the PNG review. Report with the page PNG; do not tune constants without approval, because FlowDocument uses the same constants.

## Maintenance notes

- `FlowDocument.DrawLine` and `BuiltInOfdRenderer.DrawLine` are near-duplicates, and that duplication is why this bug existed. A later refactor should extract a shared Flow line emitter in `Ofdrw.Net.Layout`. Doing it here would change public FlowDocument output, so it is deferred.
- Plan 004 changes glyph widths in justified lines. Underline width is computed from the same `groupWidth`, so it will follow automatically; re-run these tests after 004.
