# Plan 004: Honor justified (`w:jc both/distribute`) paragraphs in Native DOCX→OFD

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs`
> If any in-scope file changed since this plan was written (other than by plan 003's underline change), compare the "Current state" excerpts against the live code before proceeding; on a mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: MED (changes line geometry of every justified paragraph; visual re-acceptance needed)
- **Depends on**: plans/003-native-docx-underline.md (same function, `DrawLine`)
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Justified text (`<w:jc w:val="both"/>`) is the default body alignment in most Chinese official and business documents. The reader maps it to `BuiltInParagraphAlignment.Justify`, and the BuiltIn PDF renderer honors it (`BuiltInPdfRenderer.cs:235`). The Native OFD renderer silently turns it into left alignment, so every justified paragraph gets a ragged right edge and looks different from the source document, with no diagnostic.

## Current state

- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs:600-603` maps `JustificationValues.Both` and `JustificationValues.Distribute` to `BuiltInParagraphAlignment.Justify`.
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs:570-575` drops it:
  ```csharp
  var alignment = paragraph.Format.Alignment switch
  {
      BuiltInParagraphAlignment.Center => FlowAlignment.Center,
      BuiltInParagraphAlignment.Right => FlowAlignment.Right,
      _ => FlowAlignment.Left
  };
  ```
- `src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs`:
  - Line 11: `internal enum FlowAlignment { Left, Center, Right }`
  - Lines 59-92, `FlowLine`: constructor `FlowLine(List<FlowGlyph> glyphs, double height, double indent, FlowAlignment alignment, bool pageBreak = false, bool terminalNewline = false, bool automaticWrap = false)`. It stores `Glyphs`, `Height`, `Indent`, `Alignment`, `PageBreak`, `TerminalNewline`, `AlignmentWidth` (sum of glyph widths). `automaticWrap` is **not stored**. When it is true, trailing whitespace glyphs are replaced with zero-width glyphs.
  - Lines 228 and 240: `Flush(automaticWrap: true)` is called for lines ended by wrapping. Explicit `\n` breaks and the paragraph's final line use `automaticWrap: false`.
  - Lines 265-271, `Align(left, width, contentWidth, alignment)`: the `_ => left` arm means a new enum member automatically aligns left.
  - `FlowGlyph(string text, FlowTextStyle style, double width, object? image = null, double imageHeight = 0)` is immutable, with an internal constructor.
- `BuiltInOfdRenderer.DrawLine` (line 584) computes `x` via `FlowParagraphLayout.Align(left + line.Indent, width - line.Indent, line.AlignmentWidth, line.Alignment)` and emits each glyph's `Width` into the OFD `DeltaX`. So widening glyph widths is all that justification needs.
- `FlowAlignment` is also used by the public `FlowDocument` (`src/Ofdrw.Net.Layout/FlowDocument.cs`). Adding an enum member must not change FlowDocument output: FlowDocument never produces the new value, so it doesn't.
- Test exemplar: `tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs:367-393`. Page geometry there: `<w:pgSz w:w="1800" w:h="6000"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/>`. The right content edge is `(1800-200)/20 pt` = 80 pt = `80 * 25.4 / 72` mm.

## Required behavior

1. Add `Justify` to `FlowAlignment`, appended last: `{ Left, Center, Right, Justify }`.
2. Store `automaticWrap` on `FlowLine` as `internal bool AutomaticWrap { get; }`.
3. The renderer maps `BuiltInParagraphAlignment.Justify` → `FlowAlignment.Justify`.
4. In `BuiltInOfdRenderer.DrawLine`, when `line.Alignment == FlowAlignment.Justify && line.AutomaticWrap`:
   - `available = width - line.Indent`; `slack = available - line.AlignmentWidth`. If `slack <= 0.001` → no change.
   - **Gaps**: the indexes of visible glyphs (non-image, `Width > 0`) excluding the last visible glyph.
     - If any of them is a whitespace glyph (`Text.All(char.IsWhiteSpace)` and not `FlowParagraphLayout.IsNonbreakingSpace`), the gap set is those whitespace glyphs only.
     - Otherwise (typical CJK lines) it is every visible glyph except the last.
   - If the gap set is empty → no change.
   - Build a new glyph list where each gap glyph is replaced by `new FlowGlyph(g.Text, g.Style, g.Width + slack / gapCount, g.Image, g.ImageHeight)`. Then lay out exactly as today, starting at `x = left + line.Indent`; the content now exactly fills the available width.
   - Lines that are not automatic wraps (the final line, or a line ended by `\n`) stay left-aligned. This matches Word `both`. `distribute` also stretches the last line in Word; that is out of scope (see Maintenance notes).
5. Do not change `FlowParagraphLayout.Layout` line breaking.

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
| New tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F --filter "FullyQualifiedName~DocxJustifyTests"` | pass |
| Suites | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Layout.Tests -c Release $F`; `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests` | exit 0 |

## Scope

**In scope**: `src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs` (enum member + `AutomaticWrap` property only); `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs` (alignment mapping and `DrawLine`); `tests/Ofdrw.Net.Converter.Docx.Tests/DocxJustifyTests.cs` (create); `scripts/validation/docx-justify/generate-sample.py` (create); one bullet in `docs/content/reference/conversion-contracts.md`.

**Out of scope**: the public `FlowDocument` API (no public Justify option in this plan); `BuiltInPdfRenderer.cs`; line-breaking logic; `distribute` last-line stretching; CJK punctuation compression.

## Git workflow

Branch `advisor/004-native-docx-justify`; commit e.g. `fix(docx): justify wrapped lines in Native OFD output`. No push or PR unless instructed.

## Steps

### Step 1: Failing tests (`DocxJustifyTests.cs`)

Use the narrow page geometry above. `rightEdge = 80 * 25.4 / 72`.
1. `Latin_JustifiedParagraph_ShouldFillWrappedLines` (explicit Native and default): `<w:jc w:val="both"/>` with a 30-word English sentence. Group the read-back `OfdTextElement`s by `YMillimeters`. For every line except the last, `max(X + Width)` is within 0.01 mm of `rightEdge`. The last line's `max(X + Width)` is `< rightEdge - 1`. Concatenated text equals the source text, so original text is preserved.
2. `Cjk_JustifiedParagraph_ShouldFillWrappedLines`: the same with 80 Chinese characters and no spaces.
3. `ExplicitBreak_ShouldNotBeStretched`: `both` paragraph `Alpha` `<w:br/>` `Beta`. Both lines stay left-aligned (X equals the left content edge `10 * 25.4 / 72`, ±0.001).
4. `LeftAlignment_Unchanged`: the same text without `jc` gives identical element X/Width values as before the change. Compute the expected values on the base commit first and hard-code them, or assert that the lines are not filled to `rightEdge`.

**Verify**: tests 1 and 2 fail; 3 and 4 pass.

### Step 2: Implement "Required behavior" items 1-4

**Verify**: build exits 0; new tests pass; all three suites pass. If plan 003 has landed, its `DocxUnderlineTests` still pass.

### Step 3: Visual acceptance (AGENTS.md)

Write `scripts/validation/docx-justify/generate-sample.py` (stdlib `zipfile`, fixed timestamps, output `argv[1]`). Content:
- 3 justified paragraphs in English, Chinese and mixed text, with **bold** and *italic* runs and one underlined run
- one centered and one right-aligned paragraph for contrast
- a justified paragraph inside a table cell
- enough text for 2 pages

Run:
```sh
mkdir -p artifacts/visual/004
python3 scripts/validation/docx-justify/generate-sample.py artifacts/visual/004/justify.docx
CLI="dotnet run --project src/Ofdrw.Net.Cli -c Release --"
$CLI docx-to-ofd -i artifacts/visual/004/justify.docx -o artifacts/visual/004/native.ofd --docx-ofd-mode native
$CLI docx-to-ofd -i artifacts/visual/004/justify.docx -o artifacts/visual/004/default.ofd
for m in native default; do $CLI ofd-to-pdf -i artifacts/visual/004/$m.ofd -o artifacts/visual/004/$m.pdf; pdftoppm -r 144 -png artifacts/visual/004/$m.pdf artifacts/visual/004/$m; done
$CLI extract-text artifacts/visual/004/native.ofd artifacts/visual/004/native.txt
sha256sum artifacts/visual/004/* > artifacts/visual/004/SHA256SUMS
```
Also convert `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` the same way into `artifacts/visual/004/layout-*`, and compare against base-commit renders. Changes are expected only in justified paragraphs.

Review every page. Check:
- the right edges of wrapped justified lines align
- last lines are ragged-left
- no excessive gaps on short lines
- no overlap or clipping, no text beyond the margin
- underline (plan 003) follows the stretched words
- English word spacing looks natural, Chinese is evenly spaced

**Verify**: PNGs exist; the checklist is in the report.

### Step 4: Docs

Add one bullet to `docs/content/reference/conversion-contracts.md`: Native distributes the slack of automatically wrapped lines in `both`/`distribute` paragraphs, across spaces or (for lines without spaces) between characters. Last lines and explicit breaks are not stretched. Main only.

## Done criteria

- [ ] Build exits 0; `DocxJustifyTests` pass; Docx, Layout and Pdf suites pass
- [ ] `grep -n "Justify" src/Ofdrw.Net.Layout/Internal/Flow/FlowParagraphLayout.cs src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs` shows the enum member and the mapping
- [ ] Visual record for the sample and `generated-layout.docx`, both modes. macOS Preview reported as **not done** unless performed.
- [ ] Storage gate exits 0; `plans/README.md` updated

## STOP conditions

- Any Layout test (public `FlowDocument`) changes output. FlowDocument must be unaffected.
- An existing test with hard-coded coordinates for a `both`-justified paragraph fails. Report it; do not update expectations without approval.
- The PNG review shows lines exceeding the right margin.

## Maintenance notes

- Exposing `Justify` on the public `FlowDocument` paragraph API is a cheap follow-up, now that the layout primitive exists.
- Word `distribute` also stretches the last line, and Word compresses CJK punctuation. Neither is implemented.
- Keyword location and text replacement read `DeltaX`. Justified lines now carry larger inter-word deltas, so watch `KeywordLocator*` tests if they use DOCX-produced fixtures.
