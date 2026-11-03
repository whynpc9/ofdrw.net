# Plan 002: Render correct DOCX list markers in Native/BuiltIn conversion

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `docs/engineering/plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs tests/Ofdrw.Net.Converter.Docx.Tests`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW
- **Depends on**: none (do not run in parallel with plans 003/004, which edit neighbouring DOCX files)
- **Category**: bug
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

DOCX→OFD defaults to **Native** mode, which reads OpenXML through `DocxModelReader`. BuiltIn DOCX→PDF uses the same reader. Its list-marker logic is wrong for ordinary Word documents:
- `numId=0` (Word's "numbering turned off", often used to cancel numbering inherited from a style) gets a fake `1.` marker.
- Counters never reset when a higher level advances, so you get 1, 1.1, 1.2, 2, **1.3**.
- A multi-level pattern such as `%1.%2.` leaves a literal `%1`.
- Every non-bullet format (Roman, letters, Chinese `一、`) renders as decimal.
- `w:start` and `w:lvlOverride` are ignored.

Wrong list markers are visible on almost every structured document (headings, legal clauses, Chinese government documents). The marker text is also written into the OFD as real text, so text extraction carries the wrong numbers too.

## Current state

- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs`
  - Line 30: `private readonly Dictionary<string, int> _numberingCounters = new(StringComparer.Ordinal);`
  - Lines 213-224, `ReadParagraph`: the marker is prepended as text:
    ```csharp
    model.Format.ListMarker = ResolveListMarker(paragraph.ParagraphProperties);
    if (!string.IsNullOrEmpty(model.Format.ListMarker))
    {
        model.Inlines.Add(new BuiltInTextModel { Text = model.Format.ListMarker + " " });
    }
    ```
  - Lines 761-800, the whole method to replace:
    ```csharp
    private string? ResolveListMarker(ParagraphProperties? properties)
    {
        NumberingProperties? numbering = null;
        foreach (var layer in GetParagraphPropertyLayers(properties))
            if (layer.NumberingProperties is not null) numbering = layer.NumberingProperties;
        var numberingId = numbering?.NumberingId?.Val?.Value;
        var level = numbering?.NumberingLevelReference?.Val?.Value ?? 0;
        if (numberingId is null) return null;
        var key = numberingId.Value.ToString(CultureInfo.InvariantCulture) + ":" + level;
        _numberingCounters.TryGetValue(key, out var count);
        count++;
        _numberingCounters[key] = count;
        var numberingPart = _mainPart.NumberingDefinitionsPart?.Numbering;
        var instance = numberingPart?.Elements<NumberingInstance>()
            .FirstOrDefault(x => x.NumberID?.Value == numberingId.Value);
        var abstractId = instance?.AbstractNumId?.Val?.Value;
        var abstractNumbering = numberingPart?.Elements<AbstractNum>()
            .FirstOrDefault(x => x.AbstractNumberId?.Value == abstractId);
        var levelDefinition = abstractNumbering?.Elements<Level>()
            .FirstOrDefault(x => x.LevelIndex?.Value == level);
        var format = levelDefinition?.NumberingFormat?.Val?.Value;
        if (format == NumberFormatValues.Bullet)
            return levelDefinition?.LevelText?.Val?.Value ?? "•";
        var pattern = levelDefinition?.LevelText?.Val?.Value ?? "%1.";
        return pattern.Replace("%" + (level + 1), count.ToString(CultureInfo.InvariantCulture));
    }
    ```
    (Braces are condensed here; the live code has the same statements.)
  - Diagnostics pattern: add warnings directly, **not** through `ReportUnsupported` (line 893), which throws when `UnsupportedFeatureBehavior == Throw`. Exemplar at line 151:
    ```csharp
    _diagnostics.Add(new DocxConversionDiagnostic("DOCX_SUPPLEMENTAL_TEXT_APPENDED",
        "BuiltIn preserves footnote, endnote and comment bodies as labeled content after the main body.",
        DocxConversionDiagnosticSeverity.Information));
    ```
- OpenXML SDK 3.5.1: `NumberFormatValues` values are static properties, not constants. Compare with `==` in `if` chains; `switch` `case NumberFormatValues.X:` does not compile.
- No existing test touches numbering (`grep -rln "numPr\|NumberingDefinitionsPart" tests` is empty).
- Test helper: `Ofdrw.Net.TestSupport.DocxReleaseFixtures.CreateMinimalDocx(string bodyContent, ...)` (`tests/TestSupport/DocxReleaseFixtures.cs:13`) builds a DOCX from body XML. Parts are added afterwards like this (same file, lines 133-137):
  ```csharp
  var input = CreateMinimalDocx(body);
  using (var document = WordprocessingDocument.Open(input, true))
      document.MainDocumentPart!.AddImagePart(ImagePartType.Jpeg, "rId2").FeedData(jpeg);
  input.Position = 0;
  ```
  Use `document.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>().FeedData(new MemoryStream(Encoding.UTF8.GetBytes(numberingXml)))` the same way.
- Test conventions: `tests/Ofdrw.Net.Converter.Docx.Tests/DocxConversionTests.cs` (namespace `Ofdrw.Net.Converter.Docx.Tests`; it uses `DocxToOfdConverter`, `OfdReader`, `OfdTextExtractor`). Model the new test file on `Native_WrappedSeparatorSpaces_ShouldNotShiftVisibleAlignment` (lines 367-393), which converts with both explicit Native and the default converter.

## Required behavior (implement exactly this)

1. Effective `numPr` is resolved as today (last layer wins).
2. `numId` null **or 0** → no marker.
3. Build lookup dictionaries lazily, once per reader: `numId → NumberingInstance` and `abstractNumId → AbstractNum`. If the instance or its abstract definition is missing → no marker, plus one `DOCX_NUMBERING_UNRESOLVED` warning per distinct numId (message names the numId).
4. Level definition for `ilvl` L (0–8): if the instance has `w:lvlOverride` with `w:ilvl=L` containing a `w:lvl`, use that `Level`; otherwise the abstract's `Level` with `LevelIndex == L`. A missing level definition → no marker, plus `DOCX_NUMBERING_UNRESOLVED`.
5. Start value for level L = `lvlOverride/startOverride` if present, else `lvl/start`, else 1.
6. Counters are kept **per numId** (this matches the current keying; see Maintenance notes) as an array of 9 nullable ints. On a paragraph at level L: if `counter[L]` is null, set it to the start value, else increment it. Then set `counter[L+1..8] = null`.
7. Bullet (`numFmt == bullet`) → return `lvlText` or `"•"` (unchanged behavior). Counters are still updated per rule 6.
8. `numFmt == none` → return the pattern with all placeholders replaced by "" (usually empty).
9. Otherwise take `lvlText` (default `"%{L+1}."`). For each k = 1..9 replace `%k` with level k-1's counter (if null, that level's start value), formatted with **that level's** `numFmt`:
   - `decimal` (or missing) → `1`
   - `decimalZero` → `01`, `09`, `10`
   - `lowerRoman` / `upperRoman` → standard Roman numerals 1-3999; outside that range, decimal
   - `lowerLetter` / `upperLetter` → Word style: 1→a … 26→z, 27→aa, 28→bb … 52→zz, 53→aaa (the letter repeats ⌈n/26⌉ times)
   - `chineseCounting` and `chineseCountingThousand` (but NOT `chineseLegalSimplified`, which goes to the "any other format" rule): 1-99 as 一 二 三 四 五 六 七 八 九 十, 十一 … 十九, 二十, 二十一 … 九十九. ≥100 falls back to decimal, plus one `DOCX_NUMBERING_FORMAT_DEGRADED` warning per (numId, level).
   - `ideographDigital` → each decimal digit mapped to 〇一二三四五六七八九
   - Any other format → decimal, plus one `DOCX_NUMBERING_FORMAT_DEGRADED` warning per (numId, level, format)
   Replace `%k` for k = 9 down to 1 so that a hypothetical `%10` is never a concern. Only `%1`–`%9` exist in OOXML.
10. All diagnostics use `DocxConversionDiagnosticSeverity.Warning` and the `_diagnostics.Add(...)` pattern above. Never throw for numbering.

Put the formatting helpers in a new `internal static class DocxNumberFormatter` in `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxNumberFormatter.cs` (pure functions, unit-testable).

## Commands you will need

Once per shell, from the repo root:
```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

| Purpose | Command | Expected |
|---|---|---|
| Build | `dotnet build Ofdrw.Net.sln -c Release $F` | exit 0 |
| New tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F --filter "FullyQualifiedName~DocxListNumberingTests"` | all pass |
| DOCX suite | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests -c Release $F` | all pass (LibreOffice-dependent tests may need `soffice`; record environment failures) |
| Layout/PDF suites | `dotnet test tests/Ofdrw.Net.Layout.Tests -c Release $F` ; `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests -c Release $F` | all pass |
| CLI | `dotnet run --project src/Ofdrw.Net.Cli -c Release -- docx-to-ofd -i IN.docx -o OUT.ofd [--docx-ofd-mode native]` ; `... ofd-to-pdf -i OUT.ofd -o OUT.pdf` | exit 0 |
| PNG | `pdftoppm -r 144 -png OUT.pdf PREFIX` | PNG per page |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py --require-materialized e2e --require-materialized tests` | exit 0 |

## Scope

**In scope**:
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs`: `ResolveListMarker`, the `_numberingCounters` field and new lookup fields only
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxNumberFormatter.cs` (create)
- `tests/Ofdrw.Net.Converter.Docx.Tests/DocxListNumberingTests.cs` (create)
- `scripts/validation/docx-lists/generate-sample.py` (create; deterministic visual sample generator, output under `artifacts/`)
- `docs/content/reference/conversion-contracts.md`: one bullet under "## 原始 DOCX 文本与分页"

**Out of scope**:
- `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` and `scripts/generate-docx-test-fixture.py`. They are shared LFS fixtures with recorded evidence; do not regenerate them.
- Marker indentation, tab after the marker (`w:suff`), `w:lvlRestart`, `w:isLgl`, picture bullets.
- `BuiltInOfdRenderer.cs` / `BuiltInPdfRenderer.cs`.

## Git workflow

- Branch `advisor/002-docx-list-numbering`; commit e.g. `fix(docx): correct list numbering levels, formats and numId=0`.
- Do not push or open a PR unless instructed.

## Steps

### Step 1: Write the failing tests

Create `DocxListNumberingTests.cs` with a helper `static MemoryStream Docx(string body, string numberingXml)`. Use `CreateMinimalDocx(body)`, add the numbering part as shown above, and rewind. Numbering XML root: `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">…</w:numbering>`. Helper `static async Task<string> NativeText(Stream docx, bool explicitNative)` converts and returns `string.Join("\n", new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(ofd), includeTemplates: true))`.

Cases. Each is a `[Theory]` over `explicitNative` true/false, or loops both. Assert with `Assert.Contains`. Paragraph text is unique, so match `marker + " " + text`:
1. **Multi-level reset and pattern**: abstractNum 1, lvl0 `decimal` `%1.`, lvl1 `decimal` `%1.%2.`; paragraphs at ilvl 0,1,1,0,1 with texts A..E → `1. A`, `1.1. B`, `1.2. C`, `2. D`, `2.1. E`.
2. **numId=0 suppresses**: a paragraph with `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>` and text `Plain` → the extracted text contains `Plain` and does not contain `1. Plain`.
3. **Unresolved numId**: numId 7 not defined → no marker, and the result diagnostics (`ConvertWithResultAsync`) contain code `DOCX_NUMBERING_UNRESOLVED`.
4. **Formats**: lvl0 `upperRoman` `%1.` (4 paragraphs → `I.`..`IV.`); a second num with `lowerLetter` `(%1)` and a `w:start w:val="26"` → `(z)`, `(aa)`; a third with `chineseCounting` `%1、` → `一、`, `二、`.
5. **startOverride**: num 3 → abstract 1 with `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>` → first marker `5.`.
6. **Bullet unchanged**: `bullet` with lvlText `•` → `• Item`.

Also add direct unit tests for `DocxNumberFormatter` (Roman 1/4/9/14/40/90/400/1994/3999; letters 1/26/27/52/53; Chinese 1/10/11/19/20/21/99; decimalZero 1/10). Since the class does not exist yet, write these after Step 2 or temporarily comment them out.

**Verify**: the filter command → cases 1, 2, 3, 4 and 5 fail on the current code (record the failures). Case 6 may pass.

### Step 2: Implement `DocxNumberFormatter`

Pure static methods: `Format(int value, NumberFormatValues? format, out bool degraded)` plus `ToRoman`, `ToLetters`, `ToChineseCounting`. Use `CultureInfo.InvariantCulture` for decimal output.

**Verify**: build exits 0; the formatter unit tests pass.

### Step 3: Rewrite `ResolveListMarker` per "Required behavior"

Replace `_numberingCounters` with `Dictionary<int, int?[]>` keyed by numId. Add lazily built `Dictionary<int, NumberingInstance>?` and `Dictionary<int, AbstractNum>?` fields. Keep the method signature `private string? ResolveListMarker(ParagraphProperties? properties)` so `ReadParagraph` stays unchanged. Use `HashSet<string>` keys to emit each diagnostic only once.

**Verify**: the filter command → all `DocxListNumberingTests` pass. The full DOCX suite → passes.

### Step 4: Visual acceptance (required by AGENTS.md)

Write `scripts/validation/docx-lists/generate-sample.py`, a standard-library-only script (`zipfile`) that writes a deterministic DOCX to the path in `argv[1]`. Content:
- A heading.
- A 3-level legal list (`%1.`, `%1.%2.`, `%1.%2.%3.`) with mixed Chinese/English text, including one **bold** and one *italic* run.
- An upperRoman list, a `chineseCounting` `%1、` list, a bullet list, and one paragraph with `numId=0`.
- Enough items (≥ 60 paragraphs) to force a second page.

Use fixed ZIP timestamps (`ZipInfo(date_time=(2026,1,1,0,0,0))`). Then:
```sh
mkdir -p artifacts/visual/002
python3 scripts/validation/docx-lists/generate-sample.py artifacts/visual/002/lists.docx
CLI="dotnet run --project src/Ofdrw.Net.Cli -c Release --"
$CLI docx-to-ofd -i artifacts/visual/002/lists.docx -o artifacts/visual/002/native.ofd --docx-ofd-mode native
$CLI docx-to-ofd -i artifacts/visual/002/lists.docx -o artifacts/visual/002/default.ofd
$CLI ofd-to-pdf -i artifacts/visual/002/native.ofd -o artifacts/visual/002/native.pdf
$CLI ofd-to-pdf -i artifacts/visual/002/default.ofd -o artifacts/visual/002/default.pdf
pdftoppm -r 144 -png artifacts/visual/002/native.pdf artifacts/visual/002/native
pdftoppm -r 144 -png artifacts/visual/002/default.pdf artifacts/visual/002/default
$CLI extract-text artifacts/visual/002/native.ofd artifacts/visual/002/native.txt
sha256sum artifacts/visual/002/* > artifacts/visual/002/SHA256SUMS
```
Open every PNG page and check:
- markers match the expected sequence
- no literal `%`
- bold/italic does not spread to the marker or the next paragraph
- no overlap, clipping or blank pages
- Chinese characters render (if a CJK font is installed; otherwise record the limitation)

Check `native.txt` contains every source paragraph's text.

**Verify**: PNG files exist for both modes; a written checklist of pages reviewed and observations is in your report.

### Step 5: Docs

Add one bullet to `docs/content/reference/conversion-contracts.md` under "## 原始 DOCX 文本与分页". In the existing tone, say that Native/BuiltIn list markers support multi-level counters with level reset, `%1`–`%9` patterns, `start`/`startOverride`, and the listed formats. Say that `numId=0` disables numbering, that unresolved definitions or unsupported formats return `DOCX_NUMBERING_UNRESOLVED` / `DOCX_NUMBERING_FORMAT_DEGRADED`, and that the behavior is not yet released (main only).

**Verify**: `git diff --stat` shows only in-scope files.

## Test plan

Covered in Step 1: 6 conversion cases × 2 modes, plus formatter unit tests. Use `DocxConversionTests.Native_WrappedSeparatorSpaces_ShouldNotShiftVisibleAlignment` as the structural pattern.

## Done criteria

- [ ] Build exits 0; `DocxListNumberingTests` pass; full `Ofdrw.Net.Converter.Docx.Tests` passes (environment-only failures listed)
- [ ] `grep -n "_numberingCounters\[key\]" src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/DocxModelReader.cs` returns nothing (old keying removed)
- [ ] Visual record in `artifacts/visual/002/` (not committed) with PNG review notes for both explicit native and default modes. macOS Preview acceptance reported as **not done** unless actually performed on macOS.
- [ ] Storage gate exits 0
- [ ] `docs/engineering/plans/README.md` row updated

## STOP conditions

- `ResolveListMarker` or `ReadParagraph` differ from the excerpts.
- Any existing DOCX test that previously passed now fails because its expected text included an old (wrong) marker. Report it; do not edit its expectation without approval.
- Satisfying a case seems to require changing `BuiltInOfdRenderer` or the Layout project.

## Maintenance notes

- Counters are per numId, matching previous behavior. Word continues numbering across different `w:num` instances that share an `abstractNum` (unless restarted via `startOverride`). Supporting that is a follow-up that changes markers in existing documents, so it needs its own decision.
- Indentation and the tab suffix are not rendered; the marker is plain text plus a space.
- Plans 003/004 touch the same converter; land them one after another and re-run this plan's tests after each.
