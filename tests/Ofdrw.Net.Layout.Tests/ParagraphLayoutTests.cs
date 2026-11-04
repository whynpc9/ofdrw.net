using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Layout.Text.Unicode;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class ParagraphLayoutTests
{
    private static readonly Lazy<TextFontFace> NotoFace = new(() => Face("Noto Sans", "NotoSans-Regular.ttf"));
    private static readonly Lazy<TextFontFace> WenKaiFace = new(() => Face("LXGW WenKai", "LXGWWenKai-Regular.ttf"));
    private static readonly Lazy<TextFontFace> DejaVuFace = new(() => Face("R22 DejaVu Sans", "DejaVuSans-R22.ttf"));
    private static readonly Lazy<TextFontFace> TamilFace = new(() => Face("Noto Sans Tamil", "NotoSansTamil-Regular.ttf"));

    internal static TextStyle Style(double size = 4, TextPaint? paint = null) => new(new[]
    {
        new TextFontFamily("Noto Sans", NotoFace.Value),
        new TextFontFamily("LXGW WenKai", WenKaiFace.Value),
        new TextFontFamily("DejaVu Sans", DejaVuFace.Value),
        new TextFontFamily("Noto Sans Tamil", TamilFace.Value),
    }, size, paint);

    private static ParagraphLayoutResult Layout(string text, double width = 80, TextAlignment alignment = TextAlignment.Start,
        TextParagraphDirection direction = TextParagraphDirection.Auto, double? pageHeight = null,
        TextOverflowPolicy overflow = TextOverflowPolicy.Reject, TextStyle? style = null) =>
        HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, style ?? Style(),
            new ParagraphLayoutOptions(width, pageHeight, direction, alignment, overflow: overflow)));

    [Fact]
    public void FallbackSelectsCoveringFacesAndEveryCodeUnitHasOneRun()
    {
        const string text = "Hello 世界, שלום עולם and مرحبا بالعالم. தமிழ் 2026!\nSecond paragraph.";
        var result = Layout(text, 200);
        AssertCoverage(result);
        Assert.Equal(2, result.Lines.Count);
        Assert.Equal("17.0.0", result.UnicodeVersion);
        Assert.StartsWith("HarfBuzzSharp/", result.EngineVersion);
        var runs = result.Lines.SelectMany(line => line.Runs).ToArray();
        Assert.Contains(runs, run => run.Text.Contains("世界") && run.Face == WenKaiFace.Value && run.Script == "Hani");
        Assert.Contains(runs, run => run.Text.Contains("שלום") && run.Face == DejaVuFace.Value &&
            run.Direction == TextDirection.RightToLeft && run.BidiLevel == 1);
        Assert.Contains(runs, run => run.Text.Contains("مرحبا") && run.Face == DejaVuFace.Value && run.Script == "Arab");
        Assert.Contains(runs, run => run.Text.Contains("தமிழ்") && run.Face == TamilFace.Value && run.Script == "Taml");
        Assert.All(runs.Where(run => run.Text.Contains("Hello")), run => Assert.Equal(NotoFace.Value, run.Face));
        Assert.All(runs, run => Assert.Equal(TextSyntheticStyle.None, run.Style.Synthetic));
    }

    [Fact]
    public void LinesEndOnlyAtUax14OpportunitiesAndFitTheFrame()
    {
        const string text = "The quick brown fox jumps over the lazy dog while 敏捷的棕色狐狸跳过了懒狗，然后继续奔跑。";
        var result = Layout(text, 30);
        AssertCoverage(result);
        Assert.True(result.Lines.Count > 3);
        var opportunities = UnicodeTextAnalysis.LineBreakOpportunities(text).Select(value => value.Offset).ToHashSet();
        foreach (var line in result.Lines)
        {
            Assert.True(line.ContentWidthMillimeters <= 30 + 1e-6, $"line {line.SourceStart} is {line.ContentWidthMillimeters} mm");
            Assert.Contains(line.SourceStart + line.SourceLength, opportunities);
            Assert.False(line.Runs.Count > 0 && line.ContentLeftMillimeters < -1e-9);
        }
        // Latin words are never split, and the ideographic full stop never starts a line (LB13/LB21 prohibit it).
        Assert.DoesNotContain(result.Lines, line => text[line.SourceStart] is '，' or '。');
    }

    [Fact]
    public void MixedBidiLineIsReorderedVisuallyAndRtlParagraphIsRightAligned()
    {
        var ltr = Layout("abc שלום עולם def", 100);
        var runs = ltr.Lines.Single().Runs;
        var texts = runs.Select(run => run.Text).ToArray();
        // The neutral space between scripts is common script and stays with the preceding Hebrew face.
        Assert.Equal(new[] { "abc ", "שלום עולם", " ", "def" }, texts);
        Assert.Equal(TextDirection.LeftToRight, runs[2].Direction);
        Assert.True(runs[0].OriginXMillimeters < runs[1].OriginXMillimeters);
        // The first logical Hebrew letter is drawn at the right edge of its run.
        var hebrew = runs[1];
        var first = hebrew.Glyphs.Single(glyph => glyph.ClusterStart == hebrew.SourceStart);
        Assert.Equal(first, hebrew.Glyphs[^1]);

        var rtl = Layout("שלום abc", 100);
        var line = rtl.Lines.Single();
        Assert.Equal(1, line.BaseLevel);
        Assert.Equal(100, line.ContentLeftMillimeters + line.ContentWidthMillimeters, 6);
        Assert.Equal("abc", line.Runs[0].Text);
        Assert.Equal("שלום ", line.Runs[1].Text);
    }

    [Fact]
    public void ArabicJoiningAndLineEdgesAreReshapedPerLine()
    {
        const string text = "مرحبا بالعالم مرحبا بالعالم مرحبا بالعالم";
        var wide = Layout(text, 200);
        var narrow = Layout(text, 25);
        AssertCoverage(narrow);
        Assert.True(narrow.Lines.Count > 1);
        var nominal = (uint)DejaVuFace.Value.NominalGlyph('م');
        // Joining forms are actual GIDs from GSUB, not cmap glyphs.
        Assert.Contains(wide.Lines[0].Runs.SelectMany(run => run.Glyphs), glyph => glyph.GlyphId != nominal &&
            text[glyph.ClusterStart] == 'م');
        // Each line's last word is shaped exactly like the standalone word (no join across the break).
        var standalone = Layout("بالعالم", 200).Lines[0].Runs[0].Glyphs.Select(glyph => glyph.GlyphId).ToArray();
        foreach (var line in narrow.Lines)
        {
            var content = text.Substring(line.SourceStart, line.SourceLength).TrimEnd();
            if (!content.EndsWith("بالعالم", StringComparison.Ordinal)) continue;
            var start = line.SourceStart + content.Length - "بالعالم".Length;
            var ids = line.Runs.SelectMany(run => run.Glyphs).Where(glyph => glyph.ClusterStart >= start &&
                glyph.ClusterStart < start + 7).Select(glyph => glyph.GlyphId).ToArray();
            Assert.Equal(standalone, ids);
        }
        Assert.All(narrow.Lines, line => Assert.Equal(1, line.BaseLevel));
    }

    [Fact]
    public void TamilKeepsManyToManyClustersInsideOneGrapheme()
    {
        const string text = "தமிழ் கொண்டு க்ஷ";
        var result = Layout(text, 100);
        AssertCoverage(result);
        var glyphs = result.Lines.SelectMany(line => line.Runs).SelectMany(run => run.Glyphs).ToArray();
        Assert.Contains(glyphs, glyph => glyph.ClusterLength > 1);
        var boundaries = UnicodeTextAnalysis.GraphemeBoundaries(text).ToHashSet();
        Assert.All(glyphs, glyph => Assert.Contains(glyph.ClusterStart, boundaries));
    }

    [Fact]
    public void OverflowIsRejectedOrBrokenOnlyAtGraphemeBoundaries()
    {
        const string word = "Pneumonoultramicroscopicsilicovolcanoconiosis";
        var error = Assert.Throws<TextLayoutException>(() => Layout(word, 20));
        Assert.Equal(TextLayoutError.LayoutOverflow, error.Error);
        Assert.Equal(0, error.SourceStart);

        const string marked = "éééééééééééé";
        var broken = Layout(marked, 6, overflow: TextOverflowPolicy.BreakGraphemes);
        Assert.True(broken.Lines.Count > 1);
        var boundaries = UnicodeTextAnalysis.GraphemeBoundaries(marked).ToHashSet();
        Assert.All(broken.Lines, line =>
        {
            Assert.Contains(line.SourceStart, boundaries);
            Assert.True(line.ContentWidthMillimeters <= 6 + 1e-6);
        });
    }

    [Fact]
    public void JustifiedLinesFillTheFrameExceptParagraphEnds()
    {
        const string text = "Justified text spreads its interword spaces so that every line except the last fills the frame.";
        var result = Layout(text, 40, TextAlignment.Justify);
        Assert.True(result.Lines.Count > 2);
        foreach (var line in result.Lines.Take(result.Lines.Count - 1))
        {
            Assert.Equal(40, line.ContentWidthMillimeters, 6);
            var last = line.Runs.SelectMany(run => run.Glyphs).Where(glyph => glyph.InkBoundsMillimeters is not null).Max(glyph =>
                glyph.InkBoundsMillimeters!.X + glyph.InkBoundsMillimeters.Width);
            Assert.InRange(last, 39, 40.5);
        }
        Assert.True(result.Lines[^1].ContentWidthMillimeters < 40);
        Assert.Contains(result.Lines[0].Runs.SelectMany(run => run.Glyphs), glyph => glyph.JustificationMillimeters > 0);

        var cjk = Layout("汉字两端对齐时在字符之间均匀分配剩余空间，最后一行保持起始对齐。", 30, TextAlignment.Justify);
        Assert.All(cjk.Lines.Take(cjk.Lines.Count - 1), line => Assert.Equal(30, line.ContentWidthMillimeters, 6));
    }

    [Fact]
    public void PaginationKeepsEveryLineInsideItsPage()
    {
        var text = string.Join("\n", Enumerable.Range(1, 12).Select(index =>
            $"Paragraph {index} has enough words to wrap across more than one line of the narrow frame."));
        var result = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, Style(),
            new ParagraphLayoutOptions(40, 50, paragraphSpacingMillimeters: 2, firstLineIndentMillimeters: 4)));
        AssertCoverage(result);
        Assert.True(result.Pages.Count > 3);
        foreach (var page in result.Pages)
        {
            Assert.NotEmpty(page.Lines);
            Assert.True(page.UsedHeightMillimeters <= 50 + 1e-9);
            Assert.All(page.Lines, line => Assert.Equal(page.Index, line.PageIndex));
            Assert.Equal(0, page.Lines[0].TopMillimeters);
        }
        // First lines are indented on the start side; continuation lines are not.
        var firstLines = result.Lines.Where((line, index) => index == 0 || result.Lines[index - 1].EndsParagraph);
        Assert.All(firstLines, line => Assert.Equal(4, line.ContentLeftMillimeters, 9));
        Assert.Contains(result.Lines, line => !line.EndsParagraph && line.ContentLeftMillimeters == 0);
        var tall = Assert.Throws<TextLayoutException>(() => Layout("x", 40, pageHeight: 1));
        Assert.Equal(TextLayoutError.LayoutOverflow, tall.Error);
    }

    [Fact]
    public void StylesAreSnapshotAndSplitRunsAtPaintChanges()
    {
        var red = new TextPaint(new TextColor(200, 0, 0));
        var normal = Style();
        var highlighted = Style(5, red);
        const string text = "plain RED plain";
        var result = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text,
            new[] { new TextStyleSpan(0, 6, normal), new TextStyleSpan(6, 3, highlighted), new TextStyleSpan(9, 6, normal) },
            new ParagraphLayoutOptions(100)));
        var runs = result.Lines.Single().Runs;
        Assert.Equal(3, runs.Count);
        Assert.Equal(new TextColor(200, 0, 0), runs[1].Style.Paint.FillColor);
        Assert.Equal(5, runs[1].FontSizeMillimeters);
        Assert.Equal(TextColor.Black, runs[0].Style.Paint.FillColor);
        // A common baseline across sizes; the line ascent is the larger run's ascent.
        Assert.All(runs, run => Assert.Equal(result.Lines[0].BaselineMillimeters, run.BaselineYMillimeters));
        Assert.Equal(5.0 * NotoFace.Value.Ascender / NotoFace.Value.UnitsPerEm, result.Lines[0].AscentMillimeters, 9);
    }

    [Theory]
    [InlineData("tab\there", TextLayoutError.InvalidText, 3)]
    [InlineData("a‮b", TextLayoutError.UnsupportedFontProfile, 1)]
    [InlineData("a‍b", TextLayoutError.UnsupportedFontProfile, 1)]
    [InlineData("a\u0007b", TextLayoutError.InvalidText, 1)]
    public void UnadmittedControlsAreRejectedBeforeShaping(string text, TextLayoutError error, int offset)
    {
        var exception = Assert.Throws<TextLayoutException>(() => Layout(text));
        Assert.Equal(error, exception.Error);
        Assert.Equal(offset, exception.SourceStart);
    }

    [Fact]
    public void MissingCoverageAndBoldWithoutFaceFail()
    {
        var latinOnly = new TextStyle(new TextFontFamily("Noto Sans", NotoFace.Value), 4);
        var missing = Assert.Throws<TextLayoutException>(() => Layout("ab世", style: latinOnly));
        Assert.Equal(TextLayoutError.MissingGlyph, missing.Error);
        Assert.Equal(2, missing.SourceStart);
        var bold = new TextStyle(new TextFontFamily("Noto Sans", NotoFace.Value), 4, bold: true);
        Assert.Equal(TextLayoutError.UnsupportedFontProfile,
            Assert.Throws<TextLayoutException>(() => Layout("ab", style: bold)).Error);
    }

    [Fact]
    public void HardBreaksAndBudgetsAreHonored()
    {
        var result = Layout("one\u2028two\r\nthree\n\nfive", 100);
        Assert.Equal(new[] { "one", "two", "three", "", "five" }, result.Lines.Select(line =>
            line.Runs.Count == 0 ? "" : string.Concat(line.Runs.Select(run => run.Text))));
        Assert.Equal(new[] { 1, 2, 1, 1, 0 }, result.Lines.Select(line => line.TerminatorLength));
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(
            new ParagraphLayoutRequest("abc", Style(), new ParagraphLayoutOptions(50)), cancelled.Token));
        var budget = Assert.Throws<TextLayoutException>(() => HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(
            new ParagraphLayoutRequest("a b c d e f g", Style(), new ParagraphLayoutOptions(3, maximumLines: 2))));
        Assert.Equal(TextLayoutError.BudgetExceeded, budget.Error);
    }

    [Fact]
    public void PaintOnlyStyleChangeInsideAWordKeepsShapingContext()
    {
        const string word = "ببب";
        var plain = Layout(word, 100).Lines.Single().Runs.SelectMany(run => run.Glyphs).Select(glyph => glyph.GlyphId).ToArray();
        var red = Style(4, new TextPaint(new TextColor(200, 0, 0)));
        var split = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(word,
            new[] { new TextStyleSpan(0, 1, Style()), new TextStyleSpan(1, 1, red), new TextStyleSpan(2, 1, Style()) },
            new ParagraphLayoutOptions(100)));
        var runs = split.Lines.Single().Runs;
        Assert.Equal(3, runs.Count);
        // Visual order is reversed for RTL; joined initial/medial/final forms stay identical.
        Assert.Equal(plain, runs.SelectMany(run => run.Glyphs).Select(glyph => glyph.GlyphId).ToArray());
        Assert.Equal(3, plain.Distinct().Count());
        Assert.Equal(new TextColor(200, 0, 0), runs[1].Style.Paint.FillColor);
    }

    [Fact]
    public void EmergencyBreakRebuildsCutsInsideALigature()
    {
        var dejavuOnly = new TextStyle(new TextFontFamily("DejaVu Sans", DejaVuFace.Value), 4);
        var whole = Layout("ffi", 100, style: dejavuOnly).Lines.Single();
        Assert.Single(whole.Runs.Single().Glyphs);
        Assert.True(whole.ContentWidthMillimeters > 2);
        var broken = Layout("ffi", 2, overflow: TextOverflowPolicy.BreakGraphemes, style: dejavuOnly);
        Assert.Equal(new[] { "f", "f", "i" }, broken.Lines.Select(line => line.Runs.Single().Text));
        Assert.All(broken.Lines, line => Assert.True(line.ContentWidthMillimeters <= 2));
    }

    [Fact]
    public void EmptyTextHasOneStyledEmptyLine()
    {
        var result = Layout("", 50);
        var line = Assert.Single(result.Lines);
        Assert.Empty(line.Runs);
        Assert.Equal(0, line.SourceLength);
        Assert.True(line.HeightMillimeters > 0);
        Assert.True(line.BaselineMillimeters > 0);
    }

    internal static void AssertCoverage(ParagraphLayoutResult result)
    {
        var covered = new int[result.Text.Length];
        foreach (var line in result.Lines)
        {
            var lineEnd = line.SourceStart + line.SourceLength - line.TerminatorLength;
            foreach (var run in line.Runs)
            {
                Assert.Equal(result.Text.Substring(run.SourceStart, run.SourceLength), run.Text);
                Assert.InRange(run.SourceStart, line.SourceStart, lineEnd);
                Assert.Same(run.ShapedRun.Request.Face, run.Face);
                for (var index = run.SourceStart; index < run.SourceStart + run.SourceLength; index++) covered[index]++;
                foreach (var glyph in run.Glyphs)
                {
                    Assert.InRange(glyph.ClusterStart, run.SourceStart, run.SourceStart + run.SourceLength - 1);
                    Assert.True(glyph.GlyphId > 0);
                }
            }
            for (var index = lineEnd; index < line.SourceStart + line.SourceLength; index++) covered[index]++;
        }
        Assert.All(covered, count => Assert.Equal(1, count));
    }

    internal static TextFontFace Face(string name, string file) => new(name, File.ReadAllBytes(FontPath(file)));

    internal static string FontPath(string file)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "e2e", "Ofdrw.Net.FontSubset.E2E", "testdata", "fonts", file);
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        throw new FileNotFoundException(file);
    }
}
