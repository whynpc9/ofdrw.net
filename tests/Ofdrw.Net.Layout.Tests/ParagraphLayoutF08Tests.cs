using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

// Issue #109 F08 items needed by the opt-in engine switch: CFF faces, tab stops, synthetic
// bold/oblique, supplementary characters, decorations, letter spacing and joiners/selectors.
public sealed class ParagraphLayoutF08Tests
{
    private static readonly Lazy<TextFontFace> Noto = new(() => ParagraphLayoutTests.Face("Noto Sans", "NotoSans-Regular.ttf"));
    private static readonly Lazy<TextFontFace> WenKai = new(() => ParagraphLayoutTests.Face("LXGW WenKai", "LXGWWenKai-Regular.ttf"));
    private static readonly Lazy<TextFontFace> DejaVu = new(() => ParagraphLayoutTests.Face("R22 DejaVu Sans", "DejaVuSans-R22.ttf"));
    private static readonly Lazy<TextFontFace> HanCff = new(() => new TextFontFace("Ofdrw CI Han CFF", File.ReadAllBytes(CffPath("OfdrwCIHanCFF-Regular.otf"))));

    internal static string CffPath(string file)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "tests", "Ofdrw.Net.Layout.Tests", "testdata", "cff-han", file);
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        throw new FileNotFoundException(file);
    }

    private static TextStyle Style(double size = 4, bool bold = false, bool italic = false, TextPaint? paint = null) =>
        new(new[] { new TextFontFamily("Noto Sans", Noto.Value), new TextFontFamily("LXGW WenKai", WenKai.Value),
            new TextFontFamily("DejaVu Sans", DejaVu.Value) }, size, paint, bold, italic);

    private static ParagraphLayoutResult Layout(string text, TextStyle? style = null, ParagraphLayoutOptions? options = null) =>
        HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, style ?? Style(),
            options ?? new ParagraphLayoutOptions(150)));

    private static double Advance(ParagraphLayoutRun run) => run.Glyphs.Sum(glyph => glyph.AdvanceXMillimeters);

    // ---- Tabs -------------------------------------------------------------------------------

    [Fact]
    public void TabsAdvanceToDefaultAndExplicitStopsAndKeepTheirSource()
    {
        var defaults = Layout("a\tb\tc");
        var runs = defaults.Lines.Single().Runs;
        Assert.Equal(5, runs.Count);
        var tabs = runs.Where(run => run.IsTab).ToArray();
        Assert.Equal(2, tabs.Length);
        Assert.All(tabs, tab => Assert.Equal("\t", tab.Text));
        Assert.All(tabs, tab => Assert.Equal((uint)Noto.Value.NominalGlyph(' '), tab.Glyphs.Single().GlyphId));
        // Text after each tab starts on the next 12.7 mm default stop.
        Assert.Equal(12.7, runs[2].OriginXMillimeters, 6);
        Assert.Equal(25.4, runs[4].OriginXMillimeters, 6);

        var stops = new ParagraphLayoutOptions(150).WithTabStops(new[]
        {
            new TextTabStop(30), new TextTabStop(60, TextTabAlignment.Center), new TextTabStop(100, TextTabAlignment.End),
            new TextTabStop(130, TextTabAlignment.Decimal),
        }, 10);
        var aligned = Layout("x\tStart\tCenter\tEnd\t12.50", options: stops).Lines.Single().Runs.Where(run => !run.IsTab).ToArray();
        Assert.Equal(30, aligned[1].OriginXMillimeters, 6);
        Assert.Equal(60, aligned[2].OriginXMillimeters + Advance(aligned[2]) / 2, 6);
        Assert.Equal(100, aligned[3].OriginXMillimeters + Advance(aligned[3]), 6);
        // The decimal separator of "12.50" sits on the stop.
        var decimalRun = aligned[4];
        var beforeDot = decimalRun.Glyphs.TakeWhile(glyph => decimalRun.Text[glyph.ClusterStart - decimalRun.SourceStart] != '.')
            .Sum(glyph => glyph.AdvanceXMillimeters);
        Assert.Equal(130, decimalRun.OriginXMillimeters + beforeDot, 6);

        // Past the last stop, implicit stops continue from the configured origin; a hanging first line
        // (negative indent) reaches the frame's start edge with its first tab.
        var origin = new ParagraphLayoutOptions(150, firstLineIndentMillimeters: -8).WithTabStops(null, 10, -3);
        var hanging = Layout("1.\tItem", options: origin).Lines.Single().Runs;
        Assert.Equal(-8, hanging[0].OriginXMillimeters, 6);
        Assert.Equal(-3 + 10 * Math.Ceiling((Advance(hanging[0]) - 8 + 3) / 10 + 1e-9), hanging[2].OriginXMillimeters, 6);
    }

    [Fact]
    public void AMinimumTabAdvanceSkipsAStopTooCloseToTheText()
    {
        // Issue #115: text ending 0.5 mm before a stop left a near-zero gap.
        var width = Advance(Layout("abc").Lines.Single().Runs.Single());
        var near = width + 0.5;
        string Positions(ParagraphLayoutOptions options) => string.Join(",",
            Layout("abc	def", options: options).Lines.Single().Runs.Select(run => run.OriginXMillimeters.ToString("0.###", CultureInfo.InvariantCulture)));
        double After(ParagraphLayoutOptions options) => Layout("abc	def", options: options).Lines.Single().Runs[2].OriginXMillimeters;
        var implicitStops = new ParagraphLayoutOptions(150).WithTabStops(null, 10, near);
        var explicitStop = new ParagraphLayoutOptions(150).WithTabStops(new[] { new TextTabStop(near) }, 10, near);
        // Word's rule (the default) uses any stop past the text.
        Assert.Equal(near, After(implicitStops), 6);
        Assert.Equal(near, After(explicitStop), 6);
        // Half an em (2 mm at 4 mm) skips both the implicit and the start-aligned explicit stop.
        Assert.Equal(near + 10, After(implicitStops.WithMinimumTabAdvance(0.5)), 6);
        Assert.Equal(near + 10, After(explicitStop.WithMinimumTabAdvance(0.5)), 6);
        // A gap that is already wide enough is unchanged, and so are center/end/decimal stops.
        Assert.Equal(Positions(new ParagraphLayoutOptions(150)), Positions(new ParagraphLayoutOptions(150).WithMinimumTabAdvance(0.5)));
        var end = new ParagraphLayoutOptions(150).WithTabStops(new[] { new TextTabStop(near + 0.1 + Advance(Layout("def").Lines.Single().Runs.Single()), TextTabAlignment.End) });
        Assert.Equal(Positions(end), Positions(end.WithMinimumTabAdvance(0.5)));
        Assert.Throws<TextLayoutException>(() => new ParagraphLayoutOptions(150).WithMinimumTabAdvance(-1));
        Assert.Throws<TextLayoutException>(() => new ParagraphLayoutOptions(150).WithMinimumTabAdvance(double.NaN));
    }

    [Fact]
    public void TabsDoNotJustifyAndClampToTheFrame()
    {
        var options = new ParagraphLayoutOptions(40, alignment: TextAlignment.Justify);
        var result = Layout("Name\tValue and more words that wrap onto another line", options: options);
        Assert.True(result.Lines.Count > 1);
        var first = result.Lines[0];
        // Text before the last tab is never stretched.
        Assert.All(first.Runs.TakeWhile(run => !run.IsTab).SelectMany(run => run.Glyphs),
            glyph => Assert.Equal(0, glyph.JustificationMillimeters));
        // A tab beyond the frame clamps to the frame edge instead of overflowing.
        var far = Layout("abc\tdef", options: new ParagraphLayoutOptions(20).WithTabStops(new[] { new TextTabStop(500) }));
        var line = far.Lines[0];
        Assert.True(line.ContentWidthMillimeters <= 20 + 1e-6);
    }

    // ---- Synthetic bold / oblique ------------------------------------------------------------

    [Fact]
    public void SyntheticStylesAreOptInAndRecordedWithInkAndPaint()
    {
        var bold = Style(4, bold: true, italic: true);
        Assert.Equal(TextLayoutError.UnsupportedFontProfile,
            Assert.Throws<TextLayoutException>(() => Layout("Bold 粗体", bold)).Error);
        var regular = Layout("Bold 粗体", Style(4));
        var options = new ParagraphLayoutOptions(150).WithSyntheticStyles(TextSyntheticStylePolicy.Synthesize);
        var result = Layout("Bold 粗体", bold, options);
        var runs = result.Lines.Single().Runs;
        Assert.All(runs, run => Assert.Equal(TextSyntheticStyle.Bold | TextSyntheticStyle.Oblique, run.Style.Synthetic));
        var latin = runs.Single(run => run.Script == "Latn");
        var han = runs.Single(run => run.Script == "Hani");
        // The Han stroke is thinner (dense counters); the paint is a fill-colored outline stroke.
        Assert.Equal(FontEmphasisMetrics.BoldStrokeWidth(4), latin.Style.SyntheticBoldStrokeMillimeters, 9);
        Assert.Equal(FontEmphasisMetrics.HanBoldStrokeWidth(4), han.Style.SyntheticBoldStrokeMillimeters, 9);
        Assert.Equal(TextColor.Black, latin.Style.Paint.StrokeColor);
        Assert.Equal(TextResolvedStyle.SyntheticObliqueSkew, latin.Style.ObliqueSkew);
        // Advances are unchanged; ink grows by the half stroke and leans right above the baseline.
        Assert.Equal(regular.Lines[0].ContentWidthMillimeters, result.Lines[0].ContentWidthMillimeters, 9);
        var plainInk = regular.Lines[0].Runs.Single(run => run.Script == "Latn").InkBoundsMillimeters!;
        var synthInk = latin.InkBoundsMillimeters!;
        Assert.True(synthInk.X + synthInk.Width > plainInk.X + plainInk.Width + 0.1);
        Assert.True(synthInk.Height > plainInk.Height);
        // A requested stroke cannot be combined with synthetic bold.
        var stroked = Style(4, bold: true, paint: new TextPaint(TextColor.Black, new TextColor(255, 0, 0), 0.1));
        Assert.Equal(TextLayoutError.UnsupportedFontProfile,
            Assert.Throws<TextLayoutException>(() => Layout("x", stroked, options)).Error);
    }

    [Fact]
    public async Task SyntheticStylesAreWrittenAsStrokeAndShearAndExport()
    {
        var options = new ParagraphLayoutOptions(90).WithSyntheticStyles(TextSyntheticStylePolicy.Synthesize);
        var layout = Layout("Bold italic 粗斜体", Style(5, bold: true, italic: true), options);
        var bytes = await ParagraphLayoutOfdTests.Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 110, 40, 10, 10));
        var objects = TextObjects(bytes);
        Assert.All(objects, node =>
        {
            Assert.Equal("true", node.Attribute("Stroke")?.Value);
            Assert.NotNull(node.Elements().SingleOrDefault(child => child.Name.LocalName == "StrokeColor"));
            Assert.Null(node.Attribute("Weight"));
            Assert.Null(node.Attribute("Italic"));
            var ctm = node.Attribute("CTM")!.Value.Split(' ').Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            Assert.Equal(new[] { 1, 0, -0.2, 1 }, ctm.Take(4));
            Assert.True(ctm[4] > 0);
        });
        await AssertExportsAsync(bytes, layout);
    }

    // ---- Supplementary characters --------------------------------------------------------------

    [Fact]
    public async Task SupplementaryCharactersCountOnceInCodePositions()
    {
        const string text = "姓名：张\U00020087三 地址：\U0002000B路";
        var layout = Layout(text);
        ParagraphLayoutTests.AssertCoverage(layout);
        var bytes = await ParagraphLayoutOfdTests.Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 40, 10, 10));
        var han = TextObjects(bytes).Single(node => node.Value.Contains("\U00020087"));
        var transforms = han.Elements().Where(node => node.Name.LocalName == "CGTransform").ToArray();
        var code = han.Elements().Single(node => node.Name.LocalName == "TextCode").Value;
        var scalars = new StringInfo(code).LengthInTextElements; // no combining marks here
        Assert.Equal(scalars, transforms.Sum(node => int.Parse(node.Attribute("CodeCount")!.Value)));
        Assert.Equal(transforms.Length, transforms.Select(node => node.Attribute("CodePosition")!.Value).Distinct().Count());
        Assert.True(scalars < code.Length);
        // Read back: the logical text is exact and every original GID is projected.
        var document = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var elements = document.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).ToArray();
        Assert.Equal(text.Replace(" ", ""), string.Concat(elements.Select(element => element.Text)).Replace(" ", ""));
        var run = layout.Lines.SelectMany(line => line.Runs).Single(value => value.Text.Contains("\U00020087"));
        var element = elements.Single(value => value.Text == run.Text);
        Assert.Equal(run.Glyphs.GroupBy(glyph => glyph.ClusterStart).OrderBy(group => group.Key).SelectMany(group => group)
            .Select(glyph => (int)glyph.GlyphId), OfdExplicitGlyphRunProbe.GlyphIds(element,
                document.Fonts.Single(font => font.Id == element.FontResourceId)));
        await AssertExportsAsync(bytes, layout);
    }

    // ---- Decorations, highlight, letter spacing, line height ------------------------------------

    [Fact]
    public async Task DecorationsUseFaceMetricsAndAreWrittenInPaintOrder()
    {
        var style = Style(4).WithDecorations(TextDecorationLines.Underline | TextDecorationLines.Strikethrough,
            new TextColor(0, 0, 255)).WithHighlight(new TextColor(255, 255, 0));
        var layout = Layout("Marked 标记 text   ", style);
        var line = layout.Lines.Single();
        var latin = line.Runs.First();
        var kinds = latin.Decorations.Select(value => value.Kind).ToArray();
        Assert.Equal(new[] { ParagraphLayoutDecorationKind.Highlight, ParagraphLayoutDecorationKind.Underline,
            ParagraphLayoutDecorationKind.Strikethrough }, kinds);
        var scale = 4.0 / Noto.Value.UnitsPerEm;
        var underline = latin.Decorations[1].BoundsMillimeters;
        Assert.Equal(line.BaselineMillimeters - Noto.Value.UnderlinePosition * scale, underline.Y, 9);
        Assert.Equal(Noto.Value.UnderlineThickness * scale, underline.Height, 9);
        Assert.Equal(new TextColor(0, 0, 255), latin.Decorations[1].Color);
        var strike = latin.Decorations[2].BoundsMillimeters;
        Assert.True(strike.Y < line.BaselineMillimeters && strike.Y > line.BaselineMillimeters - 4 * 0.6);
        // Hanging trailing spaces are not decorated.
        var last = line.Runs.Last();
        Assert.True(last.Decorations[0].BoundsMillimeters.Width < Advance(last) - 1);

        var bytes = await ParagraphLayoutOfdTests.Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 40, 10, 10));
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var content = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal) &&
            entry.FullName.Contains("Page", StringComparison.Ordinal));
        using var stream = content.Open();
        var order = XDocument.Load(stream).Descendants().Where(node => node.Name.LocalName is "PathObject" or "TextObject")
            .Select(node => node.Name.LocalName).ToArray();
        var firstText = Array.IndexOf(order, "TextObject");
        Assert.True(firstText > 0);
        Assert.All(order.Take(firstText), name => Assert.Equal("PathObject", name));
        Assert.Equal("PathObject", order.Last());
        // Issue #115: font fallback splits the line into Latin and Han runs, but each decoration is one
        // path spanning all of them, so viewers draw no seam where the pieces meet.
        Assert.True(line.Runs.Count >= 3);
        Assert.Equal(new[] { "PathObject", "TextObject" }, order.Take(2));
        Assert.Equal(3, order.Count(name => name == "PathObject"));
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var highlight = read.Pages.Single().Elements.OfType<OfdPathElement>().First();
        var pieces = line.Runs.SelectMany(run => run.Decorations).Where(value => value.Kind == ParagraphLayoutDecorationKind.Highlight)
            .Select(value => value.BoundsMillimeters).ToArray();
        Assert.Equal(10 + pieces.Min(value => value.X), highlight.XMillimeters, 3);
        Assert.Equal(pieces.Max(value => value.X + value.Width) - pieces.Min(value => value.X), highlight.WidthMillimeters, 3);
    }

    [Fact]
    public void DecorationsOfDifferentColorsOrSizesStaySeparate()
    {
        var text = "red blue big";
        var red = Style(4).WithHighlight(new TextColor(255, 0, 0));
        var blue = Style(4).WithHighlight(new TextColor(0, 0, 255));
        var big = Style(6).WithHighlight(new TextColor(0, 0, 255));
        var layout = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, new[]
        {
            new TextStyleSpan(0, 4, red), new TextStyleSpan(4, 5, blue), new TextStyleSpan(9, 3, big),
        }, new ParagraphLayoutOptions(150)));
        var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 40, 10, 10);
        var paths = package.Pages.Single().Elements.OfType<OfdPathElement>().ToArray();
        Assert.Equal(3, paths.Length);
        Assert.Equal(new[] { 255, 0, 0 }, new[] { paths[0].FillColor.Red, paths[1].FillColor.Red, paths[2].FillColor.Red });
    }

    [Fact]
    public void LetterSpacingAndAtLeastLineHeight()
    {
        var plain = Layout("abcd");
        var spaced = Layout("abcd", Style(4).WithLetterSpacing(0.5));
        // Spacing follows every grapheme followed by same-style content on the line: three gaps.
        Assert.Equal(plain.Lines[0].ContentWidthMillimeters + 1.5, spaced.Lines[0].ContentWidthMillimeters, 9);
        var normal = plain.Lines[0];
        var tall = Layout("abcd", options: new ParagraphLayoutOptions(150, lineHeight: TextLineHeight.AtLeast(20))).Lines[0];
        Assert.Equal(20, tall.HeightMillimeters, 9);
        Assert.Equal(normal.BaselineMillimeters, tall.BaselineMillimeters, 9);
        var unchanged = Layout("abcd", options: new ParagraphLayoutOptions(150, lineHeight: TextLineHeight.AtLeast(1))).Lines[0];
        Assert.Equal(normal.HeightMillimeters, unchanged.HeightMillimeters, 9);
    }

    [Fact]
    public void LetterSpacingCountsEveryGraphemeOfALigatureAndWidensRtlGaps()
    {
        var dejavu = new TextFontFamily("DejaVu Sans", DejaVu.Value);
        double Width(TextStyle style, string text) => Layout(text, style).Lines.Single().ContentWidthMillimeters;
        // Optional ligatures are off for letter-spaced text: "ffiX" keeps four glyphs and three gaps.
        var noLiga = new TextStyle(new[] { dejavu }, 4, features: new[] { new TextFontFeature("liga", 0) });
        var spaced = new TextStyle(new[] { dejavu }, 4).WithLetterSpacing(0.5);
        var spacedRun = Layout("ffiX", spaced).Lines.Single().Runs.Single();
        Assert.Equal(4, spacedRun.Glyphs.Count);
        Assert.Equal(Width(noLiga, "ffiX") + 1.5, Width(spaced, "ffiX"), 9);
        // An explicitly requested ligature keeps its glyph and still receives every grapheme's gap.
        var forced = new TextStyle(new[] { dejavu }, 4, features: new[] { new TextFontFeature("liga", 1) });
        var forcedSpaced = new TextStyle(new[] { dejavu }, 4, features: new[] { new TextFontFeature("liga", 1) }).WithLetterSpacing(0.5);
        Assert.True(Layout("ffiX", forced).Lines.Single().Runs.Single().Glyphs.Count < 4);
        Assert.Equal(Width(forced, "ffiX") + 1.5, Width(forcedSpaced, "ffiX"), 9);
        // RTL: the visible gap between the two Hebrew letters widens; no tracking at the run's outer edge.
        var plain = Layout("אב", new TextStyle(new[] { dejavu }, 4)).Lines.Single().Runs.Single();
        var hebrew = Layout("אב", new TextStyle(new[] { dejavu }, 4).WithLetterSpacing(0.8)).Lines.Single().Runs.Single();
        Assert.Equal(1, hebrew.BidiLevel % 2);
        var plainGap = plain.Glyphs[1].XMillimeters - plain.Glyphs[0].XMillimeters;
        Assert.Equal(plainGap + 0.8, hebrew.Glyphs[1].XMillimeters - hebrew.Glyphs[0].XMillimeters, 9);
        Assert.Equal(plain.AdvanceMillimeters + 0.8, hebrew.AdvanceMillimeters, 9);
        Assert.Equal(hebrew.OriginXMillimeters + hebrew.AdvanceMillimeters,
            hebrew.Glyphs[1].XMillimeters + hebrew.Glyphs[1].AdvanceXMillimeters, 9);
        Assert.Equal(plain.Glyphs[1].AdvanceXMillimeters, hebrew.Glyphs[1].AdvanceXMillimeters, 9);
    }

    // ---- ZWJ / ZWNJ / variation selectors -----------------------------------------------------

    [Fact]
    public void JoinersAndVariationSelectorsShapeInsideTheirCluster()
    {
        // Persian "mi-khaham": ZWNJ keeps the first letter in its final form instead of joining.
        var joined = Layout("میخواهم").Lines.Single().Runs.Single();
        var separated = Layout("می‌خواهم").Lines.Single().Runs.Single();
        Assert.NotEqual(joined.Glyphs.Select(glyph => glyph.GlyphId), separated.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Contains(separated.Glyphs, glyph => separated.Text.Substring(glyph.ClusterStart - separated.SourceStart,
            glyph.ClusterLength).Contains('‌'));
        // A selector the face does not map is resolved by the shaper and stays in the base cluster.
        var selected = Layout("中︀文").Lines.Single();
        ParagraphLayoutTests.AssertCoverage(selected.Runs.Count == 0 ? throw new InvalidOperationException() :
            Layout("中︀文"));
        Assert.Equal("中︀文", string.Concat(selected.Runs.Select(run => run.Text)));
        // Other format controls stay rejected.
        Assert.Equal(TextLayoutError.UnsupportedFontProfile,
            Assert.Throws<TextLayoutException>(() => Layout("a‎b")).Error);
    }

    // ---- CFF faces ------------------------------------------------------------------------------

    [Fact]
    public void CidKeyedCffFaceIsAdmittedAndShaped()
    {
        var face = HanCff.Value;
        Assert.Equal(TextFontOutlineFormat.Cff, face.OutlineFormat);
        Assert.Equal(1000, face.UnitsPerEm);
        var style = new TextStyle(new TextFontFamily("Ofdrw CI Han CFF", face), 5);
        var result = Layout("思源黑体 CFF 字体测试 \U00020087", style);
        ParagraphLayoutTests.AssertCoverage(result);
        Assert.All(result.Lines.SelectMany(line => line.Runs), run => Assert.Same(face, run.Face));
        Assert.All(result.Lines.SelectMany(line => line.Runs).SelectMany(run => run.Glyphs),
            glyph => Assert.True(glyph.GlyphId > 0));
        // Variable and CFF2 faces stay outside the profile.
        var bytes = File.ReadAllBytes(CffPath("OfdrwCIHanCFF-Regular.otf"));
        var renamed = (byte[])bytes.Clone();
        ReplaceTag(renamed, "CFF ", "CFF2");
        Assert.Equal(TextLayoutError.UnsupportedFontProfile,
            Assert.Throws<TextLayoutException>(() => new TextFontFace("x", renamed)).Error);
    }

    [Theory]
    [InlineData(72, "26b227058cc1f865410c8b5162eb4ad2eb4d64a40703428146d5ccaa77a6ff35")]
    [InlineData(257, "238d6200989ca2a4d362dba967bbe4e624d8c785d6f62c9d6118edb31a710bbb")]
    [InlineData(263, "f14acd84d2833f486067b197e0c446511dbf259bdf2799095cec2de8cc522888")]
    [InlineData(459, "945f1c42231b66379abe429a8a34f0f09475d231072310c0a44cccff3740a3f3")]
    public void ManagedCharStringOutlinesMatchTheIndependentOracle(int gid, string oracleSha256)
    {
        var oracleText = File.ReadAllText(CffPath($"gid-{gid}.path"));
        Assert.Equal(oracleSha256, Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(oracleText))).ToLowerInvariant());
        var expected = oracleText.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var outlines = new CffCharStringOutlines(File.ReadAllBytes(CffPath("OfdrwCIHanCFF-Regular.otf")));
        var actual = Format(outlines.Read(gid).ToPathCommands());
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void PlainCffFaceMatchesTheExistingSourceCodeProOracle()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, "tests", "Ofdrw.Net.Graphics.SkiaSharp.Tests")))
            directory = directory.Parent;
        var root = Path.Combine(directory!.FullName, "tests", "Ofdrw.Net.Graphics.SkiaSharp.Tests", "testdata", "cff");
        var outlines = new CffCharStringOutlines(File.ReadAllBytes(Path.Combine(root, "SourceCodePro-Regular.otf")));
        foreach (var gid in new[] { 2, 34, 298, 998 })
        {
            var expected = File.ReadAllText(Path.Combine(root, $"gid-{gid}.path")).Split('\n', StringSplitOptions.RemoveEmptyEntries);
            var actual = Format(outlines.Read(gid).ToPathCommands());
            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public async Task CffFaceIsSubsetInPlaceAndExportsThroughATrueTypeRendition()
    {
        var face = HanCff.Value;
        var style = new TextStyle(new TextFontFamily("Ofdrw CI Han CFF", face), 6);
        var layout = Layout("思源黑体 CFF 字体测试：张\U00020087", style, new ParagraphLayoutOptions(120));
        var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 140, 40, 10, 10);
        var bytes = await ParagraphLayoutOfdTests.Write(package);
        var document = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var font = document.Fonts.Single();
        var embedded = new OpenTypeFace(font.Data!);
        Assert.Equal(0x4F54544Fu, embedded.Version);
        // GID-preserving: same glyph count; unused CharStrings are a bare endchar.
        var original = CffFontProgram.Parse(new OpenTypeFace(face.GetFaceBytes()));
        var subset = CffFontProgram.Parse(embedded);
        Assert.Equal(original.GlyphCount, subset.GlyphCount);
        var used = layout.Lines.SelectMany(line => line.Runs).SelectMany(run => run.Glyphs).Select(glyph => (int)glyph.GlyphId).ToHashSet();
        var blank = Enumerable.Range(1, subset.GlyphCount).Where(gid => gid < subset.GlyphCount && !used.Contains(gid))
            .Count(gid => subset.Bytes[subset.CharStrings.ItemStart(gid)] == 14 &&
                subset.Bytes.Skip(subset.CharStrings.ItemStart(gid) + 1).Take(subset.CharStrings.ItemLength(gid) - 1).All(value => value == 0));
        Assert.True(blank > subset.GlyphCount / 2);
        var outlines = new CffCharStringOutlines(font.Data!);
        foreach (var gid in used) Assert.Equal(Format(new CffCharStringOutlines(face.GetFaceBytes()).Read(gid).ToPathCommands()),
            Format(outlines.Read(gid).ToPathCommands()));
        // The TrueType rendition keeps GIDs, metrics and cmap for PDF.
        var rendition = new OpenTypeFace(CffTrueTypeConversion.Convert(font.Data!));
        Assert.Equal(0x00010000u, rendition.Version);
        Assert.True(rendition.Tables.ContainsKey("glyf"));
        Assert.False(rendition.Tables.ContainsKey("CFF "));
        Assert.Equal(new OpenTypeFace(font.Data!).Table("hmtx"), rendition.Table("hmtx"));
        await AssertExportsAsync(bytes, layout);
    }

    // Deterministic visual sample of every F08 output feature in one document.
    internal const string VisualSample =
        "Tabs:\tName\t金额\t12.50\n" +
        "Synthetic bold 合成粗体, oblique 合成斜体 and both.\n" +
        "Underline 下划线, strikethrough 删除线 and highlight 高亮.\n" +
        "Letter spacing 字距 and Ext-B 姓名：张\U00020087 \U0002000B路.\n" +
        "Arabic مرحبا بالعالم and CFF 思源黑体字体测试.";

    internal static ParagraphLayoutResult VisualLayout()
    {
        var text = VisualSample;
        var plain = Style(4.5);
        var bold = Style(4.5, bold: true);
        var italic = Style(4.5, italic: true);
        var both = Style(4.5, bold: true, italic: true);
        var decorated = Style(4.5).WithDecorations(TextDecorationLines.Underline, new TextColor(200, 0, 0));
        var struck = Style(4.5).WithDecorations(TextDecorationLines.Strikethrough);
        var highlighted = Style(4.5).WithHighlight(new TextColor(255, 230, 0));
        var spaced = Style(4.5).WithLetterSpacing(0.8);
        var cff = new TextStyle(new[] { new TextFontFamily("Ofdrw CI Han CFF", HanCff.Value), new TextFontFamily("Noto Sans", Noto.Value) }, 4.5);
        var spans = new List<TextStyleSpan>();
        var marks = new (string Token, TextStyle Style)[]
        {
            ("Synthetic bold 合成粗体", bold), ("oblique 合成斜体", italic), ("both", both), ("Underline 下划线", decorated),
            ("strikethrough 删除线", struck), ("highlight 高亮", highlighted), ("Letter spacing 字距", spaced), ("CFF 思源黑体字体测试", cff),
        };
        var cursor = 0;
        foreach (var (token, style) in marks.OrderBy(mark => text.IndexOf(mark.Token, StringComparison.Ordinal)))
        {
            var start = text.IndexOf(token, StringComparison.Ordinal);
            if (start > cursor) spans.Add(new TextStyleSpan(cursor, start - cursor, plain));
            spans.Add(new TextStyleSpan(start, token.Length, style));
            cursor = start + token.Length;
        }
        spans.Add(new TextStyleSpan(cursor, text.Length - cursor, plain));
        var options = new ParagraphLayoutOptions(150, lineHeight: TextLineHeight.Multiple(1.3), paragraphSpacingMillimeters: 2)
            .WithSyntheticStyles(TextSyntheticStylePolicy.Synthesize)
            .WithTabStops(new[] { new TextTabStop(25), new TextTabStop(80, TextTabAlignment.Center), new TextTabStop(140, TextTabAlignment.Decimal) });
        return HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(text, spans, options));
    }

    [Fact]
    public async Task EveryF08OutputFeatureRoundTripsAndExports()
    {
        var layout = VisualLayout();
        ParagraphLayoutTests.AssertCoverage(layout);
        var bytes = await ParagraphLayoutOfdTests.Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 70, 10, 10));
        var document = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var texts = document.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(text => text.Text).ToArray();
        Assert.Equal(layout.Lines.SelectMany(line => line.Runs).Select(run => run.Text), texts);
        Assert.Contains(texts, text => text == "\t");
        Assert.Contains(document.Fonts, font => new OpenTypeFace(font.Data!).Version == 0x4F54544F);
        await AssertExportsAsync(bytes, layout, "f08-visual-sample");
    }

    [Fact]
    public async Task JoinersAndSelectorsRoundTripThroughOfdAndSvgButPdfRefusesToDropThem()
    {
        const string text = "ZWNJ می\u200cخواهم and selector 中\ufe00文.";
        var layout = Layout(text);
        var bytes = await ParagraphLayoutOfdTests.Write(ParagraphLayoutOfdWriter.CreateDocument(layout, 170, 30, 10, 10));
        var document = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Equal(layout.Lines.SelectMany(line => line.Runs).Select(run => run.Text),
            document.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(element => element.Text));
        using var svg = new MemoryStream();
        await new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        Assert.Contains("data-original-gid=", System.Text.Encoding.UTF8.GetString(svg.ToArray()));
        // PDFsharp's text layer cannot carry ZWNJ or variation selectors; export fails closed instead of dropping them.
        await Assert.ThrowsAnyAsync<NotSupportedException>(() =>
            new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
    }

    // ---- helpers --------------------------------------------------------------------------------

    internal static async Task AssertExportsAsync(byte[] ofd, ParagraphLayoutResult layout, string? name = null)
    {
        using var pdf = new MemoryStream();
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 1000);
        using (var document = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
        {
            var extracted = string.Concat(document.GetPages().Select(page => page.Text)).Replace(" ", "");
            foreach (var run in layout.Pages[0].Lines.SelectMany(line => line.Runs).Where(run => !run.IsTab && run.Text.Trim().Length > 0))
                Assert.Contains(run.Text.Replace(" ", ""), extracted);
        }
        using var svg = new MemoryStream();
        await new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var svgText = System.Text.Encoding.UTF8.GetString(svg.ToArray());
        var visible = layout.Pages[0].Lines.SelectMany(line => line.Runs).SelectMany(run => run.Glyphs)
            .Count(glyph => glyph.InkBoundsMillimeters is not null);
        Assert.True(System.Text.RegularExpressions.Regex.Matches(svgText, "data-original-gid=").Count >= visible);
        if (Environment.GetEnvironmentVariable("OFDRW_PARAGRAPH_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            var stem = name ?? "f08-" + Convert.ToHexString(SHA256.HashData(ofd))[..12].ToLowerInvariant();
            await File.WriteAllBytesAsync(Path.Combine(directory, stem + ".ofd"), ofd);
            await File.WriteAllBytesAsync(Path.Combine(directory, stem + ".pdf"), pdf.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(directory, stem + ".svg"), svg.ToArray());
        }
    }

    private static XElement[] TextObjects(byte[] ofd)
    {
        using var zip = new ZipArchive(new MemoryStream(ofd), ZipArchiveMode.Read);
        return zip.Entries.Where(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal) &&
                entry.FullName.Contains("Page", StringComparison.Ordinal))
            .SelectMany(entry => { using var stream = entry.Open(); return XDocument.Load(stream).Descendants()
                .Where(node => node.Name.LocalName == "TextObject").ToArray(); }).ToArray();
    }

    private static string[] Format(IReadOnlyList<TrueTypePathCommand> commands)
    {
        static string P(double x, double y) => string.Format(CultureInfo.InvariantCulture, "{0:0.000},{1:0.000}", x, y);
        return commands.Select(command => command.Verb switch
        {
            TrueTypePathVerb.Move => "M:" + P(command.X, command.Y),
            TrueTypePathVerb.Line => "L:" + P(command.X, command.Y),
            TrueTypePathVerb.Cubic => "C:" + P(command.ControlX, command.ControlY) + ":" + P(command.Control2X, command.Control2Y) +
                ":" + P(command.X, command.Y),
            TrueTypePathVerb.Close => "Z",
            _ => throw new InvalidOperationException(command.Verb.ToString()),
        }).ToArray();
    }

    private static void ReplaceTag(byte[] font, string from, string to)
    {
        var count = (font[4] << 8) | font[5];
        for (var index = 0; index < count; index++)
        {
            var record = 12 + index * 16;
            if (System.Text.Encoding.ASCII.GetString(font, record, 4) == from)
                System.Text.Encoding.ASCII.GetBytes(to).CopyTo(font, record);
        }
    }
}
