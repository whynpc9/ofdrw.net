using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextPositionedConsumerTests
{
    [Theory]
    [InlineData("DejaVuSans-R22.ttf", 2048, 0.0004)]
    [InlineData("NotoSans-Regular.ttf", 1000, 0.0004)]
    [InlineData("DejaVuSans-R22.ttf", 2048, 0.0004999999)]
    [InlineData("NotoSans-Regular.ttf", 1000, 0.0004999999)]
    public void BothWriterEntrancesRejectUnwritableSizeWhileQueryRemainsAvailable(string font, int upem, double size)
    {
        var face = Load(font);
        Assert.Equal(upem, face.UnitsPerEm);
        Assert.True(size / face.UnitsPerEm > 0);
        var layout = HarfBuzzTextPrelayout.Create().Layout(new("A", new[] { face }, size, 100));
        Assert.Single(layout.Lines);
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePage(layout, 210, 297, 10, 10)).Error);
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout, 210, 297, 10, 10)).Error);
    }

    [Theory]
    [InlineData("DejaVuSans-R22.ttf", 0.0005, "0.001")]
    [InlineData("NotoSans-Regular.ttf", 0.0005, "0.001")]
    [InlineData("DejaVuSans-R22.ttf", 0.001, "0.001")]
    [InlineData("NotoSans-Regular.ttf", 0.001, "0.001")]
    [InlineData("DejaVuSans-R22.ttf", 4, "4")]
    [InlineData("NotoSans-Regular.ttf", 4, "4")]
    public async Task BothWriterEntrancesKeepWritableSizeAndSourceAfterSerialization(string font, double size, string literal)
    {
        var layout = HarfBuzzTextPrelayout.Create().Layout(new("A", new[] { Load(font) }, size, 100));
        foreach (var positioned in new[] { false, true })
        {
            var package = positioned
                ? TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout, 210, 297, 10, 10)
                : TextPrelayoutOfdWriter.CreateSinglePage(layout, 210, 297, 10, 10);
            using var stream = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, stream);
            stream.Position = 0;
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
            using (var page = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal)).Open())
                Assert.Equal(literal, XDocument.Load(page).Descendants()
                    .Single(node => node.Name.LocalName == "TextObject").Attribute("Size")!.Value);
            stream.Position = 0;
            var read = await new OfdReader().ReadAsync(stream);
            var text = Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(read.Pages).Elements));
            Assert.Equal("A", text.Text);
            Assert.Equal(double.Parse(literal, CultureInfo.InvariantCulture), text.FontSizeMillimeters);
            Assert.True(text.FontSizeMillimeters > 0);
        }
    }

    [Theory]
    [InlineData("AV", "DejaVuSans-R22.ttf", 4)]
    [InlineData("ffi", "DejaVuSans-R22.ttf", 4)]
    [InlineData("fi", "DejaVuSans-R22.ttf", 6)]
    [InlineData("a\u0323", "DejaVuSans-R22.ttf", 4)]
    [InlineData("jAV", "DejaVuSans-R22.ttf", 4)]
    [InlineData("AV\nfi", "DejaVuSans-R22.ttf", 4)]
    [InlineData("AV", "NotoSans-Regular.ttf", 6)]
    public async Task NativeTypedConsumerPreservesGeometryBitsSourceAndSerializedDeltas(
        string source, string font, double size)
    {
        var face = Load(font);
        var layout = HarfBuzzTextPrelayout.Create().Layout(new(source, new[] { face }, size, 100));
        var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout, 210, 297, 10, 10);
        var elements = package.Pages.Single().Elements.OfType<OfdTextElement>().ToArray();
        var lines = layout.Lines.Where(line => line.SourceLength > 0).ToArray();
        Assert.Equal(lines.Length, elements.Length);
        var expectedDeltas = new List<string?>();
        for (var lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            var line = lines[lineIndex];
            var ink = line.InkBoundsMillimeters;
            var boundaryLeft = Math.Min(0, ink?.X ?? 0);
            var boundaryTop = Math.Min(0, ink?.Y - line.TopMillimeters ?? 0);
            var output = elements[lineIndex].Runs.Single();
            var typed = output.PositionedGlyphs!;
            Assert.Equal(source.Substring(line.SourceStart, line.SourceLength), typed.LogicalText);
            Assert.Equal(line.SourceStart, typed.SourceStart);
            Assert.Equal(face.FaceSha256, typed.FaceSha256);
            var strict = TextPositionedRun.Create(line.ShapedRun, size, -boundaryLeft,
                line.BaselineMillimeters - line.TopMillimeters - boundaryTop);
            double oldPenX = 0, oldPenY = 0;
            var scale = size / face.UnitsPerEm;
            var expectedOrigins = new List<double>();
            for (var index = 0; index < typed.Glyphs.Count; index++)
            {
                var raw = line.ShapedRun.Glyphs[index];
                var actual = typed.Glyphs[index];
                Assert.Equal((raw.GlyphId, raw.ClusterStart, raw.ClusterLength),
                    (actual.GlyphId, actual.ClusterStart, actual.ClusterLength));
                // Independent legacy formulas, including their original double pen and zero signs.
                var x = -boundaryLeft + (oldPenX + raw.XOffset) * scale;
                var y = line.BaselineMillimeters - line.TopMillimeters - boundaryTop - (oldPenY + raw.YOffset) * scale;
                SameBits(x, actual.OriginXMillimeters); SameBits(y, actual.OriginYMillimeters);
                SameBits(raw.XAdvance * scale, actual.XAdvanceMillimeters);
                SameBits(raw.YAdvance * scale, actual.YAdvanceMillimeters);
                SameBits(strict.Glyphs[index].OriginXMillimeters, actual.OriginXMillimeters);
                SameBits(strict.Glyphs[index].OriginYMillimeters, actual.OriginYMillimeters);
                SameBits(strict.Glyphs[index].XAdvanceMillimeters, actual.XAdvanceMillimeters);
                Assert.Equal(strict.Glyphs[index].YAdvanceMillimeters, actual.YAdvanceMillimeters);
                expectedOrigins.Add(x);
                oldPenX += raw.XAdvance; oldPenY += raw.YAdvance;
            }
            var delta = expectedOrigins.Count > 1 ? string.Join(" ", expectedOrigins.Skip(1)
                .Select((x, index) => (x - expectedOrigins[index]).ToString("0.###############", CultureInfo.InvariantCulture))) : null;
            Assert.Equal(delta, output.DeltaX);
            Assert.Null(output.DeltaY);
            // Typed packaging reconstructs the Delta values with its existing three-decimal formatter.
            expectedDeltas.Add(expectedOrigins.Count > 1 ? string.Join(" ", expectedOrigins.Skip(1)
                .Select((x, index) => (x - expectedOrigins[index]).ToString("0.###", CultureInfo.InvariantCulture))) : null);
        }
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
        {
            using var page = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal)).Open();
            var xml = XDocument.Load(page);
            var codes = xml.Descendants().Where(node => node.Name.LocalName == "TextCode").ToArray();
            Assert.Equal(lines.Length, codes.Length);
            for (var index = 0; index < lines.Length; index++)
            {
                Assert.Equal(elements[index].Text, codes[index].Value);
                Assert.Equal(expectedDeltas[index], codes[index].Attribute("DeltaX")?.Value);
            }
            var gids = xml.Descendants().Where(node => node.Name.LocalName == "Glyphs").Select(node => node.Value);
            Assert.Equal(lines.Select(line => string.Join(" ", line.ShapedRun.Glyphs.Select(glyph => glyph.GlyphId))), gids);
        }
        stream.Position = 0;
        var read = await new OfdReader().ReadAsync(stream);
        Assert.Equal(lines.Select(line => source.Substring(line.SourceStart, line.SourceLength)),
            read.Pages.Single().Elements.OfType<OfdTextElement>().Select(element => element.Text));
    }

    [Fact]
    public async Task ControlledCompensatedPenRetainsWriterCompatibilityWithoutWeakeningStrictQuery()
    {
        // A replacement shaper boundary, not a claim about native font geometry or valid visuals.
        var layout = new TextPrelayoutEngine(new CompensatedShaper()).Layout(
            new("AVA", new[] { Load("DejaVuSans-R22.ttf") }, 4, 1_000_000));
        var shape = layout.Lines.Single().ShapedRun;
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            TextPositionedRun.Create(shape, 4)).Error);
        var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout, 1_000_000, 297, 0, 10);
        var run = package.Pages.Single().Elements.OfType<OfdTextElement>().Single().Runs.Single();
        Assert.Equal(new[] { 0d, 0d, 750_000d }, run.PositionedGlyphs!.Glyphs.Select(glyph => glyph.OriginXMillimeters));
        Assert.Equal(new[] { 750_000d, 750_000d, -750_000d }, run.PositionedGlyphs.Glyphs.Select(glyph => glyph.XAdvanceMillimeters));
        Assert.All(run.PositionedGlyphs.Glyphs, glyph => SameBits(0d, glyph.YAdvanceMillimeters));
        Assert.Equal("0 750000", run.DeltaX);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        Assert.True(stream.Length > 0);
        stream.Position = 0;
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        using var page = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal)).Open();
        var code = XDocument.Load(page).Descendants().Single(node => node.Name.LocalName == "TextCode");
        Assert.Equal("AVA", code.Value); Assert.Equal("0 750000", code.Attribute("DeltaX")?.Value);
    }

    [Fact]
    public void StrictVerticalIntMinimumIsNegatedAfterConversionToDouble()
    {
        var shape = new TextShapedRun(new("AV", 0, Load("DejaVuSans-R22.ttf")), new[]
        {
            new TextShapedGlyph(36, 0, 1, 0, int.MinValue, 0, 0, null, false),
            new TextShapedGlyph(57, 1, 1, 0, int.MaxValue, 0, 0, null, false),
        }, "controlled-vertical-min");
        var result = TextPositionedRun.Create(shape, 0.0001);
        SameBits(-((double)int.MinValue) * (0.0001 / 2048), result.Glyphs[0].YAdvanceMillimeters);
        Assert.True(result.YAdvanceMillimeters > 0);
    }

    private static void SameBits(double expected, double actual) =>
        Assert.Equal(BitConverter.DoubleToInt64Bits(expected), BitConverter.DoubleToInt64Bits(actual));

    private sealed class CompensatedShaper : ITextShaper
    {
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default) => new(request,
            new[]
            {
                new TextShapedGlyph(36, request.SourceStart, 1, 384_000_000, 0, 0, 0, null, false),
                new TextShapedGlyph(57, request.SourceStart + 1, 1, 384_000_000, 0, -384_000_000, 0, null, false),
                new TextShapedGlyph(36, request.SourceStart + 2, 1, -384_000_000, 0, -384_000_000, 0, null, false),
            }, "controlled-nonnative-compensated-pen");
    }

    private static TextFontFace Load(string name)
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", name);
            if (File.Exists(path)) return new TextFontFace(name, File.ReadAllBytes(path));
        }
        throw new FileNotFoundException("Materialize existing deterministic font: " + name);
    }
}
