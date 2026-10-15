using System.Text;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SkiaSharp;
using UglyToad.PdfPig.Core;
using Binary = Ofdrw.Net.Converter.Pdf.Vector.PathFloatPrecision.Binary;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class TextPrecisionTests
{
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static PdfVectorToOfdOptions Raster() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    [Theory]
    [InlineData("translation")][InlineData("advance")][InlineData("affine")][InlineData("accumulation")][InlineData("size")]
    public async Task ActualTextPlacementAndShapeLossHonorAtomicWholePagePolicy(string kind)
    {
        using var fixture = Fixture(); var glyph = fixture.Hex("A");
        // Real package E2E additionally binds these losses to the unchanged full Noto source.
        var normal = fixture.Text("A  B中文", 20, 85);
        var bad = kind switch {
            "translation" => fixture.Text("A", 16777217, 350),
            "advance" => $"BT /F1 16 Tf 1 0 0 1 100000000 350 Tm [<{glyph}> 6249997250 <{glyph}>] TJ ET",
            "affine" => $"BT /F1 .125 Tf 128 0 0 128 16777216 350 Tm [<{glyph}> 1048573234.375 <{glyph}>] TJ ET",
            "size" => fixture.Text("A", 60, 350, 8192.0003),
            _ => Accumulation(fixture.Hex("中"))
        };
        var source = fixture.Create(new[] { normal, normal + "10 10 20 20 re f " + bad, normal });
        using var strictInput = new MemoryStream(source); using var sentinel = new MemoryStream(); sentinel.Write(new byte[] { 7, 8 });
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(strictInput, sentinel));
        Assert.Contains("TEXT_FLOAT_PRECISION", error.Message); Assert.Equal(new byte[] { 7, 8 }, sentinel.ToArray());
        using var input = new MemoryStream(source); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Raster()).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative));
        Assert.Equal(0, result.Pages[1].PathObjects); Assert.Equal(1, result.Pages[1].ImageObjects); Assert.Contains("TEXT_FLOAT_PRECISION", result.Pages[1].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), t => Assert.Equal(0, t.FillColor.Alpha));
        Assert.Equal("A  B中文", package.Pages[2].Elements.OfType<OfdTextElement>().Single().Text);
        var nativeFonts = package.Pages.Where((_, i) => i != 1).SelectMany(p => p.Elements.OfType<OfdTextElement>()).Select(t => t.FontResourceId).ToHashSet();
        Assert.All(package.Fonts.Where(f => f.Data.Length > 0), f => Assert.Contains(f.Id, nativeFonts));
    }
    private static string Accumulation(string glyph)
    {
        var s = new StringBuilder($"BT /F1 16 Tf 1 0 0 1 60 350 Tm [<{glyph}>");
        for (var i = 0; i < 1024; i++) s.Append($" -5250.0001875 <{glyph}> 7250 <{glyph}>");
        return s.Append("] TJ ET").ToString();
    }
    [Theory]
    [InlineData("ordinary")][InlineData("shear")][InlineData("reflection")][InlineData("rotation")][InlineData("large-exact")][InlineData("spacing")]
    public async Task OrdinaryAndExactlyRepresentableTransformsAndAdvancesStayNative(string kind)
    {
        using var fixture = Fixture();
        var body = kind switch {
            "ordinary" => fixture.Text("Ordinary A  B中文", 34.1, 535.2, 15.3),
            "shear" => fixture.Text("Sheared A  B中文", 34, 480, 15, .2),
            "reflection" => "q -1 0 0 1 400 0 cm " + fixture.Text("A  B中文", 34, 480) + " Q",
            "rotation" => "q 0 1 -1 0 350 10 cm " + fixture.Text("A  B中文", 34, 180) + " Q",
            "large-exact" => $"BT /F1 16 Tf 1 0 0 1 16777216 350 Tm [<{fixture.Hex("中")}> 1048574250 <{fixture.Hex("中")}>] TJ ET",
            _ => $"BT /F1 16 Tf 1 0 0 1 34 380 Tm 2 Tc [<{fixture.Hex("A  B中文")}> 125 <{fixture.Hex(" C")}>] TJ ET"
        };
        using var input = new MemoryStream(fixture.Create(new[] { body })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(0, result.Pages[0].ImageObjects); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output); Assert.Contains("中", package.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
    }
    [Fact]
    public async Task TextCharacterLimitStaysFatalUnderRasterPolicy()
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("A  B中文", 20, 85) })); using var output = new MemoryStream();
        var options = Raster(); options.MaxTextCharactersPerPage = 4;
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }

    private static TransformationMatrix T(double x = 0, double y = 0) => TransformationMatrix.FromValues(1, 0, 0, 1, x, y);
    private static readonly PdfRectangle Em = new(0, 0, 1, 1);
    private static TextFloatPrecision Guard(float size = 16) => new(new SKMatrix(1, 0, 0, 0, 1, 595, 0, 0, 1), size, 595);
    [Theory]
    [InlineData(0.00006103515625, true)][InlineData(0.0001220703125, false)]
    public void DyadicPositionErrorsOnEitherSideOfAbsoluteBound(double error, bool expected) =>
        Assert.Equal(expected, Guard().Glyph(T(error), T(), 16, Em, 0));
    [Theory]
    [InlineData(0, true)][InlineData(0.0000000001, false)]
    public void ZeroSpacingIsAllowedAndTinyDistinctSpacingCannotCollapse(double x, bool expected)
    {
        var guard = Guard(); Assert.True(guard.Glyph(T(), T(), 16, Em, 0)); Assert.Equal(expected, guard.Glyph(T(x), T(), 16, Em, 0));
    }
    [Fact]
    public void ExactCompositionAndSequentialDoublePrefixCannotSelfCertifyLostUnit()
    {
        var guard = Guard(); var huge = Math.Pow(2, 60);
        Assert.True(guard.Glyph(T(), T(), 16, Em, 0)); Assert.True(guard.Glyph(T(huge), T(), 16, Em, (float)huge));
        Assert.False(guard.Glyph(T(huge), T(1), 16, Em, 1));
        var exact = Binary.From(huge) + Binary.One - Binary.From(huge); Assert.Equal(0, exact.CompareTo(Binary.One));
        Assert.Equal(0, huge + 1 - huge);
    }
    [Fact]
    public void IndividuallySmallAdvancesAreCheckedAtEachCumulativeAnchor()
    {
        var guard = Guard(); Assert.True(guard.Glyph(T(), T(), 16, Em, 0)); var x = 0d; var rejected = false;
        for (var i = 0; i < 40; i++)
        {
            x += 100.000003; if (!guard.Glyph(T(x), T(), 16, Em, 100)) { rejected = true; break; }
            x -= 100; Assert.True(guard.Glyph(T(x), T(), 16, Em, -100));
        }
        Assert.True(rejected);
    }
    [Theory]
    [InlineData(8192.0003)][InlineData(1e-50)][InlineData(1e39)]
    public void ShapeEnvelopeOrFontRangeRejectsLossEvenWithExactBaseline(double size) =>
        Assert.False(Guard((float)size).Glyph(T(), T(), size, Em, 0));
    [Fact]
    public void NearSingularMapCannotDistortGlyphDirections()
    {
        var text = TransformationMatrix.FromValues(1, 1, 1, 1.00001, 0, 0);
        var matrix = new SKMatrix(1, -1, 0, -1, (float)1.00001, 595, 0, 0, 1);
        Assert.False(new TextFloatPrecision(matrix, 16, 595).Glyph(text, T(), 16, Em, 0));
    }
}
