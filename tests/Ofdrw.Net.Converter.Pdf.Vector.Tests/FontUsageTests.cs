using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class FontUsageTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));
    private const string PathContent = "0 0 1 rg 20 20 30 40 re f ";
    private const string Aliases = "/F1 3 0 R /F2 3 0 R";
    private static PdfVectorToOfdOptions Raster() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage };
    private static async Task<OfdDocumentPackage> Read(byte[] pdf, PdfVectorToOfdOptions options, int[]? selection = null)
    {
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output, selection);
        Assert.All(result.Pages, page => Assert.True(page.IsNative));
        output.Position = 0; return await new OfdReader().ReadAsync(output);
    }
    private static async Task Fatal(byte[] pdf, PdfVectorToOfdOptions options, int[]? selection = null)
    {
        using var input = new MemoryStream(pdf); using var output = new MemoryStream(); output.Write(new byte[] { 7, 8, 9 });
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output, selection));
        Assert.Equal(new byte[] { 7, 8, 9 }, output.ToArray());
    }
    [Theory]
    [InlineData("BT /F1 12 Tf <> Tj ET")]
    [InlineData("BT /F1 12 Tf () Tj ET")]
    [InlineData("BT /F1 12 Tf [] TJ ET")]
    [InlineData("BT /F1 12 Tf [<> ()] TJ ET")]
    [InlineData("BT /F1 12 Tf ET")]
    [InlineData("BT /F1 12 Tf /F2 24 Tf ET")]
    [InlineData("BT /F1 12 Tf [100 <> -100 () 0] TJ ET")]
    public async Task NonpaintingTextDoesNotReadOrEmbedFontPrograms(string text)
    {
        using var fixture = new PdfFixture(Font);
        var options = Raster(); options.MaxFontBytes = 1; options.MaxFontReadBytesPerPage = 1; options.MaxTotalFontBytes = 1;
        var package = await Read(fixture.Create(new[] { PathContent + text }, fontResources: Aliases), options);
        Assert.Empty(package.Fonts); Assert.Single(package.Pages[0].Elements.OfType<OfdPathElement>());
        Assert.Empty(package.Pages[0].Elements.OfType<OfdTextElement>());
    }
    [Fact]
    public async Task RepeatedUnusedSelectionsDoNotConsumeActualOutputAllowance()
    {
        using var fixture = new PdfFixture(Font);
        var pdf = fixture.Create(new[] { PathContent + "BT /F1 12 Tf <> Tj ET", PathContent, fixture.Text("A  B中文", 30, 90) });
        var options = Raster(); options.MaxTotalFontBytes = Font.Length; options.MaxFontReadBytesPerPage = Font.Length;
        var package = await Read(pdf, options, new[] { 0, 0, 1, 2 });
        Assert.Equal(4, package.Pages.Count); Assert.Equal(Font, package.Fonts.Single().Data);
        Assert.Equal("A  B中文", package.Pages[3].Elements.OfType<OfdTextElement>().Single().Text);
        await Fatal(pdf, options, new[] { 2, 2 });
    }
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task OnlySuccessfulTextBindingsReachTheOutput(bool both)
    {
        using var fixture = new PdfFixture(Font); var a = fixture.Hex("A");
        var content = "BT /F1 12 Tf 1 0 0 1 30 90 Tm " + (both ? $"<{a}> Tj " : "<> Tj ") + $"/F2 12 Tf <{a}> Tj ET";
        var package = await Read(fixture.Create(new[] { content }, fontResources: Aliases), new());
        Assert.Equal(both ? 2 : 1, package.Fonts.Count); Assert.All(package.Fonts, font => Assert.Equal(Font, font.Data));
        Assert.Equal(both ? 2 : 1, package.Pages[0].Elements.OfType<OfdTextElement>().Count());
    }
    [Fact]
    public async Task AliasLoadsHaveIndependentAttemptAllowanceAndSameNameReuseDoesNotRecharge()
    {
        using var fixture = new PdfFixture(Font); var a = fixture.Hex("A");
        var one = fixture.Create(new[] { $"BT /F1 12 Tf 1 0 0 1 30 90 Tm <{a}> Tj <{a}> Tj ET" });
        var options = Raster(); options.MaxFontReadBytesPerPage = Font.Length; options.MaxTotalFontBytes = Font.Length * 2L;
        var package = await Read(one, options); Assert.Single(package.Fonts); Assert.Equal(2, package.Pages[0].Elements.OfType<OfdTextElement>().Count());
        var two = fixture.Create(new[] { $"BT /F1 12 Tf 1 0 0 1 30 90 Tm <{a}> Tj /F2 12 Tf <{a}> Tj ET" }, fontResources: Aliases);
        await Fatal(two, options);
        options.MaxFontReadBytesPerPage = Font.Length * 2L;
        Assert.Equal(2, (await Read(two, options)).Fonts.Count);
    }
    [Theory]
    [InlineData(-1)] [InlineData(0)]
    public async Task AttemptAllowanceMustBePositive(long value)
    {
        using var input = new MemoryStream(); using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => new PdfVectorToOfdConverter(new() { MaxFontReadBytesPerPage = value }).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task AttemptAllowanceIsSnapshottedAndExactBoundaryIsAccepted()
    {
        using var fixture = new PdfFixture(Font); var pdf = fixture.Create(new[] { fixture.Text("A", 30, 90) });
        var options = Raster(); options.MaxFontReadBytesPerPage = Font.Length - 1;
        await Fatal(pdf, options); options.MaxFontReadBytesPerPage = Font.Length;
        var snapshot = VectorLimits.Snapshot(options); options.MaxFontReadBytesPerPage = 1;
        Assert.Equal(Font.Length, snapshot.MaxFontReadBytesPerPage);
        var converter = new PdfVectorToOfdConverter(snapshot);
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        await converter.ConvertAsync(input, output); Assert.True(output.Length > 0);
    }
    [Theory]
    [InlineData("[100 <> -50 () 25]", "59.1")]
    [InlineData("[100 -100]", "60")]
    [InlineData("[0 <> ()]", "60")]
    public async Task NumericOnlyTextArrayPreservesFollowingTextPosition(string array, string expectedX)
    {
        using var fixture = new PdfFixture(Font); var a = fixture.Hex("A");
        var candidate = await Read(fixture.Create(new[] { $"BT /F1 12 Tf 1 0 0 1 60 90 Tm {array} TJ <{a}> Tj ET" }), new());
        var control = await Read(fixture.Create(new[] { $"BT /F1 12 Tf 1 0 0 1 {expectedX} 90 Tm <{a}> Tj ET" }), new());
        var actual = candidate.Pages[0].Elements.OfType<OfdTextElement>().Single(); var expected = control.Pages[0].Elements.OfType<OfdTextElement>().Single();
        Assert.Equal(expected.Text, actual.Text); Assert.Equal(expected.Transform, actual.Transform);
        Assert.Equal(expected.FontSizeMillimeters, actual.FontSizeMillimeters); Assert.Equal(expected.Runs[0].XMillimeters, actual.Runs[0].XMillimeters, 6);
    }
    [Fact]
    public async Task NumericTextAndGraphicsRestoreUsesTheActualCurrentFontState()
    {
        using var fixture = new PdfFixture(Font); var a = fixture.Hex("A");
        var body = $"BT /F1 12 Tf 1 0 0 1 60 90 Tm q /F2 24 Tf [100] TJ Q <{a}> Tj ET";
        var candidate = await Read(fixture.Create(new[] { body }, fontResources: Aliases), new());
        var control = await Read(fixture.Create(new[] { fixture.Text("A", 57.6, 90) }), new());
        Assert.Single(candidate.Fonts);
        var actual = candidate.Pages[0].Elements.OfType<OfdTextElement>().Single(); var expected = control.Pages[0].Elements.OfType<OfdTextElement>().Single();
        Assert.Equal(expected.Transform, actual.Transform); Assert.Equal(expected.FontSizeMillimeters, actual.FontSizeMillimeters);
    }
    [Theory]
    [InlineData("100 0 R", null)]
    [InlineData("11 0 R", new[] { "12 0 R", "11 0 R" })]
    [InlineData("99", null)]
    public async Task UnusedBodyIsDeferredButNumericOrGlyphConsumptionKeepsStrictStructure(string encoding, string[]? objects)
    {
        using var fixture = new PdfFixture(Font);
        var unused = fixture.Create(new[] { PathContent + "BT /F1 12 Tf <> Tj ET" }, encodingOverride: encoding, trailingObjects: objects);
        Assert.Empty((await Read(unused, Raster())).Fonts);
        foreach (var body in new[] { PathContent + "BT /F1 12 Tf [100] TJ ET", fixture.Text("A", 30, 90) })
            await Fatal(fixture.Create(new[] { body }, encodingOverride: encoding, trailingObjects: objects), Raster());
    }
    [Fact]
    public async Task UnknownWritingDirectionIsUnsupportedEvenWithoutGlyphs()
    {
        using var fixture = new PdfFixture(Font);
        using var input = new MemoryStream(fixture.Create(new[] { PathContent + "BT /F1 12 Tf [100] TJ ET" }, encodingOverride: "/Identity-V"));
        using var output = new MemoryStream(); output.Write(new byte[] { 7, 8, 9 });
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Equal(new byte[] { 7, 8, 9 }, output.ToArray());
    }
    [Fact]
    public async Task DeclaredFontCountRemainsFatalOnUnusedFonts()
    {
        using var fixture = new PdfFixture(Font); var options = Raster(); options.MaxFontsPerPage = 1;
        await Fatal(fixture.Create(new[] { PathContent + "BT /F1 12 Tf ET" }, fontResources: Aliases), options);
    }
    [Theory]
    [InlineData("\\376\\377")]
    [InlineData("\\377\\376")]
    public async Task BomOnlyLiteralHasRealGlyphBytesAndCannotBeDroppedAsAnEmptyToken(string literal)
    {
        using var fixture = new PdfFixture(Font);
        using var input = new MemoryStream(fixture.Create(new[] { PathContent + "BT /F1 12 Tf [(" + literal + ")] TJ ET" }));
        using var output = new MemoryStream(); output.Write(new byte[] { 7, 8, 9 });
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("UNICODE_MAPPING", error.Message); Assert.Equal(new byte[] { 7, 8, 9 }, output.ToArray());
    }
}
