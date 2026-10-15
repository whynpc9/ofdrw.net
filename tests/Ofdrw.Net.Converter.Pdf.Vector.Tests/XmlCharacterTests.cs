using System.Xml;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SkiaSharp;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class XmlCharacterTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "xml-boundary.ttf"));
    private static byte[] Source(char value, int copies = 1, bool leadingPath = false)
    {
        using var fixture = new PdfFixture(Font);
        var middle = (leadingPath ? "0 0 1 rg 20 20 100 40 re f " : "") + fixture.Text(new string(value, copies), 30, 120);
        return fixture.Create(new[] { fixture.Text("Before 中文", 30, 160), middle, fixture.Text("End 字体", 30, 160) });
    }
    private static MemoryStream Sentinel() { var s = new MemoryStream(); s.Write(new byte[] { 4, 5, 6 }); s.Position = 0; return s; }
    private static void Unchanged(MemoryStream s) { Assert.Equal(new byte[] { 4, 5, 6 }, s.ToArray()); Assert.Equal(0, s.Position); }
    [Theory]
    [InlineData('\uFFFE', PdfTextLayerMode.Invisible)] [InlineData('\uFFFF', PdfTextLayerMode.Invisible)]
    [InlineData('\uFFFE', PdfTextLayerMode.None)] [InlineData('\uFFFF', PdfTextLayerMode.None)]
    public async Task FailClassifiesActualXmlInvalidGlyphAndPreservesOutput(char value, PdfTextLayerMode layer)
    {
        var source = Source(value);
        using var data = SKData.CreateCopy(Font); using var face = SKTypeface.FromData(data);
        var gid = face.GetGlyph(value); Assert.NotEqual(0, gid);
        using (var pdf = PdfDocument.Open(source))
        {
            Assert.Equal(value.ToString(), pdf.GetPage(2).Text);
            using var font = new SKFont(face, 1000); using var outline = font.GetGlyphPath(gid);
            Assert.NotNull(outline); Assert.False(outline.IsEmpty);
        }
        var options = new PdfVectorToOfdOptions(); options.Compatibility.TextLayerMode = layer;
        using var input = new MemoryStream(source); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.Contains("TEXT_XML_CHARACTER", error.Message); Unchanged(output);
    }
    [Theory] [InlineData('\uFFFE')] [InlineData('\uFFFF')]
    public async Task ExplicitNoTextLayerRasterizesWholeMiddlePageAndPreservesNativeNeighbors(char value)
    {
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage };
        options.Compatibility.TextLayerMode = PdfTextLayerMode.None;
        using var input = new MemoryStream(Source(value)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(p => p.IsNative));
        Assert.Contains("TEXT_XML_CHARACTER", result.Pages[1].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(3, package.Pages.Count);
        Assert.Single(package.Pages[1].Elements.OfType<OfdImageElement>());
        Assert.Empty(package.Pages[1].Elements.OfType<OfdTextElement>()); Assert.Empty(package.Pages[1].Elements.OfType<OfdPathElement>());
        Assert.Equal("Before 中文", string.Concat(package.Pages[0].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
        Assert.Equal("End 字体", string.Concat(package.Pages[2].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
    }
    [Theory] [InlineData('\uFFFE')] [InlineData('\uFFFF')]
    public async Task DefaultInvisibleFallbackStillCannotSerializeXmlInvalidOriginalText(char value)
    {
        using var input = new MemoryStream(Source(value)); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<ArgumentException>(() => new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage }).ConvertAsync(input, output));
        Assert.Contains("invalid character", error.Message); Unchanged(output);
    }
    [Theory] [InlineData('\uD7FF')] [InlineData('\uE000')] [InlineData('\uFDD0')] [InlineData('\uFFFD')]
    public async Task LegalBmpBoundariesRemainExactNativeText(char value)
    {
        Assert.True(XmlConvert.IsXmlChar(value)); using var input = new MemoryStream(Source(value)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output); Assert.All(result.Pages, p => Assert.True(p.IsNative));
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(value.ToString(), string.Concat(package.Pages[1].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
        Assert.All(package.Fonts, f => Assert.Equal(Font, f.Data));
    }
    [Theory]
    [InlineData('\uFFFE', false)] [InlineData('\uFFFF', false)] [InlineData('\uFFFE', true)] [InlineData('\uFFFF', true)]
    public async Task ExhaustedCharacterOrEventBudgetRemainsHardFailureBeforeXmlProfile(char value, bool eventBudget)
    {
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            MaxTextCharactersPerPage = eventBudget ? 100000 : 1, MaxEventsPerPage = eventBudget ? 1 : 10000 };
        options.Compatibility.TextLayerMode = PdfTextLayerMode.None;
        using var input = new MemoryStream(Source(value, eventBudget ? 1 : 2, eventBudget)); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output, new[] { 1 }));
        Assert.Contains(eventBudget ? "event limit" : "character limit", error.Message); Unchanged(output);
    }
    [Theory] [InlineData(PdfVectorUnsupportedPagePolicy.Fail)] [InlineData(PdfVectorUnsupportedPagePolicy.RasterizePage)]
    public async Task CancellationRemainsHardFailure(PdfVectorUnsupportedPagePolicy policy)
    {
        using var input = new MemoryStream(Source('\uFFFF')); using var output = Sentinel(); using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = policy }).ConvertAsync(input, output, cancellationToken: cancellation.Token));
        Unchanged(output);
    }
    [Theory] [InlineData(9)] [InlineData(10)] [InlineData(13)] [InlineData(31)]
    public async Task ExistingControlPolicyRemainsUnsupported(int code)
    {
        using var input = new MemoryStream(Source((char)code)); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("UNICODE_MAPPING", error.Message); Unchanged(output);
    }
    [Theory] [InlineData("D800DC00")] [InlineData("DBFFDFFF")]
    public async Task ActualSurrogatePairMappingRetainsExistingProfile(string unicodeHex)
    {
        using var fixture = new PdfFixture(Font); var source = fixture.Create(new[] { fixture.Text("A", 30, 120) }, aUnicodeHex: unicodeHex);
        using (var pdf = PdfDocument.Open(source)) Assert.Equal(2, pdf.GetPage(1).Letters.Single().Value.Length);
        using var input = new MemoryStream(source); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("UNICODE_MAPPING", error.Message); Unchanged(output);
    }
}
