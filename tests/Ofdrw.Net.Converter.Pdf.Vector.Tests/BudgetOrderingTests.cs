using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class BudgetOrderingTests
{
    private const string Normal = "10 10 20 20 re f ";
    private const string Collapsed = "q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 100 l S Q";
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static PdfVectorToOfdOptions Options(PdfVectorUnsupportedPagePolicy policy, int events) => new()
    { MaxEventsPerPage = events, UnsupportedPagePolicy = policy, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    private static readonly PdfVectorUnsupportedPagePolicy[] Policies = { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage };

    public static IEnumerable<object[]> PathProfiles()
    {
        yield return new object[] { Collapsed, "PATH_FLOAT_PRECISION", 1 };
        yield return new object[] { "q 1 0 0 1 -16777216 0 cm 16777216 100 1 40 re f Q", "PATH_FLOAT_PRECISION", 1 };
        yield return new object[] { "q 1 0 0 1 -16777216 0 cm 16777216 100 1 40 re B Q", "PATH_FLOAT_PRECISION", 2 };
        yield return new object[] { "q 1 2 2 4 0 0 cm 20 20 m 40 40 l S Q", "SINGULAR_SERIALIZED_MATRIX", 1 };
        yield return new object[] { "1 J 5 M 20 20 m 40 40 l S", "STROKE_PROFILE", 1 };
        yield return new object[] { "0 w 20 20 m 40 40 l S", "STROKE_PROFILE", 1 };
        yield return new object[] { ".0004 w 20 20 m 40 40 l S", "STROKE_WRITER_PRECISION", 1 };
        yield return new object[] { ".0004 w 20 20 30 30 re B", "STROKE_WRITER_PRECISION", 2 };
        yield return new object[] { "20 20 m h f", "DEGENERATE_POINT_FILL", 1 };
        yield return new object[] { "20 20 m h B", "DEGENERATE_POINT_FILL", 2 };
    }

    [Theory]
    [MemberData(nameof(PathProfiles))]
    public async Task RealPaintCapacityPrecedesItsProfileAndEnoughCapacityKeepsPolicy(string body, string reason, int required)
    {
        using var fixture = Fixture(); var pdf = fixture.Create(new[] { Normal + body });
        foreach (var policy in Policies)
        {
            // One prior fill leaves required-1 slots, including the one-slot B case.
            await Hard(pdf, Options(policy, required), "draw event limit");
            var options = Options(policy, required + 1);
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            if (policy == PdfVectorUnsupportedPagePolicy.Fail)
            {
                output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
                Assert.Contains(reason, error.Message); Unchanged(output);
            }
            else
            {
                var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
                Assert.False(result.Pages[0].IsNative); Assert.Contains(reason, result.Pages[0].Diagnostic);
                Assert.Equal(1, result.Pages[0].ImageObjects); Assert.Equal(0, result.Pages[0].PathObjects);
                output.Position = 0; var page = (await new OfdReader().ReadAsync(output)).Pages[0];
                Assert.Single(page.Elements.OfType<OfdImageElement>()); Assert.Empty(page.Elements.OfType<OfdPathElement>());
            }
        }
    }

    [Theory]
    [InlineData("B", 0)][InlineData("B", 1)][InlineData("B", 2)]
    [InlineData("B*", 0)][InlineData("B*", 1)][InlineData("B*", 2)]
    [InlineData("b", 0)][InlineData("b", 1)][InlineData("b", 2)]
    [InlineData("b*", 0)][InlineData("b*", 1)][InlineData("b*", 2)]
    public async Task CombinedPaintRequiresBothSlotsAndPreservesOrder(string operation, int remaining)
    {
        using var fixture = Fixture();
        var prefix = string.Concat(Enumerable.Repeat(Normal, 2 - remaining));
        var pdf = fixture.Create(new[] { prefix + "20 20 m 50 20 l 50 50 l " + operation });
        foreach (var policy in Policies)
        {
            var options = Options(policy, 2);
            if (remaining < 2) { await Hard(pdf, options, "draw event limit"); continue; }
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
            Assert.True(result.Pages[0].IsNative); Assert.Equal(2, result.Pages[0].PathObjects);
            output.Position = 0; var paths = (await new OfdReader().ReadAsync(output)).Pages[0].Elements.OfType<OfdPathElement>().ToArray();
            Assert.Equal(2, paths.Length); Assert.True(paths[0].Fill); Assert.False(paths[0].Stroke);
            Assert.False(paths[1].Fill); Assert.True(paths[1].Stroke);
        }
    }

    [Theory]
    [InlineData("S f B")][InlineData("20 20 m S")][InlineData("20 20 m h S")][InlineData("20 20 m s")]
    [InlineData("20 20 m 40 40 m f")][InlineData("20 20 m B")][InlineData("20 20 m 40 40 l n")]
    [InlineData("BT /F1 12 Tf () Tj ET")][InlineData("BT /F1 12 Tf [] TJ ET")][InlineData("BT /F1 12 Tf ET")]
    [InlineData("BT /F1 12 Tf [125 -125] TJ ET")]
    public async Task NoOpAfterExactCapacityDoesNotRequireAnotherSlot(string noOp)
    {
        using var fixture = Fixture(); var pdf = fixture.Create(new[] { Normal + noOp });
        foreach (var policy in Policies)
        {
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(Options(policy, 1)).ConvertWithResultAsync(input, output);
            Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects);
            Assert.Equal(0, result.Pages[0].TextObjects); Assert.Equal(0, result.Pages[0].ImageObjects);
            output.Position = 0; Assert.Empty((await new OfdReader().ReadAsync(output)).Fonts);
        }
    }

    [Theory]
    [InlineData("1 J 20 20 m S")][InlineData("1 J 20 20 m h S")][InlineData("0 w 20 20 m S")]
    public async Task NonsegmentUnsupportedProfileRetainsItsPolicyWithoutEventCharge(string move)
    {
        using var fixture = Fixture(); var pdf = fixture.Create(new[] { Normal + move });
        using var failInput = new MemoryStream(pdf); using var sentinel = new MemoryStream(); sentinel.Write(new byte[] { 4, 5, 6 }); sentinel.Position = 0;
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.Fail, 1)).ConvertAsync(failInput, sentinel));
        Assert.Contains("STROKE_PROFILE", error.Message); Unchanged(sentinel);
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage, 1)).ConvertWithResultAsync(input, output);
        Assert.False(result.Pages[0].IsNative); Assert.Contains("STROKE_PROFILE", result.Pages[0].Diagnostic);
    }

    [Theory]
    [InlineData("unicode")][InlineData("combining")][InlineData("cmap")][InlineData("state")]
    [InlineData("matrix")][InlineData("float")][InlineData("whitespace")][InlineData("writer")]
    public async Task NonemptyGlyphAndRunProfilesCannotBypassExhaustedEvents(string kind)
    {
        using var fixture = Fixture();
        var body = kind switch
        {
            "unicode" => "BT /F1 12 Tf 1 0 0 1 30 100 Tm <0000> Tj ET",
            "state" => fixture.Text("A", 30, 100, 0),
            "matrix" => $"BT /F1 12 Tf 1 2 2 4 30 100 Tm <{fixture.Hex("A")}> Tj ET",
            "float" => fixture.Text("A", 16777217, 350),
            "whitespace" => fixture.Text(" ", 30, 100),
            "writer" => fixture.Text("A", 30, 100, .0004),
            _ => fixture.Text("A", 30, 100)
        };
        var pdf = fixture.Create(new[] { Normal + body }, aUnicodeHex: kind == "combining" ? "0301" : kind == "cmap" ? "0042" : null);
        foreach (var policy in Policies) await Hard(pdf, Options(policy, 1), "draw event limit");
    }

    [Theory]
    [InlineData(false)][InlineData(true)]
    public async Task AlreadyConsumedCharactersPrecedeTheNextGlyphProfile(bool unmapped)
    {
        using var fixture = Fixture();
        var bad = unmapped ? "BT /F1 12 Tf 1 0 0 1 30 120 Tm <0000> Tj ET" : fixture.Text("A", 30, 120, 0);
        var pdf = fixture.Create(new[] { fixture.Text("A", 30, 100) + bad });
        foreach (var policy in Policies)
        {
            var options = Options(policy, 2); options.MaxTextCharactersPerPage = 1;
            await Hard(pdf, options, "original text character limit");
        }
    }

    [Fact]
    public async Task SeveralGlyphsStillUseOneEventAndOnlyAdmittedGlyphsCount()
    {
        using var fixture = Fixture(); var pdf = fixture.Create(new[] { fixture.Text("A  B中文", 30, 100) });
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.Fail, 1)).ConvertWithResultAsync(input, output);
        Assert.Equal(1, result.Pages[0].TextObjects); Assert.True(result.Pages[0].IsNative);
        var options = Options(PdfVectorUnsupportedPagePolicy.RasterizePage, 1); options.MaxTextCharactersPerPage = 1;
        await Hard(fixture.Create(new[] { fixture.Text("AA", 30, 100) }), options, "original text character limit");
    }

    [Fact]
    public async Task CloseAndConstructionCommandsAndFontAttemptKeepTheirHardGates()
    {
        using var fixture = Fixture();
        foreach (var body in new[] { Normal + "20 20 m s", Normal + "20 20 m 40 40 l n" })
        {
            var options = Options(PdfVectorUnsupportedPagePolicy.RasterizePage, 1); options.MaxPathCommandsPerPage = 6;
            await Hard(fixture.Create(new[] { body }), options, "path command limit");
        }
        var fontOptions = Options(PdfVectorUnsupportedPagePolicy.RasterizePage, 1); fontOptions.MaxFontReadBytesPerPage = 1;
        await Hard(fixture.Create(new[] { Normal + fixture.Text("A", 30, 100) }), fontOptions, "font");
        using var input = new MemoryStream(fixture.Create(new[] { Normal + Collapsed })); using var output = new MemoryStream();
        output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage, 1))
            .ConvertAsync(input, output, cancellationToken: new CancellationToken(true))); Unchanged(output);
    }

    private static async Task Hard(byte[] pdf, PdfVectorToOfdOptions options, string message)
    {
        using var input = new MemoryStream(pdf); using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.Contains(message, error.Message); Unchanged(output);
    }
    private static void Unchanged(MemoryStream output) { Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray()); Assert.Equal(0, output.Position); }
}
