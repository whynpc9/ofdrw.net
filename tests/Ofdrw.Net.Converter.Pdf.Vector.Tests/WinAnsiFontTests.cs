using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class WinAnsiFontTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));

    [Fact]
    public async Task EmbeddedNonsymbolicWinAnsiRetainsActualUnicodeAndFontBytes()
    {
        using var fixture = new PdfFixture(Font);
        var text = fixture.TextWinAnsi("A  B", 25, 95, 14);
        using var input = new MemoryStream(fixture.Create(new[] { "0.1 0.2 0.3 rg " + text }, simpleWinAnsi: true));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(Assert.Single(result.Pages).IsNative);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("A  B", Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Equal(Font, Assert.Single(package.Fonts).Data);
    }

    [Theory]
    [InlineData("/MacRomanEncoding", 32, false, "FONT_PROFILE")]
    [InlineData("/WinAnsiEncoding", 4, false, "FONT_PROFILE")]
    [InlineData("/WinAnsiEncoding", 32, true, "SOURCE_GLYPH_CMAP_MISMATCH")]
    public async Task EncodingOrUnicodeMismatchRejectsWholePage(string encoding, int flags, bool wrongUnicode, string code)
    {
        using var fixture = new PdfFixture(Font);
        var text = fixture.TextWinAnsi("AB", 25, 95);
        using var input = new MemoryStream(fixture.Create(new[] { "20 20 30 30 re f " + text }, simpleWinAnsi: true,
            encodingOverride: encoding, fontDescriptorFlags: flags, wrongUnicode: wrongUnicode));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains(code, error.Message); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData((byte)0xA0, ' ')]
    [InlineData((byte)0xAD, '-')]
    public async Task DuplicateWinAnsiCodesCannotCertifySourceGlyph(byte code, char unicode)
    {
        using var fixture = new PdfFixture(Font);
        var content = fixture.TextWinAnsi("A", 25, 95) + fixture.TextWinAnsiCode(code, unicode, 50, 95);
        using var input = new MemoryStream(fixture.Create(new[] { content }, simpleWinAnsi: true));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("SOURCE_GLYPH_CMAP_MISMATCH", error.Message);
        Assert.Equal(0, output.Length);
    }
}
