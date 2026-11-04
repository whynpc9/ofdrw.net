using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class DocxUnderlineTests
{
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Native_UnderlinedRun_GetsOneMatchingStrokePath(bool explicitNative)
    {
        var page = await ConvertFirstPage("""
            <w:p><w:r><w:t xml:space="preserve">Plain </w:t></w:r><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Under</w:t></w:r><w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>
            <w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>NoLine</w:t></w:r></w:p>
            """, explicitNative);
        var texts = page.Elements.OfType<OfdTextElement>().ToArray();
        var strokes = page.Elements.OfType<OfdPathElement>().Where(path => path.Stroke).ToArray();
        var under = Assert.Single(texts, text => text.Text == "Under");

        var line = Assert.Single(strokes);
        Assert.Equal(under.XMillimeters, line.XMillimeters, 3);
        Assert.Equal(under.WidthMillimeters, line.WidthMillimeters, 3);
        Assert.Equal(under.YMillimeters, line.YMillimeters, 3);
        foreach (var other in texts.Where(text => text.Text is "Plain " or "NoLine"))
            Assert.DoesNotContain(strokes, path => Math.Abs(path.XMillimeters - other.XMillimeters) < 0.001 &&
                Math.Abs(path.YMillimeters - other.YMillimeters) < 0.001);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Native_WrappedUnderlinedRun_GetsOnePathPerLine(bool explicitNative)
    {
        var page = await ConvertFirstPage("""
            <w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underlined words wrap here</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="1800" w:h="6000"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>
            """, explicitNative);
        var underlined = page.Elements.OfType<OfdTextElement>().Where(text => text.Text.Trim().Length > 0).ToArray();
        var strokes = page.Elements.OfType<OfdPathElement>().Where(path => path.Stroke).ToArray();

        Assert.True(underlined.Select(text => text.YMillimeters).Distinct().Count() >= 2, "The run must wrap for this test.");
        Assert.Equal(underlined.Length, strokes.Length);
        foreach (var text in underlined)
            Assert.Single(strokes, path => Math.Abs(path.XMillimeters - text.XMillimeters) < 0.001 &&
                Math.Abs(path.YMillimeters - text.YMillimeters) < 0.001);
    }

    private static async Task<OfdPage> ConvertFirstPage(string body, bool explicitNative)
    {
        await using var input = DocxReleaseFixtures.CreateMinimalDocx(body);
        await using var output = new MemoryStream();
        var converter = explicitNative
            ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
            : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output);
        output.Position = 0;
        return (await new OfdReader().ReadAsync(output)).Pages[0];
    }
}
