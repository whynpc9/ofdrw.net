using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class DocxJustifyTests
{
    private const string Section = """<w:sectPr><w:pgSz w:w="1800" w:h="6000"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>""";
    private const double RightEdge = 80 * 25.4 / 72;
    private const double LeftEdge = 10 * 25.4 / 72;
    private const string Latin = "The quick brown fox jumps over the lazy dog while seven wizards quietly judge boxing matches and " +
        "a jovial sphinx of black quartz hears every vow before the final bell rings tonight.";
    private const string Cjk = "中华人民共和国国家标准规定了电子文件版式文档格式的技术要求适用于版式文档的生成交换与长期保存" +
        "本标准由全国信息技术标准化技术委员会提出并归口管理起草单位包括多家科研院所和软件企业共同完成";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Latin_JustifiedParagraph_ShouldFillWrappedLines(bool explicitNative) =>
        AssertJustified(await Convert(Paragraph("both", Latin), explicitNative), Latin);

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Cjk_JustifiedParagraph_ShouldFillWrappedLines(bool explicitNative) =>
        AssertJustified(await Convert(Paragraph("both", Cjk), explicitNative), Cjk);

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ExplicitBreak_ShouldNotBeStretched(bool explicitNative)
    {
        var page = await Convert("""<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>Alpha</w:t><w:br/><w:t>Beta</w:t></w:r></w:p>""", explicitNative);
        var lines = Lines(page);

        Assert.Equal(2, lines.Length);
        Assert.All(lines, line => Assert.Equal(LeftEdge, line.Min(text => text.XMillimeters), 3));
        Assert.All(lines, line => Assert.True(line.Max(text => text.XMillimeters + text.WidthMillimeters) < RightEdge - 1));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task LeftAlignment_Unchanged(bool explicitNative)
    {
        var lines = Lines(await Convert(Paragraph(null, Latin), explicitNative));

        Assert.True(lines.Length > 2);
        Assert.All(lines, line => Assert.Equal(LeftEdge, line.Min(text => text.XMillimeters), 3));
        Assert.Contains(lines.Take(lines.Length - 1), line => line.Max(text => text.XMillimeters + text.WidthMillimeters) < RightEdge - 0.5);
    }

    private static void AssertJustified(OfdPage page, string source)
    {
        var lines = Lines(page);
        Assert.True(lines.Length > 2, "The paragraph must wrap for this test.");
        var filled = 0;
        foreach (var line in lines.Take(lines.Length - 1))
        {
            var end = line.Max(text => text.XMillimeters + text.WidthMillimeters);
            // A wrapped line holding one Latin word has no gap to stretch and stays left, as in Word.
            if (string.Concat(line.Select(text => text.Text)).Trim().Contains(' ') || line.Any(text => text.Text.Any(c => c >= 0x2E80)))
            {
                Assert.InRange(end, RightEdge - 0.01, RightEdge + 0.01);
                filled++;
            }
            else Assert.Equal(LeftEdge, line.Min(text => text.XMillimeters), 3);
        }
        Assert.True(filled > 0, "At least one wrapped line must be stretched.");
        Assert.True(lines[^1].Max(text => text.XMillimeters + text.WidthMillimeters) < RightEdge - 1);
        Assert.Equal(source, string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text)));
    }

    private static OfdTextElement[][] Lines(OfdPage page) => page.Elements.OfType<OfdTextElement>()
        .Where(text => text.Text.Trim().Length > 0)
        .GroupBy(text => Math.Round(text.YMillimeters, 3)).OrderBy(group => group.Key)
        .Select(group => group.ToArray()).ToArray();

    private static string Paragraph(string? alignment, string text) =>
        $"""<w:p>{(alignment is null ? "" : $"<w:pPr><w:jc w:val=\"{alignment}\"/></w:pPr>")}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>""";

    private static async Task<OfdPage> Convert(string body, bool explicitNative)
    {
        await using var input = DocxReleaseFixtures.CreateMinimalDocx(body + Section);
        await using var output = new MemoryStream();
        var converter = explicitNative
            ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
            : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output);
        output.Position = 0;
        return (await new OfdReader().ReadAsync(output)).Pages[0];
    }
}
