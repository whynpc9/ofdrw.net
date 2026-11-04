using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class DocxListNumberingTests
{
    private const string Decimal2Levels = """
        <w:abstractNum w:abstractNumId="1">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
          <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/></w:lvl>
        </w:abstractNum>
        """;

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task MultiLevelCountersResetAndExpandPatterns(bool explicitNative)
    {
        var text = await NativeText(Docx(
            Item(1, 0, "Alpha") + Item(1, 1, "Bravo") + Item(1, 1, "Charlie") + Item(1, 0, "Delta") + Item(1, 1, "Echo"),
            Decimal2Levels + Num(1, 1)), explicitNative);

        Assert.Contains("1. Alpha", text);
        Assert.Contains("1.1. Bravo", text);
        Assert.Contains("1.2. Charlie", text);
        Assert.Contains("2. Delta", text);
        Assert.Contains("2.1. Echo", text);
        Assert.DoesNotContain("%", text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task NumIdZeroSuppressesTheMarker(bool explicitNative)
    {
        var text = await NativeText(Docx(Item(0, 0, "Plain"), Decimal2Levels + Num(1, 1)), explicitNative);

        Assert.Contains("Plain", text);
        Assert.DoesNotContain("1. Plain", text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task UnresolvedNumIdHasNoMarkerAndWarns(bool explicitNative)
    {
        await using var input = Docx(Item(7, 0, "Orphan") + Item(7, 0, "Second"), Decimal2Levels + Num(1, 1));
        await using var output = new MemoryStream();
        var result = await Converter(explicitNative).ConvertWithResultAsync(input, output);
        output.Position = 0;
        var text = Extract(await new OfdReader().ReadAsync(output));

        Assert.Contains("Orphan", text);
        Assert.DoesNotContain("1. Orphan", text);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NUMBERING_UNRESOLVED");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task RomanLetterAndChineseFormats(bool explicitNative)
    {
        var numbering = """
            <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
            <w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="26"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/></w:lvl></w:abstractNum>
            <w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%1、"/></w:lvl></w:abstractNum>
            """ + Num(1, 1) + Num(2, 2) + Num(3, 3);
        var body = Item(1, 0, "RomanA") + Item(1, 0, "RomanB") + Item(1, 0, "RomanC") + Item(1, 0, "RomanD") +
            Item(2, 0, "LetterA") + Item(2, 0, "LetterB") + Item(3, 0, "HanA") + Item(3, 0, "HanB");

        var text = await NativeText(Docx(body, numbering), explicitNative);

        Assert.Contains("I. RomanA", text);
        Assert.Contains("II. RomanB", text);
        Assert.Contains("III. RomanC", text);
        Assert.Contains("IV. RomanD", text);
        Assert.Contains("(z) LetterA", text);
        Assert.Contains("(aa) LetterB", text);
        Assert.Contains("一、 HanA", text);
        Assert.Contains("二、 HanB", text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task StartOverrideSetsTheFirstValue(bool explicitNative)
    {
        var numbering = Decimal2Levels +
            """<w:num w:numId="3"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>""";

        var text = await NativeText(Docx(Item(3, 0, "Fifth") + Item(3, 0, "Sixth"), numbering), explicitNative);

        Assert.Contains("5. Fifth", text);
        Assert.Contains("6. Sixth", text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task BulletKeepsItsLevelText(bool explicitNative)
    {
        var numbering = """<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>""" + Num(1, 1);

        var text = await NativeText(Docx(Item(1, 0, "Item"), numbering), explicitNative);

        Assert.Contains("• Item", text);
    }

    [Theory]
    [InlineData(1, "I")]
    [InlineData(4, "IV")]
    [InlineData(9, "IX")]
    [InlineData(14, "XIV")]
    [InlineData(40, "XL")]
    [InlineData(90, "XC")]
    [InlineData(400, "CD")]
    [InlineData(1994, "MCMXCIV")]
    [InlineData(3999, "MMMCMXCIX")]
    public void RomanNumerals(int value, string expected)
    {
        Assert.Equal(expected, DocxNumberFormatter.Format(value, NumberFormatValues.UpperRoman, out var degraded));
        Assert.False(degraded);
        Assert.Equal(expected.ToLowerInvariant(), DocxNumberFormatter.Format(value, NumberFormatValues.LowerRoman, out _));
    }

    [Fact]
    public void RomanOutsideRangeFallsBackToDecimal() =>
        Assert.Equal("4000", DocxNumberFormatter.Format(4000, NumberFormatValues.UpperRoman, out _));

    [Theory]
    [InlineData(1, "a")]
    [InlineData(26, "z")]
    [InlineData(27, "aa")]
    [InlineData(52, "zz")]
    [InlineData(53, "aaa")]
    public void Letters(int value, string expected)
    {
        Assert.Equal(expected, DocxNumberFormatter.Format(value, NumberFormatValues.LowerLetter, out _));
        Assert.Equal(expected.ToUpperInvariant(), DocxNumberFormatter.Format(value, NumberFormatValues.UpperLetter, out _));
    }

    [Theory]
    [InlineData(1, "一")]
    [InlineData(10, "十")]
    [InlineData(11, "十一")]
    [InlineData(19, "十九")]
    [InlineData(20, "二十")]
    [InlineData(21, "二十一")]
    [InlineData(99, "九十九")]
    public void ChineseCounting(int value, string expected)
    {
        Assert.Equal(expected, DocxNumberFormatter.Format(value, NumberFormatValues.ChineseCounting, out var degraded));
        Assert.False(degraded);
    }

    [Fact]
    public void ChineseCountingAbove99DegradesToDecimal()
    {
        Assert.Equal("100", DocxNumberFormatter.Format(100, NumberFormatValues.ChineseCounting, out var degraded));
        Assert.True(degraded);
    }

    [Theory]
    [InlineData(1, "01")]
    [InlineData(9, "09")]
    [InlineData(10, "10")]
    public void DecimalZero(int value, string expected) =>
        Assert.Equal(expected, DocxNumberFormatter.Format(value, NumberFormatValues.DecimalZero, out _));

    [Fact]
    public void IdeographDigitalMapsEachDigit() =>
        Assert.Equal("一〇五", DocxNumberFormatter.Format(105, NumberFormatValues.IdeographDigital, out _));

    [Fact]
    public void UnsupportedFormatDegradesToDecimal()
    {
        Assert.Equal("3", DocxNumberFormatter.Format(3, NumberFormatValues.ChineseLegalSimplified, out var degraded));
        Assert.True(degraded);
    }

    private static string Item(int numId, int level, string text) =>
        $"""<w:p><w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{numId}"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>""";

    private static string Num(int numId, int abstractId) =>
        $"""<w:num w:numId="{numId}"><w:abstractNumId w:val="{abstractId}"/></w:num>""";

    private static MemoryStream Docx(string body, string numberingXml)
    {
        var input = DocxReleaseFixtures.CreateMinimalDocx(body);
        using (var document = WordprocessingDocument.Open(input, true))
        {
            var xml = """<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">""" +
                numberingXml + "</w:numbering>";
            document.MainDocumentPart!.AddNewPart<NumberingDefinitionsPart>()
                .FeedData(new MemoryStream(Encoding.UTF8.GetBytes(xml)));
        }

        input.Position = 0;
        return input;
    }

    private static DocxToOfdConverter Converter(bool explicitNative) => explicitNative
        ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
        : new DocxToOfdConverter();

    private static async Task<string> NativeText(Stream docx, bool explicitNative)
    {
        await using (docx)
        {
            await using var output = new MemoryStream();
            await Converter(explicitNative).ConvertAsync(docx, output);
            output.Position = 0;
            return Extract(await new OfdReader().ReadAsync(output));
        }
    }

    private static string Extract(Ofdrw.Net.Core.Models.OfdDocumentPackage package) =>
        string.Join("\n", new OfdTextExtractor().ExtractPages(package, includeTemplates: true));
}
