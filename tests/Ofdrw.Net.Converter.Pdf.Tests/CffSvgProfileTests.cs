using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class CffSvgProfileTests
{
    private const string Text = "A g é @";
    private const string FontSha256 = "9f9664e2edf6f045c11e774f9bd0be6993971f2544a39061a5ce478b96b051f8";

    [Fact]
    public async Task EmbeddedCff1OtfPreservesOfdAndExportsActualUnicodeSvgText()
    {
        var font = Fixture();
        Assert.Equal(FontSha256, Convert.ToHexString(SHA256.HashData(font)).ToLowerInvariant());
        var ofd = await Source(font);
        using (var input = new MemoryStream(ofd))
        {
            var reopened = await new OfdReader().ReadAsync(input);
            Assert.Equal(font, Assert.Single(reopened.Fonts).Data);
            Assert.Equal(Text, Assert.Single(reopened.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        }

        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var css = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "style").Value;
        Assert.Contains("font/otf;base64," + Convert.ToBase64String(font), css);
        var drawn = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "text");
        Assert.Equal(Text, drawn.Value);
        Assert.Equal("ofd-font-" + FontSha256, drawn.Attribute("font-family")?.Value);

        using var pdf = new MemoryStream();
        var refusal = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf));
        Assert.Contains("CFF", refusal.ToString());
        Assert.Equal(0, pdf.Length);

        var artifacts = Environment.GetEnvironmentVariable("OFDRW_R08_CFF_SVG_ARTIFACT_DIR");
        if (!string.IsNullOrWhiteSpace(artifacts))
        {
            Directory.CreateDirectory(artifacts);
            File.WriteAllBytes(Path.Combine(artifacts, "cff-text.ofd"), ofd);
            File.WriteAllBytes(Path.Combine(artifacts, "cff-text.svg"), svg.ToArray());
        }
    }

    [Fact]
    public async Task CffCharStringsCountMismatchFailsBeforeSvgOutput()
    {
        var face = new OpenTypeFace(Fixture());
        var maxp = face.Table("maxp", 6);
        var count = OpenTypeFace.U16(maxp, 4);
        OpenTypeFace.Put16(maxp, 4, count + 1);
        var ofd = await Source(face.Build());
        using var svg = new MemoryStream();
        var refusal = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg));
        Assert.Contains("CharStrings count", refusal.Message);
        Assert.Equal(0, svg.Length);
    }

    [Fact]
    public async Task SelectedCff2TableFailsBeforeSvgOutput()
    {
        var face = new OpenTypeFace(Fixture());
        var program = face.Table("CFF ");
        face.Tables.Remove("CFF ");
        face.Tables.Add("CFF2", program);
        var ofd = await Source(face.Build());
        using var svg = new MemoryStream();
        var refusal = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg));
        Assert.Contains("CFF1", refusal.Message);
        Assert.Equal(0, svg.Length);
    }

    [Fact]
    public void RepeatedHighSegmentCffLookupChargesSharedBudgetBeforeEachLookup()
    {
        const int scalar = 0xE1FE;
        var original = new OpenTypeFace(Fixture());
        var glyph = new OpenTypeCmap(original).Glyph('A');
        var mappings = Enumerable.Range(0, 256).ToDictionary(index => 0xE000 + index * 2, _ => glyph);
        var built = OpenTypeCmap.Build(mappings);
        var offset = (int)OpenTypeFace.U32(built, 8);
        var length = OpenTypeFace.U16(built, offset + 2);
        var format4Only = new byte[12 + length];
        OpenTypeFace.Put16(format4Only, 2, 1);
        OpenTypeFace.Put16(format4Only, 4, 3);
        OpenTypeFace.Put16(format4Only, 6, 1);
        OpenTypeFace.Put32(format4Only, 8, 12);
        Buffer.BlockCopy(built, offset, format4Only, 12, length);
        original.Tables["cmap"] = format4Only;
        var bytes = original.Build();
        var face = new OpenTypeFace(bytes);
        var count = Cff1SvgProfile.Validate(bytes, face);
        var probeBudget = new FontParseBudget(maximum: 1_000_000);
        var probe = new OpenTypeCmap(face, probeBudget);
        Assert.True(probe.GlyphLookupWorkUpperBound >= 256);
        Assert.Equal(glyph, probe.Glyph(scalar));

        var maximum = probeBudget.Consumed + probe.GlyphLookupWorkUpperBound;
        var controlBudget = new FontParseBudget(maximum);
        var control = new OpenTypeCmap(face, controlBudget);
        Cff1SvgProfile.ValidateSelectedText(new OfdTextElement { Text = "\uE1FE" },
            new OfdFontResource(), control, count, controlBudget);
        Assert.Equal(maximum, controlBudget.Consumed);

        var repeatedBudget = new FontParseBudget(maximum);
        var repeated = new OpenTypeCmap(face, repeatedBudget);
        var failure = Assert.Throws<InvalidDataException>(() =>
            Cff1SvgProfile.ValidateSelectedText(new OfdTextElement { Text = "\uE1FE\uE1FE" },
                new OfdFontResource(), repeated, count, repeatedBudget));
        Assert.Contains("MaximumCmapParseWork", failure.Message);
        Assert.True(repeatedBudget.IsExhausted);
    }

    private static byte[] Fixture()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        return File.ReadAllBytes(Path.Combine(directory!.FullName,
            "tests/Ofdrw.Net.Graphics.SkiaSharp.Tests/testdata/cff/SourceCodePro-Regular.otf"));
    }

    private static async Task<byte[]> Source(byte[] font)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Source Code Pro", Data = font });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 45 };
        page.Elements.Add(new OfdTextElement
        {
            ObjectId = "11", FontResourceId = "10", FontName = "Source Code Pro",
            Text = Text, FontSizeMillimeters = 10, XMillimeters = 8, YMillimeters = 8,
            WidthMillimeters = 84, HeightMillimeters = 20
        });
        package.Pages.Add(page);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }
}
