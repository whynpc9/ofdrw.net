using System.Security.Cryptography;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class CffHardeningTests
{
    private static byte[] ReadHanFont() =>
        File.ReadAllBytes(ParagraphLayoutF08Tests.CffPath("OfdrwCIHanCFF-Regular.otf"));

    [Fact]
    public void ValidFontRenditionIsUnchanged()
    {
        var rendition = CffTrueTypeConversion.Convert(ReadHanFont());
        var hash = Convert.ToHexString(SHA256.HashData(rendition)).ToLowerInvariant();
        Assert.Equal("f2ff29b500bf799ff7acbbc7ca684297ddf5a71c3e5ebdb2b9a153c826ac5ee8", hash);
    }

    private static byte[] OverflowingIndex() =>
        [0, 1, 4, 0, 0, 0, 1, 0x7f, 0xff, 0xff, 0xff, 0];

    [Fact]
    public void IndexWithOverflowingLastOffsetIsInvalidData()
    {
        var error = Assert.Throws<InvalidDataException>(() => CffIndex.Read(OverflowingIndex(), 0));
        Assert.Equal("CFF INDEX data exceeds the font program.", error.Message);
    }

    [Fact]
    public void ArithmeticOverflowInCharStringIsInvalidData()
    {
        var face = new OpenTypeFace(ReadHanFont());
        var charStrings = CffFontProgram.Parse(face).CharStrings;
        var gid = Enumerable.Range(1, charStrings.Count - 1).First(id => charStrings.ItemLength(id) >= 48);
        var charString = new List<byte> { 28, 0x7f, 0xff }; // shortint 32767
        for (var square = 0; square < 7; square++) charString.AddRange([12, 27, 12, 24]); // dup, mul
        charString.AddRange([12, 18, 14]); // drop, endchar
        charString.CopyTo(face.Table("CFF "), charStrings.ItemStart(gid));

        var outlines = new CffCharStringOutlines(face.Build());
        var error = Assert.Throws<InvalidDataException>(() => outlines.Read(gid));
        Assert.Equal("CFF CharString produced a non-finite operand.", error.Message);
    }

    [Fact]
    public void TotalOperationBudgetStopsTheRendition()
    {
        var source = ReadHanFont();
        var error = Assert.Throws<InvalidDataException>(() =>
            CffTrueTypeConversion.Convert(source, default, maximumOperations: 1000, maximumGlyfBytes: long.MaxValue));
        Assert.Equal("CFF font exceeded its total interpretation budget.", error.Message);

        // Give every glyph a one-operation endchar and have two also call the same global subroutine.
        var face = new OpenTypeFace(source);
        var program = CffFontProgram.Parse(face);
        var charStrings = program.CharStrings;
        var globalSubrs = program.GlobalSubrs;
        var cff = face.Table("CFF ");
        byte[] subroutine = [139, 12, 18, 11]; // push 0, drop, return
        var subr = Enumerable.Range(0, globalSubrs.Count).First(id => globalSubrs.ItemLength(id) >= subroutine.Length);
        subroutine.CopyTo(cff, globalSubrs.ItemStart(subr));
        var bias = globalSubrs.Count < 1240 ? 107 : globalSubrs.Count < 33900 ? 1131 : 32768;
        var operand = subr - bias;
        byte[] call = [28, (byte)(operand >> 8), (byte)operand, 29, 14]; // shortint, callgsubr, endchar
        for (var gid = 1; gid < charStrings.Count; gid++)
            cff[charStrings.ItemStart(gid)] = 14;
        var callers = Enumerable.Range(1, charStrings.Count - 1)
            .Where(id => charStrings.ItemLength(id) >= call.Length).Take(2).ToArray();
        Assert.Equal(2, callers.Length);
        foreach (var gid in callers) call.CopyTo(cff, charStrings.ItemStart(gid));

        var repeated = face.Build();
        var exactOperations = charStrings.Count - 1 + callers.Length * 5;
        Assert.NotEmpty(CffTrueTypeConversion.Convert(repeated, default, exactOperations, maximumGlyfBytes: 0));
        error = Assert.Throws<InvalidDataException>(() =>
            CffTrueTypeConversion.Convert(repeated, default, exactOperations - 1, maximumGlyfBytes: 0));
        Assert.Equal("CFF font exceeded its total interpretation budget.", error.Message);
    }

    [Fact]
    public void TotalSizeBudgetStopsTheRendition()
    {
        var source = ReadHanFont();
        var rendition = CffTrueTypeConversion.Convert(source);
        var exactLength = new OpenTypeFace(rendition).Table("glyf").LongLength;
        Assert.Equal(rendition, CffTrueTypeConversion.Convert(source, default, long.MaxValue, exactLength));
        foreach (var limit in new[] { 1024L, exactLength - 1 })
        {
            var error = Assert.Throws<InvalidDataException>(() =>
                CffTrueTypeConversion.Convert(source, default, maximumOperations: long.MaxValue, maximumGlyfBytes: limit));
            Assert.Equal("CFF TrueType rendition exceeds its size budget.", error.Message);
        }
    }

    [Fact]
    public void DefaultBudgetsAdmitTheValidFont() =>
        Assert.NotEmpty(CffTrueTypeConversion.Convert(ReadHanFont()));

    [Fact]
    public void CancelledRenditionPropagatesCancellation()
    {
        var source = ReadHanFont();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => CffTrueTypeConversion.Convert(source, cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PdfExportIsolatesAnEmbeddedFontWithAnOverflowingIndex(bool selectHostileFont)
    {
        const string text = "CFF 思源黑体字体测试";
        var source = ReadHanFont();
        var hostile = new OpenTypeFace(source);
        var cff = hostile.Table("CFF ");
        OverflowingIndex().CopyTo(cff, cff[2]); // Replace the Name INDEX at the end of the CFF header.
        var hostileBytes = hostile.Build();
        var package = new OfdDocumentPackage();
        package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Hostile CFF", Data = hostileBytes });
        package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "Ofdrw CI Han CFF", Data = source });
        var body = new OfdTextElement
        {
            Text = text, FontResourceId = selectHostileFont ? "10" : "11",
            XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 95, HeightMillimeters = 10,
        };
        body.Runs.Add(new OfdTextRun { Text = text, YMillimeters = 5 });
        var page = new OfdPage { Id = "1", WidthMillimeters = 120, HeightMillimeters = 40 };
        page.Elements.Add(body);
        package.Pages.Add(page);
        var ofd = await ParagraphLayoutOfdTests.Write(package);
        var readback = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Equal(hostileBytes, Assert.Single(readback.Fonts, font => font.FontName == "Hostile CFF").Data);

        using var pdf = new MemoryStream();
        var converter = new OfdToPdfConverter();
        if (selectHostileFont)
        {
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => converter.ConvertAsync(new MemoryStream(ofd), pdf));
            Assert.Contains("Hostile CFF", error.Message);
            Assert.IsType<InvalidDataException>(error.InnerException?.InnerException);
            return;
        }

        await converter.ConvertAsync(new MemoryStream(ofd), pdf);
        using var document = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Equal(text.Replace(" ", ""), document.GetPage(1).Text.Replace(" ", ""));
    }
}
