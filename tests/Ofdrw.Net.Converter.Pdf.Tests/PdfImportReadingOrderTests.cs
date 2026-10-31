using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using UglyToad.PdfPig;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class PdfImportReadingOrderTests
{
    public PdfImportReadingOrderTests() => PdfFontRegistry.EnsureInstalled();

    // Correctly shaped and unshaped references share the same embedded font and
    // logical source. A matching GetWords string on the latter is a false positive.
    [Theory]
    [InlineData("cell", "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8", "abc םוקמ םולש def")]
    [InlineData("fixed", "7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5", "abc שלום מקום def")]
    public async Task UnverifiedRtlSemanticImportRejectsBeforeWritingOutput(string variant, string hash, string raw)
    {
        var bytes = ReadFixture(variant, hash);
        using var parsed = PdfDocument.Open(bytes);
        Assert.Equal(raw, string.Join(" ", parsed.GetPage(1).GetWords().Select(word => word.Text)));

        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new PdfToOfdConverter().ConvertAsync(new MemoryStream(bytes), output));
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Contains("TextLayerMode=None", error.Message);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("cell", "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8")]
    [InlineData("fixed", "7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5")]
    public async Task ExplicitRasterOnlyImportRetainsImageWithoutClaimingRtlSemantics(string variant, string hash)
    {
        using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions { TextLayerMode = PdfTextLayerMode.None })
            .ConvertAsync(new MemoryStream(ReadFixture(variant, hash)), output);
        output.Position = 0;
        var ofd = await new OfdReader().ReadAsync(output);
        Assert.NotEmpty(Assert.Single(Assert.Single(ofd.Pages).Elements.OfType<OfdImageElement>()).Data);
        Assert.Empty(ofd.Pages.Single().Elements.OfType<OfdTextElement>());
    }

    [Fact]
    public void NonRasterTextFallbackDoesNotSwallowReadingOrderRejection()
    {
        using var parsed = PdfDocument.Open(ReadFixture("cell", "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8"));
        var page = new OfdPage();
        // Invoke the fallback itself, without pretending this forces an actual
        // renderer outage or proves its end-to-end selection on this host.
        var method = typeof(PdfToOfdConverter).GetMethod("AddTextFallbackPage", BindingFlags.NonPublic | BindingFlags.Static)!;
        var invocation = Assert.Throws<TargetInvocationException>(() => method.Invoke(null, new object[] { page, parsed.GetPage(1), 100d, 100d, 0 }));
        var error = Assert.IsType<NotSupportedException>(invocation.InnerException);
        Assert.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED", error.Message);
        Assert.Empty(page.Elements);
    }

    [Fact]
    public async Task LaterRtlPageLeavesOutputUntouchedWhileSelectingOnlyLtrPageSucceeds()
    {
        using var document = new PdfSharpCore.Pdf.PdfDocument();
        var ltr = document.AddPage();
        using (var gfx = XGraphics.FromPdfPage(ltr))
            gfx.DrawString("LTR control", new XFont("Arial", 12), XBrushes.Black, new XPoint(40, 60));
        using var source = PdfSharpCore.Pdf.IO.PdfReader.Open(
            new MemoryStream(ReadFixture("cell", "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8")),
            PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        document.AddPage(source.Pages[0]);
        using var pdf = new MemoryStream();
        document.Save(pdf, false);

        var priorOutput = Encoding.ASCII.GetBytes("existing output");
        using var failed = new MemoryStream();
        failed.Write(priorOutput);
        failed.Position = 0;
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter()
            .ConvertAsync(new MemoryStream(pdf.ToArray()), failed));
        Assert.Equal(priorOutput, failed.ToArray());

        using var selected = new MemoryStream();
        await new PdfToOfdConverter().ConvertAsync(new MemoryStream(pdf.ToArray()), selected, new[] { 0 });
        selected.Position = 0;
        var ofd = await new OfdReader().ReadAsync(selected);
        Assert.Single(ofd.Pages);
        Assert.Equal("LTR control ", string.Concat(ofd.Pages[0].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
    }

    [Theory]
    [InlineData("\U0001E900", true)] // Supplementary Adlam, bidi R.
    [InlineData("A\U0001E900B", true)] // Do not miss a pair inside one mapped value.
    [InlineData("\U00010400", false)] // Supplementary Deseret, bidi L.
    [InlineData("A", false)]
    public void ReadingOrderGateChecksDecodedScalarsWithoutRejectingSupplementaryLtr(string value, bool reject)
    {
        // A deliberately synthetic ToUnicode carrier tests decoded scalar
        // admission, not physical font identity or rendered-script support.
        using var parsed = PdfDocument.Open(MappedScalarPdf(value));
        var page = parsed.GetPage(1);
        Assert.Equal(value, Assert.Single(page.Letters).Value);
        if (reject) Assert.Throws<NotSupportedException>(() => PdfImportReadingOrder.Validate(page));
        else PdfImportReadingOrder.Validate(page);
    }

    private static byte[] MappedScalarPdf(string value)
    {
        var unicode = Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes(value));
        var cmap = $"/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Test def /CMapType 2 def\n1 begincodespacerange <0000> <FFFF> endcodespacerange\n1 beginbfchar <0041> <{unicode}> endbfchar\nendcmap CMapName currentdict /CMap defineresource pop end end\n";
        static string Stream(string s) => $"<< /Length {Encoding.ASCII.GetByteCount(s)} >>\nstream\n{s}endstream";
        var objects = new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>",
            "<< /Type /Font /Subtype /Type0 /BaseFont /Synthetic /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 5 0 R >>",
            Stream(cmap), Stream("BT /F1 12 Tf 20 100 Td <0041> Tj ET\n"),
            "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Synthetic /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 8 0 R /DW 1000 /CIDToGIDMap /Identity >>",
            "<< /Type /FontDescriptor /FontName /Synthetic /Flags 4 /FontBBox [0 0 1000 1000] /ItalicAngle 0 /Ascent 1000 /Descent 0 /CapHeight 1000 /StemV 80 >>" };
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var i = 0; i < objects.Length; i++)
        {
            offsets.Add(pdf.Length);
            pdf.Append($"{i + 1} 0 obj\n{objects[i]}\nendobj\n");
        }
        var xref = pdf.Length;
        pdf.Append($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n");
        foreach (var offset in offsets) pdf.Append($"{offset:0000000000} 00000 n \n");
        pdf.Append($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static byte[] ReadFixture(string variant, string sha256)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "D3", $"fpdf2-two-word-{variant}.pdf"));
        Assert.Equal(sha256, Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        return bytes;
    }
}
