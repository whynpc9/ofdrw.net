using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Drawing;
using PdfSharpDocument = PdfSharpCore.Pdf.PdfDocument;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class PdfConversionTests
{
    [Fact]
    public async Task PdfToOfd_AndBack_ShouldPreservePageCountForSelectedPages()
    {
        await using var inputPdf = new MemoryStream();
        CreateSamplePdf(inputPdf, 3);
        inputPdf.Position = 0;

        var pdfToOfd = new PdfToOfdConverter();
        await using var ofdStream = new MemoryStream();
        await pdfToOfd.ConvertAsync(inputPdf, ofdStream, new[] { 2, 0, 2 });

        ofdStream.Position = 0;
        var ofdReader = new OfdReader();
        var ofdDoc = await ofdReader.ReadAsync(ofdStream);
        Assert.Equal(3, ofdDoc.Pages.Count);

        ofdStream.Position = 0;
        var ofdToPdf = new OfdToPdfConverter();
        await using var outputPdf = new MemoryStream();
        await ofdToPdf.ConvertAsync(ofdStream, outputPdf);

        outputPdf.Position = 0;
        using var pdf = PdfPigDocument.Open(outputPdf);
        Assert.Equal(3, pdf.NumberOfPages);
    }

    private static void CreateSamplePdf(Stream output, int pageCount)
    {
        using var document = new PdfSharpDocument();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.AddPage();
            page.Width = XUnit.FromMillimeter(210);
            page.Height = XUnit.FromMillimeter(297);
            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawString($"page-{i + 1}", new XFont("Arial", 18), XBrushes.Black, new XPoint(40, 80));
        }

        document.Save(output, false);
    }
}
