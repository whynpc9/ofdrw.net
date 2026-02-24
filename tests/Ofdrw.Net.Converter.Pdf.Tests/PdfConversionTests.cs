using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
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

    [Fact]
    public async Task PdfToOfd_ShouldNormalizeInvalidPageSelection_AndPreserveDuplicates()
    {
        await using var inputPdf = new MemoryStream();
        CreateSamplePdf(inputPdf, 3);
        inputPdf.Position = 0;

        var converter = new PdfToOfdConverter();
        await using var ofdStream = new MemoryStream();
        await converter.ConvertAsync(inputPdf, ofdStream, new[] { 99, -1, 1, 1 });

        ofdStream.Position = 0;
        var reader = new OfdReader();
        var ofd = await reader.ReadAsync(ofdStream);

        Assert.Equal(2, ofd.Pages.Count);
        Assert.All(ofd.Pages, page =>
        {
            var text = page.Elements.OfType<OfdTextElement>().SingleOrDefault();
            Assert.NotNull(text);
            Assert.Contains("page-2", text!.Text);
        });
    }

    [Fact]
    public async Task PdfToOfd_ShouldProduceRenderableFallback_ForBlankPdfPage()
    {
        await using var inputPdf = new MemoryStream();
        CreateBlankPdf(inputPdf, 1);
        inputPdf.Position = 0;

        var converter = new PdfToOfdConverter();
        await using var ofdStream = new MemoryStream();
        await converter.ConvertAsync(inputPdf, ofdStream);

        ofdStream.Position = 0;
        var reader = new OfdReader();
        var ofd = await reader.ReadAsync(ofdStream);

        Assert.Single(ofd.Pages);
        Assert.NotEmpty(ofd.Pages[0].Elements);
        Assert.Contains(ofd.Pages[0].Elements, e => e is OfdTextElement or OfdImageElement);
    }

    [Fact]
    public async Task OfdToPdf_ShouldHonorRequestedPageOrderAndDuplicates()
    {
        var ofdBytes = await CreateSampleOfdAsync("Alpha", "Beta", "Gamma");

        var converter = new OfdToPdfConverter();
        await using var pdfOutput = new MemoryStream();
        await using var ofdInput = new MemoryStream(ofdBytes);
        await converter.ConvertAsync(ofdInput, pdfOutput, new[] { 2, 0, 2 });

        pdfOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(pdfOutput);

        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("Gamma", pdf.GetPage(1).Text);
        Assert.Contains("Alpha", pdf.GetPage(2).Text);
        Assert.Contains("Gamma", pdf.GetPage(3).Text);
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

    private static void CreateBlankPdf(Stream output, int pageCount)
    {
        using var document = new PdfSharpDocument();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.AddPage();
            page.Width = XUnit.FromMillimeter(210);
            page.Height = XUnit.FromMillimeter(297);
        }

        document.Save(output, false);
    }

    private static async Task<byte[]> CreateSampleOfdAsync(params string[] labels)
    {
        var builder = new OfdDocumentBuilder();
        for (var i = 0; i < labels.Length; i++)
        {
            builder.AddPage(new OfdPage
            {
                Index = i,
                WidthMillimeters = 210,
                HeightMillimeters = 297,
                Elements =
                {
                    new OfdTextElement
                    {
                        Text = labels[i],
                        FontName = "SimSun",
                        FontSizeMillimeters = 4,
                        XMillimeters = 10,
                        YMillimeters = 12,
                        WidthMillimeters = 80,
                        HeightMillimeters = 10
                    }
                }
            });
        }

        var writer = new OfdPackageWriter();
        await using var ms = new MemoryStream();
        await writer.WriteAsync(builder.Build(), ms);
        return ms.ToArray();
    }
}
