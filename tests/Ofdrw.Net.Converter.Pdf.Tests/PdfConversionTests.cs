using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Layout.Editing;
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
            var image = page.Elements.OfType<OfdImageElement>().SingleOrDefault();
            Assert.NotNull(image);
            Assert.NotEmpty(image!.Data);
            Assert.Equal("image/png", image.MediaType);
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
        var image = Assert.Single(ofd.Pages[0].Elements.OfType<OfdImageElement>());
        Assert.NotEmpty(image.Data);
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

    [Fact]
    public async Task OfdToPdf_ShouldRenderUpstreamTemplatePaths()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
            "999.ofd");

        await using var ofdInput = File.OpenRead(samplePath);
        await using var pdfOutput = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofdInput, pdfOutput);

        pdfOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(pdfOutput);
        Assert.Equal(5, pdf.NumberOfPages);
        Assert.True(
            pdf.GetPage(1).Operations.Count(operation =>
                operation.Operator == "Do") >= 2,
            "The first page should draw both its QR code and the signed seal form.");
        Assert.All(Enumerable.Range(1, pdf.NumberOfPages), pageNumber =>
            Assert.Contains(
                pdf.GetPage(pageNumber).Operations,
                operation => operation.Operator is "m" or "l" or "c"));
        Assert.Contains(
            "/Subtype /Form",
            Encoding.ASCII.GetString(pdfOutput.ToArray()));
    }

    [Fact]
    public async Task OfdToPdf_ShouldApplyCropBoxAsRenderingOrigin()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "Inside crop",
                    FontName = "Arial",
                    FontSizeMillimeters = 4,
                    XMillimeters = 50,
                    YMillimeters = 60,
                    WidthMillimeters = 40,
                    HeightMillimeters = 8
                }
            }
        };
        package.Pages.Add(page);
        OfdDocumentEditor.CropPage(page, 40, 50, 100, 120);

        await using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        await using var pdfOutput = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdfOutput);

        pdfOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(pdfOutput);
        var renderedPage = pdf.GetPage(1);
        Assert.Equal(100d * 72d / 25.4d, renderedPage.Width, 1);
        Assert.Equal(120d * 72d / 25.4d, renderedPage.Height, 1);
        Assert.Contains("Inside crop", renderedPage.Text);
    }

    [Fact]
    public async Task OfdToSvg_ShouldPreserveTemplateVectorsTextAndImages()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
            "999.ofd");

        await using var ofdInput = File.OpenRead(samplePath);
        await using var svgOutput = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(ofdInput, svgOutput);

        await using var modelInput = File.OpenRead(samplePath);
        var package = await new OfdReader().ReadAsync(modelInput);
        var firstPage = package.Pages.OrderBy(page => page.Index).First();
        var expectedPathCount = firstPage.Elements.OfType<OfdPathElement>().Count() +
            firstPage.Templates.Sum(template =>
                template.Elements.OfType<OfdPathElement>().Count());

        svgOutput.Position = 0;
        var svg = System.Xml.Linq.XDocument.Load(svgOutput);
        Assert.Equal("svg", svg.Root?.Name.LocalName);
        Assert.Equal(
            expectedPathCount,
            svg.Descendants().Count(element => element.Name.LocalName == "path"));
        Assert.Contains(svg.Descendants(), element => element.Name.LocalName == "text");
        Assert.Single(svg.Descendants(), element => element.Name.LocalName == "image");
        Assert.All(
            svg.Descendants().Where(element => element.Name.LocalName == "path"),
            path => Assert.DoesNotContain(" B ", $" {path.Attribute("d")?.Value} "));
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

    private static string ResolveRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate repository root.");
    }
}
