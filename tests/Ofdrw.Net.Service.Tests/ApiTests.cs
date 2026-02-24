using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Testing;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Drawing;
using PdfSharpDocument = PdfSharpCore.Pdf.PdfDocument;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Service.Tests;

public sealed class ApiTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly WebApplicationFactory<Program> _factory;

    public ApiTests(WebApplicationFactory<Program> factory)
    {
        _factory = factory;
    }

    [Fact]
    public async Task GetProfile_ShouldReturnOk()
    {
        var client = _factory.CreateClient();
        var response = await client.GetAsync("/api/v1/ofd/validate/profiles/emr-ofd-h-202x");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("emr-ofd-h-202x", body);
    }

    [Fact]
    public async Task ValidateEndpoint_ShouldReturnReport()
    {
        var client = _factory.CreateClient();
        var content = new MultipartFormDataContent();

        var bytes = await CreateSampleOfdAsync();
        var file = new ByteArrayContent(bytes);
        file.Headers.ContentType = new MediaTypeHeaderValue("application/ofd");
        content.Add(file, "ofd", "sample.ofd");

        var response = await client.PostAsync("/api/v1/ofd/validate/emr-tech-spec", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var json = await response.Content.ReadAsStringAsync();
        Assert.Contains("profileVersion", json, System.StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConverterEndpoints_ShouldRoundTrip_AndHonorPageSelections()
    {
        var client = _factory.CreateClient();
        var pdfBytes = CreateSamplePdf(3);

        using var pdfToOfdRequest = new MultipartFormDataContent();
        var pdfFile = new ByteArrayContent(pdfBytes);
        pdfFile.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
        pdfToOfdRequest.Add(pdfFile, "pdf", "sample.pdf");
        pdfToOfdRequest.Add(new StringContent("2,0"), "pages");

        using var pdfToOfdResponse = await client.PostAsync("/api/v1/ofd/convert/pdf-to-ofd", pdfToOfdRequest);
        Assert.Equal(HttpStatusCode.OK, pdfToOfdResponse.StatusCode);
        var ofdBytes = await pdfToOfdResponse.Content.ReadAsByteArrayAsync();
        Assert.NotEmpty(ofdBytes);

        await using (var ofdStream = new MemoryStream(ofdBytes))
        {
            var ofdReader = new OfdReader();
            var ofd = await ofdReader.ReadAsync(ofdStream);
            Assert.Equal(2, ofd.Pages.Count);
            Assert.Contains("page-3", ofd.Pages[0].Elements.OfType<OfdTextElement>().First().Text);
            Assert.Contains("page-1", ofd.Pages[1].Elements.OfType<OfdTextElement>().First().Text);
        }

        using var ofdToPdfRequest = new MultipartFormDataContent();
        var ofdFile = new ByteArrayContent(ofdBytes);
        ofdFile.Headers.ContentType = new MediaTypeHeaderValue("application/ofd");
        ofdToPdfRequest.Add(ofdFile, "ofd", "sample.ofd");
        ofdToPdfRequest.Add(new StringContent("1,1,0"), "pages");

        using var ofdToPdfResponse = await client.PostAsync("/api/v1/ofd/convert/ofd-to-pdf", ofdToPdfRequest);
        Assert.Equal(HttpStatusCode.OK, ofdToPdfResponse.StatusCode);
        var outputPdfBytes = await ofdToPdfResponse.Content.ReadAsByteArrayAsync();

        using var pdfStream = new MemoryStream(outputPdfBytes);
        using var pdf = PdfPigDocument.Open(pdfStream);
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("page-1", pdf.GetPage(1).Text);
        Assert.Contains("page-1", pdf.GetPage(2).Text);
        Assert.Contains("page-3", pdf.GetPage(3).Text);
    }

    [Fact]
    public async Task ConverterEndpoints_ShouldReturnBadRequest_WhenFileMissing()
    {
        var client = _factory.CreateClient();

        using var pdfToOfdRequest = new MultipartFormDataContent();
        using var pdfToOfdResponse = await client.PostAsync("/api/v1/ofd/convert/pdf-to-ofd", pdfToOfdRequest);
        Assert.Equal(HttpStatusCode.BadRequest, pdfToOfdResponse.StatusCode);

        using var ofdToPdfRequest = new MultipartFormDataContent();
        using var ofdToPdfResponse = await client.PostAsync("/api/v1/ofd/convert/ofd-to-pdf", ofdToPdfRequest);
        Assert.Equal(HttpStatusCode.BadRequest, ofdToPdfResponse.StatusCode);
    }

    private static async Task<byte[]> CreateSampleOfdAsync()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new Ofdrw.Net.Core.Models.OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "sample",
                    XMillimeters = 10,
                    YMillimeters = 10,
                    WidthMillimeters = 50,
                    HeightMillimeters = 8,
                    FontSizeMillimeters = 4
                }
            }
        });

        await using var ms = new MemoryStream();
        var writer = new OfdPackageWriter();
        await writer.WriteAsync(builder.Build(), ms);
        return ms.ToArray();
    }

    private static byte[] CreateSamplePdf(int pageCount)
    {
        using var ms = new MemoryStream();
        using var document = new PdfSharpDocument();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.AddPage();
            page.Width = XUnit.FromMillimeter(210);
            page.Height = XUnit.FromMillimeter(297);
            using var gfx = XGraphics.FromPdfPage(page);
            gfx.DrawString($"page-{i + 1}", new XFont("Arial", 18), XBrushes.Black, new XPoint(40, 80));
        }

        document.Save(ms, false);
        return ms.ToArray();
    }
}
