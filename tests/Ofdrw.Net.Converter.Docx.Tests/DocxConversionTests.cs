using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Docx.Tests;

/// <summary>
/// Covers the generated, non-sensitive DOCX conversion fixture.
/// </summary>
public sealed class DocxConversionTests
{
    /// <summary>
    /// Verifies the real LibreOffice PDF and OFD pipeline.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_ShouldConvertToPdfAndOfd()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Docx.E2E",
            "testdata",
            "generated-layout.docx");

        var options = new DocxConversionOptions
        {
            Engine = DocxConversionEngine.LibreOffice
        };
        var docxToPdf = new DocxToPdfConverter(options);
        await using var pdf = new MemoryStream();
        await using (var docx = File.OpenRead(samplePath))
        {
            await docxToPdf.ConvertAsync(docx, pdf);
        }

        AssertPdfHasTwoPages(pdf);

        var docxToOfd = new DocxToOfdConverter(options);
        await using var ofd = new MemoryStream();
        await using (var docx = File.OpenRead(samplePath))
        {
            await docxToOfd.ConvertAsync(docx, ofd);
        }

        ofd.Position = 0;
        var package = await new OfdReader().ReadAsync(ofd);
        Assert.Equal(2, package.Pages.Count);
        Assert.All(package.Pages, page =>
        {
            var image = Assert.Single(page.Elements.OfType<OfdImageElement>());
            Assert.Equal("image/png", image.MediaType);
            Assert.NotEmpty(image.Data);
        });
    }

    /// <summary>
    /// Verifies process timeout validation.
    /// </summary>
    [Fact]
    public void DocxToPdf_ShouldRejectNonPositiveTimeout()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new DocxToPdfConverter(new DocxConversionOptions
            {
                ProcessTimeout = TimeSpan.Zero
            }));
    }

    private static void AssertPdfHasTwoPages(MemoryStream stream)
    {
        stream.Position = 0;
        using var pdf = PdfPigDocument.Open(stream);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.All(Enumerable.Range(1, pdf.NumberOfPages), pageNumber =>
        {
            var page = pdf.GetPage(pageNumber);
            Assert.InRange(page.Width, 590, 600);
            Assert.InRange(page.Height, 840, 850);
        });
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

        throw new DirectoryNotFoundException("Unable to locate the repository root.");
    }
}
