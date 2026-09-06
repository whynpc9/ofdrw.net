using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.TestSupport;
using PdfSharpCore.Pdf;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ResourceBudgetTests
{
    [Fact]
    public async Task PdfInputBudget_ShouldStopBeforeStagingTheWholeInput()
    {
        using var input = new NonSeekableInput(new byte[100]);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(new PdfToOfdOptions { MaxInputBytes = 10 }).ConvertAsync(input, output));
        Assert.Equal(11, input.BytesRead);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task PixelBudget_ShouldRejectBeforeRasterizing()
    {
        using var pdf = new PdfDocument();
        var page = pdf.AddPage(); page.Width = 1000; page.Height = 1000;
        using var input = new MemoryStream();
        pdf.Save(input, false); input.Position = 0;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(new PdfToOfdOptions { MaxRasterizedPixelsPerPage = 100 }).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task InvalidImage_ShouldFailWithoutPublishingABlankPdf()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdImageElement { Data = [1, 2, 3], WidthMillimeters = 10, HeightMillimeters = 10 });
        package.Pages.Add(page);
        using var input = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter().ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }
}
