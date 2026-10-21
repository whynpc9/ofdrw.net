using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class MultiDocumentExportTests
{
    [Fact]
    public async Task PdfSvgHtmlAndImage_UseSelectedDocBody()
    {
        var source = await MultiDocumentFixture.CreateZipBytesAsync();
        using var first = new MemoryStream(source, false);
        using var second = new MemoryStream(source, false);
        using var pdf0 = new MemoryStream(); using var pdf1 = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(first, pdf0);
        await new OfdToPdfConverter(new OfdToPdfOptions { PackageLoadOptions = new OfdPackageLoadOptions { DocumentIndex = 1 } })
            .ConvertAsync(second, pdf1);
        Assert.NotEqual(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(pdf0.ToArray())),
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(pdf1.ToArray())));

        using var svgSource = new MemoryStream(source, false); using var svg = new MemoryStream();
        await new OfdToSvgConverter(new OfdPackageLoadOptions { DocumentIndex = 1 }).ConvertAsync(svgSource, svg);
        Assert.Contains("SECOND DOC1", Encoding.UTF8.GetString(svg.ToArray()));

        using var htmlSource = new MemoryStream(source, false); using var html = new MemoryStream();
        await new OfdToHtmlConverter(new OfdHtmlOptions(), new OfdPackageLoadOptions { DocumentIndex = 1 }).ConvertAsync(htmlSource, html);
        Assert.Contains("SECOND DOC1", Encoding.UTF8.GetString(html.ToArray()));

        using var image0 = new MemoryStream(); using var image1 = new MemoryStream();
        using var imageSource0 = new MemoryStream(source, false); using var imageSource1 = new MemoryStream(source, false);
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2 })
            .ConvertAsync(imageSource0, image0);
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2,
            PackageLoadOptions = new OfdPackageLoadOptions { DocumentIndex = 1 } }).ConvertAsync(imageSource1, image1);
        Assert.True(image1.Length > 100);
        Assert.NotEqual(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(image0.ToArray())),
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(image1.ToArray())));
    }

}
