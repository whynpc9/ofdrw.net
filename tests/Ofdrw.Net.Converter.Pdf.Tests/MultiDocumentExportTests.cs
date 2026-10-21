using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
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
    public async Task SignatureAppearances_AreScopedToSelectedDocBody()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = XDocument.Load(new MemoryStream(entries["OFD.xml"]));
        var bodies = root.Root!.Elements().Where(node => node.Name.LocalName == "DocBody").ToArray();
        var pageIds = new string[2];
        for (var index = 0; index < bodies.Length; index++)
        {
            var prefix = $"Doc_{index}/Signs/";
            bodies[index].Add(new XElement(root.Root.Name.Namespace + "Signatures", prefix + "Signatures.xml"));
            var document = XDocument.Load(new MemoryStream(entries[$"Doc_{index}/Document.xml"]));
            pageIds[index] = document.Descendants().First(node => node.Name.LocalName == "Page").Attribute("ID")!.Value;
            entries[prefix + "Signatures.xml"] = Encoding.UTF8.GetBytes("<Signatures><Signature BaseLoc=\"S0/Signature.xml\"/></Signatures>");
            var stamps = string.Concat(Enumerable.Range(0, index == 0 ? 2 : 1).Select(_ =>
                $"<StampAnnot PageRef=\"{pageIds[index]}\" Boundary=\"5 10 10 10\"/>"));
            entries[prefix + "S0/Signature.xml"] = Encoding.UTF8.GetBytes(
                $"<Signature><SignedInfo><Seal BaseLoc=\"../Seal.png\"/>{stamps}</SignedInfo></Signature>");
            entries[prefix + "Seal.png"] = ImageIoTests.ImageBytes(index == 1);
        }
        Assert.Equal(pageIds[0], pageIds[1]);
        entries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(input,
            new OfdPackageLoadOptions { DocumentIndex = 1 });
        var appearances = OfdSignatureAppearanceReader.Read(selected, new HashSet<string> { pageIds[1] }, 1);
        Assert.Equal(entries["Doc_1/Signs/Seal.png"], Assert.Single(appearances).Data);
        input.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter(new OfdToPdfOptions
            { PackageLoadOptions = new OfdPackageLoadOptions { DocumentIndex = 1 } }).ConvertAsync(input, pdf);
        Assert.True(pdf.Length > 100);
        input.Position = 0;
        using var image = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions
            { PixelsPerMillimeter = 1, MaxSignatureAppearanceCount = 1,
                PackageLoadOptions = new OfdPackageLoadOptions { DocumentIndex = 1 } }).ConvertAsync(input, image);
        Assert.True(image.Length > 100);
    }

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
