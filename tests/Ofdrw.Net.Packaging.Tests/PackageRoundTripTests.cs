using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class PackageRoundTripTests
{
    [Fact]
    public async Task Writer_ShouldGenerateRequiredEntries_AndReaderCanParse()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "hello ofd",
                    XMillimeters = 10,
                    YMillimeters = 12,
                    WidthMillimeters = 60,
                    HeightMillimeters = 8,
                    FontSizeMillimeters = 4
                }
            }
        });

        var package = builder.Build();
        var writer = new OfdPackageWriter();

        await using var ms = new MemoryStream();
        await writer.WriteAsync(package, ms);
        ms.Position = 0;

        var loader = new OfdPackageLoader();
        var archive = await loader.LoadAsync(ms);

        Assert.Contains("OFD.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Document.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Pages/Page_0/Content.xml", archive.EntryNames);

        var pageXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        var ns = pageXml.Root!.Name.Namespace;
        var textObject = pageXml.Descendants(ns + "TextObject").Single();
        Assert.Equal("1", textObject.Attribute("Font")?.Value);
        Assert.Equal("hello ofd", textObject.Element(ns + "TextCode")?.Value);

        var publicResXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/PublicRes.xml"));
        var publicResNs = publicResXml.Root!.Name.Namespace;
        var font = publicResXml.Descendants(publicResNs + "Font").Single();
        Assert.Equal("1", font.Attribute("ID")?.Value);
        Assert.Equal("SimSun", font.Attribute("FontName")?.Value);

        ms.Position = 0;
        var reader = new OfdReader();
        var parsed = await reader.ReadAsync(ms);

        Assert.Single(parsed.Pages);
        var firstText = parsed.Pages[0].Elements.OfType<OfdTextElement>().First();
        Assert.Equal("hello ofd", firstText.Text);
        Assert.Equal("SimSun", firstText.FontName);
    }

    [Fact]
    public async Task Writer_ShouldStoreImagesInDocumentResources_AndReaderCanResolveThem()
    {
        var imageData = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdImageElement
                {
                    XMillimeters = 0,
                    YMillimeters = 0,
                    WidthMillimeters = 210,
                    HeightMillimeters = 297,
                    Data = imageData,
                    MediaType = "image/png"
                }
            }
        });

        var writer = new OfdPackageWriter();
        await using var ms = new MemoryStream();
        await writer.WriteAsync(builder.Build(), ms);
        ms.Position = 0;

        var loader = new OfdPackageLoader();
        var archive = await loader.LoadAsync(ms);

        Assert.Contains("Doc_0/DocumentRes.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Res/Image_1.png", archive.EntryNames);
        Assert.DoesNotContain("Doc_0/Pages/Page_0/PageRes.xml", archive.EntryNames);

        var documentResXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/DocumentRes.xml"));
        var ns = documentResXml.Root!.Name.Namespace;
        Assert.Equal("Res", documentResXml.Root.Attribute("BaseLoc")?.Value);
        var media = documentResXml.Descendants(ns + "MultiMedia").Single();
        Assert.Equal("1", media.Attribute("ID")?.Value);
        Assert.Equal("PNG", media.Attribute("Format")?.Value);
        Assert.Equal("Image_1.png", media.Element(ns + "MediaFile")?.Value);

        var pageXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        var pageNs = pageXml.Root!.Name.Namespace;
        var imageObject = pageXml.Descendants(pageNs + "ImageObject").Single();
        Assert.Equal("210 0 0 297 0 0", imageObject.Attribute("CTM")?.Value);

        ms.Position = 0;
        var reader = new OfdReader();
        var parsed = await reader.ReadAsync(ms);
        var image = Assert.Single(parsed.Pages[0].Elements.OfType<OfdImageElement>());
        Assert.Equal(imageData, image.Data);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal("Image_1.png", image.FileName);
    }
}
