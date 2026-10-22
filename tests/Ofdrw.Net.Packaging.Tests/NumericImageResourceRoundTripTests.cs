using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class NumericImageResourceRoundTripTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportedImageReferenceKeepsEquivalentSpellingAndOwner(bool annotation)
    {
        var source = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        var image = new OfdImageElement { Data = [1, 2, 3, 4], FileName = "fixture.png", MediaType = "image/png",
            WidthMillimeters = 20, HeightMillimeters = 20 };
        if (annotation)
        {
            page.Annotations = new();
            var item = new OfdAnnotation { Type = OfdAnnotationType.Stamp, Creator = "test",
                LastModifiedDate = new DateTime(2026, 10, 22),
                Appearance = new() { WidthMillimeters = 30, HeightMillimeters = 30 } };
            item.Appearance.Elements.Add(image);
            page.Annotations.Items.Add(item);
        }
        else
        {
            source.CompositeResources = new();
            var resource = new OfdCompositeResource { WidthMillimeters = 30, HeightMillimeters = 30 };
            resource.Elements.Add(image);
            source.CompositeResources.Items.Add(resource);
            page.Elements.Add(new OfdCompositeElement { Resource = resource, WidthMillimeters = 30, HeightMillimeters = 30 });
        }
        source.Pages.Add(page);

        using var seed = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(source, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var mediaPath = entries.Keys.Single(path => path.EndsWith("DocumentRes.xml", StringComparison.Ordinal));
        var mediaXml = Parse(entries[mediaPath]);
        var media = Assert.Single(mediaXml.Descendants(), node => node.Name.LocalName == "MultiMedia");
        media.SetAttributeValue("ID", "030");
        media.SetAttributeValue(XName.Get("keep", "urn:vendor"), "media");
        entries[mediaPath] = Encoding.UTF8.GetBytes(mediaXml.ToString(SaveOptions.DisableFormatting));
        var imagePath = entries.Keys.Single(path => path.EndsWith(".xml", StringComparison.Ordinal) &&
            Parse(entries[path]).Descendants().Any(node => node.Name.LocalName == "ImageObject"));
        var imageXml = Parse(entries[imagePath]);
        var imageNode = Assert.Single(imageXml.Descendants(), node => node.Name.LocalName == "ImageObject");
        imageNode.SetAttributeValue("ResourceID", "30");
        imageNode.SetAttributeValue(XName.Get("keep", "urn:vendor"), "image");
        entries[imagePath] = Encoding.UTF8.GetBytes(imageXml.ToString(SaveOptions.DisableFormatting));

        using var importedBytes = MultiDocumentFixture.Zip(entries);
        var imported = await new OfdReader().ReadAsync(importedBytes);
        var importedImage = annotation
            ? Assert.IsType<OfdImageElement>(Assert.Single(imported.Pages[0].Annotations!.Items[0].Appearance.Elements))
            : Assert.IsType<OfdImageElement>(Assert.Single(imported.CompositeResources!.Items[0].Elements));
        Assert.Equal("30", importedImage.ResourceId);
        Assert.Equal(image.Data, importedImage.Data);

        using var savedBytes = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(imported, savedBytes);
        var savedEntries = MultiDocumentFixture.Unzip(savedBytes);
        var savedMedia = Parse(savedEntries[mediaPath]).Descendants()
            .Single(node => node.Name.LocalName == "MultiMedia");
        var savedImage = Parse(savedEntries[imagePath]).Descendants()
            .Single(node => node.Name.LocalName == "ImageObject");
        Assert.Equal("030", savedMedia.Attribute("ID")?.Value);
        Assert.Equal("media", savedMedia.Attribute(XName.Get("keep", "urn:vendor"))?.Value);
        Assert.Equal("30", savedImage.Attribute("ResourceID")?.Value);
        Assert.Equal("image", savedImage.Attribute(XName.Get("keep", "urn:vendor"))?.Value);
        savedBytes.Position = 0;
        var readback = await new OfdReader().ReadAsync(savedBytes);
        var roundTripImage = annotation
            ? Assert.IsType<OfdImageElement>(Assert.Single(readback.Pages[0].Annotations!.Items[0].Appearance.Elements))
            : Assert.IsType<OfdImageElement>(Assert.Single(readback.CompositeResources!.Items[0].Elements));
        Assert.Equal("30", roundTripImage.ResourceId);
        Assert.Equal(image.Data, roundTripImage.Data);

        if (!annotation)
        {
            var collidingEntries = new Dictionary<string, byte[]>(entries);
            var collidingXml = Parse(entries[mediaPath]);
            var first = collidingXml.Descendants().Single(node => node.Name.LocalName == "MultiMedia");
            var duplicate = new XElement(first);
            duplicate.SetAttributeValue("ID", "30");
            first.AddAfterSelf(duplicate);
            collidingEntries[mediaPath] = Encoding.UTF8.GetBytes(collidingXml.ToString(SaveOptions.DisableFormatting));
            using var collidingBytes = MultiDocumentFixture.Zip(collidingEntries);
            await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(collidingBytes));
        }

        importedImage.ResourceId = "31";
        using var rejected = new MemoryStream();
        if (annotation)
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(imported, rejected));
        else
            await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(imported, rejected));
        Assert.Equal(0, rejected.Length);
    }

    private static XDocument Parse(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return XDocument.Load(stream);
    }
}
