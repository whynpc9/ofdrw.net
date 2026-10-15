using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx.Internal;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class StagedRasterNamespaceAdapterTests
{
    [Fact]
    public async Task ClosedRasterRebuildUsesFreshDefinitionsAndLeavesInputUnchanged()
    {
        var source = await Source(); var originalPage = source.Pages[0]; var original = (OfdImageElement)originalPage.Elements[0];
        var entries = source.PreservedEntries.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray()); var xml = original.SourceXml;
        var rebuilt = StagedRasterNamespaceAdapter.Rebuild(source, default, OfdConstants.Namespace);
        Assert.Equal(OfdConstants.Namespace, rebuilt.Options.Namespace);
        Assert.NotSame(source, rebuilt); Assert.NotSame(originalPage, rebuilt.Pages[0]);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(rebuilt.Pages[0].Elements));
        Assert.Null(rebuilt.Pages[0].Id); Assert.Null(image.ObjectId); Assert.NotSame(original.Data, image.Data); Assert.Equal(original.Data, image.Data);
        Assert.Equal(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(8)), rebuilt.Options.Metadata.CreationDate);
        Assert.Same(originalPage, source.Pages[0]); Assert.Equal(xml, original.SourceXml); Assert.Equal(OfdConstants.StandardNamespace, source.Options.Namespace);
        foreach (var entry in entries) Assert.Equal(entry.Value, source.PreservedEntries[entry.Key]);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(rebuilt, output); Assert.True(output.Length > 0);
    }

    [Theory]
    [InlineData("root-foreign")]
    [InlineData("root-unqualified")]
    [InlineData("foreign-image-reference")]
    [InlineData("unknown-image-reference")]
    [InlineData("page-comment")]
    [InlineData("layer-qname")]
    [InlineData("media-format")]
    [InlineData("opaque-entry")]
    public async Task UnknownSourceIsExplicitlyRejectedWithoutChangingInput(string mutation)
    {
        var source = await Source(); var path = source.Pages[0].SourceEntryPath!; var ns = XNamespace.Get(source.Options.Namespace);
        var entryPath = mutation.StartsWith("root-") ? "OFD.xml" : mutation == "media-format" ? "Doc_0/DocumentRes.xml" : path;
        using var xmlInput = new MemoryStream(source.PreservedEntries[entryPath]);
        var root = XDocument.Load(xmlInput);
        switch (mutation)
        {
            case "root-foreign": root.Root!.SetAttributeValue(XName.Get("ObjectRef", "urn:vendor"), "7"); break;
            case "root-unqualified": root.Root!.SetAttributeValue("ObjectRef", "7"); break;
            case "foreign-image-reference": root.Descendants(ns + "ImageObject").Single().SetAttributeValue(XName.Get("ResourceID", "urn:vendor"), "7"); break;
            case "unknown-image-reference": root.Descendants(ns + "ImageObject").Single().SetAttributeValue("RefID", "7"); break;
            case "layer-qname": root.Descendants(ns + "Layer").Single().SetAttributeValue("Type", "v:Layer"); break;
            case "media-format": root.Descendants(ns + "MultiMedia").Single().SetAttributeValue("Format", "v:Codec"); break;
            case "page-comment": root.Descendants(ns + "Layer").Single().AddFirst(new XComment("retained vendor marker")); break;
            case "opaque-entry": source.PreservedEntries["Doc_0/vendor.dat"] = Encoding.UTF8.GetBytes("<Vendor RefID='7'/>"); break;
        }
        if (mutation != "opaque-entry") source.PreservedEntries[entryPath] = Encoding.UTF8.GetBytes(root.ToString());
        using (var archive = Zip(source.PreservedEntries)) source = await new OfdReader().ReadAsync(archive);
        var before = source.PreservedEntries.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray()); var image = (OfdImageElement)source.Pages[0].Elements[0]; var sourceXml = image.SourceXml;
        Assert.Throws<NotSupportedException>(() => StagedRasterNamespaceAdapter.Rebuild(source, default));
        Assert.Equal(sourceXml, image.SourceXml); Assert.Equal(OfdConstants.StandardNamespace, source.Options.Namespace);
        foreach (var entry in before) Assert.Equal(entry.Value, source.PreservedEntries[entry.Key]);
    }

    [Fact]
    public async Task QualifiedResourceOwnerIsNotAnAliasOfUnqualifiedOwner()
    {
        var source = await Source(); var image = (OfdImageElement)source.Pages[0].Elements[0]; image.ResourceId = "v:" + image.ResourceId;
        Assert.Throws<NotSupportedException>(() => StagedRasterNamespaceAdapter.Rebuild(source, default));
    }

    [Theory]
    [InlineData("Foo", "ofd:Unknown")]
    [InlineData("ResourceID", "999")]
    public async Task EffectiveImageFragmentCannotBypassArchiveClosure(string name, string value)
    {
        var source = await Source(); var image = (OfdImageElement)source.Pages[0].Elements[0]; var root = XElement.Parse(image.SourceXml!);
        root.SetAttributeValue(name, value); image.SourceXml = root.ToString();
        Assert.Throws<NotSupportedException>(() => StagedRasterNamespaceAdapter.Rebuild(source, default));
        Assert.Equal(value, XElement.Parse(image.SourceXml).Attribute(name)!.Value);
    }

    private static async Task<OfdDocumentPackage> Source()
    {
        var package = new OfdDocumentPackage(); package.Options.Namespace = OfdConstants.StandardNamespace;
        package.Options.Metadata.CreationDate = new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.FromHours(8));
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdImageElement { Data = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j9AoAAAAASUVORK5CYII="), MediaType = "image/png", FileName = "image.png", WidthMillimeters = 100, HeightMillimeters = 100 });
        package.Pages.Add(page); using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); output.Position = 0;
        return await new OfdReader().ReadAsync(output);
    }
    private static MemoryStream Zip(IDictionary<string, byte[]> entries)
    {
        var output = new MemoryStream(); using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var pair in entries)
        { using var entry = zip.CreateEntry(pair.Key).Open(); entry.Write(pair.Value); }
        output.Position = 0; return output;
    }
}
