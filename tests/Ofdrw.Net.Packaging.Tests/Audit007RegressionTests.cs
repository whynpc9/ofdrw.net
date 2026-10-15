using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit007RegressionTests
{
    [Theory]
    [InlineData("")]
    [InlineData("Doc_0/")]
    public async Task Audit007_RootDocumentSidecars_AndNestedControl(string prefix)
    {
        var package = new OfdDocumentPackage { DocumentEntryPath = prefix + "Document.xml" };
        package.Attachments.Add(new OfdAttachment { Name = "payload.bin", Data = [0, 1, 2, 255] });
        package.CustomTags.Add("patient", "root-value");
        package.Pages.Add(new OfdPage
        {
            Index = 0, WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdPathElement { AbbreviatedData = "M 0 0 L 10 10", WidthMillimeters = 10, HeightMillimeters = 10 } },
            Annotations = new OfdAnnotationCollection
            {
                Items = { new OfdAnnotation
                {
                    Creator = "Audit007", LastModifiedDate = new DateTime(2026, 10, 15), Remark = "annotation-value",
                    Appearance = new() { WidthMillimeters = 10, HeightMillimeters = 10 }
                } }
            }
        });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        using (var zip = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read))
            Assert.All(zip.Entries, entry => Assert.False(entry.FullName.StartsWith('/'), entry.FullName));
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.Equal(new byte[] { 0, 1, 2, 255 }, archive.GetBytes(prefix + "Attachs/Attach_1_payload.bin"));
        var document = Xml(archive, prefix + "Document.xml");
        var ns = document.Root!.Name.Namespace;
        Assert.Equal("Attachs/Attachments.xml", document.Root.Element(ns + "Attachments")!.Value);
        Assert.Equal("Tags/CustomTags.xml", document.Root.Element(ns + "CustomTags")!.Value);
        Assert.Equal("Pages/Page_0/Content.xml", document.Descendants(ns + "Page").Single().Attribute("BaseLoc")!.Value);
        var indexPath = prefix + document.Root.Element(ns + "Annotations")!.Value;
        var annotationIndex = Xml(archive, indexPath);
        var annotationPath = prefix + "Annots/" + annotationIndex.Descendants(ns + "FileLoc").Single().Value;
        Assert.Equal("annotation-value", Xml(archive, annotationPath).Descendants(ns + "Remark").Single().Value);
        Assert.Equal("root-value", Xml(archive, prefix + "Tags/CustomTag_EMR.xml").Descendants(ns + "Tag").Single().Attribute("Value")!.Value);
        output.Position = 0;
        var read = await new OfdReader().ReadAsync(output);
        Assert.Equal(prefix + "Document.xml", read.DocumentEntryPath);
        Assert.Equal(new byte[] { 0, 1, 2, 255 }, Assert.Single(read.Attachments).Data);
        Assert.Equal("root-value", read.CustomTags["patient"]);
        Assert.Equal("annotation-value", Assert.Single(Assert.Single(read.Pages).Annotations!.Items).Remark);
        Assert.Equal("M 0 0 L 10 10", Assert.IsType<OfdPathElement>(read.Pages[0].Elements.Single()).AbbreviatedData);
        using var rewritten = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(read, rewritten);
        rewritten.Position = 0;
        var reread = await new OfdReader().ReadAsync(rewritten);
        Assert.Equal(read.Attachments[0].Data, Assert.Single(reread.Attachments).Data);
        Assert.Equal("annotation-value", Assert.Single(reread.Pages[0].Annotations!.Items).Remark);
    }

    [Fact]
    public async Task Audit007_AbsoluteXmlReferencesAndRootPageRemainReadable()
    {
        var package = new OfdDocumentPackage { DocumentEntryPath = "Document.xml", PublicResourceLocation = "/PublicRes.xml" };
        package.Pages.Add(new OfdPage
        {
            Index = 0, SourceEntryPath = "Content.xml", WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdPathElement { AbbreviatedData = "M 1 1 L 2 2", WidthMillimeters = 10, HeightMillimeters = 10 } }
        });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.Contains("Content.xml", archive.EntryNames);
        Assert.Contains("PublicRes.xml", archive.EntryNames);
        Assert.DoesNotContain(archive.EntryNames, name => name.StartsWith('/'));
        var document = Xml(archive, "Document.xml");
        Assert.Equal("/PublicRes.xml", document.Descendants(document.Root!.Name.Namespace + "PublicRes").Single().Value);
        output.Position = 0;
        Assert.Equal("M 1 1 L 2 2", Assert.IsType<OfdPathElement>((await new OfdReader().ReadAsync(output)).Pages[0].Elements.Single()).AbbreviatedData);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Doc_0/")]
    public async Task Audit007_SidecarResourceRoleCollisionStillRejects(string prefix)
    {
        var package = new OfdDocumentPackage { DocumentEntryPath = prefix + "Document.xml", PublicResourceLocation = "Tags/CustomTags.xml" };
        package.CustomTags.Add("key", "value");
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Audit007_LeadingSlashZipMemberStillRejects()
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var member = zip.CreateEntry("/payload.bin").Open();
            member.WriteByte(42);
        }
        output.Position = 0;
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(output));
    }

    private static XDocument Xml(OfdPackageArchive archive, string path) => XDocument.Parse(archive.ReadUtf8Text(path));
}
