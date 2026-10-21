using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class MultiDocumentTests
{
    [Fact]
    public async Task Reader_ListsAndSelectsDistinctDocuments()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        using var input = MultiDocumentFixture.Zip(entries);
        var reader = new OfdReader();
        var listed = await reader.ListDocumentsAsync(input);
        Assert.Equal(["Doc_0/Document.xml", "Doc_1/Document.xml"], listed.Select(item => item.DocumentEntryPath));
        Assert.Equal(["First", "Second"], listed.Select(item => item.Title));
        input.Position = 0;
        var selected = await reader.ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Equal("Doc_1/Document.xml", selected.DocumentEntryPath);
        Assert.Contains("SECOND DOC1", selected.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text));
        input.Position = 0;
        var first = await reader.ReadAsync(input);
        Assert.Equal("Doc_0/Document.xml", first.DocumentEntryPath);
        input.Position = 0;
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => reader.ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 2 }));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public async Task SelectedWatermarkSave_PreservesOtherDocumentEntries(int index)
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        using var input = MultiDocumentFixture.Zip(entries);
        var package = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = index });
        OfdWatermark.AddText(package, [0], "SELECTED", new OfdWatermarkOptions { LayerId = "selected" });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        var actual = MultiDocumentFixture.Unzip(output);
        var other = index == 0 ? "Doc_1/" : "Doc_0/";
        Assert.All(entries.Where(entry => entry.Key.StartsWith(other, StringComparison.Ordinal)), entry => Assert.Equal(entry.Value, actual[entry.Key]));
        var ns = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"])).Root!.Name.Namespace;
        var before = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"])).Root!.Elements(ns + "DocBody").ToArray();
        var after = XDocument.Parse(Encoding.UTF8.GetString(actual["OFD.xml"])).Root!.Elements(ns + "DocBody").ToArray();
        Assert.True(XNode.DeepEquals(before[1 - index], after[1 - index]));
        using var reread = MultiDocumentFixture.Zip(actual);
        var edited = await new OfdReader().ReadAsync(reread, new OfdPackageLoadOptions { DocumentIndex = index });
        Assert.Contains("SELECTED", edited.Pages[0].Elements.OfType<OfdTextElement>().Select(text => text.Text));
    }

    [Fact]
    public async Task SignedAndCrossDocumentInputs_FailBeforeWriting()
    {
        foreach (var signed in new[] { true, false })
        {
            var entries = await MultiDocumentFixture.CreateEntriesAsync();
            var ns = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"])).Root!.Name.Namespace;
            if (signed)
            {
                var root = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"]));
                root.Root!.Elements(ns + "DocBody").Last().Add(new XElement(ns + "Signatures", "Doc_1/Signs/Signatures.xml"));
                entries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
                entries["Doc_1/Signs/Signatures.xml"] = Encoding.UTF8.GetBytes("signed marker");
            }
            else
            {
                var doc = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_1/Document.xml"]));
                doc.Descendants(ns + "PublicRes").Single().Value = "../Doc_0/PublicRes.xml";
                entries["Doc_1/Document.xml"] = Encoding.UTF8.GetBytes(doc.ToString());
            }
            using var input = MultiDocumentFixture.Zip(entries);
            var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 0 });
            using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, output));
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CompoundResourceBaseLoc_RejectsCrossDocumentFontDependency(bool nestedResource)
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var resource = XDocument.Load(new MemoryStream(entries["Doc_0/PublicRes.xml"]));
        resource.Root!.SetAttributeValue("BaseLoc", nestedResource ? ".." : "../Doc_1/Res");
        var font = resource.Descendants().First(node => node.Name.LocalName == "Font");
        font.Elements().Where(node => node.Name.LocalName == "FontFile").Remove();
        font.Add(new XElement(resource.Root.Name.Namespace + "FontFile", nestedResource ? "../../Doc_1/Res/font.ttf" : "font.ttf"));
        if (nestedResource)
        {
            entries.Remove("Doc_0/PublicRes.xml");
            entries["Doc_0/Res/Sub/PublicRes.xml"] = Encoding.UTF8.GetBytes(resource.ToString());
            var document = XDocument.Load(new MemoryStream(entries["Doc_0/Document.xml"]));
            document.Descendants().First(node => node.Name.LocalName == "PublicRes").Value = "Res/Sub/PublicRes.xml";
            entries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString());
        }
        else entries["Doc_0/PublicRes.xml"] = Encoding.UTF8.GetBytes(resource.ToString());
        entries["Doc_1/Res/font.ttf"] = [1, 2, 3];
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task CompoundResourceBaseLoc_RejectsCrossDocumentMediaDependency()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var resource = XDocument.Load(new MemoryStream(entries["Doc_0/PublicRes.xml"]));
        var ns = resource.Root!.Name.Namespace;
        resource.Root.SetAttributeValue("BaseLoc", "..");
        resource.Root.Add(new XElement(ns + "MultiMedias",
            new XElement(ns + "MultiMedia", new XAttribute("ID", "987"),
                new XAttribute("MediaFile", "../../Doc_1/Res/image.png"))));
        entries.Remove("Doc_0/PublicRes.xml");
        entries["Doc_0/Res/Sub/PublicRes.xml"] = Encoding.UTF8.GetBytes(resource.ToString());
        var document = XDocument.Load(new MemoryStream(entries["Doc_0/Document.xml"]));
        document.Descendants().First(node => node.Name.LocalName == "PublicRes").Value = "Res/Sub/PublicRes.xml";
        entries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString());
        entries["Doc_1/Res/image.png"] = [1, 2, 3];
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task RemovingSelectedSecondPage_PrunesOnlySecondDocument()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        var removedPath = selected.Pages[0].SourceEntryPath!;
        OfdDocumentEditor.RemovePages(selected, [0]);
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(selected, output);
        var written = MultiDocumentFixture.Unzip(output);
        Assert.DoesNotContain(removedPath, written.Keys);
        Assert.All(entries.Where(entry => entry.Key.StartsWith("Doc_0/", StringComparison.Ordinal)), entry => Assert.Equal(entry.Value, written[entry.Key]));
        using var reread = MultiDocumentFixture.Zip(written);
        var actual = await new OfdReader().ReadAsync(reread, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Single(actual.Pages);
        Assert.Contains("PAGE2", Assert.Single(actual.Pages[0].Elements.OfType<OfdTextElement>()).Text);
    }

    [Fact]
    public async Task ListingIncludesSeparateSignatureLists_WhileSignedEditIsRejected()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"]));
        var ns = root.Root!.Name.Namespace;
        var bodies = root.Root.Elements(ns + "DocBody").ToArray();
        for (var index = 0; index < bodies.Length; index++)
        {
            var path = "Doc_" + index + "/Signs/Signatures.xml";
            bodies[index].Add(new XElement(ns + "Signatures", path));
            entries[path] = Encoding.UTF8.GetBytes("signature marker " + index);
        }
        entries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
        using var input = MultiDocumentFixture.Zip(entries);
        var listed = await new OfdReader().ListDocumentsAsync(input);
        Assert.Equal("Doc_0/Signs/Signatures.xml", Assert.Single(listed[0].SignatureListPaths));
        Assert.Equal("Doc_1/Signs/Signatures.xml", Assert.Single(listed[1].SignatureListPaths));
        input.Position = 0;
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicateOrMissingDocRoot_IsRejected(bool missing)
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"]));
        var ns = root.Root!.Name.Namespace;
        var second = root.Root.Elements(ns + "DocBody").Last();
        var docRoot = second.Element(ns + "DocRoot")!;
        if (missing) docRoot.Remove(); else docRoot.Value = "doc_0/Document.xml";
        entries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
        using var input = MultiDocumentFixture.Zip(entries);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ListDocumentsAsync(input));
    }

    [Fact]
    public async Task ListingRejectsMissingDocumentEntry()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        entries.Remove("Doc_1/Document.xml");
        using var input = MultiDocumentFixture.Zip(entries);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ListDocumentsAsync(input));
    }

}
