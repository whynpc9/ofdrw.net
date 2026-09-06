using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Validation;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class EditingIntegrityTests
{
    [Fact]
    public async Task RemovedPageContentUsedAsALiveTemplate_ShouldRemainReachable()
    {
        var original = await RoundTrip(CreateDocument());
        var ns = (XNamespace)original.Options.Namespace;
        XDocument Xml(string path) => XDocument.Parse(Encoding.UTF8.GetString(original.PreservedEntries[path]).TrimStart('\uFEFF'));
        var document = Xml("Doc_0/Document.xml");
        var common = document.Root!.Element(ns + "CommonData")!;
        common.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "500"), new XAttribute("BaseLoc", "Pages/Page_1/Content.xml")));
        common.Element(ns + "MaxUnitID")!.Value = "500";
        original.PreservedEntries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString());
        var first = Xml("Doc_0/Pages/Page_0/Content.xml");
        first.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "500"), new XAttribute("ZOrder", "Background")));
        original.PreservedEntries["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes(first.ToString());
        using var input = new MemoryStream();
        using (var zip = new ZipArchive(input, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in original.PreservedEntries)
            {
                using var stream = zip.CreateEntry(entry.Key).Open();
                stream.Write(entry.Value);
            }
        }
        input.Position = 0;
        var loaded = await new OfdReader().ReadAsync(input);
        OfdDocumentEditor.RemovePages(loaded, [1]);
        using var output = new MemoryStream();
        var report = await new OfdPackageWriter().WriteWithResultAsync(loaded, output);
        output.Position = 0;
        var result = await new OfdReader().ReadAsync(output);
        Assert.Single(result.Pages);
        Assert.Contains(Assert.Single(result.Pages[0].Templates).Elements.OfType<OfdTextElement>(), text => text.Text == "REMOVED_PAGE_TEXT");
        Assert.Contains(report.Diagnostics, warning => warning.Contains("shared content"));
    }

    [Fact]
    public async Task RemovedPage_ShouldLeaveNoPageTextOrExclusivePayloads()
    {
        var package = CreateDocument();
        var loaded = await RoundTrip(package);
        var sharedImage = loaded.Pages[0].Elements.OfType<OfdImageElement>().Single().Data;
        var deletedPath = loaded.Pages[1].SourceEntryPath!;
        OfdDocumentEditor.RemovePages(loaded, [1]);
        using var output = new MemoryStream();
        var result = await new OfdPackageWriter().WriteWithResultAsync(loaded, output);
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.Contains(deletedPath, result.RemovedEntries);
        Assert.DoesNotContain(archive.EntryNames, path => Encoding.UTF8.GetString(archive.GetBytes(path)).Contains("REMOVED_PAGE_TEXT"));
        Assert.DoesNotContain(archive.EntryNames, path => archive.GetBytes(path).SequenceEqual(new byte[] { 8, 8, 8 }));
        Assert.DoesNotContain(archive.EntryNames, path => archive.GetBytes(path).SequenceEqual(new byte[] { 9, 9, 9 }));
        Assert.Contains(archive.EntryNames, path => archive.GetBytes(path).SequenceEqual(sharedImage));
        Assert.Equal(new byte[] { 7, 6, 5 }, archive.GetBytes("Doc_0/Extensions/opaque.bin"));
        Assert.DoesNotContain(OfdPackageStructureChecker.Check(archive), issue => issue.IsError);
        output.Position = 0;
        Assert.Single((await new OfdReader().ReadAsync(output)).Pages);
    }

    [Fact]
    public async Task RemovingFirstPage_ShouldRetainSurvivingSourcePaths()
    {
        var package = CreateDocument();
        package.DocumentEntryPath = "Doc_0/Book.xml";
        var loaded = await RoundTrip(package);
        var path = loaded.Pages[1].SourceEntryPath;
        OfdDocumentEditor.RemovePages(loaded, [0]);
        var result = await RoundTrip(loaded);
        Assert.Equal("Doc_0/Book.xml", result.DocumentEntryPath);
        Assert.Equal(path, Assert.Single(result.Pages).SourceEntryPath);
        Assert.Contains(result.Pages[0].Elements.OfType<OfdTextElement>(), text => text.Text == "REMOVED_PAGE_TEXT");
    }

    [Fact]
    public async Task UnknownXml_ShouldRetainUncertainResourcesAndReportWhy()
    {
        var loaded = await RoundTrip(CreateDocument());
        loaded.PreservedEntries["Doc_0/Extensions/opaque.xml"] = Encoding.UTF8.GetBytes("<unparseable");
        OfdDocumentEditor.RemovePages(loaded, [1]);
        using var output = new MemoryStream();
        var result = await new OfdPackageWriter().WriteWithResultAsync(loaded, output);
        Assert.Contains(result.Diagnostics, warning => warning.Contains("opaque.xml"));
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain(archive.EntryNames, path => path == "Doc_0/Pages/Page_1/Content.xml");
        Assert.Contains(archive.EntryNames, path => archive.GetBytes(path).SequenceEqual(new byte[] { 8, 8, 8 }));
    }

    [Fact]
    public async Task RewrittenSignedPackage_ShouldRemoveInvalidatedSignatureDeclarations()
    {
        var loaded = await RoundTrip(CreateDocument());
        var ns = (XNamespace)loaded.Options.Namespace;
        var root = XDocument.Parse(Encoding.UTF8.GetString(loaded.PreservedEntries["OFD.xml"]).TrimStart('\uFEFF'));
        root.Descendants(ns + "DocBody").Single().Add(new XElement(ns + "Signatures", "Doc_0/Signs/Signatures.xml"));
        loaded.PreservedEntries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
        loaded.PreservedDocBodyElements.Add(new XElement(ns + "Signatures", "Doc_0/Signs/Signatures.xml").ToString());
        loaded.PreservedEntries["Doc_0/Signs/Signatures.xml"] = Encoding.UTF8.GetBytes(
            $"<Signatures xmlns='{ns}'><Signature ID='100' BaseLoc='Sign_0/Signature.xml'/></Signatures>");
        var exclusivePayload = loaded.PreservedEntries.Single(pair => pair.Value.SequenceEqual(new byte[] { 8, 8, 8 })).Key;
        loaded.PreservedEntries["Doc_0/Signs/Sign_0/Signature.xml"] = Encoding.UTF8.GetBytes(
            $"<Signature xmlns='{ns}'><SignedInfo><References><Reference FileRef='/{exclusivePayload}'/></References></SignedInfo><SignedValue>SignedValue.dat</SignedValue></Signature>");
        loaded.PreservedEntries["Doc_0/Signs/Sign_0/SignedValue.dat"] = [1, 2, 3];
        OfdDocumentEditor.RemovePages(loaded, [1]);
        using var output = new MemoryStream();
        var result = await new OfdPackageWriter().WriteWithResultAsync(loaded, output);
        Assert.True(result.SignaturesInvalidated);
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain("Signatures", archive.ReadUtf8Text("OFD.xml"));
        Assert.DoesNotContain(archive.EntryNames, path => path.Contains("/Signs/"));
        Assert.False(archive.Contains(exclusivePayload));
    }

    [Fact]
    public async Task Merge_ShouldKeepDistinctSameNamedFontsAndImageStyles()
    {
        OfdDocumentPackage Source(byte marker)
        {
            var source = new OfdDocumentPackage();
            source.Fonts.Add(new OfdFontResource { Id = "20", FontName = "SameFamily", FileName = "same.ttf", Data = [marker] });
            var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
            page.Elements.Add(new OfdTextElement { FontName = "SameFamily", FontResourceId = "20", Text = marker.ToString() });
            page.Elements.Add(new OfdImageElement
            {
                Data = [1, 2, 3], WidthMillimeters = 20, HeightMillimeters = 20,
                Alpha = 128, Transform = [0, 20, -20, 0, 20, 0],
                SourceXml = "<ImageObject VendorStyle='retained' />",
                ClipsXml = "<Clips><Clip><Area><Path ID='100'><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>"
            });
            source.Pages.Add(page);
            return source;
        }

        var merged = OfdDocumentMerger.Merge([Source(1), Source(2)]);
        var result = await RoundTrip(merged);
        Assert.Equal(2, result.Fonts.Count);
        for (var index = 0; index < 2; index++)
        {
            var text = result.Pages[index].Elements.OfType<OfdTextElement>().Single();
            var font = result.Fonts.Single(font => font.Id == text.FontResourceId);
            Assert.Equal(new byte[] { (byte)(index + 1) }, font.Data);
            var image = result.Pages[index].Elements.OfType<OfdImageElement>().Single();
            Assert.Equal(128, image.Alpha);
            Assert.Equal(new double[] { 0, 20, -20, 0, 20, 0 }, image.Transform);
            Assert.Contains("VendorStyle=\"retained\"", image.SourceXml);
            Assert.Contains("AbbreviatedData", image.ClipsXml);
        }
        var clipIds = result.Pages.SelectMany(page => page.Elements.OfType<OfdImageElement>())
            .Select(image => XElement.Parse(image.ClipsXml!).Descendants().Single(node => node.Name.LocalName == "Path").Attribute("ID")!.Value).ToList();
        Assert.Equal(2, clipIds.Distinct().Count());
    }

    [Fact]
    public void Merge_ShouldRejectUnmappedReferencesOrReportExplicitlySkippedObjects()
    {
        var source = new OfdDocumentPackage();
        var page = new OfdPage();
        page.Elements.Add(new OfdPathElement { SourceXml = "<PathObject DrawParam='123' />" });
        source.Pages.Add(page);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
        var result = OfdDocumentMerger.MergeWithResult([source], new OfdDocumentMergeOptions { SkipUnsupportedRawElements = true });
        Assert.Empty(result.Package.Pages[0].Elements);
        Assert.Contains("DrawParam", Assert.Single(result.Diagnostics));
    }

    private static OfdDocumentPackage CreateDocument()
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "20", FontName = "Shared", Data = [4, 4, 4] });
        package.Fonts.Add(new OfdFontResource { Id = "21", FontName = "Removed", Data = [9, 9, 9] });
        for (var index = 0; index < 2; index++)
        {
            var page = new OfdPage { Index = index, WidthMillimeters = 210, HeightMillimeters = 297 };
            page.Elements.Add(new OfdTextElement { Text = index == 0 ? "KEEP" : "REMOVED_PAGE_TEXT", FontName = index == 0 ? "Shared" : "Removed", FontResourceId = index == 0 ? "20" : "21" });
            page.Elements.Add(new OfdImageElement { Data = [1, 1, 1], WidthMillimeters = 10, HeightMillimeters = 10 });
            if (index == 1) page.Elements.Add(new OfdImageElement { Data = [8, 8, 8], WidthMillimeters = 10, HeightMillimeters = 10 });
            package.Pages.Add(page);
        }
        package.PreservedEntries["Doc_0/Extensions/opaque.bin"] = [7, 6, 5];
        return package;
    }

    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
}
