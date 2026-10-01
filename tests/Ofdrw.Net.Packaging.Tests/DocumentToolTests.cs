using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class DocumentToolTests
{
    [Fact]
    public async Task Watermarks_KeepSelectedLayersAndSurviveMergeAndSave()
    {
        var source = await RoundTrip(Source());
        OfdWatermark.AddText(source, [1], "DRAFT 测试", new OfdWatermarkOptions { LayerId = "chosen", LayerType = "Foreground" });
        OfdWatermark.AddImage(source, [1], Png, "image/png", new OfdWatermarkOptions { LayerId = "chosen", LayerType = "Foreground", YMillimeters = 70 });
        Assert.DoesNotContain(source.Pages[0].Elements.OfType<OfdTextElement>(), text => text.Text.StartsWith("DRAFT"));
        var merged = await RoundTrip(OfdDocumentMerger.Merge([source, source]));
        Assert.Equal(2, merged.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Count(text => text.Text == "DRAFT 测试"));
        Assert.All(merged.Pages.Where(page => page.Index % 2 == 1), page =>
        {
            var watermark = page.Elements.OfType<OfdTextElement>().Single(text => text.Text.StartsWith("DRAFT"));
            Assert.Equal("Foreground", watermark.LayerType);
            Assert.Equal(watermark.LayerId, page.Elements.OfType<OfdImageElement>().Last().LayerId);
        });
    }

    [Fact]
    public void Watermark_PreflightRejectsBadLastPageOrCancellationWithoutMutation()
    {
        var source = Source();
        var count = source.Pages[0].Elements.Count;
        Assert.Throws<ArgumentException>(() => OfdWatermark.AddText(source, [0, 9], "draft"));
        Assert.Throws<ArgumentException>(() => OfdWatermark.AddText(source, [0], "draft", new OfdWatermarkOptions { XMillimeters = double.NaN }));
        using var cts = new CancellationTokenSource(); cts.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdWatermark.AddText(source, [0], "draft", cancellationToken: cts.Token));
        Assert.Throws<ArgumentException>(() => OfdWatermark.AddText(source, [0, 1], "draft", new OfdWatermarkOptions { MaxGeneratedTextCharacters = 5 }));
        Assert.Throws<InvalidDataException>(() => OfdWatermark.AddImage(source, [0], new byte[24], "image/png"));
        Assert.Equal(count, source.Pages[0].Elements.Count);
    }

    [Fact]
    public async Task Split_PreservesSelectionOrderAndSharedResourcesButRemovesPrivatePayloads()
    {
        var source = await RoundTrip(Source());
        var split = OfdDocumentSplitter.Split(source, [1]);
        var saved = await RoundTrip(split);
        Assert.Equal(2, source.Pages.Count);
        Assert.Equal("SECOND", saved.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
        Assert.DoesNotContain(saved.PreservedEntries.Values, data => data.SequenceEqual(new byte[] { 8, 8, 8 }));
        Assert.Contains(saved.PreservedEntries.Values, data => data.SequenceEqual(Png));
        Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(saved.Attachments).Data);
        var reversed = await RoundTrip(OfdDocumentSplitter.Split(source, [1, 0]));
        Assert.Equal(new[] { "SECOND", "FIRST" }, reversed.Pages.Select(page => page.Elements.OfType<OfdTextElement>().Single().Text));
        Assert.Throws<ArgumentException>(() => OfdDocumentSplitter.Split(source, [0, 0]));
        source.PreservedEntries["opaque.xml"] = Encoding.UTF8.GetBytes("<broken");
        Assert.ThrowsAny<Exception>(() => OfdDocumentSplitter.Split(source, [0]));
    }

    [Fact]
    public async Task Mix_PreservesFirstPhysicalBoxAndPerSourceTemplateAnnotationOrder()
    {
        var source = Source();
        source.Pages[0].XMillimeters = 5;
        source.Pages[0].Templates.Add(new OfdTemplateContent { ZOrder = "Background", Elements = { new OfdTextElement { Text = "BACKGROUND", LayerId = "same" } } });
        source.Pages[0].Templates.Add(new OfdTemplateContent { ZOrder = "Foreground", Elements = { new OfdTextElement { Text = "FOREGROUND", LayerId = "same" } } });
        source.Pages[0].AnnotationAppearances.Add(new OfdTextElement { Text = "NOTE", LayerId = "same", XMillimeters = 17 });
        var mixed = await RoundTrip(OfdDocumentMixer.Mix([new(source, 0), new(source, 1)]));
        var page = Assert.Single(mixed.Pages);
        Assert.Equal(5, page.XMillimeters);
        Assert.Equal(210, page.WidthMillimeters);
        Assert.Equal(new[] { "BACKGROUND", "FIRST", "FOREGROUND", "NOTE", "SECOND" }, page.Elements.OfType<OfdTextElement>().Select(text => text.Text));
        Assert.Equal(17, page.Elements.OfType<OfdTextElement>().Single(text => text.Text == "NOTE").XMillimeters);
        Assert.Equal(2, source.Pages.Count);
        Assert.Throws<ArgumentException>(() => OfdDocumentMixer.Mix([new(source, 0)], maxObjectCount: 1));
        Assert.Throws<ArgumentException>(() => OfdDocumentMixer.Mix([new(source, 0)], maxExpandedBytes: 1));
        source.Pages[0].Elements.Add(new OfdRawElement { Xml = "<Unknown Ref='99'/>" });
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
    }

    [Theory]
    [InlineData("/OFD.xml")]
    [InlineData("/Doc_0/Extensions/ordinary.bin")]
    [InlineData("SignedValue.dat")]
    public async Task CleanSignatures_OnlyDeletesOwnedUnreferencedPayloadsAcrossAllDocuments(string signedValue)
    {
        var source = await RoundTrip(Source());
        AddSignatures(source, signedValue);
        var root = Xml(source, "OFD.xml");
        var second = new XElement(root.Root!.Elements().Single());
        root.Root.Add(second);
        Put(source, "OFD.xml", root);
        source.PreservedEntries["Doc_0/Extensions/ordinary.bin"] = [77];
        using var input = Zip(source.PreservedEntries);
        using var output = new MemoryStream();
        var result = await OfdPackageSignatureCleaner.CleanAsync(input, output);
        output.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain("Signatures", archive.ReadUtf8Text("OFD.xml"));
        Assert.Contains("DocBody", archive.ReadUtf8Text("OFD.xml"));
        Assert.Equal(2, XDocument.Parse(archive.ReadUtf8Text("OFD.xml")).Root!.Elements().Count());
        Assert.Equal(new byte[] { 77 }, archive.GetBytes("Doc_0/Extensions/ordinary.bin"));
        Assert.False(archive.Contains("Doc_0/Signs/Sign_0/Seal.esl"));
        Assert.False(archive.Contains("Doc_0/Signs/Signatures.xml"));
        Assert.Equal(source.PreservedEntries["Doc_0/Document.xml"], archive.GetBytes("Doc_0/Document.xml"));
        if (signedValue == "SignedValue.dat") Assert.False(archive.Contains("Doc_0/Signs/Sign_0/SignedValue.dat"));
        else Assert.Contains(result.Diagnostics, warning => warning.Contains("Unowned"));
    }

    [Fact]
    public async Task CleanSignatures_PreservesSharedSealAndUnknownXml()
    {
        var source = await RoundTrip(Source()); AddSignatures(source, "SignedValue.dat");
        source.PreservedEntries["Doc_0/Extensions/shared.xml"] = Encoding.UTF8.GetBytes("<Extension File='/Doc_0/Signs/Sign_0/Seal.esl'/>");
        using var output = new MemoryStream();
        await OfdPackageSignatureCleaner.CleanAsync(Zip(source.PreservedEntries), output);
        output.Position = 0;
        Assert.True((await new OfdPackageLoader().LoadAsync(output)).Contains("Doc_0/Signs/Sign_0/Seal.esl"));
    }

    [Theory]
    [InlineData("", 32, 43, 50, 10)]
    [InlineData(" ID='777' CTM='1 0 0 1 0 0'", 32, 43, 50, 10)]
    [InlineData(" ID='777' CTM='2 0 0 2 0 0'", 34, 46, 100, 20)]
    [InlineData(" ID='777' CTM='0 1 -1 0 0 0'", 17, 42, 10, 50)]
    public async Task AnnotationReader_TranslatesAppearanceAndMixSavesItExactlyOnce(string attributes, double x, double y, double width, double height)
    {
        var source = await RoundTrip(Source());
        var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Add(new XElement(XName.Get("Annotations", ns), "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{source.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance{attributes} Boundary='30 40 80 20'><TextObject ID='901' Font='{source.Fonts[0].Id}' Size='4' Boundary='2 3 50 10'><TextCode X='0' Y='4'>ANNOTATION</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        var note = Assert.IsType<OfdTextElement>(Assert.Single(read.Pages[0].AnnotationAppearances));
        Assert.Equal(x, note.XMillimeters); Assert.Equal(y, note.YMillimeters);
        Assert.Equal(width, note.WidthMillimeters); Assert.Equal(height, note.HeightMillimeters);
        var mixed = await RoundTrip(OfdDocumentMixer.Mix([new(read, 0)]));
        var text = mixed.Pages[0].Elements.OfType<OfdTextElement>().Single(element => element.Text == "ANNOTATION");
        Assert.Equal(x, text.XMillimeters); Assert.Equal(y, text.YMillimeters);
        Assert.Equal(width, text.WidthMillimeters); Assert.Equal(height, text.HeightMillimeters);
        Assert.Empty(mixed.Pages[0].AnnotationAppearances);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Split_PrunesExclusiveTemplateAndPageResourcesButKeepsLiveSharedTemplate(bool keepTemplate)
    {
        var source = await RoundTrip(Source());
        var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Elements().Single(node => node.Name.LocalName == "CommonData").Add(
            new XElement(XName.Get("TemplatePage", ns), new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        Put(source, "Doc_0/Document.xml", document);
        var page = Xml(source, source.Pages[keepTemplate ? 1 : 0].SourceEntryPath!);
        page.Root!.Add(new XElement(XName.Get("Template", ns), new XAttribute("TemplateID", "700"), new XAttribute("ZOrder", "Background")));
        Put(source, source.Pages[keepTemplate ? 1 : 0].SourceEntryPath!, page);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='701'><ImageObject ID='702' ResourceID='703' Boundary='1 1 5 5'/></Layer></Content></Page>");
        source.PreservedEntries["Doc_0/Templates/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='703' Format='PNG'><MediaFile>Private.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Templates/Private.png"] = Png.Concat(new byte[] { 77 }).ToArray();
        // An unselected page has its own implicit PageRes with a private image.
        source.PreservedEntries["Doc_0/Pages/Page_0/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='710' Format='PNG'><MediaFile>Private.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Pages/Page_0/Private.png"] = [55, 55];
        var loaded = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        var result = await RoundTrip(OfdDocumentSplitter.Split(loaded, [1]));
        Assert.Equal(keepTemplate, result.PreservedEntries.ContainsKey("Doc_0/Templates/Content.xml"));
        Assert.Equal(keepTemplate, result.PreservedEntries.ContainsKey("Doc_0/Templates/PageRes.xml"));
        Assert.Equal(keepTemplate, result.PreservedEntries.ContainsKey("Doc_0/Templates/Private.png"));
        Assert.False(result.PreservedEntries.ContainsKey("Doc_0/Pages/Page_0/PageRes.xml"));
        Assert.False(result.PreservedEntries.ContainsKey("Doc_0/Pages/Page_0/Private.png"));
    }

    [Fact]
    public async Task Split_RemovingOneTemplateKeepsImplicitResourcesSharedWithAnotherTemplate()
    {
        var source = await RoundTrip(Source()); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        for (var i = 0; i < 2; i++)
        {
            document.Root!.Elements().Single(node => node.Name.LocalName == "CommonData").Add(
                new XElement(XName.Get("TemplatePage", ns), new XAttribute("ID", (700 + i).ToString()), new XAttribute("BaseLoc", $"Templates/T{i}.xml")));
            var page = Xml(source, source.Pages[i].SourceEntryPath!);
            page.Root!.Add(new XElement(XName.Get("Template", ns), new XAttribute("TemplateID", (700 + i).ToString())));
            Put(source, source.Pages[i].SourceEntryPath!, page);
            source.PreservedEntries[$"Doc_0/Templates/T{i}.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='{720 + i}'><ImageObject ID='{730 + i}' ResourceID='710' Boundary='1 1 5 5'/></Layer></Content></Page>");
        }
        Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Templates/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='710' Format='PNG'><MediaFile>Shared.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Templates/Shared.png"] = Png;
        var loaded = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        var result = await RoundTrip(OfdDocumentSplitter.Split(loaded, [1]));
        Assert.False(result.PreservedEntries.ContainsKey("Doc_0/Templates/T0.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/T1.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/PageRes.xml"));
        Assert.Equal(Png, Assert.IsType<OfdImageElement>(Assert.Single(Assert.Single(result.Pages[0].Templates).Elements)).Data);
    }

    [Fact]
    public async Task Split_NewTypedModelCopiesOnlyReachableFontsAndFlattensTemplates()
    {
        var source = Source();
        source.Pages[1].Templates.Add(new OfdTemplateContent { Elements = { new OfdTextElement { Text = "NEW TEMPLATE", FontResourceId = "10", FontName = "Arial" } } });
        var result = await RoundTrip(OfdDocumentSplitter.Split(source, [1]));
        Assert.DoesNotContain(result.PreservedEntries.Values, data => data.SequenceEqual(new byte[] { 8, 8, 8 }));
        Assert.Contains(result.Pages[0].Elements.OfType<OfdTextElement>(), text => text.Text == "NEW TEMPLATE");
        Assert.Contains(result.Pages[0].Elements.OfType<OfdTextElement>(), text => text.Text == "SECOND");
        Assert.Equal(2, source.Pages.Count);
    }

    [Fact]
    public async Task Split_UnknownLeafTemplateReferenceKeepsItsFullClosure()
    {
        var source = await RoundTrip(Source()); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Elements().Single(node => node.Name.LocalName == "CommonData").Add(new XElement(XName.Get("TemplatePage", ns), new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='701'><ImageObject ID='702' ResourceID='703' Boundary='0 0 5 5'/></Layer></Content></Page>");
        source.PreservedEntries["Doc_0/Templates/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='703' Format='PNG'><MediaFile>Image.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Templates/Image.png"] = Png;
        source.PreservedEntries["Doc_0/Extensions/unknown.xml"] = Encoding.UTF8.GetBytes("<Extension><TemplateRef>700</TemplateRef></Extension>");
        var loaded = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        var result = await RoundTrip(OfdDocumentSplitter.Split(loaded, [1]));
        Assert.Contains(result.PreservedCommonDataElements, xml => xml.Contains("TemplatePage"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/Content.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/PageRes.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/Image.png"));
    }

    internal static byte[] Png => Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==");
    private static OfdDocumentPackage Source()
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Arial" });
        package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "Private", Data = [8, 8, 8] });
        for (var i = 0; i < 2; i++)
        {
            var page = new OfdPage { Index = i, WidthMillimeters = i == 0 ? 210 : 100, HeightMillimeters = 150 };
            page.Elements.Add(new OfdTextElement { Text = i == 0 ? "FIRST" : "SECOND", LayerId = "same", FontName = i == 0 ? "Private" : "Arial", FontResourceId = i == 0 ? "11" : "10", XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 80, HeightMillimeters = 10 });
            page.Elements.Add(new OfdImageElement { Data = Png, XMillimeters = 10, YMillimeters = 50, WidthMillimeters = 10, HeightMillimeters = 10 });
            package.Pages.Add(page);
        }
        package.Attachments.Add(new OfdAttachment { Name = "public-test.txt", Data = [1, 2, 3] });
        return package;
    }
    private static void AddSignatures(OfdDocumentPackage source, string value)
    {
        var ns = source.Options.Namespace;
        var root = Xml(source, "OFD.xml"); root.Root!.Elements().Single().Add(new XElement(XName.Get("Signatures", ns), "Doc_0/Signs/Signatures.xml")); Put(source, "OFD.xml", root);
        source.PreservedEntries["Doc_0/Signs/Signatures.xml"] = Encoding.UTF8.GetBytes($"<Signatures xmlns='{ns}'><Signature ID='500' BaseLoc='Sign_0/Signature.xml'/></Signatures>");
        source.PreservedEntries["Doc_0/Signs/Sign_0/Signature.xml"] = Encoding.UTF8.GetBytes($"<Signature xmlns='{ns}'><SignedInfo><Seal BaseLoc='Seal.esl'/></SignedInfo><SignedValue>{value}</SignedValue></Signature>");
        source.PreservedEntries["Doc_0/Signs/Sign_0/SignedValue.dat"] = [11, 12];
        source.PreservedEntries["Doc_0/Signs/Sign_0/Seal.esl"] = Png;
    }
    private static XDocument Xml(OfdDocumentPackage package, string path) { using var input = new MemoryStream(package.PreservedEntries[path]); return XDocument.Load(input); }
    private static void Put(OfdDocumentPackage package, string path, XDocument xml) => package.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString());
    private static MemoryStream Zip(IDictionary<string, byte[]> entries)
    {
        var result = new MemoryStream();
        using (var zip = new ZipArchive(result, ZipArchiveMode.Create, true)) foreach (var entry in entries) { using var output = zip.CreateEntry(entry.Key).Open(); output.Write(entry.Value); }
        result.Position = 0; return result;
    }
    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage source)
    {
        using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, stream); stream.Position = 0; return await new OfdReader().ReadAsync(stream);
    }
}
