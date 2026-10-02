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
    [InlineData(false)]
    [InlineData(true)]
    public async Task CleanSignatures_HandlesOwnedXmlValueAndSealWithSharedReferenceClosure(bool shared)
    {
        var source = await RoundTrip(Source()); AddSignatures(source, "value.xml");
        var signature = Xml(source, "Doc_0/Signs/Sign_0/Signature.xml"); var ns = signature.Root!.Name.Namespace;
        signature.Root.Element(ns + "SignedInfo")!.Element(ns + "Seal")!.SetAttributeValue("BaseLoc", "seal.xml");
        Put(source, "Doc_0/Signs/Sign_0/Signature.xml", signature);
        source.PreservedEntries.Remove("Doc_0/Signs/Sign_0/SignedValue.dat"); source.PreservedEntries.Remove("Doc_0/Signs/Sign_0/Seal.esl");
        source.PreservedEntries["Doc_0/Signs/Sign_0/value.xml"] = Encoding.UTF8.GetBytes("<Value Seal='seal.xml'/>");
        source.PreservedEntries["Doc_0/Signs/Sign_0/seal.xml"] = Encoding.UTF8.GetBytes("<Seal/>");
        if (shared) source.PreservedEntries["Doc_0/Extensions/shared.xml"] = Encoding.UTF8.GetBytes("<Extension File='/Doc_0/Signs/Sign_0/value.xml'/>");
        using var input = Zip(source.PreservedEntries); using var output = new MemoryStream();
        await OfdPackageSignatureCleaner.CleanAsync(input, output); output.Position = 0;
        var result = await new OfdPackageLoader().LoadAsync(output);
        foreach (var path in new[] { "Doc_0/Signs/Sign_0/value.xml", "Doc_0/Signs/Sign_0/seal.xml" })
        {
            Assert.Equal(shared, result.Contains(path));
            if (shared) Assert.Equal(source.PreservedEntries[path], result.GetBytes(path));
        }
        Assert.False(result.Contains("Doc_0/Signs/Sign_0/Signature.xml"));
        Assert.DoesNotContain("Signatures", result.ReadUtf8Text("OFD.xml"));
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
        Assert.Equal(1, new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read).Split("ANNOTATION").Length - 1);
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
        source.PreservedEntries["Doc_0/Extensions/plain.dat"] = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("plain non-XML text")).ToArray();
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

    [Theory]
    [InlineData("unknown.xml", "<Extension><TemplateRef>700</TemplateRef></Extension>", false)]
    [InlineData("state.dat", "<Extension TemplateID='700'/>", false)]
    [InlineData("state.bin", "<Extension><TemplateRef>700</TemplateRef></Extension>", false)]
    [InlineData("state.dat", "<Extension TemplateID='700'/>", true)]
    [InlineData("state.dat", "<Extension TemplateID='700'", false)]
    [InlineData("state.dat", "<Extension File='/Doc_0/Templates/Content.xml'/>", false)]
    [InlineData("state.dat", "<Extension File='../Templates/Content.xml'/>", false)]
    [InlineData("state.dat", "<Res BaseLoc='../Templates'><File>Content.xml</File></Res>", false)]
    [InlineData("state.dat", "<Extension BaseLoc='../Templates'><File>Content.xml</File></Extension>", false)]
    [InlineData("state.dat", "<Wrapper><Res BaseLoc='../Templates'><File>Content.xml</File></Res></Wrapper>", false)]
    [InlineData("state.dat", "<Wrapper BaseLoc='..'><Extension BaseLoc='Templates'><File>Content.xml</File></Extension></Wrapper>", false)]
    public async Task Split_UnknownLeafTemplateReferenceKeepsItsFullClosure(string fileName, string reference, bool utf16)
    {
        var source = await RoundTrip(Source()); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Elements().Single(node => node.Name.LocalName == "CommonData").Add(new XElement(XName.Get("TemplatePage", ns), new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='701'><ImageObject ID='702' ResourceID='703' Boundary='0 0 5 5'/></Layer></Content></Page>");
        source.PreservedEntries["Doc_0/Templates/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='703' Format='PNG'><MediaFile>Image.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Templates/Image.png"] = Png;
        var extensionPath="Doc_0/Extensions/"+fileName;
        source.PreservedEntries[extensionPath] = utf16 ? Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(reference)).ToArray() : Encoding.UTF8.GetBytes(reference);
        var loaded = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        var result = await RoundTrip(OfdDocumentSplitter.Split(loaded, [1]));
        Assert.Contains(result.PreservedCommonDataElements, xml => xml.Contains("TemplatePage"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/Content.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/PageRes.xml"));
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Templates/Image.png"));
        Assert.Equal(source.PreservedEntries[extensionPath],result.PreservedEntries[extensionPath]);
    }

    [Theory]
    [InlineData(" Placement='unknown'", "")]
    [InlineData(" xmlns:v='urn:vendor' v:Transform='1 0 0 1 10 10'", "")]
    [InlineData(" xmlns:v='urn:vendor' v:TemplateID='payload.bin'", "")]
    [InlineData(" xmlns:v='urn:vendor'", "<v:FileRef>payload.bin</v:FileRef>")]
    [InlineData("", "<?vendor payload.bin?>")]
    public async Task Mix_RejectsUnmodeledTemplateReferenceWrappersWhileRoundtripPreserves(string attributes,string children)
    {
        var source=await RoundTrip(Source());var ns=XNamespace.Get(source.Options.Namespace);
        var document=Xml(source,"Doc_0/Document.xml");document.Root!.Element(ns+"CommonData")!.Add(new XElement(ns+"TemplatePage",new XAttribute("ID","700"),new XAttribute("BaseLoc","Templates/Content.xml")));Put(source,"Doc_0/Document.xml",document);
        source.PreservedEntries["Doc_0/Templates/Content.xml"]=Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='701'><TextObject ID='702' Font='10' Size='4' Boundary='0 0 20 10'><TextCode X='0' Y='4'>TEMPLATE</TextCode></TextObject></Layer></Content></Page>");
        var path=source.Pages[0].SourceEntryPath!;var page=Xml(source,path);var reference=XElement.Parse($"<Template xmlns='{ns}' TemplateID='700' ZOrder='Background'{attributes}>{children}</Template>");page.Root!.Add(reference);Put(source,path,page);
        using var input=Zip(source.PreservedEntries);var read=await new OfdReader().ReadAsync(input);Assert.Single(read.Pages[0].Templates);
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,0)]));
        var saved=await RoundTrip(read);Assert.True(XNode.DeepEquals(reference,Xml(saved,path).Root!.Element(ns+"Template")));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("<NotOFD><DocBody/></NotOFD>")]
    [InlineData("<OFD/>")]
    [InlineData("<OFD xmlns='urn:vendor'><DocBody/></OFD>")]
    public async Task CleanSignatures_RejectsOtherZipContainersBeforeWriting(string? root)
    {
        var entries = new Dictionary<string, byte[]> { ["ordinary.txt"] = [1] };
        if (root is not null) entries["OFD.xml"] = Encoding.UTF8.GetBytes(root);
        using var input = Zip(entries); using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPackageSignatureCleaner.CleanAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CleanSignatures_DeduplicatesDescriptionsAndPreservesTransitiveOrOpaqueReferences(bool opaque)
    {
        var source = await RoundTrip(Source()); AddSignatures(source, "SignedValue.dat"); var ns = source.Options.Namespace;
        source.PreservedEntries["Doc_0/Signs/Signatures.xml"] = Encoding.UTF8.GetBytes($"<Signatures xmlns='{ns}'>" +
            string.Concat(Enumerable.Range(0, 1_000).Select(i => $"<Signature ID='{1000 + i}' BaseLoc='Sign_0/Signature.xml'/>")) + "</Signatures>");
        source.PreservedEntries["Doc_0/Extensions/retained.xml"] = Encoding.UTF8.GetBytes(opaque ? "<broken" : "<Extension><List>/Doc_0/Signs/Signatures.xml</List></Extension>");
        using var input = Zip(source.PreservedEntries); using var output = new MemoryStream();
        await OfdPackageSignatureCleaner.CleanAsync(input, output); output.Position = 0;
        var result = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain("Signatures", result.ReadUtf8Text("OFD.xml"));
        Assert.True(result.Contains("Doc_0/Signs/Signatures.xml"));
        Assert.True(result.Contains("Doc_0/Signs/Sign_0/Signature.xml"));
        Assert.True(result.Contains("Doc_0/Signs/Sign_0/SignedValue.dat"));
        Assert.True(result.Contains("Doc_0/Signs/Sign_0/Seal.esl"));
    }

    [Fact]
    public async Task CleanSignatures_DoesNotTreatVendorXmlAsTypedSignatureOwnership()
    {
        var source = await RoundTrip(Source()); AddSignatures(source, "SignedValue.dat");
        source.PreservedEntries["Doc_0/Signs/Signatures.xml"] = Encoding.UTF8.GetBytes("<Signatures xmlns='urn:vendor'><Signature BaseLoc='Sign_0/Signature.xml'/></Signatures>");
        using var input = Zip(source.PreservedEntries); using var output = new MemoryStream();
        var report = await OfdPackageSignatureCleaner.CleanAsync(input, output); output.Position = 0;
        var result = await new OfdPackageLoader().LoadAsync(output);
        Assert.DoesNotContain("Signatures", result.ReadUtf8Text("OFD.xml"));
        Assert.True(result.Contains("Doc_0/Signs/Sign_0/SignedValue.dat"));
        Assert.Contains(report.Diagnostics, message => message.Contains("Unmodeled"));
    }

    [Theory]
    [InlineData(" VendorStyle='keep'")]
    [InlineData(" CTM='NaN 0 0 1 0 0'")]
    public async Task Reader_PreservesUnmodeledAnnotationMetadataWhileMixRejectsFlattening(string attributes)
    {
        var source = await RoundTrip(Source()); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Add(new XElement(XName.Get("Annotations", ns), "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{source.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance{attributes} Boundary='1 1 10 10'><TextObject ID='901' Font='{source.Fonts[0].Id}' Size='4' Boundary='0 0 10 10'><TextCode X='0' Y='4'>UNKNOWN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        Assert.Contains(read.Pages[0].AnnotationAppearances, element => element is OfdRawElement);
        if (attributes.Contains("VendorStyle")) Assert.Contains(read.Pages[0].AnnotationAppearances.OfType<OfdTextElement>(), text => text.Text == "UNKNOWN");
        OfdWatermark.AddText(read, [0], "DRAFT");
        var saved = await RoundTrip(OfdDocumentSplitter.Split(read, [0]));
        Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(saved, 0)]));
    }

    [Fact]
    public async Task AnnotationIndex_MultiplePagesAndDuplicateRecordsKeepEachPageAppearanceOnce()
    {
        var source = new OfdDocumentPackage();
        for (var i = 0; i < 200; i++) source.Pages.Add(new OfdPage { Index = i, WidthMillimeters = 210, HeightMillimeters = 297 });
        source = await RoundTrip(source); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Add(new XElement(XName.Get("Annotations", ns), "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'>" +
            string.Concat(source.Pages.SelectMany((page, i) => new[] { 0, 1 }.Select(_ => $"<Page PageID='{page.Id}'><FileLoc>P{i}.xml</FileLoc></Page>"))) + "</Annotations>");
        for (var i = 0; i < 200; i++) source.PreservedEntries[$"Doc_0/Annots/P{i}.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='{1000 + i}'><Appearance Boundary='0 0 20 20'><TextObject ID='{2000 + i}' Boundary='0 0 20 20' Size='4'><TextCode X='0' Y='4'>NOTE{i}</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        Assert.Equal(200, read.Pages.Count);
        for (var i = 0; i < 200; i++) Assert.Equal($"NOTE{i}", Assert.IsType<OfdTextElement>(Assert.Single(read.Pages[i].AnnotationAppearances)).Text);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task TypedSplit_WithoutFontIdKeepsTheRequestedEmbeddedStyle(bool bold, bool italic)
    {
        var source = new OfdDocumentPackage();
        source.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Same", Data = [1] });
        source.Fonts.Add(new OfdFontResource { Id = "11", FontName = "Same", Bold = bold, Italic = italic, Data = [2] });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { Text = "styled", FontName = "Same", Weight = bold ? 700 : 400, Italic = italic }); source.Pages.Add(page);
        var result = await RoundTrip(OfdDocumentSplitter.Split(source, [0]));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(result.Pages[0].Elements));
        Assert.Equal(new byte[] { 2 }, result.Fonts.Single(font => font.Id == text.FontResourceId).Data);
        Assert.DoesNotContain(result.PreservedEntries.Values, data => data.SequenceEqual(new byte[] { 1 }));
    }

    [Fact]
    public async Task CleanSignatures_SharedTypedXmlWithoutXmlExtensionKeepsItsReferencedValues()
    {
        var source = await RoundTrip(Source()); AddSignatures(source, "SignedValue.dat");
        source.PreservedEntries["Doc_0/Signs/Signatures.dat"] = source.PreservedEntries["Doc_0/Signs/Signatures.xml"];
        source.PreservedEntries.Remove("Doc_0/Signs/Signatures.xml");
        source.PreservedEntries["Doc_0/Signs/Sign_0/Signature.dat"] = source.PreservedEntries["Doc_0/Signs/Sign_0/Signature.xml"];
        source.PreservedEntries.Remove("Doc_0/Signs/Sign_0/Signature.xml");
        var root = Xml(source, "OFD.xml"); root.Descendants().Single(node => node.Name.LocalName == "Signatures").Value = "Doc_0/Signs/Signatures.dat"; Put(source, "OFD.xml", root);
        source.PreservedEntries["Doc_0/Signs/Signatures.dat"] = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(source.PreservedEntries["Doc_0/Signs/Signatures.dat"]).Replace("Signature.xml", "Signature.dat"));
        source.PreservedEntries["Doc_0/Extensions/shared.xml"] = Encoding.UTF8.GetBytes("<Extension File='/Doc_0/Signs/Signatures.dat'/>");
        using var input = Zip(source.PreservedEntries); using var output = new MemoryStream(); await OfdPackageSignatureCleaner.CleanAsync(input, output); output.Position = 0;
        var result = await new OfdPackageLoader().LoadAsync(output);
        Assert.True(result.Contains("Doc_0/Signs/Signatures.dat")); Assert.True(result.Contains("Doc_0/Signs/Sign_0/Signature.dat"));
        Assert.True(result.Contains("Doc_0/Signs/Sign_0/SignedValue.dat")); Assert.True(result.Contains("Doc_0/Signs/Sign_0/Seal.esl"));
    }

    [Theory]
    [InlineData("urn:vendor")]
    [InlineData("http://www.ofdspec.org")]
    public async Task Split_DoesNotMutateUndeclaredSameNamedAnnotationOrResourceExtensions(string extensionNamespace)
    {
        var source = await RoundTrip(Source()); var removedPage = source.Pages[0];
        var privateFont = source.Pages[0].Elements.OfType<OfdTextElement>().Single().FontResourceId;
        source.PreservedEntries["Doc_0/Extensions/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{extensionNamespace}'><Page PageID='{removedPage.Id}'><FileLoc>payload.bin</FileLoc></Page></Annotations>");
        source.PreservedEntries["Doc_0/Extensions/Resources.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{extensionNamespace}'><Fonts><Font ID='{privateFont}'><FontFile>font.bin</FontFile></Font></Fonts></Res>");
        source.PreservedEntries["Doc_0/Extensions/payload.bin"] = [4, 5]; source.PreservedEntries["Doc_0/Extensions/font.bin"] = [6, 7];
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries)); var split = await RoundTrip(OfdDocumentSplitter.Split(read, [1]));
        foreach (var path in new[] { "Doc_0/Extensions/Annotations.xml", "Doc_0/Extensions/Resources.xml", "Doc_0/Extensions/payload.bin", "Doc_0/Extensions/font.bin" })
            Assert.Equal(source.PreservedEntries[path], split.PreservedEntries[path]);
    }

    [Fact]
    public async Task Split_KeepsVendorTemplateDeclarationsAndImplicitPageResourceBytes()
    {
        var source = await RoundTrip(Source()); var document = Xml(source, "Doc_0/Document.xml"); var ns = document.Root!.Name.Namespace;
        document.Root.Element(ns + "CommonData")!.Add(new XElement(XName.Get("TemplatePage", "urn:vendor"), new XAttribute("ID", "800"), new XAttribute("BaseLoc", "Extensions/template.bin")));
        Put(source, "Doc_0/Document.xml", document); source.PreservedEntries["Doc_0/Extensions/template.bin"] = [8];
        source.PreservedEntries["Doc_0/Pages/Page_0/PageRes.xml"] = Encoding.UTF8.GetBytes("<Res xmlns='urn:vendor'><Fonts><Font ID='999'><FontFile>vendor.bin</FontFile></Font></Fonts></Res>");
        source.PreservedEntries["Doc_0/Pages/Page_0/vendor.bin"] = [9];
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries)); var result = await RoundTrip(OfdDocumentSplitter.Split(read, [1]));
        Assert.Contains(result.PreservedCommonDataElements, value => value.Contains("urn:vendor"));
        Assert.Equal(source.PreservedEntries["Doc_0/Pages/Page_0/PageRes.xml"], result.PreservedEntries["Doc_0/Pages/Page_0/PageRes.xml"]);
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Pages/Page_0/vendor.bin"));
    }

    [Theory]
    [InlineData("Annotations", false)]
    [InlineData("TemplatePage", true)]
    public void Mix_RejectsSameNamedVendorDocumentDeclarations(string name, bool common)
    {
        var source = Source(); var xml = $"<{name} xmlns='urn:vendor' />";
        if (common) source.PreservedCommonDataElements.Add(xml); else source.PreservedDocumentElements.Add(xml);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
        Assert.Equal(2, source.Pages.Count);
    }

    [Theory]
    [InlineData("Page")]
    [InlineData("FileLoc")]
    [InlineData("Annot")]
    [InlineData("Appearance")]
    [InlineData("TextObject")]
    [InlineData("TextCode")]
    [InlineData("Clips")]
    [InlineData("TextWithoutCode")]
    [InlineData("AbbreviatedData")]
    public async Task Reader_DoesNotInterpretSameNamedVendorAnnotationNodesAsVisibleContent(string vendorNode)
    {
        var source = await RoundTrip(Source()); var ns = source.Options.Namespace;
        var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Add(new XElement(XName.Get("Annotations", ns), "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        string QName(string name) => vendorNode == name ? "v:" + name : name;
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}' xmlns:v='urn:vendor'><{QName("Page")} PageID='{source.Pages[0].Id}'><{QName("FileLoc")}>Page.xml</{QName("FileLoc")}></{QName("Page")}></Annotations>");
        source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}' xmlns:v='urn:vendor'><{QName("Annot")} ID='900'><{QName("Appearance")} Boundary='0 0 20 20'><{QName("TextObject")} ID='901' Boundary='0 0 20 20' Size='4'><{QName("TextCode")} X='0' Y='4'>VENDOR SECRET</{QName("TextCode")}></{QName("TextObject")}></{QName("Appearance")}></{QName("Annot")}></PageAnnot>");
        if (vendorNode == "Clips")
            source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}' xmlns:v='urn:vendor'><Annot ID='900'><Appearance Boundary='0 0 20 20'><TextObject ID='901' Boundary='0 0 20 20' Size='4'><v:Clips><v:Clip><v:Area><v:Path><v:AbbreviatedData>M 0 0 L 10 0 L 10 10 C</v:AbbreviatedData></v:Path></v:Area></v:Clip></v:Clips><TextCode X='0' Y='4'>VENDOR SECRET</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        if (vendorNode == "TextWithoutCode")
            source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}' xmlns:v='urn:vendor'><Annot ID='900'><Appearance Boundary='0 0 20 20'><TextObject ID='901' Boundary='0 0 20 20' Size='4'><v:Data>VENDOR SECRET</v:Data></TextObject></Appearance></Annot></PageAnnot>");
        if (vendorNode == "AbbreviatedData")
            source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}' xmlns:v='urn:vendor'><Annot ID='900'><Appearance Boundary='0 0 20 20'><PathObject ID='901' Boundary='0 0 20 20'><v:AbbreviatedData>M 0 0 L 10 0</v:AbbreviatedData></PathObject></Appearance></Annot></PageAnnot>");
        var read = await new OfdReader().ReadAsync(Zip(source.PreservedEntries));
        Assert.DoesNotContain("VENDOR SECRET", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read));
        Assert.DoesNotContain(read.Pages[0].AnnotationAppearances, element => element is OfdTextElement);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        var saved = await RoundTrip(read); Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
    }

    [Fact]
    public void ImageWatermark_MutablePayloadsAreIsolatedBetweenPagesAndBudgeted()
    {
        var source = Source(); var data = Png;
        OfdWatermark.AddImage(source, [0,1], data, "image/png");
        var first = source.Pages[0].Elements.OfType<OfdImageElement>().Last(); var second = source.Pages[1].Elements.OfType<OfdImageElement>().Last();
        Assert.NotSame(first.Data, second.Data); Assert.NotSame(data, first.Data);
        first.Data[0] = 0; Assert.Equal(data[0], second.Data[0]);
        Assert.Throws<ArgumentException>(() => OfdWatermark.AddImage(source, [0,1], data, "image/png", new OfdWatermarkOptions { MaxGeneratedImageBytes = data.Length }));
    }

    [Fact]
    public void TextWatermark_IgnoresImageOnlyPolicyWhileImagesRemainForbiddenAtomically()
    {
        var source = Source(); var options = new OfdWatermarkOptions { MaxImageBytes = 0, MaxGeneratedImageBytes = 0, MaxDecodedImagePixels = 0 };
        OfdWatermark.AddText(source, [0,1], "TEXT", options);
        Assert.All(source.Pages, page => Assert.Equal("TEXT", Assert.IsType<OfdTextElement>(page.Elements.Last()).Text));
        var counts = source.Pages.Select(page => page.Elements.Count).ToArray();
        Assert.Throws<ArgumentException>(() => OfdWatermark.AddImage(source, [0,1], Png, "image/png", options));
        Assert.Equal(counts, source.Pages.Select(page => page.Elements.Count));
    }

    [Fact]
    public async Task Writer_ClipsPrecedeObjectSpecificTextAndPathChildren()
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        var clip = "<Clips><Clip><Area><Path ID='clip'><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>";
        page.Elements.Add(new OfdTextElement { Text = "CLIP", ClippingXml = clip, SourceXml = "<TextObject><TextCode X='0' Y='4'>CLIP</TextCode></TextObject>" });
        page.Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 20 0", ClippingXml = clip }); package.Pages.Add(page);
        var result = await RoundTrip(package); var content = Xml(result, result.Pages[0].SourceEntryPath!);
        foreach (var node in content.Descendants().Where(node => node.Name.LocalName is "TextObject" or "PathObject"))
        {
            var children = node.Elements().ToList(); var clips = children.FindIndex(child => child.Name.LocalName == "Clips");
            var specific = children.FindIndex(child => child.Name.LocalName is "TextCode" or "AbbreviatedData");
            Assert.True(clips >= 0 && clips < specific);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task OrdinaryClips_PreserveVendorSiblingAndSelectOnlyStandardNamespace(bool vendorFirst)
    {
        var package = await RoundTrip(Source()); var ns = XNamespace.Get(package.Options.Namespace);
        var path = package.Pages[0].SourceEntryPath!; var content = Xml(package, path);
        foreach (var node in content.Descendants().Where(node => node.Name == ns + "TextObject" || node.Name == ns + "ImageObject").ToList())
        {
            var standard = XElement.Parse($"<Clips xmlns='{ns}'><Clip><Area><Path ID='800'><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>");
            var vendor = new XElement(XNamespace.Get("urn:vendor") + "Clips", new XAttribute("Keep", "yes"));
            node.AddFirst(vendorFirst ? new[] { vendor, standard } : new[] { standard, vendor });
        }
        Put(package, path, content); using (var input = Zip(package.PreservedEntries)) package = await new OfdReader().ReadAsync(input);
        Assert.All(package.Pages[0].Elements.Where(element => element is OfdTextElement or OfdImageElement),
            element => Assert.Equal(ns + "Clips", XElement.Parse(element.ClippingXml!).Name));
        var result = await RoundTrip(package); content = Xml(result, result.Pages[0].SourceEntryPath!);
        foreach (var node in content.Descendants().Where(node => node.Name == ns + "TextObject" || node.Name == ns + "ImageObject"))
        {
            Assert.Single(node.Elements(ns + "Clips"));
            Assert.Equal("yes", Assert.Single(node.Elements(XNamespace.Get("urn:vendor") + "Clips")).Attribute("Keep")!.Value);
        }
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(result, 0)]));
    }

    internal static byte[] Png => Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==");

    [Theory]
    [InlineData(false, "attribute")]
    [InlineData(false, "child")]
    [InlineData(true, "attribute")]
    [InlineData(true, "child")]
    [InlineData(true, "text")]
    public async Task Mix_ExtendedDeclarationWrappersRefuseWhileKnownDrawingAndXmlSurvive(bool template, string extra)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XNamespace.Get("urn:vendor");
        var document = Xml(source, "Doc_0/Document.xml"); XElement declaration;
        if (template)
        {
            declaration = new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml"));
            document.Root!.Element(ns + "CommonData")!.Add(declaration);
            var path = source.Pages[0].SourceEntryPath!; var page = Xml(source, path); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700"))); Put(source, path, page);
            source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='702'><TextObject ID='703' Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Layer></Content></Page>");
        }
        else
        {
            declaration = new XElement(ns + "Annotations", "Annots/Annotations.xml"); document.Root!.Add(declaration);
            source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{source.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
            source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot><Appearance Boundary='0 0 20 20'><TextObject Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        }
        if (extra == "attribute") declaration.SetAttributeValue(vendor + "Payload", "Extensions/data.bin");
        if (extra == "child") declaration.Add(new XElement(vendor + "Extra", new XAttribute("File", "Extensions/data.bin")));
        if (extra == "text") declaration.Add("KEEP");
        Put(source, "Doc_0/Document.xml", document); using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        var drawings = template ? Assert.Single(read.Pages[0].Templates).Elements : read.Pages[0].AnnotationAppearances;
        Assert.Equal("KNOWN", Assert.Single(drawings.OfType<OfdTextElement>()).Text);
        var saved = await RoundTrip(read); var wrapper = Assert.Single(Xml(saved, "Doc_0/Document.xml").Descendants(declaration.Name));
        if (extra == "attribute") Assert.Equal("Extensions/data.bin", wrapper.Attribute(vendor + "Payload")!.Value);
        if (extra == "child") Assert.Equal("Extensions/data.bin", wrapper.Element(vendor + "Extra")!.Attribute("File")!.Value);
        if (extra == "text") Assert.Equal("KEEP", wrapper.Value);
    }

    [Theory]
    [InlineData("attribute")]
    [InlineData("child")]
    [InlineData("text")]
    [InlineData("file-child")]
    [InlineData("file-attribute")]
    [InlineData("duplicate")]
    public async Task AnnotationIndexExtendedPageRecordsKeepScopedMarkersAndUnambiguousArtwork(string extra)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XNamespace.Get("urn:vendor");
        var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        var file = new XElement(ns + "FileLoc", "Page.xml"); var record = new XElement(ns + "Page", new XAttribute("PageID", source.Pages[0].Id!), file);
        if (extra == "attribute") record.SetAttributeValue(vendor + "Payload", "../Extensions/data.bin");
        if (extra == "child") record.Add(new XElement(vendor + "Extra", "../Extensions/data.bin"));
        if (extra == "text") record.Add("KEEP");
        if (extra == "file-attribute") file.SetAttributeValue(vendor + "Payload", "../Extensions/data.bin");
        if (extra == "file-child") file.Add(new XElement(vendor + "Extra", "SECRET"));
        if (extra == "duplicate") record.Add(new XElement(file));
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes(new XDocument(new XElement(ns + "Annotations", record)).ToString());
        source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot><Appearance Boundary='0 0 20 20'><TextObject Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Contains(read.Pages[0].AnnotationAppearances, element => element is OfdRawElement); Assert.Empty(read.Pages[1].AnnotationAppearances);
        Assert.Equal(extra is "file-child" or "duplicate" ? 0 : 1, read.Pages[0].AnnotationAppearances.OfType<OfdTextElement>().Count());
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)])); Assert.Single(OfdDocumentMixer.Mix([new(read, 1)]).Pages);
        var saved = await RoundTrip(read); Assert.Equal(source.PreservedEntries["Doc_0/Annots/Annotations.xml"], saved.PreservedEntries["Doc_0/Annots/Annotations.xml"]);
    }

    [Theory]
    [InlineData(null, "Foreground")]
    [InlineData("Background", "Background")]
    [InlineData("Foreground", "Foreground")]
    public async Task TemplateDeclarationStandardNameAndDefaultZOrderSupportFlattening(string? pageOrder, string expectedOrder)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml"), new XAttribute("Name", "Label"), new XAttribute("ZOrder", "Foreground"))); Put(source, "Doc_0/Document.xml", document);
        var path = source.Pages[0].SourceEntryPath!; var page = Xml(source, path); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700"), pageOrder is null ? null : new XAttribute("ZOrder", pageOrder))); Put(source, path, page);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='702'><TextObject ID='703' Size='3'><TextCode X='0' Y='3'>TEMPLATE</TextCode></TextObject></Layer></Content></Page>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input); Assert.Equal(expectedOrder, Assert.Single(read.Pages[0].Templates).ZOrder);
        var mixed = OfdDocumentMixer.Mix([new(read, 0)]); var labels = mixed.Pages[0].Elements.OfType<OfdTextElement>().Select(text => text.Text).ToList();
        Assert.Equal(expectedOrder == "Background" ? new[] { "TEMPLATE", "FIRST" } : new[] { "FIRST", "TEMPLATE" }, labels);
    }

    [Theory]
    [InlineData("Annotations", false)]
    [InlineData("PageAnnot", false)]
    [InlineData("Annot", false)]
    [InlineData("Annot", true)]
    public async Task AnnotationRootAndAnnotExtensionsPreserveKnownArtAndRefuseMix(string wrapper, bool hidden)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XNamespace.Get("urn:vendor");
        var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        var index = XDocument.Parse($"<Annotations xmlns='{ns}'><Page PageID='{source.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        var annots = XDocument.Parse($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance Boundary='0 0 20 20'><TextObject Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        var node = wrapper == "Annotations" ? index.Root! : wrapper == "PageAnnot" ? annots.Root! : annots.Root!.Element(ns + "Annot")!;
        node.SetAttributeValue(vendor + "Payload", "../Extensions/data.bin"); node.Add(new XElement(vendor + "Extra", "keep"));
        if (hidden) node.SetAttributeValue("Visible", "false");
        Put(source, "Doc_0/Annots/Annotations.xml", index); Put(source, "Doc_0/Annots/Page.xml", annots);
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Contains(read.Pages[0].AnnotationAppearances, element => element is OfdRawElement);
        Assert.Equal(hidden ? 0 : 1, read.Pages[0].AnnotationAppearances.OfType<OfdTextElement>().Count());
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        if (wrapper != "Annotations") Assert.Single(OfdDocumentMixer.Mix([new(read, 1)]).Pages);
        var saved = await RoundTrip(read); Assert.Equal(source.PreservedEntries["Doc_0/Annots/Page.xml"], saved.PreservedEntries["Doc_0/Annots/Page.xml"]);
        Assert.Equal(source.PreservedEntries["Doc_0/Annots/Annotations.xml"], saved.PreservedEntries["Doc_0/Annots/Annotations.xml"]);
    }

    [Theory]
    [InlineData("", "Foreground", "Foreground")]
    [InlineData(" ", "Foreground", "Foreground")]
    [InlineData("Watermark", "Foreground", "Foreground")]
    [InlineData(null, "", "Background")]
    [InlineData(null, "Watermark", "Background")]
    [InlineData(null, "Body", "Background")]
    public async Task UnmodeledTemplateOrderingFallsBackForReadingButRefusesMix(string? referenceOrder, string declarationOrder, string expected)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var document = Xml(source, "Doc_0/Document.xml");
        document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml"), new XAttribute("ZOrder", declarationOrder))); Put(source, "Doc_0/Document.xml", document);
        var path = source.Pages[0].SourceEntryPath!; var page = Xml(source, path); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700"), referenceOrder is null ? null : new XAttribute("ZOrder", referenceOrder))); Put(source, path, page);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='702'><TextObject Size='3'><TextCode X='0' Y='3'>TEMPLATE</TextCode></TextObject></Layer></Content></Page>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Equal(expected, Assert.Single(read.Pages[0].Templates).ZOrder); Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
    }

    [Theory]
    [InlineData(false, "Page", "attribute")]
    [InlineData(false, "Area", "attribute")]
    [InlineData(false, "PhysicalBox", "attribute")]
    [InlineData(false, "Content", "attribute")]
    [InlineData(false, "Layer", "attribute")]
    [InlineData(false, "Content", "child")]
    [InlineData(false, "Layer", "DrawParam")]
    [InlineData(true, "Page", "attribute")]
    [InlineData(true, "Area", "attribute")]
    [InlineData(true, "Content", "attribute")]
    [InlineData(true, "Layer", "attribute")]
    [InlineData(true, "Content", "child")]
    [InlineData(true, "Page", "nested-template")]
    public async Task MixPreflightsSelectedPageAndTemplateContainers(bool template, string container, string extra)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XNamespace.Get("urn:vendor");
        var path = source.Pages[0].SourceEntryPath!;
        if (template)
        {
            var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml"))); Put(source, "Doc_0/Document.xml", document);
            var page = Xml(source, path); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700"))); Put(source, path, page);
            path = "Doc_0/Templates/Content.xml";
            source.PreservedEntries[path] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Area><PhysicalBox>0 0 210 297</PhysicalBox></Area><Content><Layer ID='702'><TextObject Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Layer></Content></Page>");
        }
        var xml = Xml(source, path); var node = xml.Descendants(ns + container).First();
        if (extra == "attribute") node.SetAttributeValue(vendor + "Payload", "Extensions/data.bin");
        if (extra == "child") node.Add(new XElement(vendor + "Extra", "keep"));
        if (extra == "DrawParam") node.SetAttributeValue("DrawParam", "999");
        if (extra == "nested-template") node.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700")));
        Put(source, path, xml); using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        var drawings = template ? Assert.Single(read.Pages[0].Templates).Elements : read.Pages[0].Elements;
        Assert.NotEmpty(drawings.OfType<OfdTextElement>()); Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        Assert.Single(OfdDocumentMixer.Mix([new(read, 1)]).Pages);
        Assert.Equal(source.PreservedEntries[path], read.PreservedEntries[path]);
    }

    [Theory]
    [InlineData("../../../external")]
    [InlineData("")]
    [InlineData("/")]
    public async Task VendorAnnotations_TextIsPreservedWithoutInterpretingItAsAPackagePath(string value)
    {
        var source = await RoundTrip(Source()); var document = Xml(source, "Doc_0/Document.xml");
        var vendorName = XName.Get("Annotations", "urn:vendor"); document.Root!.Add(new XElement(vendorName, value));
        Put(source, "Doc_0/Document.xml", document);
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.All(read.Pages, page => Assert.IsType<OfdRawElement>(Assert.Single(page.AnnotationAppearances)));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        var saved = await RoundTrip(read); Assert.Equal(value, Assert.Single(Xml(saved, "Doc_0/Document.xml").Root!.Elements(vendorName)).Value);
    }

    [Fact]
    public async Task VendorAnnotationDeclaration_DoesNotSuppressLaterStandardDeclarationWithTheSameText()
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XName.Get("Annotations", "urn:vendor");
        var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Add(new XElement(vendor, "Annots/Annotations.xml"), new XElement(ns + "Annotations", "Annots/Annotations.xml")); Put(source, "Doc_0/Document.xml", document);
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{source.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        source.PreservedEntries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot><Appearance Boundary='0 0 20 20'><TextObject Size='3'><TextCode X='0' Y='3'>KNOWN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Equal("KNOWN", Assert.Single(read.Pages[0].AnnotationAppearances.OfType<OfdTextElement>()).Text);
        Assert.All(read.Pages, page => Assert.Contains(page.AnnotationAppearances, element => element is OfdRawElement));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        var saved = await RoundTrip(read); Assert.Equal("KNOWN", Assert.Single(saved.Pages[0].AnnotationAppearances.OfType<OfdTextElement>()).Text);
        Assert.Equal("Annots/Annotations.xml", Assert.Single(Xml(saved, "Doc_0/Document.xml").Root!.Elements(vendor)).Value);
    }

    [Fact]
    public async Task StandardAnnotationDeclaration_StillRejectsEscapingPackagePaths()
    {
        var source = await RoundTrip(Source()); var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Add(new XElement(document.Root.Name.Namespace + "Annotations", "../../../external")); Put(source, "Doc_0/Document.xml", document);
        using var input = Zip(source.PreservedEntries); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(input));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public async Task VendorTemplates_DoNotEnterTheStandardDeclarationOrReferenceTables(bool standardReference, bool vendorFirst)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var vendor = XName.Get("TemplatePage", "urn:vendor");
        var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Element(ns + "CommonData")!.Add(
            new XElement(vendor, new XAttribute("ID", "700"), new XAttribute("BaseLoc", "../../../external")),
            new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        Put(source, "Doc_0/Document.xml", document); var path = source.Pages[0].SourceEntryPath!; var page = Xml(source, path);
        var vendorReference = new XElement(XName.Get("Template", "urn:vendor"), new XAttribute("TemplateID", "700"), new XAttribute("ZOrder", "Foreground"), new XAttribute("Keep", "yes"));
        var standard = new XElement(ns + "Template", new XAttribute("TemplateID", "700"));
        if (standardReference) page.Root!.Add(vendorFirst ? new[] { vendorReference, standard } : new[] { standard, vendorReference });
        else page.Root!.Add(vendorReference);
        Put(source, path, page);
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='702'><TextObject ID='703' Size='3'><TextCode X='0' Y='3'>TEMPLATE</TextCode></TextObject></Layer></Content></Page>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Equal(standardReference ? 1 : 0, read.Pages[0].Templates.Count);
        if (standardReference) Assert.Equal("TEMPLATE", Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(read.Pages[0].Templates).Elements)).Text);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        var saved = await RoundTrip(read); Assert.Equal("../../../external", Assert.Single(Xml(saved, "Doc_0/Document.xml").Descendants(vendor)).Attribute("BaseLoc")!.Value);
        Assert.Equal(standardReference ? 1 : 0, saved.Pages[0].Templates.Count);
        if (standardReference) Assert.Equal("TEMPLATE", Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(saved.Pages[0].Templates).Elements)).Text);
        var preserved = Assert.Single(Xml(saved, path).Root!.Elements(vendorReference.Name));
        Assert.Equal("700", preserved.Attribute("TemplateID")!.Value); Assert.Equal("Foreground", preserved.Attribute("ZOrder")!.Value); Assert.Equal("yes", preserved.Attribute("Keep")!.Value);
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    public async Task Mix_CustomTagsKeepUniqueKeysAndCoalesceIdenticalValues(bool savedInput, int taggedSource)
    {
        var sources = new[] { Source(), Source() }; sources[taggedSource].CustomTags["owner"] = "测试 keep";
        foreach (var source in sources) { source.CustomTags["source-text-origin"] = "DOCX/OpenXML"; source.CustomTags["source-text-kind"] = "machine-readable"; source.CustomTags["docx-ofd-mode"] = "Native"; }
        if (savedInput) sources[taggedSource] = await RoundTrip(sources[taggedSource]);
        var page = sources[taggedSource].Pages[0]; var count = page.Elements.Count;
        var mixed = await RoundTrip(OfdDocumentMixer.Mix(sources.Select(source => new OfdMixSource(source, 0))));
        Assert.Equal(4, mixed.CustomTags.Count); Assert.Equal("测试 keep", mixed.CustomTags["owner"]); Assert.Equal("Native", mixed.CustomTags["docx-ofd-mode"]);
        Assert.Equal("machine-readable", mixed.CustomTags["source-text-kind"]); Assert.Equal("DOCX/OpenXML", mixed.CustomTags["source-text-origin"]);
        Assert.Equal("测试 keep", sources[taggedSource].CustomTags["owner"]);
        Assert.Same(page, sources[taggedSource].Pages[0]); Assert.Equal(count, page.Elements.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Mix_CustomTagsConflictsFailWithoutChangingEitherSource(bool savedInput)
    {
        var first = Source(); var second = Source(); first.CustomTags["docx-ofd-mode"] = "Native"; second.CustomTags["docx-ofd-mode"] = "DualLayer";
        if (savedInput) { first = await RoundTrip(first); second = await RoundTrip(second); }
        var error = Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(first, 0), new(second, 0)]));
        Assert.Contains("docx-ofd-mode", error.Message); Assert.Contains("source 2", error.Message);
        Assert.Equal("Native", first.CustomTags["docx-ofd-mode"]); Assert.Equal("DualLayer", second.CustomTags["docx-ofd-mode"]);
    }

    [Theory]
    [InlineData("attribute")]
    [InlineData("child")]
    [InlineData("vendor")]
    [InlineData("conflict")]
    [InlineData("schema")]
    public async Task Mix_CustomTagsUnmodeledXmlAndConflictingSourceRecordsRemainRejected(string kind)
    {
        var source = Source(); source.CustomTags["fixture"] = "public"; source = await RoundTrip(source);
        var path = "Doc_0/Tags/CustomTag_EMR.xml"; var detail = Xml(source, path); var tag = detail.Root!.Elements().Single();
        if (kind == "attribute") tag.SetAttributeValue("Unknown", "keep");
        if (kind == "child") tag.Add(new XElement(detail.Root.Name.Namespace + "Note", "keep"));
        if (kind == "vendor") tag.Name = XName.Get("Tag", "urn:vendor");
        if (kind == "conflict") detail.Root.Add(new XElement(tag.Name, new XAttribute("Key", "fixture"), new XAttribute("Value", "other")));
        if (kind == "schema") detail.Root.Name = detail.Root.Name.Namespace + "OtherSchema";
        Put(source, path, detail); using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
        Assert.Equal(source.PreservedEntries[path], read.PreservedEntries[path]);
    }

    [Theory]
    [InlineData("path")]
    [InlineData("id")]
    public async Task Mix_CustomTagsPotentialPackageReferencesRefuseUntilRemappingIsDefined(string kind)
    {
        var source = await RoundTrip(Source());
        source.CustomTags["source"] = kind == "path" ? "/" + source.PreservedEntries.Keys.Single(path => path.Contains("Attach_1_")) : source.Attachments[0].Id!;
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(source, 0)]));
        Assert.Single(source.Attachments); Assert.Single(source.CustomTags);
    }

    [Theory]
    [InlineData("PublicRes", ".xml")]
    [InlineData("PublicRes", ".dat")]
    [InlineData("DocumentRes", ".xml")]
    [InlineData("DocumentRes", ".bin")]
    public async Task ResourceIdsAboveTheDocumentHint_AreReservedBeforeAllocatingWatermarkObjects(string declaration, string suffix)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source, "Doc_0/Document.xml"); var reference = document.Root!.Element(ns + "CommonData")!.Element(ns + declaration)!;
        var oldPath = "Doc_0/" + reference.Value; var resources = Xml(source, oldPath); var kind = declaration == "PublicRes" ? "Font" : "MultiMedia";
        var resource = resources.Descendants(ns + kind).First(); var oldId = resource.Attribute("ID")!.Value;
        var max = source.PreservedEntries.Where(pair => pair.Key.EndsWith(".xml") && pair.Key != oldPath)
            .SelectMany(pair => Xml(source, pair.Key).Descendants().Attributes("ID"))
            .Concat(resources.Descendants().Attributes("ID").Where(attribute => attribute != resource.Attribute("ID")))
            .Select(attribute => long.TryParse(attribute.Value, out var id) ? id : 0).Max();
        var reserved = (max + 1).ToString(System.Globalization.CultureInfo.InvariantCulture); resource.SetAttributeValue("ID", reserved);
        foreach (var path in source.Pages.Select(page => page.SourceEntryPath!))
        {
            var page = Xml(source, path);
            foreach (var attribute in page.Descendants().Attributes(declaration == "PublicRes" ? "Font" : "ResourceID").Where(attribute => attribute.Value == oldId)) attribute.Value = reserved;
            Put(source, path, page);
        }
        var location = "Resources" + suffix; reference.Value = location; document.Root.Element(ns + "CommonData")!.Element(ns + "MaxUnitID")!.Value = max.ToString(System.Globalization.CultureInfo.InvariantCulture);
        source.PreservedEntries.Remove(oldPath); Put(source, "Doc_0/" + location, resources); Put(source, "Doc_0/Document.xml", document);
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        OfdWatermark.AddText(read, [0], "RESERVE", fontName: "Arial"); OfdWatermark.AddImage(read, [0], Png, "image/png");
        var saved = await RoundTrip(read); var ids = new List<string>();
        foreach (var pair in saved.PreservedEntries)
        {
            XDocument xml; try { xml = Xml(saved, pair.Key); } catch (System.Xml.XmlException) { continue; }
            ids.AddRange(xml.Descendants().Where(node => node.Name.Namespace == ns).Attributes("ID").Select(attribute => attribute.Value));
        }
        Assert.Equal(ids.Count, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Single(Xml(saved, "Doc_0/" + location).Descendants(ns + kind), node => node.Attribute("ID")?.Value == reserved);
        if (declaration == "PublicRes") Assert.Equal(reserved, Assert.Single(saved.Fonts, font => font.FontName == "Arial").Id);
        else Assert.All(saved.Pages.SelectMany(page => page.Elements).OfType<OfdImageElement>(), image => Assert.Equal(reserved, image.ResourceId));
    }

    [Theory]
    [InlineData("PublicRes", "save")]
    [InlineData("DocumentRes", "save")]
    [InlineData("PublicRes", "watermark")]
    [InlineData("DocumentRes", "watermark")]
    [InlineData("PublicRes", "split")]
    [InlineData("DocumentRes", "split")]
    public async Task DeclaredResourceXml_NonXmlSuffixSurvivesSaveAndTools(string declaration, string operation)
    {
        var fixture = Source(); fixture.Fonts.Single(font => font.Id == "10").Data = [9, 9, 9];
        var source = await RoundTrip(fixture); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source, "Doc_0/Document.xml"); var reference = document.Root!.Element(ns + "CommonData")!.Element(ns + declaration)!;
        var oldPath = "Doc_0/" + reference.Value; var path = declaration == "PublicRes" ? "PublicResources.dat" : "ImageResources.bin";
        var resource = Xml(source, oldPath); resource.Root!.SetAttributeValue(XNamespace.Get("urn:vendor") + "Keep", "yes");
        var unowned = source.PreservedEntries[oldPath];
        source.PreservedEntries.Remove(oldPath); Put(source, "Doc_0/" + path, resource);
        reference.Value = path; Put(source, "Doc_0/Document.xml", document);
        // An undeclared table with duplicate IDs must neither be indexed nor modified.
        source.PreservedEntries["Doc_0/Extensions/Unowned.dat"] = unowned;
        using var input = Zip(source.PreservedEntries); source = await new OfdReader().ReadAsync(input);
        if (operation == "watermark")
        {
            OfdWatermark.AddText(source, [1], "DRAFT"); OfdWatermark.AddImage(source, [1], Png, "image/png");
        }
        if (operation == "split") source = OfdDocumentSplitter.Split(source, [1]);
        var saved = await RoundTrip(source);
        Assert.Equal(path, Xml(saved, "Doc_0/Document.xml").Root!.Element(ns + "CommonData")!.Element(ns + declaration)!.Value);
        Assert.Equal("yes", Xml(saved, "Doc_0/" + path).Root!.Attribute(XNamespace.Get("urn:vendor") + "Keep")!.Value);
        Assert.Equal(unowned, saved.PreservedEntries["Doc_0/Extensions/Unowned.dat"]);
        Assert.Contains(saved.Fonts, font => font.Data.SequenceEqual(new byte[] { 9, 9, 9 }));
        Assert.Contains(saved.Pages.SelectMany(page => page.Elements).OfType<OfdImageElement>(), image => image.Data.SequenceEqual(Png));
        Assert.Contains(saved.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(), text => text.Text == "SECOND");
        if (operation == "watermark") Assert.Contains(saved.Pages[1].Elements.OfType<OfdTextElement>(), text => text.Text == "DRAFT");
    }

    [Theory]
    [InlineData("PublicRes", false)]
    [InlineData("DocumentRes", false)]
    [InlineData("PublicRes", true)]
    [InlineData("DocumentRes", true)]
    public async Task DeclaredNonXmlResource_UnsupportedExistingPayloadStillRefusesOverwrite(string declaration, bool malformed)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source, "Doc_0/Document.xml"); var reference = document.Root!.Element(ns + "CommonData")!.Element(ns + declaration)!;
        source.PreservedEntries["Doc_0/Unsupported.dat"] = source.PreservedEntries["Doc_0/" + reference.Value]; reference.Value = "Unsupported.dat";
        Put(source, "Doc_0/Document.xml", document);
        var payload = Encoding.UTF8.GetBytes(malformed ? "<Res" : "<Res xmlns='urn:vendor'><Fonts/></Res>");
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        read.PreservedEntries["Doc_0/Unsupported.dat"] = payload;
        using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(read, output));
        Assert.Equal(0, output.Length); Assert.Equal(payload, read.PreservedEntries["Doc_0/Unsupported.dat"]);
    }
    [Theory]
    [InlineData("Extension","urn:vendor")]
    [InlineData("Extension","http://www.ofdspec.org")]
    [InlineData("Signatures","urn:vendor")]
    public async Task Mix_RejectsDocBodyExtensionsWhileOrdinarySavePreserves(string name,string extensionNamespace)
    {
        var source=await RoundTrip(Source());var root=Xml(source,"OFD.xml");var extension=new XElement(XName.Get(name,extensionNamespace),new XAttribute("File","Doc_0/Extensions/payload.bin"));root.Root!.Elements().Single().Add(extension);Put(source,"OFD.xml",root);source.PreservedEntries["Doc_0/Extensions/payload.bin"]=[4,5,6];
        using var input=Zip(source.PreservedEntries);var read=await new OfdReader().ReadAsync(input);Assert.Contains(read.PreservedDocBodyElements,xml=>XElement.Parse(xml).Name==extension.Name);
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,0)]));
        var saved=await RoundTrip(read);var preserved=Assert.Single(Xml(saved,"OFD.xml").Descendants().Where(node=>node.Name==extension.Name));Assert.Equal(extension.Attribute("File")!.Value,preserved.Attribute("File")!.Value);Assert.Equal(new byte[]{4,5,6},saved.PreservedEntries["Doc_0/Extensions/payload.bin"]);
    }

    [Theory]
    [InlineData("page")]
    [InlineData("font")]
    [InlineData("media")]
    [InlineData("font-path")]
    [InlineData("media-path")]
    [InlineData("opaque")]
    public async Task Split_ExtensionlessXmlProtectsPageAndResourceClosure(string kind)
    {
        var source=Source();source.Pages[0].Elements.OfType<OfdImageElement>().Single().Data=[7,7,7];source=await RoundTrip(source);
        var pagePath=source.Pages[0].SourceEntryPath!;var font=source.Pages[0].Elements.OfType<OfdTextElement>().Single().FontResourceId!;var media=source.Pages[0].Elements.OfType<OfdImageElement>().Single().ResourceId;
        var fontPath=source.PreservedEntries.Single(entry=>entry.Value.SequenceEqual(new byte[]{8,8,8})).Key;var mediaPath=source.PreservedEntries.Single(entry=>entry.Value.SequenceEqual(new byte[]{7,7,7})).Key;
        var xml=kind switch {"page"=>$"<Extension File='/{pagePath}'/>","font"=>$"<Extension Font='{font}'/>","media"=>$"<Extension ResourceID='{media}'/>","font-path"=>$"<Extension File='/{fontPath}'/>","media-path"=>$"<Extension File='/{mediaPath}'/>",_=>"<Extension"};
        source.PreservedEntries["Doc_0/Extensions/state.dat"]=Encoding.UTF8.GetBytes(xml);using var input=Zip(source.PreservedEntries);source=await new OfdReader().ReadAsync(input);
        var result=await RoundTrip(OfdDocumentSplitter.Split(source,[1]));Assert.Equal(source.PreservedEntries["Doc_0/Extensions/state.dat"],result.PreservedEntries["Doc_0/Extensions/state.dat"]);
        if(kind is "page" or "opaque")Assert.Equal(source.PreservedEntries[pagePath],result.PreservedEntries[pagePath]);
        if(kind is "page" or "font" or "font-path" or "opaque")Assert.True(result.PreservedEntries.ContainsKey(fontPath));
        if(kind is "page" or "media" or "media-path" or "opaque")Assert.True(result.PreservedEntries.ContainsKey(mediaPath));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SignatureCleanup_ExtensionlessNestedBaseLocRestoresSharedPayloadClosure(bool opaque)
    {
        var source=await RoundTrip(Source());AddSignatures(source,"value.xml");source.PreservedEntries["Doc_0/Signs/Sign_0/value.xml"]=Encoding.UTF8.GetBytes("<Value Seal='Seal.esl'/>");
        source.PreservedEntries["Doc_0/Extensions/state.dat"]=Encoding.UTF8.GetBytes(opaque?"<Wrapper":"<Wrapper><Extension BaseLoc='../Signs/Sign_0'><File>value.xml</File></Extension></Wrapper>");
        using var output=new MemoryStream();await OfdPackageSignatureCleaner.CleanAsync(Zip(source.PreservedEntries),output);output.Position=0;var result=await new OfdPackageLoader().LoadAsync(output);
        Assert.Equal(source.PreservedEntries["Doc_0/Signs/Sign_0/value.xml"],result.GetBytes("Doc_0/Signs/Sign_0/value.xml"));Assert.Equal(Png,result.GetBytes("Doc_0/Signs/Sign_0/Seal.esl"));Assert.DoesNotContain("Signatures",result.ReadUtf8Text("OFD.xml"));
    }

    [Fact]
    public async Task Split_ExtensionlessPageBesideImplicitResourcesKeepsTheirClosure()
    {
        var source=await RoundTrip(Source());var ns=source.Options.Namespace;
        source.PreservedEntries["Doc_0/Pages/Page_0/PageRes.xml"]=Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><MultiMedias><MultiMedia ID='710' Format='PNG'><MediaFile>Private.png</MediaFile></MultiMedia></MultiMedias></Res>");
        source.PreservedEntries["Doc_0/Pages/Page_0/Private.png"]=[5,6,7];
        source.PreservedEntries["Doc_0/Pages/Page_0/shared.dat"]=Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='711'><ImageObject ID='712' ResourceID='710' Boundary='0 0 10 10'/></Layer></Content></Page>");
        using var input=Zip(source.PreservedEntries);source=await new OfdReader().ReadAsync(input);var result=await RoundTrip(OfdDocumentSplitter.Split(source,[1]));
        Assert.Equal(source.PreservedEntries["Doc_0/Pages/Page_0/shared.dat"],result.PreservedEntries["Doc_0/Pages/Page_0/shared.dat"]);
        Assert.True(result.PreservedEntries.ContainsKey("Doc_0/Pages/Page_0/PageRes.xml"));Assert.Equal(new byte[]{5,6,7},result.PreservedEntries["Doc_0/Pages/Page_0/Private.png"]);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Reader_UnmodeledAnnotationRecordsAreScopedAndAggregated(bool global)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source,"Doc_0/Document.xml"); document.Root!.Add(new XElement(ns+"Annotations","Annots/Annotations.xml")); Put(source,"Doc_0/Document.xml",document);
        var id = global ? "" : $" PageID='{source.Pages[1].Id}'";
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}' xmlns:v='urn:vendor'>"+string.Concat(Enumerable.Repeat($"<Page{id}><v:FileLoc>unmodeled.bin</v:FileLoc></Page>",1000))+"</Annotations>");
        using var input=Zip(source.PreservedEntries); var read=await new OfdReader().ReadAsync(input);
        Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[1].AnnotationAppearances));
        if (global)
        {
            Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].AnnotationAppearances));
            Assert.Same(((OfdRawElement)read.Pages[0].AnnotationAppearances[0]).Xml,((OfdRawElement)read.Pages[1].AnnotationAppearances[0]).Xml);
            Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,0)]));
        }
        else
        {
            Assert.Empty(read.Pages[0].AnnotationAppearances);
            Assert.Single(OfdDocumentMixer.Mix([new(read,0)]).Pages);
        }
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,1)]));
        var saved=await RoundTrip(read); Assert.Equal(source.PreservedEntries["Doc_0/Annots/Annotations.xml"],saved.PreservedEntries["Doc_0/Annots/Annotations.xml"]);
    }

    [Theory]
    [InlineData(false,false)]
    [InlineData(false,true)]
    [InlineData(true,false)]
    [InlineData(true,true)]
    public async Task Reader_OrphanAnnotationKeysStayGlobalForNullOrUnmatchedPageIds(bool nullPageIds,bool validRecord)
    {
        var source=await RoundTrip(Source());var ns=XNamespace.Get(source.Options.Namespace);var orphan=nullPageIds?source.Pages[0].Id!:"UNMATCHED";
        var document=Xml(source,"Doc_0/Document.xml");if(nullPageIds)document.Root!.Element(ns+"Pages")!.Elements(ns+"Page").Attributes("ID").Remove();
        document.Root!.Add(new XElement(ns+"Annotations","Annots/Annotations.xml"));Put(source,"Doc_0/Document.xml",document);
        var location=validRecord?"<FileLoc>Page.xml</FileLoc>":"<v:FileLoc xmlns:v='urn:vendor'>unmodeled.bin</v:FileLoc>";
        source.PreservedEntries["Doc_0/Annots/Annotations.xml"]=Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{orphan}'>{location}</Page></Annotations>");
        source.PreservedEntries["Doc_0/Annots/Page.xml"]=Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot><Appearance Boundary='0 0 20 20'><TextObject Size='4'><TextCode X='0' Y='4'>ORPHAN</TextCode></TextObject></Appearance></Annot></PageAnnot>");
        using var input=Zip(source.PreservedEntries);var read=await new OfdReader().ReadAsync(input);
        Assert.All(read.Pages,page=>Assert.IsType<OfdRawElement>(Assert.Single(page.AnnotationAppearances)));
        Assert.Same(((OfdRawElement)read.Pages[0].AnnotationAppearances[0]).Xml,((OfdRawElement)read.Pages[1].AnnotationAppearances[0]).Xml);
        Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,0)]));Assert.Throws<NotSupportedException>(()=>OfdDocumentMixer.Mix([new(read,1)]));
        var saved=await RoundTrip(read);Assert.Equal(source.PreservedEntries["Doc_0/Annots/Annotations.xml"],saved.PreservedEntries["Doc_0/Annots/Annotations.xml"]);
    }
    [Fact]
    public async Task Writer_MissingPathDataUsesSourceObjectNamespaceAndKeepsExtensions()
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdPathElement { SourceXml = "<PathObject><v:Note xmlns:v='urn:vendor'>KEEP</v:Note></PathObject>", AbbreviatedData = "M 0 0 L 10 0", Stroke = false, Fill = false }); source.Pages.Add(page);
        var saved = await RoundTrip(source); var path = saved.Pages[0].Elements.OfType<OfdPathElement>().Single();
        Assert.Equal("M 0 0 L 10 0", path.AbbreviatedData); var xml = XElement.Parse(path.SourceXml!);
        Assert.Equal("M 0 0 L 10 0", xml.Element(xml.Name.Namespace + "AbbreviatedData")!.Value);
        Assert.Equal("KEEP", xml.Element(XName.Get("Note", "urn:vendor"))!.Value);
    }
    [Fact]
    public void Mix_KeepsSupportedAreaTransformAndMergeRejectsUnmappedSubstitution()
    {
        var source = Source(); var image = source.Pages[0].Elements.OfType<OfdImageElement>().Single();
        source.Pages[0].Elements.OfType<OfdTextElement>().Single().SourceXml = "<TextObject Name='body' Visible='true' LineWidth='0.5'><TextCode X='0' Y='4'>FIRST</TextCode></TextObject>";
        image.ClippingXml = "<Clips><Clip><Area Start='0 0' CTM='1 0 0 1 2 3'><Path ID='800' Name='clip' Visible='true' Stroke='false' Fill='true' LineWidth='0.35' Alpha='255' Cap='Butt' Join='Miter' MiterLimit='3' DashOffset='0' DashPattern='1 2'><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>";
        var mixed = OfdDocumentMixer.Mix([new(source, 0)]);
        var clip = XElement.Parse(mixed.Pages[0].Elements.OfType<OfdImageElement>().Single().ClippingXml!);
        Assert.Equal("1 0 0 1 2 3", clip.Descendants(clip.Name.Namespace + "Area").Single().Attribute("CTM")!.Value);
        image.SourceXml = "<ImageObject Substitution='123'/>";
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
        image.SourceXml = "<ImageObject ImageMask='123'/>";
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Reader_PageAndTemplateLeavesUseOnlyDirectLiteralTextAndPreserveExtensions(bool template)
    {
        var source = await RoundTrip(Source()); var ns = XNamespace.Get(source.Options.Namespace); var pagePath = source.Pages[0].SourceEntryPath!;
        var fixture = XElement.Parse($"<Page xmlns='{ns}'><Content><Layer ID='700'><TextObject ID='701' Font='10' Size='4' Boundary='0 0 40 10'><TextCode X='0' Y='4'>OK<Note>SECRET</Note></TextCode><Note><TextCode>NESTED</TextCode></Note></TextObject><TextObject ID='702' Boundary='0 20 40 10'><Note>FALLBACK</Note></TextObject><PathObject ID='703' Boundary='0 30 40 10'><AbbreviatedData>M 0 0 L 1 0<Note> L 9 9</Note></AbbreviatedData></PathObject></Layer></Content></Page>");
        var fixturePath = template ? "Doc_0/Templates/Content.xml" : pagePath;
        source.PreservedEntries[fixturePath] = Encoding.UTF8.GetBytes(fixture.ToString());
        if (template)
        {
            var document = Xml(source, "Doc_0/Document.xml"); document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml"))); Put(source, "Doc_0/Document.xml", document);
            var page = Xml(source, pagePath); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700"))); Put(source, pagePath, page);
        }
        using var input = Zip(source.PreservedEntries); var read = await new OfdReader().ReadAsync(input);
        var elements = template ? read.Pages[0].Templates.Single().Elements : read.Pages[0].Elements;
        Assert.Equal(new[] { "OK", "" }, elements.OfType<OfdTextElement>().Select(text => text.Text));
        Assert.Equal("M 0 0 L 1 0", elements.OfType<OfdPathElement>().Single().AbbreviatedData);
        Assert.DoesNotContain("SECRET", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read, includeTemplates: true));
        Assert.DoesNotContain("FALLBACK", new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read, includeTemplates: true));
        var saved = await RoundTrip(read); Assert.Contains("SECRET", Encoding.UTF8.GetString(saved.PreservedEntries[fixturePath]));
        var savedXml = Xml(saved, fixturePath);
        Assert.Equal(" L 9 9", savedXml.Descendants(ns + "AbbreviatedData").Single().Element(ns + "Note")!.Value);
        if (!template)
        {
            read.Pages[0].Elements.OfType<OfdPathElement>().Single().AbbreviatedData = "M 1 1 L 2 2";
            var changed = await RoundTrip(read); Assert.Equal("M 1 1 L 2 2", changed.Pages[0].Elements.OfType<OfdPathElement>().Single().AbbreviatedData);
            Assert.Equal(" L 9 9", Xml(changed, fixturePath).Descendants(ns + "AbbreviatedData").Single().Element(ns + "Note")!.Value);
        }
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(read, 0)]));
    }
    [Theory]
    [InlineData("Payload")]
    [InlineData("v:Font")]
    [InlineData("v:ResourceID")]
    public async Task Mix_RejectsUnknownOrdinaryObjectAttributesWhileRoundtripPreserves(string attribute)
    {
        var source = await RoundTrip(Source()); var text = source.Pages[0].Elements.OfType<OfdTextElement>().First();
        var xml = XElement.Parse(text.SourceXml!); var name = attribute.StartsWith("v:") ? XName.Get(attribute.Substring(2), "urn:vendor") : XName.Get(attribute);
        xml.SetAttributeValue(name, "/Doc_0/Extensions/private.bin"); text.SourceXml = xml.ToString();
        source.PreservedEntries["Doc_0/Extensions/private.bin"] = [1, 2, 3];
        var saved = await RoundTrip(source);
        Assert.Equal("/Doc_0/Extensions/private.bin", XElement.Parse(saved.Pages[0].Elements.OfType<OfdTextElement>().First().SourceXml!).Attribute(name)!.Value);
        Assert.Equal(new byte[] {1,2,3}, saved.PreservedEntries["Doc_0/Extensions/private.bin"]);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new(saved, 0)]));
    }
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
