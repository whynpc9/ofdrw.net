using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TypedNavigationTests
{
    [Fact]
    public async Task NavigationUsesActualPagesThroughWriteReadReorderAndSplit()
    {
        var source = Document();
        source.Bookmarks = new OfdBookmarkCollection();
        var bookmark = new OfdBookmark { Name = "中英 landing", Destination = new OfdDestination { Page = source.Pages[1], Type = OfdDestinationType.XYZ, Left = 12.5, Top = 7, Zoom = 1.25 } };
        source.Bookmarks.Items.Add(bookmark);
        source.Outlines = new OfdOutlineCollection(); var outline = new OfdOutline { Title = "目录 Outline", Expanded = true, Count = 0, Actions = new OfdActionCollection() };
        outline.Actions.Items.Add(new OfdGoToAction { Bookmark = bookmark, Event = "CLICK" }); source.Outlines.Items.Add(outline);
        source.Actions = new OfdActionCollection(); source.Actions.Items.Add(new OfdGoToAction { Destination = new OfdDestination { Page = source.Pages[0], Type = OfdDestinationType.Fit } });
        var loaded = await RoundTrip(source);
        Assert.Same(loaded.Pages[1], loaded.Bookmarks!.Items[0].Destination.Page);
        Assert.Same(loaded.Bookmarks.Items[0], Assert.IsType<OfdGoToAction>(loaded.Outlines!.Items[0].Actions!.Items[0]).Bookmark);
        OfdDocumentEditor.ReorderPages(loaded, [1, 0]);
        loaded = await RoundTrip(loaded);
        Assert.Equal("A", Assert.Single(Assert.IsType<OfdGoToAction>(loaded.Actions!.Items[0]).Destination!.Page.Elements.OfType<OfdTextElement>()).Text);
        Assert.Equal("B", Assert.Single(loaded.Bookmarks!.Items[0].Destination.Page.Elements.OfType<OfdTextElement>()).Text);
        var split = OfdDocumentSplitter.Split(loaded, [1, 0]);
        Assert.NotSame(loaded.Bookmarks.Items[0].Destination.Page, split.Bookmarks!.Items[0].Destination.Page);
        Assert.Same(split.Pages[1], split.Bookmarks.Items[0].Destination.Page);
        _ = await RoundTrip(split);
    }

    [Fact]
    public async Task EmptyAndNullOwnershipPreserveOpaqueAnchorsAndRejectDanglingBookmark()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source, source.DocumentEntryPath!);
        document.Root!.Add(XElement.Parse($"<Outlines xmlns='{ns}' xmlns:v='urn:vendor' v:Keep='v:Value'><OutlineElem Title='A'/><v:Anchor/><OutlineElem Title='B'/></Outlines>"));
        document.Root.Add(XElement.Parse($"<Actions xmlns='{ns}'><Action><Goto><Bookmark Name='B'/></Goto></Action></Actions>"));
        document.Root.Add(XElement.Parse($"<Bookmarks xmlns='{ns}'><Bookmark Name='B'><Dest Type='Fit' PageID='{source.Pages[0].Id}'/></Bookmark></Bookmarks>"));
        source = await ReadChanged(source, source.DocumentEntryPath!, document);
        source.Outlines!.Items.RemoveAt(0);
        var saved = await RoundTrip(source); var outlines = Xml(saved, saved.DocumentEntryPath!).Root!.Element(ns + "Outlines")!;
        Assert.Equal(new[] { "Anchor", "OutlineElem" }, outlines.Elements().Select(node => node.Name.LocalName));
        source.Outlines.Items.Clear(); saved = await RoundTrip(source); outlines = Xml(saved, saved.DocumentEntryPath!).Root!.Element(ns + "Outlines")!;
        Assert.Single(outlines.Elements()); Assert.Equal("v:Value", outlines.Attribute(XName.Get("Keep", "urn:vendor"))!.Value);
        source.Actions = null; source.Bookmarks!.Items[0].Name = "renamed";
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidOperationException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task SplitKeepsModifiedOwnershipDirtyAndPristineOwnershipWithdrawable()
    {
        var source = Document(); source.Outlines = new OfdOutlineCollection(); source.Outlines.Items.Add(new OfdOutline { Title = "A" });
        source = await RoundTrip(source);
        var pristine = OfdDocumentSplitter.Split(source, [0]); pristine.Outlines = null; _ = await RoundTrip(pristine);
        source.Outlines!.Items[0].Title = "Edited";
        var edited = OfdDocumentSplitter.Split(source, [0]); Assert.Throws<InvalidOperationException>(() => edited.Outlines = null);
    }

    [Fact]
    public async Task NewPreferredIdCannotClaimAnOpaqueReservation()
    {
        var source = Document(); source.PreservedEntries["Doc_0/opaque.dat"] = Encoding.UTF8.GetBytes("<Vendor ID='900'><RefID>900</RefID></Vendor>");
        source.Pages[0].Elements[0].ObjectId = "900";
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output));
        Assert.Equal(0, output.Length); Assert.Equal("900", source.Pages[0].Elements[0].ObjectId);
    }

    [Fact]
    public async Task ImportedIdsKeepLexemesDespiteOpaqueDuplicateShadows()
    {
        var source = Document(); source.Pages[0].Id = "001"; source = await RoundTrip(source);
        var bytes = Encoding.UTF8.GetBytes("<Unowned ID='001' RefID='001'/>"); source.PreservedEntries["Doc_0/Unowned.dat"] = bytes;
        var saved = await RoundTrip(source); Assert.Equal("001", saved.Pages[0].Id); Assert.Equal(bytes, saved.PreservedEntries["Doc_0/Unowned.dat"]);
    }

    [Fact]
    public async Task OriginalTemplateLayerDuplicateFixtureIsRejectedAsActiveDefinitions()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        var document = Xml(source, source.DocumentEntryPath!); document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "700"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        source.PreservedEntries[source.DocumentEntryPath!] = Encoding.UTF8.GetBytes(document.ToString());
        var page = Xml(source, source.Pages[0].SourceEntryPath!); page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "700")));
        source.PreservedEntries[source.Pages[0].SourceEntryPath!] = Encoding.UTF8.GetBytes(page.ToString());
        source.PreservedEntries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='700'><TextObject ID='701' Size='4'><TextCode X='0' Y='4'>OK<Note>SECRET</Note></TextCode></TextObject><PathObject ID='703'><AbbreviatedData>M 0 0 L 1 0<Note> L 9 9</Note></AbbreviatedData></PathObject></Layer></Content></Page>");
        using var input = Zip(source.PreservedEntries); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(input));
    }

    [Fact]
    public async Task FixedClipPathExtensionIdsKeepDuplicateLiteralsAndUnknownReferenceAttributes()
    {
        OfdDocumentPackage Input()
        {
            var package = Document(); package.Pages.RemoveAt(1);
            package.Pages[0].Elements[0].ClippingXml = "<Clips><Clip><Area><Path ID='100'><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>";
            return package;
        }
        var merged = await RoundTrip(OfdDocumentMerger.Merge([Input(), Input()]));
        var ns = XNamespace.Get(merged.Options.Namespace);
        var reference = XNamespace.Get("urn:vendor");
        foreach (var page in merged.Pages)
        {
            var xml = Xml(merged, page.SourceEntryPath!); var path = xml.Descendants(ns + "Path").Single();
            Assert.Equal("100", path.Attribute("ID")!.Value);
            path.SetAttributeValue(reference + "RefID", "100"); path.SetAttributeValue(reference + "ObjectRef", "100");
            merged.PreservedEntries[page.SourceEntryPath!] = Encoding.UTF8.GetBytes(xml.ToString());
        }
        using (var input = Zip(merged.PreservedEntries)) merged = await new OfdReader().ReadAsync(input);
        var saved = await RoundTrip(merged);
        foreach (var page in saved.Pages)
        {
            var path = Xml(saved, page.SourceEntryPath!).Descendants(ns + "Path").Single();
            Assert.Equal("100", path.Attribute("ID")!.Value); Assert.Equal("100", path.Attribute(reference + "RefID")!.Value); Assert.Equal("100", path.Attribute(reference + "ObjectRef")!.Value);
        }
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([saved])); // unknown cross-package references are not proven by reservation
        saved.Pages[0].Elements.Add(new OfdPathElement { ObjectId = "100", AbbreviatedData = "M 0 0 L 1 1" });
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(saved, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task ActualPathObjectDuplicateCannotUseClipShadowPolicy()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace); var path = source.Pages[0].SourceEntryPath!; var xml = Xml(source, path);
        var layer = xml.Descendants(ns + "Layer").First(); var id = source.Pages[0].Elements[0].ObjectId!;
        layer.Add(new XElement(ns + "PathObject", new XAttribute("ID", id), new XElement(ns + "AbbreviatedData", "M 0 0 L 1 1")));
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString());
        using var input = Zip(source.PreservedEntries); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(input));
    }

    [Fact]
    public async Task ImportedFragmentCannotAddAnIdFromUnwrittenPreservedDefinitions()
    {
        var source = await RoundTrip(Document()); source.PreservedEntries["Doc_0/opaque.dat"] = Encoding.UTF8.GetBytes("<Vendor ID='900'/>");
        var text = (OfdTextElement)source.Pages[0].Elements[0]; var xml = XElement.Parse(text.SourceXml!); xml.Add(new XElement(XName.Get("Extra", "urn:vendor"), new XAttribute("ID", "900"))); text.SourceXml = xml.ToString();
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task AppendedActiveSourceDefinitionIsDetectedBeforeWriterOutput()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace); var path = source.Pages[0].SourceEntryPath!; var xml = Xml(source, path);
        xml.Descendants(ns + "Layer").First().Add(new XElement(ns + "PathObject", new XAttribute("ID", source.Pages[1].Elements[0].ObjectId!), new XElement(ns + "AbbreviatedData", "M 0 0 L 1 1")));
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString());
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task PageDeclarationOpaqueAnchorRejectsReorderBeforeMutationAndAllowsDeletion()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        var xml = Xml(source, source.DocumentEntryPath!); var pages = xml.Root!.Element(ns + "Pages")!;
        pages.Elements().First().AddAfterSelf(new XElement(XName.Get("Anchor", "urn:vendor")));
        source = await ReadChanged(source, source.DocumentEntryPath!, xml); var originals = source.Pages.ToArray();
        Assert.Throws<NotSupportedException>(() => OfdDocumentEditor.ReorderPages(source, [1, 0]));
        Assert.Equal(originals, source.Pages); Assert.Equal(new[] { 0, 1 }, source.Pages.Select(page => page.Index));
        Assert.Throws<NotSupportedException>(() => OfdDocumentSplitter.Split(source, [1, 0]));
        OfdDocumentEditor.RemovePages(source, [0]); var saved = await RoundTrip(source);
        Assert.Equal(new[] { "Anchor", "Page" }, Xml(saved, saved.DocumentEntryPath!).Root!.Element(ns + "Pages")!.Elements().Select(node => node.Name.LocalName));
    }

    [Fact]
    public async Task RemoveReferencedPageFailsBeforeInputMutation()
    {
        var source = Document(); source.Bookmarks = new OfdBookmarkCollection();
        source.Bookmarks.Items.Add(new OfdBookmark { Name = "B", Destination = new OfdDestination { Page = source.Pages[1] } });
        source = await RoundTrip(source); var originals = source.Pages.ToArray();
        Assert.Throws<InvalidOperationException>(() => OfdDocumentEditor.RemovePages(source, [1]));
        Assert.Equal(originals, source.Pages); Assert.Equal(new[] { 0, 1 }, source.Pages.Select(page => page.Index));
    }

    [Fact]
    public async Task PreservedOwnedRootConflictFailsBeforeDestinationBytes()
    {
        var source = Document(); source.Outlines = new OfdOutlineCollection(); source.Outlines.Items.Add(new OfdOutline { Title = "A" });
        source.PreservedDocumentElements.Add($"<Outlines xmlns='{source.Options.Namespace}'><OutlineElem Title='B'/></Outlines>");
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task DocumentEnvelopesRetainUnknownNodesAttributesAndMixedMetadata()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace); var v = XNamespace.Get("urn:vendor");
        var ofd = Xml(source, "OFD.xml"); ofd.Root!.SetAttributeValue(XNamespace.Xmlns + "v", v.NamespaceName); ofd.Root.SetAttributeValue(v + "Kind", "v:Root");
        var body = ofd.Root.Element(ns + "DocBody")!; body.SetAttributeValue(v + "Keep", "v:Body"); body.AddFirst(new XProcessingInstruction("vendor", "body"));
        body.Element(ns + "DocInfo")!.Add(new XElement(ns + "Title", new XAttribute(v + "Keep", "v:Title"), "Head", new XElement(v + "Part", "SECRET"), "Tail"));
        source.PreservedEntries["OFD.xml"] = Encoding.UTF8.GetBytes(ofd.ToString());
        var document = Xml(source, source.DocumentEntryPath!); document.Root!.SetAttributeValue(v + "Kind", "v:Doc");
        document.Root.AddFirst(new XComment("root comment")); document.Root.Element(ns + "CommonData")!.SetAttributeValue(v + "Keep", "v:Common");
        document.Root.Element(ns + "Pages")!.Elements().First().SetAttributeValue(v + "Keep", "v:Declaration");
        source = await ReadChanged(source, source.DocumentEntryPath!, document); var saved = await RoundTrip(source);
        var savedOfd = Xml(saved, "OFD.xml"); Assert.Equal("v:Root", savedOfd.Root!.Attribute(v + "Kind")!.Value);
        Assert.Equal("urn:vendor", savedOfd.Root.GetNamespaceOfPrefix("v")!.NamespaceName);
        var savedBody = savedOfd.Root.Element(ns + "DocBody")!; Assert.Equal("v:Body", savedBody.Attribute(v + "Keep")!.Value);
        Assert.Contains(savedBody.Nodes(), node => node is XProcessingInstruction pi && pi.Data == "body");
        Assert.True(XNode.DeepEquals(body.Element(ns + "DocInfo")!.Element(ns + "Title"), savedBody.Element(ns + "DocInfo")!.Element(ns + "Title")));
        var savedDocument = Xml(saved, saved.DocumentEntryPath!); Assert.Equal("v:Doc", savedDocument.Root!.Attribute(v + "Kind")!.Value);
        Assert.Contains(savedDocument.Root.Nodes(), node => node is XComment comment && comment.Value == "root comment");
        Assert.Equal("v:Common", savedDocument.Root.Element(ns + "CommonData")!.Attribute(v + "Keep")!.Value);
        Assert.Equal("v:Declaration", savedDocument.Root.Element(ns + "Pages")!.Elements().First().Attribute(v + "Keep")!.Value);
        source.Options.Metadata.Title = "Changed"; using var rejected = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, rejected)); Assert.Equal(0, rejected.Length);
    }

    [Fact]
    public async Task NewTemplateSourceCannotBypassActiveDefinitionOwnership()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        source.PreservedCommonDataElements.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "900"), new XAttribute("BaseLoc", "new.xml")).ToString());
        source.PreservedEntries["Doc_0/new.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='{source.Pages[0].Elements[0].LayerId}'><TextObject ID='{source.Pages[0].Elements[0].ObjectId}'/></Layer></Content></Page>");
        source.Pages[0].PreservedPageElements.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "900")).ToString());
        using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task DistinctImportedImageOwnersCountAliasedPayloadAtBothOutputPositions()
    {
        var source = Document(); var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j9AoAAAAASUVORK5CYII=");
        for (var i = 0; i < 2; i++) source.Pages[i].Elements.Add(new OfdImageElement { Data = png.Concat(new[] { (byte)i }).ToArray(), MediaType = "image/png", FileName = "image.png" });
        source = await RoundTrip(source); var payload = new byte[1024 * 1024]; png.CopyTo(payload, 0);
        foreach (var page in source.Pages) Assert.Single(page.Elements.OfType<OfdImageElement>()).Data = payload;
        var saved = await RoundTrip(source); Assert.All(saved.Pages, page => Assert.Equal(payload, Assert.Single(page.Elements.OfType<OfdImageElement>()).Data));
        Assert.NotEqual(Assert.Single(saved.Pages[0].Elements.OfType<OfdImageElement>()).ResourceId, Assert.Single(saved.Pages[1].Elements.OfType<OfdImageElement>()).ResourceId);
    }

    [Fact]
    public async Task ImportedOpaqueIdCannotBecomeActiveThroughAncestorRoleChange()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace); var path = source.Pages[0].SourceEntryPath!;
        var xml = Xml(source, path); var existing = source.Pages[0].Elements[0].ObjectId!;
        xml.Descendants(ns + "Layer").First().Add(new XElement(ns + "PageBlock", new XAttribute("ID", "900"),
            new XElement(XName.Get("Wrapper", "urn:vendor"), new XElement(ns + "PathObject", new XAttribute("ID", existing), new XElement(ns + "AbbreviatedData", "M 0 0 L 1 1")))));
        source = await ReadChanged(source, path, xml); _ = await RoundTrip(source);
        var raw = Assert.Single(source.Pages[0].Elements.OfType<OfdRawElement>()); var fragment = XElement.Parse(raw.Xml);
        fragment.Elements().Single().Name = ns + "PageBlock"; raw.Xml = fragment.ToString();
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task UnselectedOpaqueShadowOfFontIdRetainsIdAndReference()
    {
        var source = Document(); source.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Synthetic" });
        foreach (var item in source.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()) { item.FontName = "Synthetic"; item.FontResourceId = "10"; }
        source = await RoundTrip(source); var path = source.Pages[1].SourceEntryPath!; var xml = Xml(source, path); var ns = xml.Root!.Name.Namespace; XNamespace v = "urn:vendor";
        xml.Descendants(ns + "Layer").First().Add(XElement.Parse("<v:Extension xmlns:v='urn:vendor' ID='80'><v:Node ID='10'/><v:Link RefID='10' Kind='v:Node'/></v:Extension>"));
        source = await ReadChanged(source, path, xml); var owned = OfdDocumentSplitter.Split(source, [0, 1]);
        var text = Assert.Single(owned.Pages[0].Elements.OfType<OfdTextElement>()); var changed = XElement.Parse(text.SourceXml!); changed.Element(ns + "TextCode")!.Value = "EDIT";
        text.SourceXml = changed.ToString(); text.Text = "EDIT"; text.Runs[0].Text = "EDIT";
        var saved = await RoundTrip(owned); var unselected = Xml(saved, path);
        Assert.Equal("10", unselected.Descendants(v + "Node").Single().Attribute("ID")!.Value);
        var link = unselected.Descendants(v + "Link").Single(); Assert.Equal("10", link.Attribute("RefID")!.Value); Assert.Equal(v, link.GetNamespaceOfPrefix("v"));
        Assert.Equal("B", Assert.Single(source.Pages[1].Elements.OfType<OfdTextElement>()).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NewExplicitOrImplicitResourceTableCannotBypassActiveOwnership(bool implicitPageResource)
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        var path = implicitPageResource ? source.Pages[0].SourceEntryPath!.Replace("Content.xml", "PageRes.xml") : "Doc_0/new-res.xml";
        if (!implicitPageResource) source.PublicResourceLocation = "new-res.xml";
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><DrawParams><DrawParam ID='{source.Pages[0].Elements[0].ObjectId}'/></DrawParams></Res>");
        using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MaxNumericReservationAllowsAuthenticatedRewriteButRefusesAllocation(bool addObject)
    {
        var source = await RoundTrip(Document()); source.PreservedEntries["Doc_0/max.dat"] = Encoding.UTF8.GetBytes("<Vendor ID='9223372036854775807'/>");
        if (addObject) source.Pages[0].Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 1 1" });
        using var output = new MemoryStream();
        if (addObject) { await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length); }
        else { await new OfdPackageWriter().WriteAsync(source, output); Assert.True(output.Length > 0); }
    }

    [Fact]
    public async Task DestinationCallbackCannotModifyValidatedPayloadSnapshot()
    {
        var source = Document(); var bytes = Enumerable.Repeat((byte)90, 10000).ToArray();
        source.Attachments.Add(new OfdAttachment { Name = "snapshot.bin", Data = bytes });
        using var output = new CallbackStream(() => Array.Fill(bytes, (byte)0));
        await new OfdPackageWriter().WriteAsync(source, output); Assert.All(bytes, value => Assert.Equal(0, value));
        output.Position = 0; var read = await new OfdReader().ReadAsync(output);
        Assert.All(Assert.Single(read.Attachments).Data, value => Assert.Equal(90, value));
    }

    private sealed class CallbackStream(Action callback) : MemoryStream
    {
        private bool invoked;
        private void Invoke() { if (invoked) return; invoked = true; callback(); }
        public override void Write(byte[] buffer, int offset, int count) { Invoke(); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Invoke(); base.Write(buffer); }
    }

    [Fact]
    public async Task NewPageImplicitResourceTableCannotIntroduceUnauthenticatedDefinitions()
    {
        var source = await RoundTrip(Document()); var ns = XNamespace.Get(source.Options.Namespace);
        source.Pages.Add(new OfdPage { Index = 2, WidthMillimeters = 100, HeightMillimeters = 100 });
        source.PreservedEntries["Doc_0/Pages/Page_2/PageRes.xml"] = Encoding.UTF8.GetBytes($"<Res xmlns='{ns}'><DrawParams><DrawParam ID='{source.Pages[0].Elements[0].ObjectId}'/></DrawParams></Res>");
        using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task SplitEnumeratesCallerSelectionOnceAndUsesValidatedSnapshot()
    {
        var source = await RoundTrip(Document()); var selection = new SinglePassSelection();
        var split = OfdDocumentSplitter.Split(source, selection);
        Assert.Equal(1, selection.Enumerations); Assert.Equal("B", Assert.Single(Assert.Single(split.Pages).Elements.OfType<OfdTextElement>()).Text);
    }

    private sealed class SinglePassSelection : IReadOnlyList<int>
    {
        public int Enumerations { get; private set; }
        public int Count => 1;
        public int this[int index] => throw new InvalidOperationException("Selection indexer must not be reread.");
        public IEnumerator<int> GetEnumerator()
        {
            if (++Enumerations != 1) throw new InvalidOperationException("Selection was enumerated twice.");
            yield return 1;
        }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private static OfdDocumentPackage Document()
    {
        var package = new OfdDocumentPackage();
        foreach (var text in new[] { "A", "B" }) { var page = new OfdPage { Index = package.Pages.Count, WidthMillimeters = 100, HeightMillimeters = 100 }; page.Elements.Add(new OfdTextElement { Text = text }); package.Pages.Add(page); }
        return package;
    }
    private static XDocument Xml(OfdDocumentPackage package, string path) => XDocument.Parse(Encoding.UTF8.GetString(package.PreservedEntries[path]));
    private static async Task<OfdDocumentPackage> ReadChanged(OfdDocumentPackage package, string path, XDocument xml)
    { package.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString()); using var input = Zip(package.PreservedEntries); return await new OfdReader().ReadAsync(input); }
    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage package)
    { using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); output.Position = 0; return await new OfdReader().ReadAsync(output); }
    private static MemoryStream Zip(IDictionary<string, byte[]> entries)
    { var output = new MemoryStream(); using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var pair in entries) { using var entry = zip.CreateEntry(pair.Key).Open(); entry.Write(pair.Value); } output.Position = 0; return output; }
}
