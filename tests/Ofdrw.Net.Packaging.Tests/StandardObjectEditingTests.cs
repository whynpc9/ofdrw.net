using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class StandardObjectEditingTests
{
    [Fact]
    public async Task Split_RetainsEditedStandardObjectsAndTheirDocumentOwnership()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Permissions = new OfdPermissions { Edit = true };
        package.ViewingPreferences = new OfdViewingPreferences { PageMode = OfdPageMode.UseOutlines };
        package.Options.Metadata.CustomDatas = new OfdCustomDatas();
        package.Options.Metadata.CustomDatas.Add("case", "old");
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        var body = root.Root.Element(ns + "DocBody")!;
        body.Element(ns + "DocInfo")!.Add(new XElement(ns + "DocUsage", "EBook"),
            new XElement(ns + "CreatorVersion", "1.0"));
        var vendor = XNamespace.Get("urn:split-vendor");
        body.Element(ns + "DocInfo")!.Element(ns + "CustomDatas")!.Add(new XElement(vendor + "Keep", "metadata"));
        body.Add(new XElement(ns + "Versions", new XElement(ns + "Version",
            new XAttribute("ID", "v1"), new XAttribute("Index", "1"), new XAttribute("BaseLoc", "Doc_0/Versions/v1.xml")),
            new XElement(vendor + "Keep", "versions")));
        entries["OFD.xml"] = Bytes(root);
        entries["Doc_0/Versions/v1.xml"] = Encoding.UTF8.GetBytes("<VersionFile/>");
        var document = Xml(entries["Doc_0/Document.xml"]);
        document.Root!.Element(ns + "Permissions")!.Add(new XElement(vendor + "Keep", "permissions"));
        document.Root.Element(ns + "VPreferences")!.Add(new XElement(vendor + "Keep", "preferences"));
        entries["Doc_0/Document.xml"] = Bytes(document);
        using var input = MultiDocumentFixture.Zip(entries);
        var source = await new OfdReader().ReadAsync(input);
        source.Options.Metadata.DocUsage = OfdDocUsage.ENewsPaper;
        source.Options.Metadata.CreatorVersion = "2.0";
        source.Options.Metadata.CustomDatas!.Items[0].Value = "new";
        source.Permissions!.Edit = false;
        source.ViewingPreferences!.PageMode = OfdPageMode.UseBookmarks;
        source.Versions!.Items[0].Index = 2;

        var split = OfdDocumentSplitter.Split(source, [1]);
        Assert.NotSame(source.Permissions, split.Permissions);
        Assert.NotSame(source.Options.Metadata.CustomDatas, split.Options.Metadata.CustomDatas);
        Assert.NotSame(source.Versions, split.Versions);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(split, output);
        var written = MultiDocumentFixture.Unzip(output);
        Assert.Equal(new[] { "metadata", "versions" },
            Xml(written["OFD.xml"]).Descendants(vendor + "Keep").Select(element => element.Value));
        Assert.Equal(2, Xml(written["Doc_0/Document.xml"]).Descendants(vendor + "Keep").Count());
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Single(saved.Pages);
        Assert.Equal(OfdDocUsage.ENewsPaper, saved.Options.Metadata.DocUsage);
        Assert.Equal("2.0", saved.Options.Metadata.CreatorVersion);
        Assert.Equal("new", saved.Options.Metadata.CustomDatas!.Items[0].Value);
        Assert.False(saved.Permissions!.Edit);
        Assert.Equal(OfdPageMode.UseBookmarks, saved.ViewingPreferences!.PageMode);
        Assert.Equal(2, saved.Versions!.Items[0].Index);
    }

    [Fact]
    public async Task Merge_CopiesNewScalarMetadataAndExistingCoverPayload()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Element(ns + "DocBody")!.Element(ns + "DocInfo")!.Add(
            new XElement(ns + "DocUsage", "EBook"), new XElement(ns + "Cover", "Doc_0/Cover.png"),
            new XElement(ns + "CreatorVersion", "1.0"));
        entries["OFD.xml"] = Bytes(root); entries["Doc_0/Cover.png"] = [1, 2, 3];
        using var input = MultiDocumentFixture.Zip(entries);
        var source = await new OfdReader().ReadAsync(input);
        var merged = OfdDocumentMerger.Merge([source]);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(merged, output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Equal(OfdDocUsage.EBook, saved.Options.Metadata.DocUsage);
        Assert.Equal("Doc_0/Cover.png", saved.Options.Metadata.Cover);
        Assert.Equal("1.0", saved.Options.Metadata.CreatorVersion);
        Assert.Equal(new byte[] { 1, 2, 3 }, saved.PreservedEntries["Doc_0/Cover.png"]);

        var mixed = OfdDocumentMixer.Mix([new OfdMixSource(source, 0)]);
        using var mixedOutput = new MemoryStream(); await new OfdPackageWriter().WriteAsync(mixed, mixedOutput);
        mixedOutput.Position = 0; var mixedRead = await new OfdReader().ReadAsync(mixedOutput);
        Assert.Equal("Doc_0/Cover.png", mixedRead.Options.Metadata.Cover);
        Assert.Equal(new byte[] { 1, 2, 3 }, mixedRead.PreservedEntries["Doc_0/Cover.png"]);

        source.PreservedEntries.Remove("Doc_0/Cover.png");
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
        source.PreservedEntries["Doc_0/Cover.png"] = Encoding.UTF8.GetBytes("<Document/>");
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
    }

    [Fact]
    public async Task SelectedDocument_StandardObjectsRoundTripWithoutLosingUnknownOrOtherDocument()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        var bodies = root.Root.Elements(ns + "DocBody").ToArray();
        var info = bodies[1].Element(ns + "DocInfo")!;
        info.Add(new XElement(ns + "DocUsage", "EBook"), new XElement(ns + "Cover", "Doc_1/Cover.png"),
            new XElement(ns + "CreatorVersion", "1.0"),
            XElement.Parse($"<CustomDatas xmlns='{ns}' xmlns:v='urn:vendor'><CustomData Name='case'>old</CustomData><v:Keep v:Tag='x'/></CustomDatas>"));
        bodies[1].Add(XElement.Parse($"<Versions xmlns='{ns}' xmlns:v='urn:vendor'><Version ID='v1' Index='1' BaseLoc='Doc_1/Versions/v1.xml'/><v:Keep/></Versions>"));
        entries["OFD.xml"] = Bytes(root);
        entries["Doc_1/Cover.png"] = [1, 2, 3];
        entries["Doc_1/Versions/v1.xml"] = Encoding.UTF8.GetBytes("<VersionFile/>");
        var document = Xml(entries["Doc_1/Document.xml"]);
        document.Root!.Add(XElement.Parse($"<Permissions xmlns='{ns}' xmlns:v='urn:vendor'><Edit>true</Edit><v:Keep v:Flag='x'/><Print Printable='true' Copies='2'/></Permissions>"));
        document.Root.Add(XElement.Parse($"<VPreferences xmlns='{ns}' xmlns:v='urn:vendor'><PageMode>UseOutlines</PageMode><v:Keep/><ZoomMode>FitWidth</ZoomMode></VPreferences>"));
        document.Root.Add(XElement.Parse($"<Actions xmlns='{ns}' xmlns:v='urn:vendor'><Action Event='CLICK' v:Tag='x'><URI URI='https://example.invalid/a' Base='https://example.invalid/'/><v:Keep/></Action></Actions>"));
        entries["Doc_1/Document.xml"] = Bytes(document);
        var beforeOther = entries.Where(entry => entry.Key.StartsWith("Doc_0/", StringComparison.Ordinal)).ToDictionary(entry => entry.Key, entry => entry.Value);
        var beforeBody = new XElement(bodies[0]);

        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Equal(OfdDocUsage.EBook, selected.Options.Metadata.DocUsage);
        Assert.Equal("old", Assert.Single(selected.Options.Metadata.CustomDatas!.Items).Value);
        Assert.True(selected.Permissions!.Edit);
        Assert.Equal(OfdPageMode.UseOutlines, selected.ViewingPreferences!.PageMode);
        Assert.Equal("v1", Assert.Single(selected.Versions!.Items).Id);
        Assert.IsType<OfdUriAction>(Assert.Single(selected.Actions!.Items));

        selected.Options.Metadata.DocUsage = OfdDocUsage.ENewsPaper;
        selected.Options.Metadata.CreatorVersion = "2.0";
        selected.Options.Metadata.CustomDatas.Items[0].Value = "new";
        selected.Permissions.Edit = false;
        selected.Permissions.Print!.Copies = 3;
        selected.ViewingPreferences.PageMode = OfdPageMode.UseBookmarks;
        selected.ViewingPreferences.Zoom = 1.25;
        selected.Versions.Items[0].Index = 2;
        Assert.IsType<OfdUriAction>(selected.Actions.Items[0]).Uri = "https://example.invalid/b";

        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(selected, output);
        var written = MultiDocumentFixture.Unzip(output);
        Assert.All(beforeOther, pair => Assert.Equal(pair.Value, written[pair.Key]));
        var resultRoot = Xml(written["OFD.xml"]);
        Assert.True(XNode.DeepEquals(beforeBody, resultRoot.Root!.Elements(ns + "DocBody").First()));
        var resultInfo = resultRoot.Root.Elements(ns + "DocBody").Last().Element(ns + "DocInfo")!;
        Assert.Equal("ENewsPaper", resultInfo.Element(ns + "DocUsage")!.Value);
        Assert.Equal("new", resultInfo.Element(ns + "CustomDatas")!.Element(ns + "CustomData")!.Value);
        Assert.NotNull(resultInfo.Element(ns + "CustomDatas")!.Element(XName.Get("Keep", "urn:vendor")));
        var resultDoc = Xml(written["Doc_1/Document.xml"]);
        Assert.Equal("false", resultDoc.Descendants(ns + "Edit").Single().Value);
        Assert.Equal("3", resultDoc.Descendants(ns + "Print").Single().Attribute("Copies")!.Value);
        Assert.Equal("1.25", resultDoc.Descendants(ns + "Zoom").Single().Value);
        Assert.NotNull(resultDoc.Descendants(XName.Get("Keep", "urn:vendor")).FirstOrDefault());
        Assert.Equal("https://example.invalid/b", resultDoc.Descendants(ns + "URI").Single().Attribute("URI")!.Value);

        using var reread = MultiDocumentFixture.Zip(written);
        var after = await new OfdReader().ReadAsync(reread, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Equal(OfdDocUsage.ENewsPaper, after.Options.Metadata.DocUsage);
        Assert.Equal(2, after.Versions!.Items[0].Index);
        Assert.Equal("https://example.invalid/b", Assert.IsType<OfdUriAction>(after.Actions!.Items[0]).Uri);
    }

    [Fact]
    public async Task NewStandardObjects_WriteReadAndFailClosedOnUnownedPaths()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Permissions = new OfdPermissions { Edit = false, Print = new OfdPrintPermission { Printable = true, Copies = 2 } };
        package.ViewingPreferences = new OfdViewingPreferences { PageLayout = OfdPageLayout.TwoPageL, ZoomMode = OfdZoomMode.FitWidth };
        package.Options.Metadata.DocUsage = OfdDocUsage.EBook;
        package.Options.Metadata.CustomDatas = new OfdCustomDatas(); package.Options.Metadata.CustomDatas.Add("purpose", "typed");
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdUriAction { Event = "CLICK", Uri = "https://example.invalid" });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0;
        var loaded = await new OfdReader().ReadAsync(output);
        Assert.False(loaded.Permissions!.Edit);
        Assert.Equal(OfdPageLayout.TwoPageL, loaded.ViewingPreferences!.PageLayout);
        Assert.Equal("typed", loaded.Options.Metadata.CustomDatas!.Items[0].Value);
        Assert.IsType<OfdUriAction>(loaded.Actions!.Items[0]);

        loaded.Permissions.Edit = true;
        Assert.Throws<InvalidOperationException>(() => loaded.Permissions = null);
        var foreign = new OfdDocumentPackage();
        foreign.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        foreign.Permissions = loaded.Permissions;
        using var foreignOutput = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(foreign, foreignOutput));
        Assert.Equal(0, foreignOutput.Length);

        loaded.Options.Metadata.Cover = "Doc_0/missing.png";
        using var rejected = new MemoryStream();
        var missingCover = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(loaded, rejected));
        Assert.Contains("Cover", missingCover.Message);
        Assert.Equal(0, rejected.Length);
    }

    [Fact]
    public async Task MissingLocalActionTargets_RemainOpaqueAndRoundTrip()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var document = Xml(entries["Doc_0/Document.xml"]); var ns = document.Root!.Name.Namespace;
        document.Root.Add(XElement.Parse($"<Actions xmlns='{ns}'><Action Event='CLICK'><Sound ResourceID='999' Volume='50'/></Action></Actions>"));
        document.Root.Element(ns + "Actions")!.Add(XElement.Parse($"<Action xmlns='{ns}' Event='CLICK'><GotoA AttachID='999'/></Action>"));
        entries["Doc_0/Document.xml"] = Bytes(document);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input);
        Assert.All(selected.Actions!.Items, item => Assert.IsType<OfdRawAction>(item));
        Assert.Equal(2, selected.Actions.Items.Count);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        Assert.Equal("999", Xml(MultiDocumentFixture.Unzip(output)["Doc_0/Document.xml"]).Descendants(ns + "Sound").Single().Attribute("ResourceID")!.Value);
        Assert.Equal("999", Xml(MultiDocumentFixture.Unzip(output)["Doc_0/Document.xml"]).Descendants(ns + "GotoA").Single().Attribute("AttachID")!.Value);
    }

    [Fact]
    public async Task AttachmentAction_BindsToRetainedAttachmentAndRejectsDanglingId()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Attachments.Add(new OfdAttachment { Id = "51", Name = "memo.txt", Data = Encoding.UTF8.GetBytes("memo") });
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdAttachmentAction { Event = "CLICK", AttachId = "51", NewWindow = false });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        output.Position = 0; var loaded = await new OfdReader().ReadAsync(output);
        Assert.Equal("51", Assert.IsType<OfdAttachmentAction>(loaded.Actions!.Items[0]).AttachId);
        Assert.False(Assert.IsType<OfdAttachmentAction>(loaded.Actions.Items[0]).NewWindow);
        Assert.IsType<OfdAttachmentAction>(loaded.Actions.Items[0]).AttachId = "999";
        using var rejected = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(loaded, rejected));
        Assert.Equal(0, rejected.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AttachmentAction_WithExternalOrMissingPayload_RemainsOpaque(bool missingPayload)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Attachments.Add(new OfdAttachment { Id = "51", Name = "memo.txt", Data = Encoding.UTF8.GetBytes("memo") });
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdAttachmentAction { Event = "CLICK", AttachId = "51" });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var descriptorPath = entries.Keys.Single(path => path.EndsWith("/Attachments.xml", StringComparison.Ordinal));
        if (missingPayload)
        {
            var payloadPath = entries.Keys.Single(path => path.StartsWith("Doc_0/Attachs/Attach_", StringComparison.Ordinal));
            entries.Remove(payloadPath);
        }
        else
        {
            var descriptor = Xml(entries[descriptorPath]);
            descriptor.Root!.Elements().Single().SetAttributeValue("External", "true");
            entries[descriptorPath] = Bytes(descriptor);
        }
        using var input = MultiDocumentFixture.Zip(entries);
        var imported = await new OfdReader().ReadAsync(input);
        Assert.True(Assert.Single(imported.Attachments).IsExternal);
        Assert.IsType<OfdRawAction>(Assert.Single(imported.Actions!.Items));
        Assert.Throws<NotSupportedException>(() => OfdDocumentEditor.RemovePages(imported, [1]));
        Assert.Equal(2, imported.Pages.Count);
    }

    [Fact]
    public async Task AttachmentAction_RejectsNewExternalTarget()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Attachments.Add(new OfdAttachment { Id = "51", Name = "memo.txt", IsExternal = true, ExternalPath = "memo.txt" });
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdAttachmentAction { Event = "CLICK", AttachId = "51" });
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(package, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("Doc_0/Document.xml")]
    [InlineData("Doc_0/PublicRes.xml")]
    [InlineData("Doc_0/Cover.png")]
    public async Task ChangedCover_RejectsControlXmlPayload(string location)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        seed.Position = 0; var imported = await new OfdReader().ReadAsync(seed);
        imported.PreservedEntries["Doc_0/Cover.png"] = Encoding.UTF8.GetBytes("<Document/>");
        imported.Options.Metadata.Cover = location;
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(imported, output));
        Assert.Contains("Cover", error.Message);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task InsertingStandardFields_IgnoresSameNamespaceExtensionAnchors()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        var info = root.Root.Element(ns + "DocBody")!.Element(ns + "DocInfo")!;
        var docId = info.Element(ns + "DocID")?.Value ?? "doc";
        info.ReplaceNodes(new XElement(ns + "DocID", docId), new XElement(ns + "Title", "title"),
            new XElement(ns + "Extra", "keep"), new XElement(ns + "Keywords", "key"));
        entries["OFD.xml"] = Bytes(root);
        var document = Xml(entries["Doc_0/Document.xml"]);
        document.Root!.Add(new XElement(ns + "Permissions", new XElement(ns + "Extra", "keep"), new XElement(ns + "Edit", "true")));
        document.Root.Add(new XElement(ns + "VPreferences", new XElement(ns + "Extra", "keep"), new XElement(ns + "PageMode", "UseOutlines")));
        entries["Doc_0/Document.xml"] = Bytes(document);
        entries["Doc_0/Cover.png"] = [1, 2, 3];
        using var input = MultiDocumentFixture.Zip(entries);
        var imported = await new OfdReader().ReadAsync(input);
        imported.Options.Metadata.Cover = "Doc_0/Cover.png";
        imported.Options.Metadata.CreatorVersion = "1.0";
        imported.Options.Metadata.CustomDatas = new OfdCustomDatas();
        imported.Options.Metadata.CustomDatas.Add("case", "value");
        imported.Permissions!.ValidPeriod = new OfdValidPeriod { StartDate = "2026-10-21T00:00:00Z" };
        imported.ViewingPreferences!.PageLayout = OfdPageLayout.OnePage;
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, output);
        var written = MultiDocumentFixture.Unzip(output);
        var savedInfo = Xml(written["OFD.xml"]).Descendants(ns + "DocInfo").Single();
        Assert.Equal(new[] { "DocID", "Title", "Extra", "Cover", "Keywords", "CreatorVersion", "CustomDatas" },
            savedInfo.Elements().Select(child => child.Name.LocalName));
        var savedDocument = Xml(written["Doc_0/Document.xml"]);
        Assert.Equal(new[] { "Extra", "Edit", "ValidPeriod" },
            savedDocument.Descendants(ns + "Permissions").Single().Elements().Select(child => child.Name.LocalName));
        Assert.Equal(new[] { "Extra", "PageMode", "PageLayout" },
            savedDocument.Descendants(ns + "VPreferences").Single().Elements().Select(child => child.Name.LocalName));
    }

    [Fact]
    public void StandardObjectClone_ReservesLargeXmlBeforeAllocatingCopy()
    {
        var data = new OfdCustomDatas();
        data.Add("large", new string('x', 256_000));
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1_000_000 });
        context.Material.ChargeCopy(1_000_000 - 1024, default);
        var allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        Assert.Throws<InvalidDataException>(() => OfdStandardXml.Clone(data, new OfdDocumentPackage(), context.Material, default));
        var allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        Assert.True(allocated < 64_000, $"Rejected XML clone allocated {allocated} bytes.");
        Assert.Equal(1_000_000 - 1024, context.Material.CloneCopyBytes);
        Assert.Equal(256_000, data.Items[0].Value.Length);
    }

    [Fact]
    public void ControlledStandardRead_BindsCapturedBaselinesWithoutReparsing()
    {
        var ns = XNamespace.Get("http://www.ofdspec.org/2016");
        var vendor = XNamespace.Get("urn:standard-read-vendor");
        var custom = new XElement(ns + "CustomDatas", new XElement(ns + "CustomData",
            new XAttribute("Name", "large"), new string('x', 20_000)),
            new XElement(vendor + "Keep", new XAttribute("Kind", "v:Type")));
        var info = new XElement(ns + "DocInfo", new XAttribute(XNamespace.Xmlns + "v", vendor.NamespaceName), custom);
        var permissions = new XElement(ns + "Permissions", new XElement(ns + "Edit", "true"));
        var preferences = new XElement(ns + "VPreferences", new XElement(ns + "PageMode", "UseOutlines"));
        var versions = new XElement(ns + "Versions", new XElement(ns + "Version", new XAttribute("ID", "v1"),
            new XAttribute("Index", "1"), new XAttribute("BaseLoc", "Doc_0/Version.xml")));
        var body = new XElement(ns + "DocBody", versions);
        var document = new XElement(ns + "Document", permissions, preferences);
        var package = new OfdDocumentPackage();
        using var context = new OfdControlledEditContext(new());
        using var scope = context.Material.BeginControlledRead(context, default);
        OfdStandardXml.ReadMetadata(package, info, context.Material, default);
        OfdStandardXml.ReadDocument(package, body, document, context.Material, default);
        Assert.Equal(0, context.Material.ReadParserAdmissions);
        Assert.Equal(4, context.Material.ReadSourceClones);
        Assert.NotSame(custom, package.Options.Metadata.CustomDatas!.Root);
        Assert.NotSame(permissions, package.Permissions!.Root);
        Assert.NotSame(preferences, package.ViewingPreferences!.Root);
        Assert.NotSame(versions, package.Versions!.Root);
        Assert.Contains(new string('x', 20_000), package.Options.Metadata.CustomDatas.OriginalXml);
        Assert.Equal(vendor.NamespaceName, package.Options.Metadata.CustomDatas.Root.Attribute(XNamespace.Xmlns + "v")?.Value);
    }

    [Fact]
    public async Task MultiDocument_RejectsChangedRootReferencesIntoOtherDocument()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        selected.Options.Metadata.Cover = "Doc_0/Document.xml";
        using var coverOutput = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, coverOutput));
        Assert.Equal(0, coverOutput.Length);

        selected.Options.Metadata.Cover = null;
        selected.Versions = new OfdVersions();
        selected.Versions.Add("v1", 1, "Doc_0/Document.xml");
        using var versionOutput = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, versionOutput));
        Assert.Equal(0, versionOutput.Length);
    }

    [Fact]
    public async Task UnchangedMissingVersionLocation_AllowsIndexEditButChangedLocationFails()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Elements(ns + "DocBody").Last().Add(XElement.Parse($"<Versions xmlns='{ns}'><Version ID='v1' Index='1' BaseLoc='Doc_1/missing.xml'/></Versions>"));
        entries["OFD.xml"] = Bytes(root);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        selected.Versions!.Items[0].Index = 2;
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        output.Position = 0;
        var saved = await new OfdReader().ReadAsync(output, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Equal(2, saved.Versions!.Items[0].Index);
        Assert.Equal("Doc_1/missing.xml", saved.Versions.Items[0].BaseLocation);

        selected.Versions.Items[0].Id = "v2";
        using var renamed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, renamed);
        renamed.Position = 0;
        var renamedVersion = await new OfdReader().ReadAsync(renamed, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Equal("v2", renamedVersion.Versions!.Items[0].Id);
        Assert.Equal("Doc_1/missing.xml", renamedVersion.Versions.Items[0].BaseLocation);

        selected.Versions.Items[0].BaseLocation = "Doc_1/also-missing.xml";
        using var rejected = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, rejected));
        Assert.Equal(0, rejected.Length);
    }

    [Fact]
    public async Task ClearingVersions_OmitsEmptyWrapperAndRetainsExtensionChildren()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Elements(ns + "DocBody").Last().Add(XElement.Parse($"<Versions xmlns='{ns}' xmlns:v='urn:vendor'><Version ID='v1' Index='1' BaseLoc='Doc_1/missing.xml'/><v:Keep/></Versions>"));
        entries["OFD.xml"] = Bytes(root);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        selected.Versions!.Clear();
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        var body = Xml(MultiDocumentFixture.Unzip(output)["OFD.xml"]).Root!.Elements(ns + "DocBody").Last();
        Assert.Empty(body.Descendants(ns + "Version"));
        Assert.NotNull(body.Element(ns + "Versions")!.Element(XName.Get("Keep", "urn:vendor")));

        var clean = new OfdDocumentPackage(); clean.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        clean.Versions = new OfdVersions(); clean.Versions.Add("v1", 1, "missing.xml"); clean.Versions.Clear();
        using var cleanOutput = new MemoryStream(); await new OfdPackageWriter().WriteAsync(clean, cleanOutput);
        Assert.Empty(Xml(MultiDocumentFixture.Unzip(cleanOutput)["OFD.xml"]).Descendants(ns + "Versions"));
    }

    [Fact]
    public async Task InvalidImportedVersionId_RemainsPreservedXmlOnUnchangedSave()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Elements(ns + "DocBody").Last().Add(XElement.Parse($"<Versions xmlns='{ns}'><Version ID='1' Index='1' BaseLoc='Doc_1/missing.xml'/></Versions>"));
        entries["OFD.xml"] = Bytes(root);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Null(selected.Versions);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        var saved = Xml(MultiDocumentFixture.Unzip(output)["OFD.xml"]);
        Assert.Equal("1", saved.Root!.Elements(ns + "DocBody").Last().Descendants(ns + "Version").Single().Attribute("ID")!.Value);
    }

    [Fact]
    public async Task DuplicateImportedVersionIds_RemainPreservedXmlOnUnchangedSave()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Elements(ns + "DocBody").Last().Add(XElement.Parse($"<Versions xmlns='{ns}'><Version ID='v1' Index='1' BaseLoc='Doc_1/a.xml'/><Version ID='v1' Index='2' BaseLoc='Doc_1/b.xml'/></Versions>"));
        entries["OFD.xml"] = Bytes(root);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Null(selected.Versions);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        var saved = Xml(MultiDocumentFixture.Unzip(output)["OFD.xml"]);
        Assert.Equal(2, saved.Root!.Elements(ns + "DocBody").Last().Descendants(ns + "Version").Count());
    }

    [Fact]
    public async Task EmptyImportedCustomDataName_RemainsPreservedXmlOnUnchangedSave()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Elements(ns + "DocBody").Last().Element(ns + "DocInfo")!.Add(
            XElement.Parse($"<CustomDatas xmlns='{ns}'><CustomData Name=''>value</CustomData></CustomDatas>"));
        entries["OFD.xml"] = Bytes(root);
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions { DocumentIndex = 1 });
        Assert.Null(selected.Options.Metadata.CustomDatas);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        var saved = Xml(MultiDocumentFixture.Unzip(output)["OFD.xml"]);
        Assert.Equal("", saved.Root!.Elements(ns + "DocBody").Last().Descendants(ns + "CustomData").Single().Attribute("Name")!.Value);
    }

    [Theory]
    [InlineData("1", "1", "1")]
    [InlineData("ebook", "fullscreen", "play")]
    [InlineData(" EBook ", " FullScreen ", " Play ")]
    public async Task NonCanonicalEnumLexemes_RemainOpaqueAndUnchanged(string usage, string mode, string operation)
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        root.Root.Element(ns + "DocBody")!.Element(ns + "DocInfo")!.Add(new XElement(ns + "DocUsage", usage));
        entries["OFD.xml"] = Bytes(root);
        var document = Xml(entries["Doc_0/Document.xml"]);
        document.Root!.Add(XElement.Parse($"<VPreferences xmlns='{ns}'><PageMode>{mode}</PageMode></VPreferences>"));
        document.Root.Add(XElement.Parse($"<Actions xmlns='{ns}'><Action Event='CLICK'><Movie ResourceID='901' Operator='{operation}'/></Action></Actions>"));
        entries["Doc_0/Document.xml"] = Bytes(document);
        var resource = Xml(entries["Doc_0/PublicRes.xml"]);
        resource.Root!.Add(new XElement(ns + "MultiMedias", new XElement(ns + "MultiMedia",
            new XAttribute("ID", "901"), new XAttribute("Type", "Video"), new XElement(ns + "MediaFile", "video.mp4"))));
        entries["Doc_0/PublicRes.xml"] = Bytes(resource);
        entries["Doc_0/Res/video.mp4"] = [1, 2, 3];
        using var input = MultiDocumentFixture.Zip(entries);
        var loaded = await new OfdReader().ReadAsync(input);
        Assert.Null(loaded.Options.Metadata.DocUsage);
        Assert.Null(loaded.ViewingPreferences);
        Assert.IsType<OfdRawAction>(Assert.Single(loaded.Actions!.Items));
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, output);
        var written = MultiDocumentFixture.Unzip(output);
        Assert.Equal(usage, Xml(written["OFD.xml"]).Descendants(ns + "DocUsage").Single().Value);
        var saved = Xml(written["Doc_0/Document.xml"]);
        Assert.Equal(mode, saved.Descendants(ns + "PageMode").Single().Value);
        Assert.Equal(operation, saved.Descendants(ns + "Movie").Single().Attribute("Operator")!.Value);
    }

    [Fact]
    public async Task ImportedUriAction_DoesNotBlockUnrelatedPageDeletion()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdUriAction { Event = "CLICK", Uri = "https://example.invalid" });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        seed.Position = 0; var imported = await new OfdReader().ReadAsync(seed);
        Assert.IsType<OfdUriAction>(Assert.Single(imported.Actions!.Items));
        OfdDocumentEditor.RemovePages(imported, [1]);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Single(saved.Pages);
        Assert.IsType<OfdUriAction>(Assert.Single(saved.Actions!.Items));
    }

    [Fact]
    public async Task NestedPermissionFromAnotherDocumentCannotBeAssigned()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Permissions = new OfdPermissions
        {
            Print = new OfdPrintPermission { Printable = true },
            ValidPeriod = new OfdValidPeriod { StartDate = "2026-10-21T00:00:00Z" }
        };
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var bytes = seed.ToArray();
        using var first = new MemoryStream(bytes); using var second = new MemoryStream(bytes);
        var source = await new OfdReader().ReadAsync(first);
        var target = await new OfdReader().ReadAsync(second);
        Assert.Throws<NotSupportedException>(() => target.Permissions!.Print = source.Permissions!.Print);
        Assert.Throws<NotSupportedException>(() => target.Permissions!.ValidPeriod = source.Permissions!.ValidPeriod);
    }

    [Fact]
    public async Task NestedPermissionReplacement_RejectsLossOfUnknownSourceContent()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Permissions = new OfdPermissions
        {
            Print = new OfdPrintPermission { Printable = true, Copies = 2 },
            ValidPeriod = new OfdValidPeriod { StartDate = "2026-10-21T00:00:00Z" }
        };
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var document = Xml(entries["Doc_0/Document.xml"]); var ns = document.Root!.Name.Namespace;
        var vendor = XNamespace.Get("urn:permission-vendor");
        document.Descendants(ns + "Print").Single().SetAttributeValue(vendor + "Flag", "x");
        document.Descendants(ns + "ValidPeriod").Single().Add(new XElement(vendor + "Keep", "period"));
        entries["Doc_0/Document.xml"] = Bytes(document);
        using var input = MultiDocumentFixture.Zip(entries);
        var imported = await new OfdReader().ReadAsync(input);
        Assert.NotNull(imported.Permissions);
        Assert.Throws<NotSupportedException>(() => imported.Permissions!.Print = null);
        Assert.Throws<NotSupportedException>(() => imported.Permissions!.ValidPeriod = null);
        Assert.Throws<NotSupportedException>(() => imported.Permissions!.Print = new OfdPrintPermission { Printable = false });
        Assert.Throws<NotSupportedException>(() => imported.Permissions!.ValidPeriod = new OfdValidPeriod());
        imported.Permissions!.Print!.Copies = 3;
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, output);
        var written = Xml(MultiDocumentFixture.Unzip(output)["Doc_0/Document.xml"]);
        Assert.Equal("3", written.Descendants(ns + "Print").Single().Attribute("Copies")!.Value);
        Assert.Equal("x", written.Descendants(ns + "Print").Single().Attribute(vendor + "Flag")!.Value);
        Assert.Equal("period", written.Descendants(vendor + "Keep").Single().Value);
    }

    [Fact]
    public async Task RemovingStandardEntries_RejectsUnknownContentAtomically()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Options.Metadata.CustomDatas = new OfdCustomDatas();
        package.Options.Metadata.CustomDatas.Add("plain", "one");
        package.Options.Metadata.CustomDatas.Add("opaque", "two");
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var root = Xml(entries["OFD.xml"]); var ns = root.Root!.Name.Namespace;
        var vendor = XNamespace.Get("urn:entry-vendor");
        var body = root.Root.Element(ns + "DocBody")!;
        var custom = body.Element(ns + "DocInfo")!.Element(ns + "CustomDatas")!;
        custom.Elements(ns + "CustomData").Last().SetAttributeValue(vendor + "Keep", "custom");
        body.Add(new XElement(ns + "Versions",
            new XElement(ns + "Version", new XAttribute("ID", "v1"), new XAttribute("Index", "1"), new XAttribute("BaseLoc", "Doc_0/Versions/v1.xml")),
            new XElement(ns + "Version", new XAttribute("ID", "v2"), new XAttribute("Index", "2"), new XAttribute("BaseLoc", "Doc_0/Versions/v2.xml"),
                new XElement(vendor + "Keep", "version"))));
        entries["OFD.xml"] = Bytes(root);
        entries["Doc_0/Versions/v1.xml"] = Encoding.UTF8.GetBytes("<VersionFile/>");
        entries["Doc_0/Versions/v2.xml"] = Encoding.UTF8.GetBytes("<VersionFile/>");
        var originalRoot = entries["OFD.xml"].ToArray();
        using var input = MultiDocumentFixture.Zip(entries);
        var imported = await new OfdReader().ReadAsync(input);
        var data = imported.Options.Metadata.CustomDatas!; var versions = imported.Versions!;
        var dataXml = data.InventoryXml; var versionsXml = versions.InventoryXml;
        Assert.Throws<NotSupportedException>(() => data.RemoveAt(1));
        Assert.Throws<NotSupportedException>(() => data.Clear());
        Assert.Throws<NotSupportedException>(() => versions.RemoveAt(1));
        Assert.Throws<NotSupportedException>(() => versions.Clear());
        Assert.Equal(dataXml, data.InventoryXml);
        Assert.Equal(versionsXml, versions.InventoryXml);
        Assert.Equal(originalRoot, imported.PreservedEntries["OFD.xml"]);
        using var unchanged = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, unchanged);
        var unchangedRoot = Xml(MultiDocumentFixture.Unzip(unchanged)["OFD.xml"]);
        var writtenCustom = unchangedRoot.Descendants(ns + "CustomDatas").Single();
        Assert.Equal(new[] { "plain", "opaque" }, writtenCustom.Elements(ns + "CustomData").Select(item => item.Attribute("Name")!.Value));
        Assert.Equal("custom", writtenCustom.Elements(ns + "CustomData").Last().Attribute(vendor + "Keep")!.Value);
        var writtenVersions = unchangedRoot.Descendants(ns + "Versions").Single();
        Assert.Equal(new[] { "v1", "v2" }, writtenVersions.Elements(ns + "Version").Select(item => item.Attribute("ID")!.Value));
        Assert.Equal("version", writtenVersions.Descendants(vendor + "Keep").Single().Value);

        data.RemoveAt(0); versions.RemoveAt(0);
        using var edited = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, edited);
        edited.Position = 0; var saved = await new OfdReader().ReadAsync(edited);
        Assert.Equal("opaque", Assert.Single(saved.Options.Metadata.CustomDatas!.Items).Name);
        Assert.Equal("v2", Assert.Single(saved.Versions!.Items).Id);
    }

    [Fact]
    public async Task MediaAndAttachmentActions_RequireRetainedLocalTargets()
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        var resource = Xml(entries["Doc_0/PublicRes.xml"]); var ns = resource.Root!.Name.Namespace;
        resource.Root.Add(new XElement(ns + "MultiMedias",
            new XElement(ns + "MultiMedia", new XAttribute("ID", "901"), new XAttribute("Type", "Audio"), new XAttribute("Format", "WAV"), new XElement(ns + "MediaFile", "audio.wav")),
            new XElement(ns + "MultiMedia", new XAttribute("ID", "902"), new XAttribute("Type", "Video"), new XAttribute("Format", "MP4"), new XElement(ns + "MediaFile", "video.mp4"))));
        entries["Doc_0/PublicRes.xml"] = Bytes(resource);
        entries["Doc_0/Res/audio.wav"] = [1, 2, 3]; entries["Doc_0/Res/video.mp4"] = [4, 5, 6];
        using var input = MultiDocumentFixture.Zip(entries);
        var selected = await new OfdReader().ReadAsync(input);
        selected.Actions = new OfdActionCollection();
        selected.Actions.Items.Add(new OfdSoundAction { Event = "CLICK", ResourceId = "901", Volume = 70 });
        selected.Actions.Items.Add(new OfdMovieAction { Event = "CLICK", ResourceId = "902", Operator = OfdMovieOperator.Play });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(selected, output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.IsType<OfdSoundAction>(saved.Actions!.Items[0]);
        Assert.IsType<OfdMovieAction>(saved.Actions.Items[1]);
        Assert.Equal("901", Assert.IsType<OfdSoundAction>(saved.Actions.Items[0]).ResourceId);

        selected.Actions.Items.Add(new OfdSoundAction { Event = "CLICK", ResourceId = "999" });
        using var rejected = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(selected, rejected));
        Assert.Equal(0, rejected.Length);
    }

    [Fact]
    public async Task PageDeletion_PreservesLocalActionsOnlyWithClosedTargetsAndKnownXml()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
        package.Attachments.Add(new OfdAttachment { Id = "51", Name = "memo.txt", Data = [1, 2, 3] });
        package.Actions = new OfdActionCollection();
        package.Actions.Items.Add(new OfdAttachmentAction { Event = "CLICK", AttachId = "51" });
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        var entries = MultiDocumentFixture.Unzip(seed);
        OfdDocumentEditor.RemovePages(package, [1]);
        Assert.Single(package.Pages);
        var resource = Xml(entries["Doc_0/PublicRes.xml"]); var ns = resource.Root!.Name.Namespace;
        resource.Root.Add(new XElement(ns + "MultiMedias",
            new XElement(ns + "MultiMedia", new XAttribute("ID", "901"), new XAttribute("Type", "Audio"), new XElement(ns + "MediaFile", "audio.wav")),
            new XElement(ns + "MultiMedia", new XAttribute("ID", "902"), new XAttribute("Type", "Video"), new XElement(ns + "MediaFile", "video.mp4"))));
        entries["Doc_0/PublicRes.xml"] = Bytes(resource);
        entries["Doc_0/Res/audio.wav"] = [4, 5, 6]; entries["Doc_0/Res/video.mp4"] = [7, 8, 9];
        var document = Xml(entries["Doc_0/Document.xml"]);
        var actions = document.Root!.Element(ns + "Actions")!;
        actions.Add(new XElement(ns + "Action", new XAttribute("Event", "CLICK"),
            new XElement(ns + "Sound", new XAttribute("ResourceID", "901"), new XAttribute("Volume", "50"))));
        actions.Add(new XElement(ns + "Action", new XAttribute("Event", "CLICK"),
            new XElement(ns + "Movie", new XAttribute("ResourceID", "902"), new XAttribute("Operator", "Play"))));
        entries["Doc_0/Document.xml"] = Bytes(document);

        using var input = MultiDocumentFixture.Zip(entries);
        var source = await new OfdReader().ReadAsync(input);
        Assert.IsType<OfdAttachmentAction>(source.Actions!.Items[0]);
        Assert.IsType<OfdSoundAction>(source.Actions.Items[1]);
        Assert.IsType<OfdMovieAction>(source.Actions.Items[2]);
        source.DocumentResourceLocation = source.PublicResourceLocation;
        OfdDocumentEditor.RemovePages(source, [1]);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Single(saved.Pages);
        Assert.IsType<OfdAttachmentAction>(saved.Actions!.Items[0]);
        Assert.IsType<OfdSoundAction>(saved.Actions.Items[1]);
        Assert.IsType<OfdMovieAction>(saved.Actions.Items[2]);
        Assert.Equal(new byte[] { 4, 5, 6 }, saved.PreservedEntries["Doc_0/Res/audio.wav"]);
        Assert.Equal(new byte[] { 7, 8, 9 }, saved.PreservedEntries["Doc_0/Res/video.mp4"]);

        using var opaqueInput = MultiDocumentFixture.Zip(entries);
        var opaqueEntries = MultiDocumentFixture.Unzip(opaqueInput);
        var opaqueDocument = Xml(opaqueEntries["Doc_0/Document.xml"]);
        opaqueDocument.Descendants(ns + "Sound").Single().SetAttributeValue(XName.Get("Keep", "urn:vendor"), "x");
        opaqueEntries["Doc_0/Document.xml"] = Bytes(opaqueDocument);
        using var opaqueZip = MultiDocumentFixture.Zip(opaqueEntries);
        var opaque = await new OfdReader().ReadAsync(opaqueZip);
        Assert.IsType<OfdSoundAction>(opaque.Actions!.Items[1]);
        Assert.Throws<NotSupportedException>(() => OfdDocumentEditor.RemovePages(opaque, [1]));

        using var missingInput = MultiDocumentFixture.Zip(entries);
        var missing = await new OfdReader().ReadAsync(missingInput);
        missing.Attachments.Clear();
        Assert.Throws<NotSupportedException>(() => OfdDocumentEditor.RemovePages(missing, [1]));
    }

    private static XDocument Xml(byte[] bytes) => XDocument.Load(new MemoryStream(bytes));
    private static byte[] Bytes(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
}
