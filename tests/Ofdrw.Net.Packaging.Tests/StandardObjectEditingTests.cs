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

    private static XDocument Xml(byte[] bytes) => XDocument.Load(new MemoryStream(bytes));
    private static byte[] Bytes(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
}
