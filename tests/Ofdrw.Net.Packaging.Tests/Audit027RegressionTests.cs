using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit027RegressionTests
{
    private static readonly XNamespace Ns = "http://www.ofdspec.org/2016";
    private static readonly byte[] Payload = Enumerable.Range(0, 256).Select(i => (byte)i).Concat(Encoding.UTF8.GetBytes("\0附件 original payload\r\n")).ToArray();

    [Theory]
    [InlineData("", "Attachs/Attachments.xml")]
    [InlineData("Nested/Doc/", "Metadata/List.xml")]
    public async Task Audit027_NoOpPreservesActiveDescriptorAndDeclaration(string prefix, string location)
    {
        var input = Fixture(prefix, location);
        var output = await Write(await Read(input));
        var document = Xml(output, prefix + "Document.xml");
        Assert.Equal(location, document.Root!.Element(Ns + "Attachments")!.Value);
        AssertXmlEqual(Xml(input, prefix + location).Root!, Xml(output, prefix + location).Root!);
        AssertXmlEqual(Xml(input, prefix + "Document.xml").Root!.Element(Ns + "Attachments")!, document.Root.Element(Ns + "Attachments")!);
        Assert.Equal(Payload, Entry(output, prefix + "Payload/original.bin"));
        var readback = await Read(output);
        Assert.Equal(Payload, readback.Attachments[0].Data);
        Assert.True(readback.Attachments[1].IsExternal);
        Assert.Empty(readback.Attachments[1].Data);
        SaveArtifacts(prefix, location, input, output);
    }

    [Theory]
    [InlineData("Name")]
    [InlineData("ID")]
    [InlineData("Format")]
    [InlineData("MediaType")]
    [InlineData("CreationDate")]
    [InlineData("ModDate")]
    [InlineData("Size")]
    [InlineData("Visible")]
    [InlineData("Usage")]
    public async Task Audit027_FieldEditOwnsOnlyItsAttribute(string field)
    {
        var input = Fixture(); var package = await Read(input); var attachment = package.Attachments[0];
        switch (field)
        {
            case "Name": attachment.Name = "更名 & changed.bin"; break;
            case "ID": attachment.Id = "97"; break;
            case "Format": attachment.Format = "application/x-edited"; break;
            case "MediaType": attachment.MediaType = "application/x-edited"; break;
            case "CreationDate": attachment.CreationDate = null; break;
            case "ModDate": attachment.ModificationDate = new DateTimeOffset(2026, 10, 16, 2, 3, 4, TimeSpan.FromHours(8)); break;
            case "Size": attachment.SizeKilobytes = 1.234567; break;
            case "Visible": attachment.Visible = true; break;
            case "Usage": attachment.Usage = null; break;
        }
        var output = await Write(package);
        var expected = Xml(input, "Metadata/List.xml").Root!;
        var actual = Xml(output, "Metadata/List.xml").Root!;
        var changed = actual.Element(Ns + "Attachment")!.Attribute(field);
        Assert.NotEqual(expected.Element(Ns + "Attachment")!.Attribute(field)?.Value, changed?.Value);
        expected.Element(Ns + "Attachment")!.SetAttributeValue(field, changed?.Value);
        AssertXmlEqual(expected, actual);
        AssertXmlEqual(Xml(input, "Document.xml").Root!.Element(Ns + "Attachments")!, Xml(output, "Document.xml").Root!.Element(Ns + "Attachments")!);
        Assert.Equal(Payload, (await Read(output)).Attachments[0].Data);
        Assert.Equal("../Payload/original.bin", actual.Element(Ns + "Attachment")!.Element(Ns + "FileLoc")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit027_ExplicitRemovalDoesNotReviveSourceNodes(bool all)
    {
        var input = Fixture(); var package = await Read(input);
        if (all) package.Attachments.Clear(); else package.Attachments.RemoveAt(0);
        var output = await Write(package); var readback = await Read(output);
        Assert.Equal(all ? 0 : 1, readback.Attachments.Count);
        if (all) Assert.Null(Xml(output, "Document.xml").Root!.Element(Ns + "Attachments"));
        else
        {
            var expected = Xml(input, "Metadata/List.xml").Root!; expected.Elements(Ns + "Attachment").First().Remove();
            AssertXmlEqual(expected, Xml(output, "Metadata/List.xml").Root!);
            Assert.Equal("8", Assert.Single(readback.Attachments).Id);
        }
        Assert.Equal(Payload, Entry(output, "Payload/original.bin")); // Logical removal retains orphaned ZIP entries.
        Assert.Equal(readback.Attachments.Count, (await Read(await Write(readback))).Attachments.Count);
    }

    [Fact]
    public async Task Audit027_NewObjectWithSameIdAndNameCannotBorrowDeletedOrigin()
    {
        var input = Fixture(); var package = await Read(input); var removed = package.Attachments[0]; package.Attachments.RemoveAt(0);
        package.Attachments.Insert(0, new OfdAttachment { Id = removed.Id, Name = removed.Name, Data = [4, 5, 6] });
        // Occupied conventional names must not be overwritten by a new payload.
        package.PreservedEntries["Metadata/Attach_1_original.bin"] = [9, 9];
        var output = await Write(package); var nodes = Xml(output, "Metadata/List.xml").Root!.Elements(Ns + "Attachment").ToArray();
        Assert.Null(nodes[0].Attribute("vendor")); Assert.Empty(nodes[0].Elements(XName.Get("After", "urn:one")));
        Assert.Equal("external", nodes[1].Attribute("vendor")!.Value);
        Assert.Equal("Attach_1_original.bin.1", nodes[0].Element(Ns + "FileLoc")!.Value);
        Assert.Equal(new byte[] { 9, 9 }, Entry(output, "Metadata/Attach_1_original.bin"));
        Assert.Equal(new byte[] { 4, 5, 6 }, (await Read(output)).Attachments[0].Data);
        Assert.Equal(Payload, Entry(output, "Payload/original.bin"));
    }

    [Fact]
    public async Task Audit027_FreshAttachmentStillWorks()
    {
        var package = new OfdDocumentPackage(); package.Attachments.Add(new OfdAttachment { Name = "fresh.bin", Data = Payload });
        var output = await Write(package); var readback = await Read(output);
        Assert.Equal(Payload, Assert.Single(readback.Attachments).Data);
        Assert.Equal("fresh.bin", readback.Attachments[0].Name);
    }

    [Fact]
    public async Task Audit027_ExternalAttachmentMakesNoConnection()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
        try
        {
            var path = $"http://127.0.0.1:{((IPEndPoint)listener.LocalEndpoint).Port}/never-fetch.bin";
            var package = await Read(Fixture(external: path)); var output = await Write(package); var readback = await Read(output);
            Assert.Equal(path, readback.Attachments[1].ExternalPath); Assert.Empty(readback.Attachments[1].Data);
            Assert.False(listener.Pending());
            using var zip = new ZipArchive(new MemoryStream(output)); Assert.DoesNotContain(zip.Entries, e => e.Name == "external.bin");
        }
        finally { listener.Stop(); }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task Audit027_ControlledClonePreservesActiveSlotsAndTombstones(int removedCount)
    {
        var input = Fixture(); using var context = new OfdControlledEditContext(new());
        var loaded = await new OfdReader().ReadControlledAsync(new MemoryStream(input), new(), context);
        loaded.Attachments.RemoveRange(0, removedCount);
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        Assert.NotSame(loaded.AttachmentSource, owned.Package.AttachmentSource);
        Assert.NotSame(loaded.AttachmentSource!.Slots.Slots, owned.Package.AttachmentSource!.Slots.Slots);
        Assert.All(owned.Package.Attachments, a => Assert.Same(owned.Package, a.SourceOwner));
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        if (removedCount == 2) Assert.Null(Xml(output.ToArray(), "Document.xml").Root!.Element(Ns + "Attachments"));
        else
        {
            var expected = Xml(input, "Metadata/List.xml").Root!; if (removedCount == 1) expected.Elements(Ns + "Attachment").First().Remove();
            AssertXmlEqual(expected, Xml(output.ToArray(), "Metadata/List.xml").Root!);
        }
        Assert.Equal(2 - removedCount, (await Read(output.ToArray())).Attachments.Count);
    }

    [Theory]
    [InlineData(false, "source")]
    [InlineData(true, "source")]
    [InlineData(false, "origin")]
    [InlineData(true, "origin")]
    [InlineData(false, "slots")]
    [InlineData(true, "slots")]
    [InlineData(false, "baseline")]
    [InlineData(true, "baseline")]
    public async Task Audit027_PreparedWriteRejectsMutationOfNewInventory(bool candidate, string mutation)
    {
        var loaded = await Read(Fixture()); using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        var package = candidate ? owned.Package : loaded;
        switch (mutation)
        {
            case "source": package.AttachmentSource = null; break;
            case "origin": package.Attachments[0].SourceOwner = null; break;
            case "slots": package.AttachmentSource!.Slots.Slots.Clear(); break;
            case "baseline": Assert.True(package.AttachmentSource!.TryOriginal(package.Attachments[0], out var original)); original.Values["Name"] = "changed"; break;
        }
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("descriptor")]
    [InlineData("reference")]
    public async Task Audit027_RawSourceMutationRefusesBeforeWriting(string target)
    {
        var package = await Read(Fixture()); var path = target == "descriptor" ? "Metadata/List.xml" : "Document.xml";
        package.PreservedEntries[path] = Utf8(Encoding.UTF8.GetString(package.PreservedEntries[path]).Replace("vendor=", "tampered="));
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Audit027_ForeignOwnerAndUnsafePathMigrationRefuse()
    {
        var loaded = await Read(Fixture()); var other = new OfdDocumentPackage(); other.Attachments.Add(loaded.Attachments[0]);
        await Assert.ThrowsAsync<NotSupportedException>(() => Write(other));
        loaded.DocumentEntryPath = "Moved/Document.xml"; await Assert.ThrowsAsync<NotSupportedException>(() => Write(loaded));
        loaded.DocumentEntryPath = "Document.xml"; loaded.Attachments[0].IsExternal = true;
        await Assert.ThrowsAsync<NotSupportedException>(() => Write(loaded));
    }

    [Fact]
    public async Task Audit027_CrossDocumentToolsRefuseOpaqueSourceButKeepFreshSupport()
    {
        var package = await Read(Fixture());
        Assert.Throws<NotSupportedException>(() => OfdDocumentSplitter.Split(package, [0]));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([package]));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMixer.Mix([new OfdMixSource(package, 0)]));
        var fresh = new OfdDocumentPackage(); fresh.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 });
        fresh.Attachments.Add(new OfdAttachment { Name = "fresh.bin", Data = Payload });
        Assert.Equal(Payload, (await Read(await Write(OfdDocumentMerger.Merge([fresh])))).Attachments[0].Data);
        // A descriptor generated from the closed typed grammar remains supported.
        var plain = await Read(await Write(fresh)); Assert.Equal(Payload, (await Read(await Write(OfdDocumentSplitter.Split(plain, [0])))).Attachments[0].Data);
    }

    [Fact]
    public async Task Audit027_ImportedReorderAcrossOpaqueAnchorsRefuses()
    {
        var package = await Read(Fixture()); package.Attachments.Reverse(); await Assert.ThrowsAsync<NotSupportedException>(() => Write(package));
    }

    [Fact]
    public async Task Audit027_DuplicateMutableIdsStillUseObjectSlots()
    {
        var input = Fixture(); var package = await Read(input); package.Attachments[0].Id = "8"; package.Attachments[0].Name = "external.bin";
        var output = await Write(package); var expected = Xml(input, "Metadata/List.xml").Root!;
        expected.Element(Ns + "Attachment")!.SetAttributeValue("ID", "8"); expected.Element(Ns + "Attachment")!.SetAttributeValue("Name", "external.bin");
        AssertXmlEqual(expected, Xml(output, "Metadata/List.xml").Root!); Assert.Equal(Payload, (await Read(output)).Attachments[0].Data);
    }

    [Fact]
    public async Task Audit027_SourceStateParticipatesInStorageAndReadBudgets()
    {
        var package = await Read(Fixture()); var withSource = new OfdModelMaterialization(new()); withSource.InspectStorage(package, default);
        var source = package.AttachmentSource; package.AttachmentSource = null;
        var without = new OfdModelMaterialization(new()); without.InspectStorage(package, default);
        Assert.True(withSource.StorageBytes > without.StorageBytes);
        package.AttachmentSource = source;
        Assert.Throws<InvalidDataException>(() => new OfdModelMaterialization(new() { MaxMaterializedBytes = without.StorageBytes }).InspectStorage(package, default));
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 8192 });
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(new MemoryStream(Fixture()), new(), context));
    }

    [Fact]
    public async Task Audit027_OpaqueOnlyDescriptorIsStillAnActiveNoOpEnvelope()
    {
        var input = Fixture(); var descriptor = Xml(input, "Metadata/List.xml"); descriptor.Root!.Elements(Ns + "Attachment").Remove();
        descriptor.Root.Add(new XElement(XName.Get("Attachment", "urn:one"), new XAttribute("vendor", "unowned")));
        input = Replace(input, "Metadata/List.xml", Utf8(descriptor.ToString(SaveOptions.DisableFormatting)));
        var package = await Read(input); Assert.Empty(package.Attachments);
        var output = await Write(package);
        Assert.Equal("Metadata/List.xml", Xml(output, "Document.xml").Root!.Element(Ns + "Attachments")!.Value);
        AssertXmlEqual(descriptor.Root, Xml(output, "Metadata/List.xml").Root!);
    }

    [Fact]
    public async Task Audit027_ExternalPathEditPreservesLocationEnvelopeAndNeverEmbedsData()
    {
        var input = Fixture(); var descriptor = Xml(input, "Metadata/List.xml"); var location = descriptor.Root!.Elements(Ns + "Attachment").Last().Element(Ns + "FileLoc")!;
        location.SetAttributeValue(XName.Get("flag", "urn:one"), "keep"); location.Add(new XComment("location comment"));
        input = Replace(input, "Metadata/List.xml", Utf8(descriptor.ToString(SaveOptions.DisableFormatting)));
        var package = await Read(input); package.Attachments[1].ExternalPath = "https://invalid.example/changed"; package.Attachments[1].Data = [1, 2];
        var output = await Write(package);
        location.Nodes().OfType<XText>().Remove(); location.AddFirst(new XText(package.Attachments[1].ExternalPath!));
        AssertXmlEqual(descriptor.Root, Xml(output, "Metadata/List.xml").Root!);
        Assert.Empty((await Read(output)).Attachments[1].Data);
    }

    [Fact]
    public async Task Audit027_SharedPayloadRemainsSharedAndConflictingEditsRefuse()
    {
        var input = Fixture(); var descriptor = Xml(input, "Metadata/List.xml"); var other = descriptor.Root!.Elements(Ns + "Attachment").Last();
        other.Attribute("External")!.Remove(); other.Element(Ns + "FileLoc")!.Value = "../Payload/original.bin";
        input = Replace(input, "Metadata/List.xml", Utf8(descriptor.ToString(SaveOptions.DisableFormatting)));
        var package = await Read(input); Assert.All(package.Attachments, a => Assert.False(a.IsExternal));
        var output = await Write(package); AssertXmlEqual(descriptor.Root, Xml(output, "Metadata/List.xml").Root!);
        Assert.All((await Read(output)).Attachments, a => Assert.Equal(Payload, a.Data));
        package.Attachments[1].Data = [1, 2]; await Assert.ThrowsAsync<NotSupportedException>(() => Write(package));
    }

    [Fact]
    public async Task Audit027_DuplicateReferenceAndMixedLocationMigrationRefuse()
    {
        var input = Fixture(); var document = Xml(input, "Document.xml"); document.Root!.Add(new XElement(document.Root.Element(Ns + "Attachments")!));
        await Assert.ThrowsAsync<InvalidDataException>(() => Read(Replace(input, "Document.xml", Utf8(document.ToString(SaveOptions.DisableFormatting)))));
        var descriptor = Xml(input, "Metadata/List.xml"); descriptor.Root!.Elements(Ns + "Attachment").Last().Element(Ns + "FileLoc")!.Add(new XElement(XName.Get("Opaque", "urn:one")));
        var package = await Read(Replace(input, "Metadata/List.xml", Utf8(descriptor.ToString(SaveOptions.DisableFormatting))));
        await Write(package); package.Attachments[1].ExternalPath = "https://invalid.example/changed";
        await Assert.ThrowsAsync<NotSupportedException>(() => Write(package));
    }

    private static byte[] Fixture(string prefix = "", string location = "Metadata/List.xml", string external = "http://127.0.0.1:32767/never-fetch.bin")
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>{prefix}Document.xml</DocRoot></DocBody></OFD>"),
            [prefix + "Document.xml"] = Utf8($"<Document xmlns='{Ns}' xmlns:v='urn:one' xmlns:w='urn:two'><CommonData><MaxUnitID>2</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData><Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages><Attachments vendor='reference' v:flag='one' w:flag='two'>{location}<v:Reference kind='w:Binary'/><!--keep reference--></Attachments></Document>"),
            [prefix + "Pages/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>"),
            [prefix + location] = Utf8($"<Attachments xmlns='{Ns}' xmlns:v='urn:one' xmlns:w='urn:two' vendor='root' v:flag='one' w:flag='two'><v:Before/><Attachment ID='7' Name='original.bin' MediaType='application/x-original' Format='application/octet-stream' CreationDate='2026-01-02T03:04:05Z' ModDate='2026-02-03T04:05:06Z' Size='0.275' Visible='false' Usage='audit' vendor='attachment' v:flag='one' w:flag='two'><w:Before kind='w:Binary'/><FileLoc vendor='location' v:flag='one' w:flag='two'>../Payload/original.bin<w:Location kind='v:Binary'/><!--keep location--></FileLoc><v:After/></Attachment><w:Between/><Attachment ID='8' Name='external.bin' External='true' vendor='external'><FileLoc>{external}</FileLoc><v:After/></Attachment><v:After/></Attachments>"),
            [prefix + "Payload/original.bin"] = Payload
        };
        return Zip(entries);
    }

    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
    private static byte[] Replace(byte[] input, string name, byte[] value)
    {
        using var zip = new ZipArchive(new MemoryStream(input));
        return Zip(zip.Entries.ToDictionary(e => e.FullName, e => e.FullName == name ? value : Entry(input, e.FullName)));
    }
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var item in entries)
            {
                var entry = zip.CreateEntry(item.Key); entry.LastWriteTime = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var stream = entry.Open(); stream.Write(item.Value);
            }
        return output.ToArray();
    }
    private static byte[] Entry(byte[] bytes, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var input = zip.GetEntry(name)!.Open(); using var output = new MemoryStream(); input.CopyTo(output); return output.ToArray();
    }
    private static XDocument Xml(byte[] bytes, string name) => XDocument.Parse(Encoding.UTF8.GetString(Entry(bytes, name)));
    private static async Task<OfdDocumentPackage> Read(byte[] bytes) => await new OfdReader().ReadAsync(new MemoryStream(bytes));
    private static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }
    private static void AssertXmlEqual(XElement expected, XElement actual)
    {
        // Expanded names and all node kinds/order matter; prefix spelling and attribute order do not.
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString()).Select(a => (a.Name, a.Value)),
            actual.Attributes().Where(a => !a.IsNamespaceDeclaration).OrderBy(a => a.Name.ToString()).Select(a => (a.Name, a.Value)));
        Assert.Equal(expected.Nodes().Count(), actual.Nodes().Count());
        foreach (var pair in expected.Nodes().Zip(actual.Nodes()))
            if (pair.First is XElement element) AssertXmlEqual(element, Assert.IsType<XElement>(pair.Second));
            else { Assert.Equal(pair.First.NodeType, pair.Second.NodeType); Assert.Equal(pair.First.ToString(), pair.Second.ToString()); }
        foreach (var prefix in new[] { "v", "w" }) Assert.Equal(expected.GetNamespaceOfPrefix(prefix), actual.GetNamespaceOfPrefix(prefix));
    }
    private static void SaveArtifacts(string prefix, string location, byte[] input, byte[] output)
    {
        var root = Environment.GetEnvironmentVariable("OFDRW_AUDIT027_ARTIFACTS"); if (root is null) return;
        Directory.CreateDirectory(root); var label = prefix.Length == 0 ? "canonical" : "nested";
        File.WriteAllBytes(Path.Combine(root, label + "-input.ofd"), input); File.WriteAllBytes(Path.Combine(root, label + "-output.ofd"), output);
        File.WriteAllBytes(Path.Combine(root, label + "-descriptor.xml"), Entry(output, prefix + location));
    }
}
