using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit008RegressionTests
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private const string EmrNamespace = "urn:ofdrw-net:custom-tags:emr";

    [Theory]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task Audit008_OpaqueOnlyNoOp(string prefix)
    {
        var input = Fixture(prefix, mixed: false);
        var package = await Read(input);
        Assert.Empty(package.CustomTags);
        var output = await Write(package);
        AssertOpaque(input, output, prefix);
        Assert.Equal(Entry(input, prefix + "Meta/list.xml"), Entry(output, prefix + "Meta/list.xml"));
        SaveArtifacts(prefix, "opaque", input, output);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task Audit008_MixedNoOp(string prefix)
    {
        var input = Fixture(prefix, mixed: true);
        var package = await Read(input);
        Assert.Equal("原始", package.CustomTags["patient"]);
        Assert.False(package.CustomTags.ContainsKey("opaque-key"));
        var output = await Write(package);
        AssertOpaque(input, output, prefix);
        Assert.Equal(Entry(input, prefix + "Meta/list.xml"), Entry(output, prefix + "Meta/list.xml"));
        Assert.Equal(Entry(input, prefix + "Owned/emr.xml"), Entry(output, prefix + "Owned/emr.xml"));
        SaveArtifacts(prefix, "mixed", input, output);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task Audit008_ModifyOwnedEntry(string prefix)
    {
        var input = Fixture(prefix, mixed: true);
        var package = await Read(input);
        package.CustomTags["patient"] = "修改 & updated";
        var output = await Write(package);
        AssertOpaque(input, output, prefix);
        Assert.Equal(Entry(input, prefix + "Meta/list.xml"), Entry(output, prefix + "Meta/list.xml"));
        var detail = Xml(output, prefix + "Owned/emr.xml");
        Assert.Equal("keep", detail.Root!.Attribute("vendor")!.Value);
        Assert.Equal("extension", detail.Descendants(XName.Get("Extra", "urn:vendor")).Single().Value);
        Assert.Equal(new[] { "patient", "stable" }, detail.Root.Elements(XName.Get("Tag", Ns)).Select(x => x.Attribute("Key")!.Value));
        var readback = await Read(output);
        Assert.Equal("修改 & updated", readback.CustomTags["patient"]);
        Assert.Equal("unchanged", readback.CustomTags["stable"]);
        SaveArtifacts(prefix, "modified", input, output);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task Audit008_ClearModeledEntries(string prefix)
    {
        var input = Fixture(prefix, mixed: true);
        var package = await Read(input);
        package.CustomTags.Clear();
        var output = await Write(package);
        AssertOpaque(input, output, prefix);
        Assert.Empty((await Read(output)).CustomTags);
        Assert.DoesNotContain("原始", Encoding.UTF8.GetString(Entry(output, prefix + "Owned/emr.xml")));
        var again = await Write(await Read(output));
        Assert.Empty((await Read(again)).CustomTags);
        AssertOpaque(input, again, prefix);
        SaveArtifacts(prefix, "cleared", input, output);
    }

    [Fact]
    public async Task Audit008_FreshDocument()
    {
        var package = new OfdDocumentPackage();
        package.CustomTags["new"] = "value";
        var bytes = await Write(package);
        Assert.Equal("value", (await Read(bytes)).CustomTags["new"]);
        var index = Xml(bytes, "Doc_0/Tags/CustomTags.xml");
        Assert.Equal("EMR", index.Root!.Elements().Single().Attribute("TypeID")!.Value);
        Assert.Equal(EmrNamespace, index.Root.Elements().Single().Attribute("NameSpace")!.Value);
        Assert.Empty((await Read(await Write(new OfdDocumentPackage()))).CustomTags);
    }

    [Fact]
    public async Task Audit008_UnknownSchemaKeyValueAndBinaryPayloadRemainOpaque()
    {
        var input = Replace(Fixture("", false), "Opaque/data.xml", Encoding.UTF8.GetBytes($"<ofd:EMRTags xmlns:ofd='{Ns}'><ofd:Tag Key='foreign' Value='opaque'/></ofd:EMRTags>"));
        input = Replace(input, "Opaque/raw.bin", new byte[] { 0, 255, 4, 13 });
        var package = await Read(input);
        Assert.Empty(package.CustomTags);
        var output = await Write(package);
        AssertOpaque(input, output, "");
    }

    [Theory]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task Audit008_AddModeledEntryWithoutOverwritingOpaqueConventionalPath(string prefix)
    {
        var input = Replace(Fixture(prefix, false), prefix + "Tags/CustomTag_EMR.xml", new byte[] { 1, 2, 3 });
        var package = await Read(input);
        package.CustomTags["new"] = "新增";
        var output = await Write(package);
        Assert.Equal(Entry(input, prefix + "Opaque/data.xml"), Entry(output, prefix + "Opaque/data.xml"));
        Assert.Equal(new byte[] { 1, 2, 3 }, Entry(output, prefix + "Tags/CustomTag_EMR.xml"));
        Assert.Equal("新增", (await Read(output)).CustomTags["new"]);
        Assert.Equal(3, Xml(output, prefix + "Meta/list.xml").Root!.Elements().Count());
    }

    [Fact]
    public async Task Audit008_ControlledCloneAndRetainedSplitKeepClearedOrigins()
    {
        var input = Fixture("Nested/Doc/", true);
        using var stream = new MemoryStream(input);
        using var context = new OfdControlledEditContext(new());
        var loaded = await new OfdReader().ReadControlledAsync(stream, new(), context);
        loaded.CustomTags.Clear();
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        Assert.NotSame(loaded.CustomTagSource, owned.Package.CustomTagSource);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        using var output = new MemoryStream(); await prepared.WriteAsync(output);
        Assert.Empty((await Read(output.ToArray())).CustomTags);
        AssertOpaque(input, output.ToArray(), "Nested/Doc/");
        var split = OfdDocumentSplitter.Split(loaded, new[] { 0 });
        var saved = await Write(split);
        Assert.Empty((await Read(saved)).CustomTags);
        AssertOpaque(input, saved, "Nested/Doc/");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit008_PreparedWriteRejectsSourceMetadataMutation(bool mutateCandidate)
    {
        var loaded = await Read(Fixture("", true));
        using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        (mutateCandidate ? owned.Package : loaded).CustomTagSource = null;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Audit008_SharedOpaquePayloadNoOpSurvivesButAmbiguousRewriteRefuses()
    {
        var input = Fixture("", true);
        var index = Xml(input, "Meta/list.xml");
        index.Root!.Elements().First().Element(XName.Get("FileLoc", Ns))!.Value = "../Owned/emr.xml";
        input = Replace(input, "Meta/list.xml", Encoding.UTF8.GetBytes(index.ToString(SaveOptions.DisableFormatting)));
        var loaded = await Read(input);
        var output = await Write(loaded);
        Assert.Equal(Entry(input, "Meta/list.xml"), Entry(output, "Meta/list.xml"));
        Assert.Equal(Entry(input, "Owned/emr.xml"), Entry(output, "Owned/emr.xml"));
        loaded.CustomTags["patient"] = "changed";
        await Assert.ThrowsAsync<NotSupportedException>(() => Write(loaded));
    }

    [Fact]
    public async Task Audit008_UneditedDuplicateKeyValuesAndSecondIndexRemainIntact()
    {
        var input = Fixture("", true);
        var detail = Xml(input, "Owned/emr.xml");
        detail.Root!.AddFirst(new XElement(XName.Get("Tag", Ns), new XAttribute("Key", "stable"), new XAttribute("Value", "first-duplicate")));
        input = Replace(input, "Owned/emr.xml", Encoding.UTF8.GetBytes(detail.ToString(SaveOptions.DisableFormatting)));
        var document = Xml(input, "Document.xml");
        document.Root!.Add(new XElement(XName.Get("CustomTags", Ns), "Meta/second.xml"));
        input = Replace(input, "Document.xml", Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting)));
        input = Replace(input, "Meta/second.xml", Encoding.UTF8.GetBytes($"<ofd:CustomTags xmlns:ofd='{Ns}'><ofd:CustomTag TypeID='other' FileLoc='../Opaque/data.xml'/></ofd:CustomTags>"));
        var loaded = await Read(input); loaded.CustomTags["patient"] = "changed";
        var output = await Write(loaded);
        Assert.Equal(new[] { "Meta/list.xml", "Meta/second.xml" }, Xml(output, "Document.xml").Root!.Elements(XName.Get("CustomTags", Ns)).Select(node => node.Value));
        Assert.Equal(Entry(input, "Meta/second.xml"), Entry(output, "Meta/second.xml"));
        Assert.Equal(new[] { "first-duplicate", "unchanged" }, Xml(output, "Owned/emr.xml").Root!.Elements(XName.Get("Tag", Ns)).Where(node => node.Attribute("Key")?.Value == "stable").Select(node => node.Attribute("Value")!.Value));
        Assert.Equal("changed", (await Read(output)).CustomTags["patient"]);
    }

    private static void SaveArtifacts(string prefix, string name, byte[] input, byte[] output)
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_AUDIT008_ARTIFACTS");
        if (prefix.Length != 0 || string.IsNullOrWhiteSpace(directory)) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, name + "-input.ofd"), input);
        File.WriteAllBytes(Path.Combine(directory, name + "-output.ofd"), output);
    }

    private static byte[] Replace(byte[] data, string path, byte[] replacement)
    {
        using var stream = new MemoryStream(); stream.Write(data); stream.Position = 0;
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry(path)?.Delete();
            using var entry = zip.CreateEntry(path).Open(); entry.Write(replacement);
        }
        return stream.ToArray();
    }

    private static void AssertOpaque(byte[] input, byte[] output, string prefix)
    {
        var doc = Xml(output, prefix + "Document.xml");
        var declaration = doc.Root!.Element(XName.Get("CustomTags", Ns));
        Assert.NotNull(declaration);
        Assert.Equal("Meta/list.xml", declaration!.Value);
        Assert.Equal("declaration", declaration.Attribute(XName.Get("owner", "urn:vendor"))!.Value);
        var source = Xml(input, prefix + "Meta/list.xml");
        var index = Xml(output, prefix + "Meta/list.xml");
        Assert.Equal(source.Root!.Attributes().Select(x => x.ToString()), index.Root!.Attributes().Select(x => x.ToString()));
        Assert.Equal(source.Root.Elements().Select(x => x.Name), index.Root.Elements().Select(x => x.Name));
        foreach (var node in source.Root.Elements().Where(x => x.Attribute("TypeID")?.Value != "EMR"))
            Assert.True(XNode.DeepEquals(node, index.Root.Elements().Single(x => x.Name == node.Name && x.Attribute("TypeID")?.Value == node.Attribute("TypeID")?.Value)));
        Assert.Equal(Entry(input, prefix + "Opaque/data.xml"), Entry(output, prefix + "Opaque/data.xml"));
        Assert.Equal(Entry(input, prefix + "Opaque/raw.bin"), Entry(output, prefix + "Opaque/raw.bin"));
    }

    private static byte[] Fixture(string prefix, bool mixed)
    {
        var emr = mixed ? $"<ofd:CustomTag TypeID='EMR' NameSpace='{EmrNamespace}' vendor='owned-index'><ofd:FileLoc>../Owned/emr.xml</ofd:FileLoc><v:Extra>index-extra</v:Extra></ofd:CustomTag>" : "";
        var entries = new Dictionary<string, string>
        {
            ["OFD.xml"] = $"<ofd:OFD xmlns:ofd='{Ns}' Version='1.0' DocType='OFD'><ofd:DocBody><ofd:DocInfo/><ofd:DocRoot>{prefix}Document.xml</ofd:DocRoot></ofd:DocBody></ofd:OFD>",
            [prefix + "Document.xml"] = $"<ofd:Document xmlns:ofd='{Ns}' xmlns:v='urn:vendor'><ofd:CommonData><ofd:MaxUnitID>1</ofd:MaxUnitID><ofd:PageArea><ofd:PhysicalBox>0 0 100 100</ofd:PhysicalBox></ofd:PageArea></ofd:CommonData><ofd:Pages><ofd:Page ID='1' BaseLoc='Pages/page.xml'/></ofd:Pages><ofd:CustomTags v:owner='declaration'>Meta/list.xml</ofd:CustomTags></ofd:Document>",
            [prefix + "Pages/page.xml"] = $"<ofd:Page xmlns:ofd='{Ns}'><ofd:Content/></ofd:Page>",
            [prefix + "Meta/list.xml"] = $"<ofd:CustomTags xmlns:ofd='{Ns}' xmlns:v='urn:vendor' vendor='list'><ofd:CustomTag TypeID='VendorType' NameSpace='urn:vendor' v:flag='yes'><ofd:FileLoc>../Opaque/data.xml</ofd:FileLoc><v:Extra>opaque-index</v:Extra></ofd:CustomTag>{emr}<v:CustomTag TypeID='foreign' FileLoc='../Opaque/raw.bin'><v:Child/></v:CustomTag></ofd:CustomTags>",
            [prefix + "Opaque/data.xml"] = $"<v:EMRTags xmlns:v='urn:vendor'><v:VendorData keep='all'/><v:Nested keep='all'/></v:EMRTags>",
            [prefix + "Opaque/raw.bin"] = "<vendor:Payload xmlns:vendor='urn:binary-envelope'>opaque</vendor:Payload>",
            [prefix + "Owned/emr.xml"] = $"<ofd:EMRTags xmlns:ofd='{Ns}' xmlns:v='urn:vendor' vendor='keep'><ofd:Tag Key='patient' Value='原始' v:flag='yes'/><v:Extra>extension</v:Extra><ofd:Tag Key='stable' Value='unchanged'/></ofd:EMRTags>"
        };
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var entry in entries)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entry.Key).Open(), new UTF8Encoding(false));
                writer.Write(entry.Value);
            }
        return stream.ToArray();
    }

    private static async Task<OfdDocumentPackage> Read(byte[] data)
    { using var stream = new MemoryStream(data); return await new OfdReader().ReadAsync(stream); }
    private static async Task<byte[]> Write(OfdDocumentPackage package)
    { using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream); return stream.ToArray(); }
    private static byte[] Entry(byte[] data, string path)
    { using var stream = new MemoryStream(data); using var zip = new ZipArchive(stream, ZipArchiveMode.Read); using var entry = zip.GetEntry(path)!.Open(); using var copy = new MemoryStream(); entry.CopyTo(copy); return copy.ToArray(); }
    private static XDocument Xml(byte[] data, string path) => XDocument.Parse(Encoding.UTF8.GetString(Entry(data, path)));
}
