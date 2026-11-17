using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

// D2: generic tags use the neutral key/value channel; the historical EMR
// channel stays readable, byte-stable when untouched and editable in place.
public sealed class CustomTagChannelTests
{
    private const string Ofd = "http://www.ofdspec.org";
    private const string KvNamespace = "urn:ofdrw-net:custom-tags:kv:1";
    private const string EmrNamespace = "urn:ofdrw-net:custom-tags:emr";
    private const string NamedAreaKey = "ofdrw.net.named-areas.v1";
    private const string Index = "Doc_0/Tags/CustomTags.xml";
    private const string KvPayload = "Doc_0/Tags/CustomTag_KeyValue.xml";
    private const string EmrPayload = "Doc_0/Tags/CustomTag_EMR.xml";
    private static readonly XName KvTag = XName.Get("Tag", KvNamespace);
    private static readonly XName OfdTag = XName.Get("Tag", Ofd);

    [Fact]
    public async Task NewDocumentWritesNeutralChannelWhosePayloadRootIsInTheDeclaredNamespace()
    {
        var bytes = await Write(Tagged());
        var entries = Entries(bytes);
        Assert.DoesNotContain(entries.Keys, path => path.Contains("EMR", StringComparison.OrdinalIgnoreCase));
        var document = Xml(bytes, "Doc_0/Document.xml");
        Assert.Equal("Tags/CustomTags.xml", document.Root!.Element(XName.Get("CustomTags", Ofd))!.Value);
        var record = Assert.Single(Xml(bytes, Index).Root!.Elements());
        Assert.Equal(XName.Get("CustomTag", Ofd), record.Name);
        Assert.Equal("ofdrw-kv", record.Attribute("TypeID")!.Value);
        Assert.Equal(KvNamespace, record.Attribute("NameSpace")!.Value);
        Assert.Equal("CustomTag_KeyValue.xml", record.Element(XName.Get("FileLoc", Ofd))!.Value);
        var payload = Xml(bytes, KvPayload).Root!;
        Assert.Equal(XName.Get("Tags", record.Attribute("NameSpace")!.Value), payload.Name);
        Assert.All(payload.Elements(), tag => Assert.Equal(KvTag, tag.Name));
        Assert.Equal(new[] { "source-text-origin=DOCX/OpenXML", "docx-ofd-mode=Native" },
            payload.Elements().Select(tag => tag.Attribute("Key")!.Value + "=" + tag.Attribute("Value")!.Value));
        var read = await Read(bytes);
        Assert.Equal("DOCX/OpenXML", read.CustomTags["source-text-origin"]);
        Assert.Equal(2, read.CustomTags.Count);
        var again = await Write(read);
        Assert.Equal(entries[Index], Entries(again)[Index]);
        Assert.Equal(entries[KvPayload], Entries(again)[KvPayload]);
    }

    [Fact]
    public async Task HistoricalEmrChannelReadsAndRoundTripsByteForByte()
    {
        var input = await HistoricalEmr();
        var package = await Read(input);
        Assert.Equal("DOCX/OpenXML", package.CustomTags["source-text-origin"]);
        Assert.Equal("Native", package.CustomTags["docx-ofd-mode"]);
        Assert.Equal(2, package.CustomTags.Count);
        var output = Entries(await Write(package));
        var source = Entries(input);
        Assert.Equal(source[Index], output[Index]);
        Assert.Equal(source[EmrPayload], output[EmrPayload]);
        Assert.False(output.ContainsKey(KvPayload));
        Assert.Equal("Tags/CustomTags.xml", XDocument.Parse(Encoding.UTF8.GetString(output["Doc_0/Document.xml"]))
            .Root!.Element(XName.Get("CustomTags", Ofd))!.Value);
    }

    [Fact]
    public async Task EditsOnEmrSourceUpdateInPlaceAndRouteNewKeysToNeutralChannel()
    {
        var package = await Read(await HistoricalEmr());
        package.CustomTags["docx-ofd-mode"] = "DualLayer";
        package.CustomTags.Remove("source-text-origin");
        package.CustomTags["owner"] = "新增";
        var first = await Write(package);

        var emr = Xml(first, EmrPayload).Root!;
        Assert.Equal(XName.Get("EMRTags", Ofd), emr.Name);
        Assert.Equal(new[] { "docx-ofd-mode=DualLayer" }, emr.Elements(OfdTag).Select(Pair));
        var kv = Xml(first, KvPayload).Root!;
        Assert.Equal(XName.Get("Tags", KvNamespace), kv.Name);
        Assert.Equal(new[] { "owner=新增" }, kv.Elements(KvTag).Select(Pair));
        var records = Xml(first, Index).Root!.Elements().ToArray();
        Assert.Equal(new[] { "EMR|" + EmrNamespace + "|CustomTag_EMR.xml", "ofdrw-kv|" + KvNamespace + "|CustomTag_KeyValue.xml" },
            records.Select(r => r.Attribute("TypeID")!.Value + "|" + r.Attribute("NameSpace")!.Value + "|" + r.Element(XName.Get("FileLoc", Ofd))!.Value));
        Assert.Single(Xml(first, "Doc_0/Document.xml").Root!.Elements(XName.Get("CustomTags", Ofd)));

        var reread = await Read(first);
        Assert.Equal(new Dictionary<string, string> { ["docx-ofd-mode"] = "DualLayer", ["owner"] = "新增" }, reread.CustomTags);
        var unchanged = Entries(await Write(reread));
        Assert.Equal(Entries(first)[Index], unchanged[Index]);
        Assert.Equal(Entries(first)[EmrPayload], unchanged[EmrPayload]);
        Assert.Equal(Entries(first)[KvPayload], unchanged[KvPayload]);

        reread.CustomTags["owner"] = "changed"; reread.CustomTags["second"] = "2"; reread.CustomTags["docx-ofd-mode"] = "Native";
        var second = await Write(reread);
        Assert.Equal(new[] { "docx-ofd-mode=Native" }, Xml(second, EmrPayload).Root!.Elements(OfdTag).Select(Pair));
        Assert.Equal(new[] { "owner=changed", "second=2" }, Xml(second, KvPayload).Root!.Elements(KvTag).Select(Pair));
        Assert.Equal(2, Xml(second, Index).Root!.Elements().Count());
        Assert.DoesNotContain(Entries(second).Keys, path => path.Contains("CustomTag_KeyValue_", StringComparison.Ordinal));
    }

    [Fact]
    public async Task UndeclaredLegacyEmrPayloadGainsBothIndexRecordsOnEdit()
    {
        var entries = Entries(await HistoricalEmr());
        var document = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]));
        document.Root!.Element(XName.Get("CustomTags", Ofd))!.Remove();
        entries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        entries.Remove(Index);
        var package = await Read(Zip(entries));
        Assert.Equal("Native", package.CustomTags["docx-ofd-mode"]);
        package.CustomTags["owner"] = "kv";
        var output = await Write(package);
        var reread = await Read(output);
        Assert.Equal("Native", reread.CustomTags["docx-ofd-mode"]);
        Assert.Equal("kv", reread.CustomTags["owner"]);
        var indexPath = "Doc_0/" + Xml(output, "Doc_0/Document.xml").Root!.Element(XName.Get("CustomTags", Ofd))!.Value;
        Assert.Equal(new[] { "EMR", "ofdrw-kv" }, Xml(output, indexPath).Root!.Elements().Select(r => r.Attribute("TypeID")!.Value));
        Assert.DoesNotContain("owner", Encoding.UTF8.GetString(Entries(output)[EmrPayload]));
    }

    [Theory]
    [InlineData("emr-root-under-kv")]
    [InlineData("ofd-namespace-kv-root")]
    [InlineData("kv-root-under-emr")]
    public async Task PayloadWhoseRootDoesNotMatchItsDeclaredChannelStaysOpaque(string kind)
    {
        var entries = Entries(await Write(Tagged()));
        var payload = Xml(Zip(entries), KvPayload);
        if (kind == "emr-root-under-kv")
            payload = new XDocument(new XElement(XName.Get("EMRTags", Ofd), new XElement(OfdTag, new XAttribute("Key", "x"), new XAttribute("Value", "y"))));
        if (kind == "ofd-namespace-kv-root")
            payload = new XDocument(new XElement(XName.Get("Tags", Ofd), new XElement(OfdTag, new XAttribute("Key", "x"), new XAttribute("Value", "y"))));
        if (kind == "kv-root-under-emr")
        {
            var index = XDocument.Parse(Encoding.UTF8.GetString(entries[Index]));
            var record = index.Root!.Elements().Single();
            record.SetAttributeValue("TypeID", "EMR"); record.SetAttributeValue("NameSpace", EmrNamespace);
            entries[Index] = Encoding.UTF8.GetBytes(index.ToString(SaveOptions.DisableFormatting));
        }
        else entries[KvPayload] = Encoding.UTF8.GetBytes(payload.ToString(SaveOptions.DisableFormatting));
        var input = Zip(entries);
        var package = await Read(input);
        Assert.Empty(package.CustomTags);
        var output = Entries(await Write(package));
        Assert.Equal(entries[Index], output[Index]);
        Assert.Equal(entries[KvPayload], output[KvPayload]);
    }

    [Fact]
    public async Task NamedAreasWorkOnNeutralAndHistoricalEmrChannels()
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new AreaHolderBlock("logo", 15));
        var neutral = await Write(flow.Render());
        Assert.Contains(Xml(neutral, KvPayload).Root!.Elements(KvTag), tag => tag.Attribute("Key")!.Value == NamedAreaKey);
        var historical = ToEmr(neutral);
        var image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l3sAAAAASUVORK5CYII=");
        foreach (var (input, payloadPath, tagName) in new[] { (neutral, KvPayload, KvTag), (historical, EmrPayload, OfdTag) })
        {
            var package = await Read(input);
            var editor = new OfdNamedAreaEditor();
            Assert.False(Assert.Single(editor.GetAreas(package)).IsFilled);
            editor.FillImage(package, "logo", image, "image/png", 1, 1);
            var output = await Write(package);
            Assert.True(Assert.Single(editor.GetAreas(await Read(output))).IsFilled);
            var tags = Xml(output, payloadPath).Root!.Elements(tagName);
            Assert.Contains(tags, tag => tag.Attribute("Key")!.Value == NamedAreaKey);
            Assert.Single(Xml(output, Index).Root!.Elements());
        }
    }

    [Fact]
    public async Task MixAndSplitAcceptHistoricalEmrInput()
    {
        var historical = await Read(await HistoricalEmr());
        var mixed = await Write(OfdDocumentMixer.Mix([new OfdMixSource(historical, 0)]));
        Assert.Equal("Native", (await Read(mixed)).CustomTags["docx-ofd-mode"]);
        Assert.Contains(KvPayload, Entries(mixed).Keys);
        var split = Entries(await Write(OfdDocumentSplitter.Split(await Read(await HistoricalEmr()), new[] { 0 })));
        Assert.Equal(Entries(await HistoricalEmr())[EmrPayload], split[EmrPayload]);
        Assert.False(split.ContainsKey(KvPayload));
    }

    private static string Pair(XElement tag) => tag.Attribute("Key")!.Value + "=" + tag.Attribute("Value")!.Value;

    private static OfdDocumentPackage Tagged()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["docx-ofd-mode"] = "Native";
        return package;
    }

    // Reproduces the pre-D2 writer output: EMR index record and an OFD-namespace EMRTags payload.
    private static async Task<byte[]> HistoricalEmr() => ToEmr(await Write(Tagged()));

    private static byte[] ToEmr(byte[] neutral)
    {
        var entries = Entries(neutral);
        var tags = XDocument.Parse(Encoding.UTF8.GetString(entries[KvPayload])).Root!.Elements(KvTag);
        var payload = new XDocument(new XElement(XName.Get("EMRTags", Ofd), new XAttribute(XNamespace.Xmlns + "ofd", Ofd),
            tags.Select(tag => new XElement(OfdTag, new XAttribute("Key", tag.Attribute("Key")!.Value), new XAttribute("Value", tag.Attribute("Value")!.Value)))));
        var index = XDocument.Parse(Encoding.UTF8.GetString(entries[Index]));
        var record = index.Root!.Elements().Single();
        record.SetAttributeValue("TypeID", "EMR"); record.SetAttributeValue("NameSpace", EmrNamespace);
        record.Element(XName.Get("FileLoc", Ofd))!.Value = "CustomTag_EMR.xml";
        entries.Remove(KvPayload);
        entries[EmrPayload] = Encoding.UTF8.GetBytes(payload.ToString());
        entries[Index] = Encoding.UTF8.GetBytes(index.ToString());
        return Zip(entries);
    }

    private static async Task<OfdDocumentPackage> Read(byte[] data)
    { using var stream = new MemoryStream(data); return await new OfdReader().ReadAsync(stream); }

    private static async Task<byte[]> Write(OfdDocumentPackage package)
    { using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream); return stream.ToArray(); }

    private static Dictionary<string, byte[]> Entries(byte[] data)
    {
        using var zip = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry =>
        { using var source = entry.Open(); using var copy = new MemoryStream(); source.CopyTo(copy); return copy.ToArray(); });
    }

    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var entry in entries)
            { using var target = zip.CreateEntry(entry.Key).Open(); target.Write(entry.Value); }
        return stream.ToArray();
    }

    private static XDocument Xml(byte[] data, string path) => XDocument.Parse(Encoding.UTF8.GetString(Entries(data)[path]));
}
