using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class ControlledReaderAccountingTests
{
    [Fact]
    public async Task TwoReadsShareCountersAndRetainedPayloadAliasesAreAdoptedOnce()
    {
        var bytes = await Source(); using var context = new OfdControlledEditContext(new()); var reader = new OfdReader();
        using var firstInput = new MemoryStream(bytes); var first = await reader.ReadControlledAsync(firstInput, new(), context);
        var material = context.Material; var storage = material.StorageBytes; var copy = material.CloneCopyBytes; var work = material.ValidationWorkBytes;
        var font = Assert.Single(first.Fonts); Assert.Contains(first.PreservedEntries.Values, value => ReferenceEquals(value, font.Data));
        material.InspectStorage(first, default); Assert.Equal(storage, material.StorageBytes);
        using var secondInput = new MemoryStream(bytes); var second = await reader.ReadControlledAsync(secondInput, new(), context);
        Assert.NotSame(font.Data, Assert.Single(second.Fonts).Data); Assert.True(material.StorageBytes > storage);
        Assert.True(material.CloneCopyBytes > copy); Assert.True(material.ValidationWorkBytes > work);
        Assert.Equal(2, material.ReadMetadataReservations); Assert.Equal(first.PreservedEntries.Count * 2, material.ReadEntryAllocations);
    }

    [Fact]
    public async Task InitialReadOwnedClonePreparedWriteReadbackKeepOneOperationAndChangedText()
    {
        var bytes = await Source(); using var context = new OfdControlledEditContext(new()); var reader = new OfdReader(); using var input = new MemoryStream(bytes);
        var source = await reader.ReadControlledAsync(input, new(), context); var initialCopy = context.Material.CloneCopyBytes;
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var text = owned.GetMapped((OfdTextElement)source.Pages[0].Elements[0]);
        owned.ApplyTextSlots(text, [new(0, "CHANGED", text.Runs[0].XMillimeters, text.Runs[0].YMillimeters)]);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); var outputAllowance = context.Material.OutputUpperBoundBytes;
        using var output = new MemoryStream(); await prepared.WriteAsync(output); var writtenCopy = context.Material.CloneCopyBytes; var writtenWork = context.Material.ValidationWorkBytes;
        output.Position = 0; var readback = await reader.ReadControlledAsync(output, new(), context);
        Assert.Equal("CHANGED", ((OfdTextElement)readback.Pages[0].Elements[0]).Text);
        Assert.True(writtenCopy > initialCopy); Assert.True(context.Material.CloneCopyBytes > writtenCopy); Assert.True(context.Material.ValidationWorkBytes > writtenWork);
        Assert.Equal(outputAllowance, context.Material.OutputUpperBoundBytes); Assert.Equal(source.Fonts[0].Data, readback.Fonts[0].Data);
    }

    [Fact]
    public async Task CumulativeQuotaCannotBeResetBySecondReadOrLargerCallerOptions()
    {
        var bytes = await Source(); long oneCopy;
        using (var measurement = new OfdControlledEditContext(new()))
        { using var stream = new MemoryStream(bytes); await new OfdReader().ReadControlledAsync(stream, new(), measurement); oneCopy = measurement.Material.CloneCopyBytes; }
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = oneCopy + 4096 });
        using var first = new MemoryStream(bytes); await new OfdReader().ReadControlledAsync(first, new(), context);
        using var second = new MemoryStream(bytes); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(second, new() { MaxInputBytes = long.MaxValue }, context));
        Assert.Equal(1, context.Material.ReadMetadataReservations);
        using var third = new MemoryStream(bytes); await Assert.ThrowsAsync<InvalidOperationException>(() => new OfdReader().ReadControlledAsync(third, new(), context));
    }

    [Fact]
    public async Task MetadataQuotaRejectsAfterSmallFooterScreenBeforeArchiveOrPayloadAllocation()
    {
        var bytes = await Source(); using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 1024 * 1024 }); using var stream = new ObservedStream(bytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(stream, new(), context));
        Assert.Equal(22, stream.ReadBytes); Assert.Equal(0, context.Material.ReadMetadataReservations); Assert.Equal(0, context.Material.ReadEntryAllocations); Assert.Equal(0, context.Material.ReadDomAllocations);
    }

    [Theory]
    [InlineData("signature")] [InlineData("comment")] [InlineData("trailing")]
    [InlineData("disk")] [InlineData("disk-count")] [InlineData("offset")] [InlineData("size")] [InlineData("count")]
    [InlineData("zip64-disk")] [InlineData("zip64-central-disk")] [InlineData("zip64-disk-count")]
    [InlineData("zip64-count")] [InlineData("zip64-size")] [InlineData("zip64-offset")]
    [InlineData("short")]
    public async Task UnsupportedOrInconsistentFooterRejectsBeforeMetadataAndEntryAllocation(string defect)
    {
        var bytes = Zip("data", [1]); var footer = bytes.Length - 22;
        switch (defect)
        {
            case "signature": bytes[footer] = 0; break;
            case "comment": Set16(bytes, footer + 20, 1); break;
            case "trailing": bytes = [..bytes, 1]; break;
            case "disk": Set16(bytes, footer + 4, 1); break;
            case "disk-count": Set16(bytes, footer + 8, 0); break;
            case "offset": Set32(bytes, footer + 16, 0xfffffff0); break;
            case "size": Set32(bytes, footer + 12, 0xfffffff0); break;
            case "count": Set16(bytes, footer + 8, 500); Set16(bytes, footer + 10, 500); break;
            case "zip64-disk": Set16(bytes, footer + 4, ushort.MaxValue); break;
            case "zip64-central-disk": Set16(bytes, footer + 6, ushort.MaxValue); break;
            case "zip64-disk-count": Set16(bytes, footer + 8, ushort.MaxValue); break;
            case "zip64-count": Set16(bytes, footer + 10, ushort.MaxValue); break;
            case "zip64-size": Set32(bytes, footer + 12, uint.MaxValue); break;
            case "zip64-offset": Set32(bytes, footer + 16, uint.MaxValue); break;
            case "short": bytes = new byte[21]; break;
        }
        using var context = new OfdControlledEditContext(new()); using var input = new MemoryStream(bytes);
        await Assert.ThrowsAnyAsync<Exception>(() => new OfdReader().ReadControlledAsync(input, new(), context));
        Assert.Equal(0, context.Material.ReadMetadataReservations); Assert.Equal(0, context.Material.ReadEntryAllocations); Assert.Equal(0, context.Material.ReadDomAllocations);
        Assert.Throws<InvalidOperationException>(context.AssertActive);
    }

    [Theory]
    [InlineData(28)] [InlineData(30)] [InlineData(32)]
    public async Task MaliciousCentralNameExtraCommentLengthsCannotAllocatePayload(int field)
    {
        var bytes = Zip("data", [1]); var central = (int)BitConverter.ToUInt32(bytes, bytes.Length - 6); Set16(bytes, central + field, ushort.MaxValue);
        using var context = new OfdControlledEditContext(new()); using var input = new MemoryStream(bytes);
        await Assert.ThrowsAnyAsync<Exception>(() => new OfdReader().ReadControlledAsync(input, new(), context));
        Assert.Equal(1, context.Material.ReadMetadataReservations); Assert.Equal(0, context.Material.ReadEntryAllocations);
    }

    [Theory]
    [InlineData(0)] [InlineData(2)]
    public async Task DeclaredExpandedLengthMismatchRejectsBeforeModel(int declared)
    {
        var bytes = Zip("data", [99]); var central = (int)BitConverter.ToUInt32(bytes, bytes.Length - 6);
        Set32(bytes, central + 24, (uint)declared); Set32(bytes, 22, (uint)declared);
        using var context = new OfdControlledEditContext(new()); using var input = new MemoryStream(bytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(input, new(), context));
        Assert.Equal(1, context.Material.ReadEntryAllocations); Assert.Equal(0, context.Material.ReadDomAllocations);
    }

    [Fact]
    public async Task OptionsSnapshotCannotBeRelaxedFromAStreamCallback()
    {
        var bytes = await Source(); var options = new OfdPackageLoadOptions { MaxEntryUncompressedBytes = 1 };
        using var context = new OfdControlledEditContext(new()); using var input = new ObservedStream(bytes, () => options.MaxEntryUncompressedBytes = long.MaxValue);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(input, options, context)); Assert.Equal(0, context.Material.ReadEntryAllocations);
    }

    [Fact]
    public async Task CancellationDuringZipMetadataConsumesContextWithoutPayloadOrModel()
    {
        var bytes = await Source(); using var cancellation = new CancellationTokenSource(); using var context = new OfdControlledEditContext(new());
        using var input = new ObservedStream(bytes, cancellation.Cancel);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdReader().ReadControlledAsync(input, new(), context, cancellation.Token));
        Assert.Equal(0, context.Material.ReadEntryAllocations); Assert.Throws<InvalidOperationException>(context.AssertActive);
    }

    [Fact]
    public void DomAndSourceReservationsFailBeforeTheirAllocationCheckpoints()
    {
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 65536 }); var material = context.Material;
        using var scope = material.BeginControlledRead(context, default);
        Assert.Throws<InvalidDataException>(() => material.PrepareReadDom(65536, 100, 100, default)); Assert.Equal(0, material.ReadDomAllocations);
        var element = new XElement("TextObject", new XAttribute(XNamespace.Xmlns + "v", new string('x', 20000)), new XElement("TextCode", "data"));
        Assert.Throws<InvalidDataException>(() => OfdXmlEnvelope.Capture(element, material, default)); Assert.Equal(0, material.ReadSourceClones); Assert.Equal(0, material.ReadSourceStrings);
    }

    [Fact]
    public async Task ControlledReaderAndPublicReaderKeepIdsTextNamespaceAndSourceBytes()
    {
        var bytes = await Source(); using var publicInput = new MemoryStream(bytes); var baseline = await new OfdReader().ReadAsync(publicInput);
        using var context = new OfdControlledEditContext(new()); using var input = new MemoryStream(bytes); var controlled = await new OfdReader().ReadControlledAsync(input, new(), context);
        Assert.Equal(baseline.Options.Namespace, controlled.Options.Namespace); Assert.Equal(baseline.Pages[0].Id, controlled.Pages[0].Id);
        Assert.Equal(((OfdTextElement)baseline.Pages[0].Elements[0]).SourceXml, ((OfdTextElement)controlled.Pages[0].Elements[0]).SourceXml);
        Assert.Equal(baseline.PreservedEntries.Keys, controlled.PreservedEntries.Keys);
        foreach (var entry in baseline.PreservedEntries) Assert.Equal(entry.Value, controlled.PreservedEntries[entry.Key]);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void DenseAttributeGuardRejectsBeforeFirstParserEvenWhenMalformed(bool malformed)
    {
        var xml = "<root " + string.Join(" ", Enumerable.Range(0, 2000).Select(i => $"a{i}=''")) + (malformed ? ">" : "/>");
        var bytes = Encoding.UTF8.GetBytes(xml);
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 1024 * 1024 });
        using var scope = context.Material.BeginControlledRead(context, default);
        Assert.Throws<InvalidDataException>(() => context.Material.GuardEntry("dense.xml", bytes, default));
        Assert.Equal(0, context.Material.ReadParserAdmissions); Assert.Equal(0, context.Material.ReadDomAllocations);
    }

    [Fact]
    public void RepeatedMalformedXmlPaysForEachParserRatherThanCachingARejectedProbe()
    {
        var bytes = Encoding.UTF8.GetBytes("<root " + string.Join(" ", Enumerable.Range(0, 300).Select(i => $"a{i}=''")) + ">");
        var allowance = 65536L + 512L * bytes.Length;
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = allowance * 2 + 65536 });
        using var scope = context.Material.BeginControlledRead(context, default);
        Assert.Throws<InvalidDataException>(() => context.Material.GuardEntry("bad.xml", bytes, default)); var copy = context.Material.CloneCopyBytes;
        Assert.Throws<InvalidDataException>(() => context.Material.GuardEntry("bad.xml", bytes, default));
        Assert.True(context.Material.CloneCopyBytes >= copy + allowance); Assert.Equal(2, context.Material.ReadParserAdmissions);
        Assert.Throws<InvalidDataException>(() => context.Material.GuardEntry("bad.xml", bytes, default));
        Assert.Equal(2, context.Material.ReadParserAdmissions);
    }

    [Fact]
    public void DenseAttributeCloneAndLongIdAreAdmittedBeforeConstruction()
    {
        var root = new XElement("TextObject", Enumerable.Range(0, 2000).Select(i => new XAttribute("vendor" + i, "x")), new XElement("TextCode", "text"));
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 65536 });
        using var scope = context.Material.BeginControlledRead(context, default);
        Assert.Throws<InvalidDataException>(() => context.Material.PrepareReadAnnotationClone([root], default));
        Assert.Equal(0, context.Material.ReadAnnotationClones);
        Assert.Throws<InvalidDataException>(() => context.Material.ReadConcat("annotation-", new string('x', 20000), default));
    }

    [Fact]
    public void WholeIndexSerializationAndNumericSplitFailBeforeTheirAllocationCheckpoints()
    {
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 65536 });
        using var scope = context.Material.BeginControlledRead(context, default);
        var index = new XDocument(new XElement("Annotations", new XAttribute("vendor", new string('x', 20000))));
        Assert.Throws<InvalidDataException>(() => OfdXmlEnvelope.SerializeForRead(index, context.Material, default));
        Assert.Equal(0, context.Material.ReadSourceStrings);
        Assert.Throws<InvalidDataException>(() => context.Material.PrepareReadNumericParse(string.Join(" ", Enumerable.Repeat("0", 10000)), default));
        Assert.Equal(0, context.Material.ReadNumericParses);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task HiddenAndVisibleStrictBoundaryValidationUsesAccountedParser(bool hidden)
    {
        var bytes = await WithAnnotations(hidden, string.Join(" ", Enumerable.Repeat("0", 1000)), false);
        using var context = new OfdControlledEditContext(new()); using var stream = new MemoryStream(bytes);
        var controlled = await new OfdReader().ReadControlledAsync(stream, new(), context);
        using var publicStream = new MemoryStream(bytes); var baseline = await new OfdReader().ReadAsync(publicStream);
        Assert.True(context.Material.ReadNumericParses >= 3);
        Assert.Equal(baseline.Pages[0].AnnotationAppearances.OfType<OfdRawElement>().Select(x => x.Xml),
            controlled.Pages[0].AnnotationAppearances.OfType<OfdRawElement>().Select(x => x.Xml));
    }

    [Fact]
    public async Task DenseVendorAnnotationAndUnknownIndexRemainPreservedThroughAccountedCallSites()
    {
        var bytes = await WithAnnotations(false, "0 0 10 10", true);
        using var context = new OfdControlledEditContext(new()); using var stream = new MemoryStream(bytes);
        var controlled = await new OfdReader().ReadControlledAsync(stream, new(), context);
        using var publicStream = new MemoryStream(bytes); var baseline = await new OfdReader().ReadAsync(publicStream);
        Assert.Equal(1, context.Material.ReadAnnotationClones); Assert.True(context.Material.ReadSourceStrings > 0);
        Assert.Equal(baseline.Pages[0].AnnotationAppearances.Count, controlled.Pages[0].AnnotationAppearances.Count);
        Assert.Contains(controlled.Pages[0].AnnotationAppearances.OfType<OfdRawElement>(), raw => raw.Xml.Contains("vendorIndex"));
        Assert.Contains(controlled.Pages[0].AnnotationAppearances.OfType<OfdTextElement>(), text => text.Text == "annotated");
    }

    private static async Task<byte[]> WithAnnotations(bool hidden, string boundary, bool dense)
    {
        var entries = new Dictionary<string, byte[]>();
        using (var input = new MemoryStream(await Source()))
        using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
            foreach (var entry in archive.Entries)
            { using var payload = new MemoryStream(); using var stream = entry.Open(); stream.CopyTo(payload); entries.Add(entry.FullName, payload.ToArray()); }
        var document = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])); var ns = document.Root!.Name.Namespace;
        var pageId = document.Descendants(ns + "Page").First().Attribute("ID")!.Value;
        document.Root.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml"));
        entries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        var index = new XElement(ns + "Annotations", dense ? new XAttribute("vendorIndex", new string('x', 1000)) : null,
            new XElement(ns + "Page", new XAttribute("PageID", pageId), new XElement(ns + "FileLoc", "Page.xml")));
        entries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes(index.ToString(SaveOptions.DisableFormatting));
        var primitive = new XElement(ns + "TextObject", new XAttribute("ID", "201"), new XAttribute("Boundary", "0 0 10 10"),
            new XAttribute("Font", "100"), new XAttribute("Size", "4"),
            dense ? Enumerable.Range(0, 500).Select(i => new XAttribute("vendor" + i, "x")) : null,
            new XElement(ns + "TextCode", new XAttribute("X", "0"), new XAttribute("Y", "4"), "annotated"));
        var annotations = new XElement(ns + "PageAnnot", new XElement(ns + "Annot", new XAttribute("ID", dense ? new string('1', 1000) : "200"),
            new XAttribute("Visible", hidden ? "false" : "true"), new XElement(ns + "Appearance", new XAttribute("Boundary", boundary), primitive)));
        entries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes(annotations.ToString(SaveOptions.DisableFormatting));
        using var output = new MemoryStream(); using (var archive = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var stream = archive.CreateEntry(entry.Key).Open(); stream.Write(entry.Value); }
        return output.ToArray();
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ControlledXmlRejectsDtdAndMalformedInputTerminallyBeforeDom(bool dtd)
    {
        var text = dtd ? "<!DOCTYPE r [<!ELEMENT r (" + string.Join(",", Enumerable.Repeat("a", 10000)) + ")>]><r/>" : "<root>";
        var bytes = Zip("OFD.xml", Encoding.UTF8.GetBytes(text));
        using var context = new OfdControlledEditContext(new()); using var input = new MemoryStream(bytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(input, new(), context));
        Assert.Equal(1, context.Material.ReadParserAdmissions); Assert.Equal(0, context.Material.ReadDomAllocations);
        Assert.Throws<InvalidOperationException>(context.AssertActive);
    }

    [Fact]
    public void PublicMaterialStillClassifiesMalformedPreservedXmlAsOpaque()
    {
        var material = new OfdModelMaterialization(new());
        material.GuardEntry("unknown.xml", Encoding.UTF8.GetBytes("<root>"), default);
        Assert.Contains("unknown.xml", material.OpaqueEntries); Assert.Equal(0, material.ReadParserAdmissions);
    }

    [Fact]
    public void EverySerializationWriterRequiresItsOwnFixedBufferAdmission()
    {
        var source = new XElement("Small", "value");
        using var context = new OfdControlledEditContext(new() { MaxMaterializedBytes = 50000 });
        using var scope = context.Material.BeginControlledRead(context, default);
        Assert.Throws<InvalidDataException>(() => OfdXmlEnvelope.SerializeForRead(source, context.Material, default));
        Assert.Equal(1, context.Material.ReadWriterAdmissions); Assert.Equal(0, context.Material.ReadSourceStrings);
    }

    private static async Task<byte[]> Source()
    {
        var source = new OfdDocumentPackage(); source.Fonts.Add(new OfdFontResource { Id = "100", FontName = "test", FileName = "test.ttf", Data = new byte[128] });
        var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        page.Elements.Add(new OfdTextElement { Text = "ORIGINAL", FontResourceId = "100", WidthMillimeters = 40, HeightMillimeters = 10 }); source.Pages.Add(page);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, output); return output.ToArray();
    }
    private static byte[] Zip(string name, byte[] payload)
    {
        using var output = new MemoryStream(); using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
        { using var entry = zip.CreateEntry(name, CompressionLevel.NoCompression).Open(); entry.Write(payload); }
        return output.ToArray();
    }
    private static void Set16(byte[] value, int at, ushort number) { value[at] = (byte)number; value[at + 1] = (byte)(number >> 8); }
    private static void Set32(byte[] value, int at, uint number) { for (var i = 0; i < 4; i++) value[at + i] = (byte)(number >> (8 * i)); }
    private sealed class ObservedStream(byte[] bytes, Action? action = null) : MemoryStream(bytes)
    {
        internal long ReadBytes;
        private bool called;
        private void Callback() { if (called) return; called = true; action?.Invoke(); }
        public override int Read(byte[] buffer, int offset, int count) { Callback(); var read = base.Read(buffer, offset, count); ReadBytes += read; return read; }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) { Callback(); return base.ReadAsync(buffer, offset, count, token); }
    }
}
