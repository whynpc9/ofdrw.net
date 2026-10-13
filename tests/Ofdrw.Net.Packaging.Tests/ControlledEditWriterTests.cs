using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class ControlledEditWriterTests
{
    [Fact]
    public async Task SourceMutationBetweenPrepareAndWriteRejectsBeforeDestinationAndDisposeRevokesAuthority()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); source.Options.Metadata.Title = "late";
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
        using var another = new OfdControlledEditContext(new OfdModelLimits());
        using var second = OfdOwnedEditCandidate.CloneForControlledEdit(source, another.CaptureSource(source), another);
        using var secondWrite = OfdPackageWriter.PrepareOwnedWrite(second); second.Dispose();
        await Assert.ThrowsAsync<InvalidOperationException>(() => secondWrite.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task OneContextCannotIssueTwoCandidatesSharingMutableCounters()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var first = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        Assert.Throws<InvalidOperationException>(() => OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context));
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareOwnedWrite(first));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ResourcePayloadPathConflictsRejectDuringPrepare(bool controlCollision)
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        page.Elements.Add(new OfdImageElement { Data = new byte[] { 1 }, WidthMillimeters = 10, HeightMillimeters = 10 });
        page.Elements.Add(new OfdImageElement { Data = new byte[] { 2 }, WidthMillimeters = 10, HeightMillimeters = 10 }); source.Pages.Add(page);
        using var original = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, original); original.Position = 0;
        source = await new OfdReader().ReadAsync(original);
        var resourcePath = "Doc_0/DocumentRes.xml"; var resource = XDocument.Load(new MemoryStream(source.PreservedEntries[resourcePath]));
        var files = resource.Descendants().Where(node => node.Name.LocalName == "MediaFile").ToArray();
        files[1].Value = controlCollision ? "/OFD.xml" : files[0].Value;
        source.PreservedEntries[resourcePath] = Encoding.UTF8.GetBytes(resource.ToString(SaveOptions.DisableFormatting)); source = await ReadEntries(source);
        if (!controlCollision) Assert.IsType<OfdImageElement>(source.Pages[0].Elements[1]).Data = new byte[] { 99 };
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareOwnedWrite(owned));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task RetainedSlotPreservesLaterNestedIdsAndRejectsMixedContent(bool mixed)
    {
        var source = await Source(); var path = source.Pages[0].SourceEntryPath!;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[path]));
        var textRoot = xml.Descendants().First(node => node.Name.LocalName == "TextObject");
        textRoot.Add(new XElement(XName.Get("Keep", "urn:vendor"), new XAttribute("ID", "901"), new XAttribute("RefID", "901")));
        if (mixed) textRoot.Elements().First(node => node.Name.LocalName == "TextCode").Add(new XComment("opaque"));
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)); source = await ReadEntries(source);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var text = owned.GetMapped(Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0])); var before = text.SourceXml;
        if (mixed)
        { Assert.Throws<NotSupportedException>(() => owned.ApplyTextSlots(text, [new(0, "changed", 1, 4), new(1, "", 8, 4)])); Assert.Equal(before, text.SourceXml); }
        else
        {
            owned.ApplyTextSlots(text, [new(0, "longer", 1, 4), new(1, "", 8, 4)]);
            using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream(); await prepared.WriteAsync(output);
            output.Position = 0; var readback = await new OfdReader().ReadAsync(output);
            var keep = XElement.Parse(Assert.IsType<OfdTextElement>(readback.Pages[0].Elements[0]).SourceXml!).Element(XName.Get("Keep", "urn:vendor"))!;
            Assert.Equal("901", keep.Attribute("ID")!.Value); Assert.Equal("901", keep.Attribute("RefID")!.Value);
        }
    }

    [Fact]
    public async Task HistoricalPagesAreDetachedIndependentOwners()
    {
        var source = await Source(); var historical = source.Pages[1]; source.Pages.RemoveAt(1);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var mapped = owned.GetMapped(historical); Assert.DoesNotContain(mapped, owned.Package.Pages);
        Assert.Same(mapped, owned.Package.SourceIdentity!.OriginalPages[1]); Assert.NotSame(historical.Elements, mapped.Elements);
        Assert.NotSame(source.SourceIdentity!.OriginalPages, owned.Package.SourceIdentity.OriginalPages);
        Assert.IsType<OfdRawElement>(historical.Elements[0]).Xml = "changed";
        Assert.NotEqual("changed", Assert.IsType<OfdRawElement>(mapped.Elements[0]).Xml);
    }

    [Fact]
    public async Task RetainedSlotsUpdateVariableLengthAndEmptyTextWithoutMovingOpaqueIds()
    {
        var source = await Source(); var original = Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0]);
        var originalXml = original.SourceXml; var originalEntries = source.PreservedEntries.ToDictionary(pair => pair.Key, pair => pair.Value.ToArray());
        using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        var text = owned.GetMapped(original); var root = XElement.Parse(text.SourceXml!); var ns = root.Name.Namespace;
        var id = text.ObjectId; var font = text.FontResourceId;
        owned.ApplyTextSlots(text, [new(0, "ABCDE", 1.125, 4, "1 1 1 1"), new(1, "", 8, 4)]);
        Assert.Equal("ABCDE", text.Text); Assert.Equal("", text.Runs[1].Text); Assert.Null(text.Runs[1].DeltaX);
        var changed = XElement.Parse(text.SourceXml!); Assert.Equal(2, changed.Elements(ns + "TextCode").Count());
        Assert.Equal(root.Attribute("ID")!.Value, changed.Attribute("ID")!.Value); Assert.Equal(root.Attribute("Font")!.Value, changed.Attribute("Font")!.Value);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream();
        var before = context.Material.CloneCopyBytes; await prepared.WriteAsync(output); Assert.True(context.Material.CloneCopyBytes > before);
        output.Position = 0; var readback = await new OfdReader().ReadAsync(output);
        var saved = Assert.IsType<OfdTextElement>(readback.Pages[0].Elements[0]); Assert.Equal("ABCDE", saved.Text);
        Assert.Equal(id, saved.ObjectId); Assert.Equal(font, saved.FontResourceId); Assert.Equal(1.125, saved.Runs[0].XMillimeters);
        var opaque = Assert.IsType<OfdRawElement>(readback.Pages[1].Elements[0]);
        var vendor = XElement.Parse(opaque.Xml); Assert.Equal("900", vendor.Attribute("ID")!.Value); Assert.Equal("900", vendor.Attribute("RefID")!.Value);
        Assert.Equal(originalXml, original.SourceXml); guard.AssertUnchanged(source, context);
        foreach (var pair in originalEntries) Assert.Equal(pair.Value, source.PreservedEntries[pair.Key]);
        await Assert.ThrowsAsync<InvalidOperationException>(() => prepared.WriteAsync(new MemoryStream()));
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareOwnedWrite(owned));
    }

    [Fact]
    public async Task WholeCloneIsWritableAfterCertificateAndContextDisposal()
    {
        var source = await Source(); source.Attachments.Add(new OfdAttachment { Id = "exact-id", Name = "keep", Data = new byte[] { 8 },
            CreationDate = DateTimeOffset.Parse("2020-01-02T03:04:05+08:00"), Visible = false, Usage = "keep" });
        var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context); var copy = owned.Package;
        owned.Dispose(); context.Dispose(); source.PreservedEntries.Clear(); source.Pages.Clear(); source.Attachments[0].Data[0] = 9;
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(copy, output); output.Position = 0;
        var readback = await new OfdReader().ReadAsync(output); Assert.Equal(2, readback.Pages.Count);
        Assert.Equal("exact-id", Assert.Single(readback.Attachments).Id); Assert.Equal(new byte[] { 8 }, readback.Attachments[0].Data);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SourceAndCandidateMutationRejectAtActualStageBoundaries(bool changeSource)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source); using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        if (changeSource)
        {
            source.PreservedEntries["opaque.bin"][0]++;
            Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareOwnedWrite(owned));
            Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareOwnedWrite(owned));
        }
        else
        {
            using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); owned.Package.Options.Metadata.Author = "after plan";
            using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task CancellationAndLowAggregateOutputConsumePreparedAttempts()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source); using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        using var output = new MemoryStream(); await Assert.ThrowsAsync<OperationCanceledException>(() => prepared.WriteAsync(output, cancelled.Token)); Assert.Equal(0, output.Length);
        await Assert.ThrowsAsync<InvalidOperationException>(() => prepared.WriteAsync(output));
        // Fund the independently metered parser/copy stages observed above,
        // then exhaust only the logical-output ledger before preparation.
        var ceiling = Math.Max(Math.Max(context.Material.CloneCopyBytes, context.Material.ValidationWorkBytes),
            Math.Max(context.Material.StorageBytes, context.Material.OutputUpperBoundBytes)) + 1_000_000;
        using var small = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = ceiling });
        var smallGuard = small.CaptureSource(source); using var limited = OfdOwnedEditCandidate.CloneForControlledEdit(source, smallGuard, small);
        small.Material.ChargeOutput(ceiling - 1);
        var quota = Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareOwnedWrite(limited));
        Assert.Contains("logical output", quota.Message);
        Assert.True(small.Material.OutputUpperBoundBytes >= ceiling - 1); // no reset
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task DestinationCallbacksCannotChangeFrozenEntriesAndAreDetected(bool sourceMutation)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source); using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        using var output = new CallbackStream(() => (sourceMutation ? source : owned.Package).Options.DocType = "mutated");
        await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.True(output.Length > 0);
        output.Position = 0; var loaded = await new OfdReader().ReadAsync(output); Assert.NotEqual("mutated", loaded.Options.DocType);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task SlotFailuresLeaveCompleteCandidateUnchangedAndRejectForeignRoots(int failure)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var text = owned.GetMapped(Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0])); var xml = text.SourceXml;
        if (failure == 0) Assert.Throws<InvalidDataException>(() => owned.ApplyTextSlots(text, [new(0, "changed", 0, 4), new(9, "bad", 1, 4)]));
        else if (failure == 1) Assert.Throws<NotSupportedException>(() => owned.ApplyTextSlots(new OfdTextElement { ObjectId = text.ObjectId, SourceXml = xml }, [new(0, "forged", 0, 4)]));
        else { text.Runs[0].Text = "stale"; Assert.Throws<InvalidDataException>(() => owned.ApplyTextSlots(text, [new(0, "changed", 0, 4), new(1, "", 1, 4)])); }
        Assert.Equal(xml, text.SourceXml); Assert.Equal("ABCA", text.Text);
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareOwnedWrite(owned));
    }

    [Fact]
    public async Task AllExistingSlotsMayBecomeEmptyWithExplicitAnchors()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var text = owned.GetMapped(Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0]));
        owned.ApplyTextSlots(text, [new(0, "", 1, 4), new(1, "", 8, 4)]);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        output.Position = 0; var readback = await new OfdReader().ReadAsync(output); var saved = Assert.IsType<OfdTextElement>(readback.Pages[0].Elements[0]);
        Assert.Equal("", saved.Text); Assert.Equal(2, saved.Runs.Count); Assert.All(saved.Runs, run => Assert.Null(run.DeltaX));
    }

    private sealed class CallbackStream(Action callback) : MemoryStream
    {
        private bool called;
        public override void Write(byte[] buffer, int offset, int count)
        { if (!called) { called = true; callback(); } base.Write(buffer, offset, count); }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
        { if (!called) { called = true; callback(); } return base.WriteAsync(buffer, token); }
    }

    private static async Task<OfdDocumentPackage> Source()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 };
        var text = new OfdTextElement { Text = "ABCA", XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 10 };
        text.Runs.Add(new OfdTextRun { Text = "AB", XMillimeters = 1, YMillimeters = 4, DeltaX = "2" });
        text.Runs.Add(new OfdTextRun { Text = "CA", XMillimeters = 8, YMillimeters = 4, DeltaX = "2" });
        page.Elements.Add(text); package.Pages.Add(page);
        var other = new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297 };
        other.Elements.Add(new OfdRawElement { LocalName = "Vendor", Xml = "<v:Vendor xmlns:v='urn:vendor' ID='900' RefID='900'><v:Keep/></v:Vendor>" }); package.Pages.Add(other);
        package.PreservedEntries.Add("opaque.bin", new byte[] { 1, 2, 3 });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); output.Position = 0;
        return await new OfdReader().ReadAsync(output);
    }

    private static async Task<OfdDocumentPackage> ReadEntries(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true)) foreach (var pair in package.PreservedEntries)
        { using var entry = zip.CreateEntry(pair.Key).Open(); await entry.WriteAsync(pair.Value); }
        stream.Position = 0; return await new OfdReader().ReadAsync(stream);
    }
}
