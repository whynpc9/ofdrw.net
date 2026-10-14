using System.IO.Compression;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class ContinuousTextEditTests
{
    [Fact]
    public async Task CapturedConcreteSlotsAreIndependentAndActualOutputMatchesGenericRoute()
    {
        var source = await Source(); var owner = Text(source);
        using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        var slots = Slots(); var inputs = new[] { new OfdRetainedTextEdit(owner, slots) };
        using var batch = OfdRetainedTextEditBatch.Capture(source, guard, context, inputs);
        slots[0] = new(0, "MUTATED", 99, 99); inputs[0] = new(new OfdTextElement(), Slots());
        using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        using var actual = new MemoryStream(); await prepared.WriteAsync(actual);
        using var genericContext = new OfdControlledEditContext(new OfdModelLimits()); var genericGuard = genericContext.CaptureSource(source);
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(source, genericGuard, genericContext);
        candidate.ApplyTextSlots(candidate.GetMapped(owner), Slots());
        using var genericPrepared = OfdPackageWriter.PrepareOwnedWrite(candidate);
        using var expected = new MemoryStream(); await genericPrepared.WriteAsync(expected);
        AssertEntriesEqual(expected, actual);
        actual.Position = 0; var readback = await new OfdReader().ReadControlledAsync(actual, new(), context);
        Assert.Equal("ABCDE", Text(readback).Text); Assert.Equal("", Text(readback).Runs[1].Text);
        Assert.Equal(1.125, Text(readback).Runs[0].XMillimeters);
        Assert.Equal("ABCA", owner.Text); guard.AssertUnchanged(source, context);
        Assert.Equal(source.PreservedEntries["opaque.bin"], readback.PreservedEntries["opaque.bin"]);
    }

    [Fact]
    public async Task ClosedFactoryAvoidsRepeatedPayloadPassesWithoutChangingOutput()
    {
        const int size = 1024 * 1024; var source = await Source(size); var owner = Text(source);
        using var generic = new OfdControlledEditContext(new OfdModelLimits()); var genericGuard = generic.CaptureSource(source);
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(source, genericGuard, generic);
        candidate.ApplyTextSlots(candidate.GetMapped(owner), Slots()); genericGuard.AssertUnchanged(source, generic);
        _ = candidate.CaptureCandidate(); using var genericPrepared = OfdPackageWriter.PrepareOwnedWrite(candidate);
        using var closed = new OfdControlledEditContext(new OfdModelLimits()); var closedGuard = closed.CaptureSource(source);
        using var batch = Batch(source, closedGuard, closed);
        using var closedPrepared = OfdPackageWriter.PrepareContinuousTextEdit(source, closedGuard, closed, batch);
        Assert.True(generic.Material.ValidationWorkBytes - closed.Material.ValidationWorkBytes > 3L * size);
        using var expected = new MemoryStream(); await genericPrepared.WriteAsync(expected);
        using var actual = new MemoryStream(); await closedPrepared.WriteAsync(actual); AssertEntriesEqual(expected, actual);
    }

    [Fact]
    public async Task SourceMutationAfterCaptureRejectsBeforeAnyOwnedPayloadAllocation()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); source.PreservedEntries["opaque.bin"][0]++;
        Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch));
        Assert.Equal(0, context.Material.OwnedPayloadAllocations); AssertTerminal(context);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ForeignOrDuplicateRootsCannotPublishPartialPatchedCandidate(bool duplicate)
    {
        var source = await Source(); var owner = Text(source);
        using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        var second = duplicate ? owner : Text(await Source());
        using var batch = OfdRetainedTextEditBatch.Capture(source, guard, context,
            [new(owner, Slots()), new(second, Slots())]);
        Assert.ThrowsAny<InvalidOperationException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch));
        Assert.Equal("ABCA", owner.Text); AssertTerminal(context);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task ForeignContextGuardOrSourceBindingRejects(int mismatch)
    {
        var source = await Source(); using var original = new OfdControlledEditContext(new OfdModelLimits()); var guard = original.CaptureSource(source);
        using var batch = Batch(source, guard, original);
        using var other = new OfdControlledEditContext(new OfdModelLimits()); var otherGuard = other.CaptureSource(source);
        var actualContext = mismatch == 0 ? other : original;
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareContinuousTextEdit(
            mismatch == 2 ? new OfdDocumentPackage() : source, mismatch == 1 ? otherGuard : guard, actualContext, batch));
        AssertTerminal(actualContext);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task DisposedBatchAndRepeatedCloneAreTerminal(bool repeatedClone)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context);
        if (repeatedClone)
        { using var prior = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context); }
        else batch.Dispose();
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch));
        AssertTerminal(context);
    }

    [Fact]
    public async Task BatchAndPreparedWriteAreOneShotAndDisposeRevokesClosedCandidate()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context);
        using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        using var output = new MemoryStream(); await prepared.WriteAsync(output);
        await Assert.ThrowsAsync<InvalidOperationException>(() => prepared.WriteAsync(new MemoryStream()));
        Assert.Throws<InvalidOperationException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch));
        AssertTerminal(context);
        using var secondContext = new OfdControlledEditContext(new OfdModelLimits()); var secondGuard = secondContext.CaptureSource(source);
        using var secondBatch = Batch(source, secondGuard, secondContext);
        var second = OfdPackageWriter.PrepareContinuousTextEdit(source, secondGuard, secondContext, secondBatch); second.Dispose();
        using var empty = new MemoryStream(); await Assert.ThrowsAsync<InvalidOperationException>(() => second.WriteAsync(empty)); Assert.Equal(0, empty.Length);
    }

    [Fact]
    public async Task ExistingCallerGapSourceGuardStillRejectsBeforeOutput()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        source.Options.Metadata.Title = "late";
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
        AssertTerminal(context);
    }

    [Fact]
    public async Task ActualDestinationCallbackSourceMutationRemainsRejected()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        using var output = new CallbackStream(() => source.PreservedEntries["opaque.bin"][0]++);
        await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.True(output.Called); AssertTerminal(context);
    }

    [Fact]
    public async Task FreshVerificationBeforeAndAfterWriteDoesNotExposeOrReviveAuthority()
    {
        var source = await Source(64 * 1024); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        var before = context.Material.ValidationWorkBytes; prepared.AssertSourceAndCandidateUnchanged();
        Assert.True(context.Material.ValidationWorkBytes > before + 128 * 1024);
        using var output = new MemoryStream(); await prepared.WriteAsync(output);
        before = context.Material.ValidationWorkBytes; prepared.AssertSourceAndCandidateUnchanged();
        Assert.True(context.Material.ValidationWorkBytes > before + 128 * 1024);
        prepared.Dispose(); Assert.Throws<InvalidOperationException>(() => prepared.AssertSourceAndCandidateUnchanged());
        await Assert.ThrowsAsync<InvalidOperationException>(() => prepared.WriteAsync(new MemoryStream()));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ActualPostWriteMutationIsFreshlyRejectedForSourceAndCandidate(bool mutateCandidate)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(candidate);
        using var output = new MemoryStream(); await prepared.WriteAsync(output);
        (mutateCandidate ? candidate.Package : source).PreservedEntries["opaque.bin"][0]++;
        Assert.Throws<InvalidDataException>(() => prepared.AssertSourceAndCandidateUnchanged()); AssertTerminal(context);
        Assert.Throws<InvalidOperationException>(() => prepared.AssertSourceAndCandidateUnchanged());
    }

    [Fact]
    public async Task VerificationAndDisposeRejectActiveWriteWithoutGrantingAnotherConsumption()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        using var output = new CallbackStream(() =>
        {
            Assert.Throws<InvalidOperationException>(() => prepared.AssertSourceAndCandidateUnchanged());
            Assert.Throws<InvalidOperationException>(() => prepared.Dispose());
        });
        await prepared.WriteAsync(output); Assert.True(output.Called); prepared.AssertSourceAndCandidateUnchanged();
        await Assert.ThrowsAsync<InvalidOperationException>(() => prepared.WriteAsync(new MemoryStream()));
    }

    [Fact]
    public async Task CancelledFreshVerificationTerminatesWrittenOwner()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch);
        using var output = new MemoryStream(); await prepared.WriteAsync(output); using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => prepared.AssertSourceAndCandidateUnchanged(canceled.Token)); AssertTerminal(context);
    }

    [Fact]
    public async Task WriterPlanningQuotaAfterPrivatePatchPublishesNoPartialOwner()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); context.Material.ChargeOutput(512L * 1024 * 1024 - 1);
        Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch));
        Assert.True(context.Material.OwnedPayloadAllocations > 0); Assert.Equal("ABCA", Text(source).Text); AssertTerminal(context);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task CancelOrUnfundedBatchCaptureIsTerminalBeforeClone(bool quota)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new OfdModelLimits()); var guard = context.CaptureSource(source);
        if (quota)
        {
            context.Material.ChargeCopy(512L * 1024 * 1024 - context.Material.CloneCopyBytes - 511, default);
            Assert.Throws<InvalidDataException>(() => OfdRetainedTextEditBatch.Capture(source, guard, context, [new(Text(source), Slots())]));
        }
        else
        {
            using var batch = Batch(source, guard, context); using var canceled = new CancellationTokenSource(); canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfdPackageWriter.PrepareContinuousTextEdit(source, guard, context, batch, canceled.Token));
        }
        Assert.Equal(0, context.Material.OwnedPayloadAllocations); AssertTerminal(context);
    }

    private static OfdRetainedTextEditBatch Batch(OfdDocumentPackage source, OfdSerializedGuard guard, OfdControlledEditContext context)
        => OfdRetainedTextEditBatch.Capture(source, guard, context, [new(Text(source), Slots())]);
    private static OfdTextSlotEdit[] Slots() => [new(0, "ABCDE", 1.125, 4, "1 1 1 1"), new(1, "", 8, 4)];
    private static OfdTextElement Text(OfdDocumentPackage source) => Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0]);
    private static void AssertTerminal(OfdControlledEditContext context) => Assert.Throws<InvalidOperationException>(() => context.AssertActive());

    private static async Task<OfdDocumentPackage> Source(int payload = 3)
    {
        var package = new OfdDocumentPackage(); package.Options.EnableDeflateCompression = false;
        var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        var text = new OfdTextElement { Text = "ABCA", XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 10 };
        text.Runs.Add(new OfdTextRun { Text = "AB", XMillimeters = 1, YMillimeters = 4, DeltaX = "2" });
        text.Runs.Add(new OfdTextRun { Text = "CA", XMillimeters = 8, YMillimeters = 4, DeltaX = "2" });
        page.Elements.Add(text); package.Pages.Add(page); package.PreservedEntries.Add("opaque.bin", new byte[payload]);
        using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream); stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
    private static void AssertEntriesEqual(MemoryStream expected, MemoryStream actual)
    {
        expected.Position = actual.Position = 0;
        using var a = new ZipArchive(expected, ZipArchiveMode.Read, true); using var b = new ZipArchive(actual, ZipArchiveMode.Read, true);
        Assert.Equal(a.Entries.Select(e => e.FullName).Order().ToArray(), b.Entries.Select(e => e.FullName).Order().ToArray());
        foreach (var entry in a.Entries)
        { using var x = entry.Open(); using var y = b.GetEntry(entry.FullName)!.Open(); using var p = new MemoryStream(); using var q = new MemoryStream(); x.CopyTo(p); y.CopyTo(q); Assert.Equal(p.ToArray(), q.ToArray()); }
    }
    private sealed class CallbackStream(Action callback) : MemoryStream
    {
        internal bool Called { get; private set; }
        private void Invoke() { if (!Called) { Called = true; callback(); } }
        public override void Write(byte[] buffer, int offset, int count) { Invoke(); base.Write(buffer, offset, count); }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
        { Invoke(); return base.WriteAsync(buffer, token); }
    }
}
