using System.Collections;
using System.IO.Compression;
using System.Reflection;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class ContinuousTextWriteThroughTests
{
    [Fact]
    public async Task ActualArchiveMatchesPreparedRouteAndWrittenOwnerRetainsFreshChecks()
    {
        var source = await Source(64 * 1024); var original = Text(source).SourceXml;
        using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        using var written = await OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output);
        Assert.False(output.Disposed); var before = context.Material.ValidationWorkBytes;
        written.AssertSourceAndCandidateUnchanged(); Assert.True(context.Material.ValidationWorkBytes > before + 128 * 1024);
        output.Position = 0; var actual = await new OfdReader().ReadControlledAsync(output, new(), context);
        Assert.Equal("ABCDE", Text(actual).Text); Assert.Equal(original, Text(source).SourceXml);
        using var other = new OfdControlledEditContext(new()); var otherGuard = other.CaptureSource(source);
        using var otherBatch = Batch(source, otherGuard, other);
        using var prepared = OfdPackageWriter.PrepareContinuousTextEdit(source, otherGuard, other, otherBatch);
        using var expected = new MemoryStream(); await prepared.WriteAsync(expected); EqualEntries(expected, output);
        Assert.Empty(typeof(OfdPackageWriter.WrittenTextEdit).GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic));
        Assert.DoesNotContain(typeof(OfdPackageWriter.WrittenTextEdit).GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic), m => m.Name.Contains("Write"));
    }

    [Fact]
    public async Task ClosedPrefixNeverAccessesDestinationBeforeSourceAndCandidateAdmission()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); source.PreservedEntries["opaque.bin"][0]++;
        using var output = new ProbeStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Equal(0, output.PropertyCalls); Assert.Equal(0, output.WriteCalls); AssertTerminal(context);
    }

    [Fact]
    public async Task WrapperQuotaBeforeCloneAndOutputHasNoUnfundedDestinationAccess()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context);
        context.Material.ChargeCopy(512L * 1024 * 1024 - context.Material.CloneCopyBytes - 1023, default);
        using var output = new ProbeStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Equal(0, context.Material.OwnedPayloadAllocations); Assert.Equal(0, output.PropertyCalls); AssertTerminal(context);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task PropertyWriteAndZipDisposeMutationsNeverReturnWrittenAuthority(int stage)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        Action mutate = () => source.PreservedEntries["opaque.bin"][0]++;
        if (stage == 0) output.OnFirstProperty = mutate;
        else if (stage == 1) output.OnFirstWrite = mutate;
        else output.OnCentralDirectory = mutate;
        OfdPackageWriter.WrittenTextEdit? written = null;
        await Assert.ThrowsAsync<InvalidDataException>(async () => written = await OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Null(written); Assert.True(output.WriteCalls > 0);
        Assert.True(stage != 2 || output.CentralDirectorySeen); Assert.False(output.Disposed); AssertTerminal(context);
    }

    [Fact]
    public async Task WritePrimaryIsTerminalBeforeCleanupAndSecondaryDisposeCannotReplaceIt()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        var primary = new NoDataIOException("write-primary"); var secondary = new IOException("cleanup-secondary");
        output.AsyncFailure = primary;
        output.OnCentralDirectory = () => { AssertTerminal(context); throw secondary; };
        var observed = await Assert.ThrowsAsync<NoDataIOException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Same(primary, observed); Assert.Equal(0, primary.DataReads);
        Assert.True(output.CentralDirectorySeen); AssertTerminal(context); Assert.False(output.Disposed);
    }

    [Fact]
    public async Task NormalZipDisposeFailureIsPrimaryAndIsNeverRetried()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        var primary = new NoDataIOException("zip-dispose-primary"); var attempts = 0;
        output.OnCentralDirectory = () => { attempts++; throw primary; };
        var observed = await Assert.ThrowsAsync<NoDataIOException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Same(primary, observed); Assert.Equal(1, attempts); Assert.Equal(0, primary.DataReads); AssertTerminal(context);
    }

    [Fact]
    public async Task DestinationPropertyFailureHasNoWriteAndLeavesBorrowedStreamOpen()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        var primary = new NoDataIOException("property-primary"); output.OnFirstProperty = () => throw primary;
        var observed = await Assert.ThrowsAsync<NoDataIOException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Same(primary, observed); Assert.Equal(0, output.WriteCalls); Assert.False(output.Disposed); AssertTerminal(context);
    }

    [Fact]
    public async Task AsyncSuspensionPublishesNoOwnerAndCompletionMutationIsChecked()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        output.Gate = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var pending = OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output);
        await output.AsyncEntered.Task.WaitAsync(TimeSpan.FromSeconds(10)); Assert.False(pending.IsCompleted);
        source.Options.Metadata.Title = "changed-during-await"; output.Gate.SetResult();
        await Assert.ThrowsAsync<InvalidDataException>(() => pending); AssertTerminal(context);
    }

    [Fact]
    public async Task CancellationAtActualAwaitTerminatesBeforeExternalCleanup()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream(); using var cancellation = new CancellationTokenSource();
        output.OnAsync = () => cancellation.Cancel();
        output.OnCentralDirectory = () => AssertTerminal(context);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output, cancellation.Token));
        AssertTerminal(context); Assert.False(output.Disposed);
    }

    [Fact]
    public async Task WriterCannotIssueAuthorityAfterCallbackFailsItsContext()
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream { OnFirstWrite = context.Fail };
        await Assert.ThrowsAsync<InvalidOperationException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        AssertTerminal(context);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task BatchForeignReplayOrDuplicateIsTerminalWithoutDestinationAccess(int error)
    {
        var source = await Source(); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var other = new OfdControlledEditContext(new()); var otherGuard = other.CaptureSource(source);
        using var batch = error == 0 ? Batch(source, otherGuard, other) : OfdRetainedTextEditBatch.Capture(source, guard, context,
            error == 2 ? [new(Text(source), Slots()), new(Text(source), Slots())] : [new(Text(source), Slots())]);
        if (error == 1) batch.Dispose();
        using var output = new ProbeStream();
        await Assert.ThrowsAsync<InvalidOperationException>(() => OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output));
        Assert.Equal(0, output.PropertyCalls); Assert.Equal("ABCA", Text(source).Text); AssertTerminal(context);
    }

    [Fact]
    public async Task WrittenHelperRehashesEveryTimeAndCannotReviveAfterMismatchOrDispose()
    {
        var source = await Source(64 * 1024); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        using var written = await OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output);
        for (var i = 0; i < 2; i++) { var before = context.Material.ValidationWorkBytes; written.AssertSourceAndCandidateUnchanged(); Assert.True(context.Material.ValidationWorkBytes > before + 128 * 1024); }
        source.PreservedEntries["opaque.bin"][0]++;
        Assert.Throws<InvalidDataException>(() => written.AssertSourceAndCandidateUnchanged()); AssertTerminal(context);
        Assert.Throws<InvalidOperationException>(() => written.AssertSourceAndCandidateUnchanged()); written.Dispose();
        Assert.Throws<InvalidOperationException>(() => written.AssertSourceAndCandidateUnchanged());
    }

    [Fact]
    public async Task ConcurrentValidationAndDisposeEitherSerializeOrRejectContentionWithoutRevival()
    {
        var source = await Source(8 * 1024 * 1024); using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = Batch(source, guard, context); using var output = new ProbeStream();
        var written = await OfdPackageWriter.WriteContinuousTextEditAsync(source, guard, context, batch, output);
        using var barrier = new Barrier(3); var before = context.Material.ValidationWorkBytes;
        var validation = Task.Run(() =>
        {
            barrier.SignalAndWait();
            try { written.AssertSourceAndCandidateUnchanged(); return true; }
            catch (InvalidOperationException) { return false; }
        });
        var disposal = Task.Run(() =>
        {
            barrier.SignalAndWait();
            try { written.Dispose(); return true; }
            catch (InvalidOperationException) { return false; }
        });
        barrier.SignalAndWait(); await Task.WhenAll(validation, disposal);
        Assert.True(validation.Result || disposal.Result);
        if (validation.Result) Assert.True(context.Material.ValidationWorkBytes > before + 16 * 1024 * 1024);
        written.Dispose(); Assert.Throws<InvalidOperationException>(() => written.AssertSourceAndCandidateUnchanged());
    }

    private static OfdRetainedTextEditBatch Batch(OfdDocumentPackage source, OfdSerializedGuard guard, OfdControlledEditContext context)
        => OfdRetainedTextEditBatch.Capture(source, guard, context, [new(Text(source), Slots())]);
    private static OfdTextSlotEdit[] Slots() => [new(0, "ABCDE", 1.125, 4, "1 1 1 1"), new(1, "", 8, 4)];
    private static OfdTextElement Text(OfdDocumentPackage source) => Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0]);
    private static void AssertTerminal(OfdControlledEditContext context) => Assert.Throws<InvalidOperationException>(() => context.AssertActive());
    private static async Task<OfdDocumentPackage> Source(int size = 3)
    {
        var package = new OfdDocumentPackage(); package.Options.EnableDeflateCompression = false;
        var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        var text = new OfdTextElement { Text = "ABCA", XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 10 };
        text.Runs.Add(new() { Text = "AB", XMillimeters = 1, YMillimeters = 4, DeltaX = "2" }); text.Runs.Add(new() { Text = "CA", XMillimeters = 8, YMillimeters = 4, DeltaX = "2" });
        page.Elements.Add(text); package.Pages.Add(page); package.PreservedEntries.Add("opaque.bin", new byte[size]);
        using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream); stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
    private static void EqualEntries(MemoryStream expected, MemoryStream actual)
    {
        expected.Position = actual.Position = 0; using var a = new ZipArchive(expected, ZipArchiveMode.Read, true); using var b = new ZipArchive(actual, ZipArchiveMode.Read, true);
        Assert.Equal(a.Entries.Select(e => e.FullName).Order().ToArray(), b.Entries.Select(e => e.FullName).Order().ToArray());
        foreach (var entry in a.Entries) { using var x = entry.Open(); using var y = b.GetEntry(entry.FullName)!.Open(); using var p = new MemoryStream(); using var q = new MemoryStream(); x.CopyTo(p); y.CopyTo(q); Assert.Equal(p.ToArray(), q.ToArray()); }
    }
    private sealed class NoDataIOException(string message) : IOException(message)
    {
        internal int DataReads;
        public override IDictionary Data { get { DataReads++; throw new InvalidOperationException("Exception.Data must not be read."); } }
    }
    private sealed class ProbeStream : MemoryStream
    {
        internal int PropertyCalls, WriteCalls;
        internal bool Disposed, CentralDirectorySeen;
        internal Action? OnFirstProperty, OnFirstWrite, OnCentralDirectory, OnAsync;
        internal Exception? AsyncFailure;
        internal TaskCompletionSource? Gate;
        internal readonly TaskCompletionSource AsyncEntered = new(TaskCreationOptions.RunContinuationsAsynchronously);
        private void Property() { PropertyCalls++; if (PropertyCalls == 1) OnFirstProperty?.Invoke(); }
        public override bool CanWrite { get { Property(); return base.CanWrite; } }
        public override bool CanSeek { get { Property(); return base.CanSeek; } }
        public override long Position { get { Property(); return base.Position; } set { Property(); base.Position = value; } }
        public override long Length { get { Property(); return base.Length; } }
        private void Before(ReadOnlySpan<byte> bytes)
        {
            WriteCalls++; if (WriteCalls == 1) OnFirstWrite?.Invoke();
            if (!CentralDirectorySeen && bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4b && bytes[2] == 1 && bytes[3] == 2)
            { CentralDirectorySeen = true; OnCentralDirectory?.Invoke(); }
        }
        public override void Write(byte[] buffer, int offset, int count) { Before(buffer.AsSpan(offset, count)); base.Write(buffer, offset, count); }
        public override void Write(ReadOnlySpan<byte> buffer) { Before(buffer); base.Write(buffer); }
        public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            OnAsync?.Invoke(); AsyncEntered.TrySetResult(); if (Gate is not null) await Gate.Task;
            token.ThrowIfCancellationRequested(); if (AsyncFailure is not null) throw AsyncFailure;
            Before(buffer.AsSpan(offset, count)); base.Write(buffer, offset, count);
        }
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
        {
            OnAsync?.Invoke(); AsyncEntered.TrySetResult(); if (Gate is not null) await Gate.Task;
            token.ThrowIfCancellationRequested(); if (AsyncFailure is not null) throw AsyncFailure;
            Before(buffer.Span); base.Write(buffer.Span);
        }
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }
}
