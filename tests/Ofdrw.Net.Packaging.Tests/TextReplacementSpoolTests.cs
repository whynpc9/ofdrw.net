using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing.Internal;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementSpoolTests
{
    [Fact]
    public async Task WritesRewritesAndReadsArePrepaidAndPositionReservationsNeverRefund()
    {
        using var context = new OfdControlledEditContext(new());
        using var spool = new OfdTextReplacementSpool(context, 128);
        var initialCopy = context.Material.CloneCopyBytes;
        await spool.WriteAsync(new byte[] { 1, 2, 3, 4 }, 0, 4);
        spool.Position = 1; spool.Write(new byte[] { 9, 8 }, 0, 2);
        spool.SetLength(2); spool.Position = 0;
        var read = new byte[4]; Assert.Equal(2, await spool.ReadAsync(read, 0, read.Length));
        Assert.Equal(new byte[] { 1, 9, 0, 0 }, read);
        Assert.Equal(4, spool.ReservedPosition); Assert.Equal(4, context.Material.OutputUpperBoundBytes);
        Assert.Equal(6, spool.ActualWrittenBytes); Assert.Equal(6, spool.ReservedWriteBytes);
        Assert.Equal(2, spool.ActualReadBytes); Assert.Equal(4, spool.ReservedReadBytes);
        Assert.Equal(initialCopy, context.Material.CloneCopyBytes);
    }

    [Theory]
    [InlineData("write")]
    [InlineData("seek")]
    [InlineData("setlength")]
    public void PositionLimitRejectsBeforeFileGrowthAndMakesTheContextTerminal(string operation)
    {
        using var context = new OfdControlledEditContext(new()); using var spool = new OfdTextReplacementSpool(context, 3);
        var error = Assert.Throws<InvalidDataException>(() =>
        {
            if (operation == "write") spool.Write(new byte[4], 0, 4);
            else if (operation == "seek") spool.Position = 4;
            else spool.SetLength(4);
        });
        Assert.Equal("ReplacementSpoolPositionLimit", error.Message);
        Assert.Equal(0, spool.ActualWrittenBytes); Assert.Equal(0, spool.ReservedPosition);
        Assert.Throws<InvalidOperationException>(() => context.AssertActive());
    }

    [Fact]
    public void ExhaustedWorkCannotWriteOrChangeTheReservedPosition()
    {
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1_000_000 });
        using var spool = new OfdTextReplacementSpool(context, 1_000_000);
        context.Material.ChargeWork(1_000_000, default);
        var failure = Assert.Throws<InvalidDataException>(() => spool.Write(new byte[1], 0, 1));
        Assert.True(OfdTraversalBudget.IsLimit(failure));
        Assert.Equal(0, spool.ActualWrittenBytes); Assert.Equal(0, spool.ReservedPosition);
    }

    [Fact]
    public async Task OriginalCancellationIsEnforcedEvenWhenTheIoCallSuppliesDefault()
    {
        using var context = new OfdControlledEditContext(new()); using var cancellation = new CancellationTokenSource();
        using var spool = new OfdTextReplacementSpool(context, 128, cancellation.Token);
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => spool.WriteAsync(new byte[1], 0, 1));
        Assert.Equal(0, spool.ActualWrittenBytes); Assert.Equal(0, spool.ReservedPosition);
        Assert.Throws<InvalidOperationException>(() => context.AssertActive());
    }

    [Fact]
    public void TinyCopyQuotaRejectsBeforeThePrivateFileIsCreated()
    {
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1 });
        var failure = Assert.Throws<InvalidDataException>(() => new OfdTextReplacementSpool(context, 128));
        Assert.True(OfdTraversalBudget.IsLimit(failure)); Assert.Equal(0, context.Material.OutputUpperBoundBytes);
        Assert.Throws<InvalidOperationException>(() => context.AssertActive());
    }

    [Fact]
    public async Task StageCopyChargesBothFilePositionsAndBindsExactBytesOnce()
    {
        var stage = Path.Combine(Path.GetTempPath(), "ofdrw-stage-copy-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            using var context = new OfdControlledEditContext(new()); using var spool = new OfdTextReplacementSpool(context, 1_000_000);
            var bytes = Enumerable.Range(0, 100_000).Select(i => (byte)i).ToArray();
            await spool.WriteAsync(bytes); var digest = System.Security.Cryptography.SHA256.HashData(bytes);
            using (var file = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                await spool.CopyToStageAsync(file, bytes.Length, digest);
            Assert.Equal(bytes, await File.ReadAllBytesAsync(stage));
            Assert.Equal(bytes.Length, spool.StageWrittenBytes); Assert.Equal(bytes.Length, spool.StageReservedPosition);
            Assert.Equal(2L * bytes.Length, context.Material.OutputUpperBoundBytes);
            Assert.Equal(bytes.Length, spool.ActualReadBytes);
            using var another = new FileStream(stage + ".second", FileMode.CreateNew, FileAccess.Write, FileShare.None);
            await Assert.ThrowsAsync<InvalidOperationException>(() => spool.CopyToStageAsync(another, bytes.Length, digest));
        }
        finally { File.Delete(stage); File.Delete(stage + ".second"); }
    }

    [Fact]
    public async Task HashMismatchMakesTheCopyTerminalAndNeverTouchesTheExistingTarget()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-stage-fail-test-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "target.ofd"); var stage = Path.Combine(directory, "stage.tmp");
        try
        {
            await File.WriteAllTextAsync(target, "existing target sentinel");
            using var context = new OfdControlledEditContext(new()); using var spool = new OfdTextReplacementSpool(context, 128);
            await spool.WriteAsync(new byte[] { 1, 2, 3 });
            using var file = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            var failure = await Assert.ThrowsAsync<InvalidDataException>(() => spool.CopyToStageAsync(file, 3, new byte[32]));
            Assert.Equal("ReplacementStageBindingMismatch", failure.Message); Assert.Equal(3, spool.StageWrittenBytes);
            Assert.Throws<InvalidOperationException>(() => context.AssertActive());
            Assert.Equal("existing target sentinel", await File.ReadAllTextAsync(target));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task StageCopyWorkQuotaFailsBeforeReadingOrWritingAnyByte()
    {
        var stage = Path.Combine(Path.GetTempPath(), "ofdrw-stage-quota-test-" + Guid.NewGuid().ToString("N"));
        try
        {
            const long limit = 1_000_000; using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = limit });
            using var spool = new OfdTextReplacementSpool(context, limit); var bytes = new byte[100]; await spool.WriteAsync(bytes);
            context.Material.ChargeWork(limit - context.Material.ValidationWorkBytes - 32, default);
            using var file = new FileStream(stage, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            var failure = await Assert.ThrowsAsync<InvalidDataException>(() => spool.CopyToStageAsync(file, bytes.Length, System.Security.Cryptography.SHA256.HashData(bytes)));
            Assert.True(OfdTraversalBudget.IsLimit(failure)); Assert.Equal(0, spool.ActualReadBytes); Assert.Equal(0, spool.StageWrittenBytes);
            Assert.Equal(0, file.Length);
        }
        finally { File.Delete(stage); }
    }
}
