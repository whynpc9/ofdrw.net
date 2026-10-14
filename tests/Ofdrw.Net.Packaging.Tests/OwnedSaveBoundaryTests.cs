using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class OwnedSaveBoundaryTests
{
    [Fact]
    public async Task SharedSnapshotHashAndStageCopyUseActualBytesAndCumulativeFees()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-owned-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        var original = Path.Combine(directory, "original.ofd"); var stagePath = Path.Combine(directory, "stage.ofd");
        var bytes = Enumerable.Range(0, 180_000).Select(i => (byte)(i * 31)).ToArray(); await File.WriteAllBytesAsync(original, bytes);
        try
        {
            using var context = new OfdControlledEditContext(new());
            using var spool = new OfdAccountedTemporaryFile(context, 512L * 1024 * 1024);
            var before = context.Material.ValidationWorkBytes;
            var captured = await spool.CaptureOriginalAsync(original, bytes.Length, default);
            Assert.Equal(bytes.Length, captured.Length); Assert.Equal(SHA256.HashData(bytes), captured.Sha256);
            Assert.Equal(bytes.Length, spool.OriginalActualReadBytes); Assert.Equal(bytes.Length, spool.OriginalReservedReadBytes);
            Assert.Equal(bytes.Length, spool.ActualWrittenBytes); Assert.Equal(bytes.Length, spool.ReservedPosition);
            Assert.Equal(3L * bytes.Length + 81, context.Material.ValidationWorkBytes - before);
            before = context.Material.ValidationWorkBytes;
            var observed = await spool.ObserveRawZipAsync(default);
            Assert.Equal(captured.Sha256, observed.Sha256); Assert.Equal(bytes.Length, observed.Length);
            Assert.Equal(2L * bytes.Length + 64, context.Material.ValidationWorkBytes - before);
            using (var stage = new FileStream(stagePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                before = context.Material.ValidationWorkBytes;
                await spool.CopyToStageAsync(stage, captured.Length, captured.Sha256);
                Assert.Equal(3L * bytes.Length + 96, context.Material.ValidationWorkBytes - before);
                Assert.Equal(bytes.Length, spool.StageWrittenBytes); Assert.Equal(bytes.Length, spool.StageReservedPosition);
            }
            Assert.Equal(bytes, await File.ReadAllBytesAsync(stagePath));
            Assert.Equal(2L * bytes.Length, context.Material.OutputUpperBoundBytes);
            Assert.Equal(0, context.Material.StorageBytes);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
    [Fact]
    public async Task CurrentInputBindingDetectsSameLengthTamperAndTerminatesContext()
    {
        var path = Path.Combine(Path.GetTempPath(), "ofdrw-owned-" + Guid.NewGuid().ToString("N"));
        var bytes = Encoding.UTF8.GetBytes(new string('A', 8192)); await File.WriteAllBytesAsync(path, bytes);
        try
        {
            using var context = new OfdControlledEditContext(new()); var expected = SHA256.HashData(bytes);
            bytes[bytes.Length / 2] ^= 1; await File.WriteAllBytesAsync(path, bytes);
            await Assert.ThrowsAsync<InvalidDataException>(() => OfdAccountedTemporaryFile.ValidateOriginalPathAsync(context, path, bytes.Length, expected, default));
            Assert.Throws<InvalidOperationException>(context.AssertActive);
        }
        finally { File.Delete(path); }
    }
    [Fact]
    public void IssuersFromOtherAssembliesFailBeforeTheyCanAuthorizeAnyWork()
    {
        var ticket = OfdOwnedSaveAuthority.BeginFreshJob(new(), default);
        Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.BindHost(ticket, new LocalHost()));
        Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.ReserveBootstrap(ticket, 1, default));
        var noopTicket = OfdOwnedSaveAuthority.BeginFreshJob(new(), default);
        Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.BindNoopKernel(noopTicket, new LocalNoop()));
        Assert.Throws<InvalidOperationException>(() => OfdOwnedSaveAuthority.ReserveBootstrapWork(noopTicket, 1, default));
    }
    [Fact]
    public void CurrentInventoryFactsAreConsumedAndCannotReplaceALaterFreshAssert()
    {
        var bytes = new byte[2048]; var package = new OfdDocumentPackage(); package.PreservedEntries.Add("opaque.bin", bytes);
        using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(package);
        var before = context.Material.ValidationWorkBytes;
        var evidence = guard.AssertUnchangedAndObservePayloads(package, context, default);
        Assert.True(context.Material.ValidationWorkBytes >= before + bytes.Length);
        Assert.Equal(SHA256.HashData(bytes), Convert.FromHexString(evidence.RequireObserved(package, context, bytes, default).Sha256Hex));
        evidence.CompletePrivateIssuance();
        Assert.Throws<InvalidDataException>(() => evidence.RequireObserved(package, context, bytes, default));
        bytes[4]++;
        Assert.Throws<InvalidDataException>(() => guard.AssertUnchanged(package, context));
        Assert.Throws<InvalidOperationException>(context.AssertActive);
    }
    [Fact]
    public void AffectedCoverageIncludesUnchangedNeighborsInActualPageOrder()
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage();
        var first = new OfdTextElement { Text = "A" }; var second = new OfdTextElement { Text = "B" };
        page.Elements.Add(first); page.Elements.Add(second); source.Pages.Add(page);
        using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source);
        using var batch = OfdRetainedTextEditBatch.Capture(source, guard, context,
            new[] { new OfdRetainedTextEdit(first, new[] { new OfdTextSlotEdit(0, "C", 0, 4, null, null) }) });
        batch.RequireCompleteAffectedRoots(new[] { first, second }, default);
        Assert.Throws<InvalidDataException>(() => batch.RequireCompleteAffectedRoots(new[] { first }, default));
        Assert.Throws<InvalidDataException>(() => batch.RequireCompleteAffectedRoots(new[] { second, first }, default));
        Assert.Throws<InvalidDataException>(() => batch.RequireCompleteAffectedRoots(new[] { first, first }, default));
    }
    [Fact]
    public void CapabilityOwnersHavePrivateConstructorsAndNoPublicMutablePayloadAccessors()
    {
        foreach (var type in new[] { typeof(OfdOwnedInputSnapshot), typeof(OfdOwnedWrittenSpool), typeof(OfdActualReadback),
            typeof(OfdVerifiedReadback), typeof(OfdOwnedAtomicStage), typeof(OfdVerifiedStage), typeof(OfdOwnedNoopCopyClaim),
            typeof(OfdOwnedJobTicket), typeof(OfdOwnedSnapshotCapture), typeof(OfdOwnedNoopScope) })
        {
            Assert.All(type.GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic), ctor => Assert.True(ctor.IsPrivate));
            Assert.Empty(type.GetProperties(BindingFlags.Instance | BindingFlags.Public));
            Assert.Empty(type.GetFields(BindingFlags.Instance | BindingFlags.Public));
        }
    }
    private sealed class LocalHost : OfdOwnedHostIssuer { }
    private sealed class LocalNoop : OfdOwnedNoopKernelIssuer { }
}
