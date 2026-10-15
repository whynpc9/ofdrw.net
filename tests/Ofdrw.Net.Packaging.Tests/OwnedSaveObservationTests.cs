using System.Security.Cryptography;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Tests;

// Actual concrete-file event checks; these numeric probes confer no save authority.
public sealed class OwnedSaveObservationTests
{
    [Fact]
    public async Task ActualFileEventsSeparateReservationsOriginalReadsAndTemporaryFilePeak()
    {
        var directory = Path.Combine(Path.GetTempPath(), "owned-stats-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        var sourcePath = Path.Combine(directory, "source.ofd"); var stagePath = Path.Combine(directory, "stage.ofd");
        var bytes = new byte[180_000]; new Random(17).NextBytes(bytes); await File.WriteAllBytesAsync(sourcePath, bytes);
        try
        {
            using var context = new OfdControlledEditContext(new()); context.Material.ChargeCopy(65_536, default);
            var telemetry = OfdOwnedIoTelemetry.CreateAfterPrepayment();
            using var input = new OfdAccountedTemporaryFile(context, 512L * 1024 * 1024, default, telemetry.Input);
            var actual = await input.CaptureOriginalAsync(sourcePath, bytes.Length, default);
            var captured = telemetry.CaptureActual(1, 512L * 1024 * 1024, context.Material);
            Assert.True(captured.Available); Assert.True(captured.InputSnapshotCreated); Assert.False(captured.OutputSpoolCreated);
            Assert.Equal(bytes.Length, captured.InputSnapshotOriginalReadActualBytes);
            Assert.Equal(bytes.Length + 1L, captured.InputSnapshotOriginalReadReservedBytes);
            Assert.Equal(0, captured.InputSnapshotReadActualBytes); Assert.Equal(bytes.Length, captured.InputSnapshotWriteActualBytes);
            Assert.Equal(bytes.Length, captured.InputSnapshotActualLengthHighWater); Assert.Equal(bytes.Length, captured.PeakActualFileLengths);
            Assert.Equal(1, captured.InputSnapshotFlushActualCount);
            var raw = await input.ObserveRawZipAsync(default); Assert.Equal(SHA256.HashData(bytes), raw.Sha256);
            using (var stage = new FileStream(stagePath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                telemetry.Stage.RecordOpened();
                await input.CopyToStageAsync(stage, actual.Length, actual.Sha256, default, telemetry.Stage);
                var copied = telemetry.CaptureActual(2, 512L * 1024 * 1024, context.Material);
                Assert.Equal(2L * bytes.Length, copied.PeakActualFileLengths);
                Assert.Equal(2L * bytes.Length, copied.PeakReservedFilePositions);
                Assert.Equal(bytes.Length, copied.StageWriteActualBytes); Assert.Equal(bytes.Length, copied.StageReservedHighWater);
                Assert.Equal(2L * bytes.Length, copied.InputSnapshotReadActualBytes);
                Assert.Equal(0, copied.LogicalWriterReservationU); Assert.True(copied.AllActualPositionsObserved);
                Assert.Equal(bytes.Length, captured.PeakActualFileLengths); // immutable prior snapshot
            }
            File.Delete(stagePath); telemetry.Stage.RecordRemoved(); input.Dispose();
            context.Fail(); var failed = telemetry.CaptureActual(17, 512L * 1024 * 1024, context.Material);
            Assert.True(failed.Available); Assert.Equal(0, failed.ActiveActualFileLengths); Assert.Equal(0, failed.ActiveReservedFilePositions);
            Assert.Equal(2L * bytes.Length, failed.PeakActualFileLengths); Assert.True(failed.InputSnapshotRemoved); Assert.True(failed.StageRemoved);
            Assert.Equal(2L * bytes.Length, failed.OutputUpperBoundBytes); // no refund after physical cleanup
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
    [Fact]
    public void FailureAndSnapshotExhaustionDoNotAllocateRequestQuotaOrReturnStaleCounters()
    {
        var ticket = OfdOwnedSaveAuthority.BeginFreshJob(new(), default);
        var initial = OfdOwnedSaveAuthority.CaptureStatistics(ticket);
        OfdOwnedSaveAuthority.Fail(ticket);
        var failed = OfdOwnedSaveAuthority.CaptureStatistics(ticket);
        Assert.True(failed.Available); Assert.Equal(initial.CloneCopyBytes, failed.CloneCopyBytes);
        Assert.NotEqual(initial.PhaseCode, failed.PhaseCode); Assert.False(failed.InputSnapshotCreated);
        for (var i = 2; i < 32; i++) Assert.True(OfdOwnedSaveAuthority.CaptureStatistics(ticket).Available);
        Assert.False(OfdOwnedSaveAuthority.CaptureStatistics(ticket).Available);
    }
}
