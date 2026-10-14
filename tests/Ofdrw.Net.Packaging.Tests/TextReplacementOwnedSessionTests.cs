using System.Collections;
using System.Reflection;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing.Internal;

namespace Ofdrw.Net.Packaging.Tests;

// Trusted friend ownership tests use the existing deterministic fake metrics.
// They do not qualify a CLI provider or prove real-font/production acceptance.
public sealed class TextReplacementOwnedSessionTests
{
    [Fact]
    public async Task PrivateWriteThroughReturnsCopiedValuesAfterRealWriteAndRejectsLegacyConsumption()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var metrics = new TextReplacementTests.Metrics(); var report = await session.PrepareAndWriteOwnedAsync(metrics, "B", "--");
        Assert.True(report.Changed); Assert.Single(report.Replacements); Assert.Single(report.Statistics.Analyses);
        Assert.Equal(1, metrics.Opens); Assert.Equal(2, metrics.Measures); Assert.Equal(1, metrics.StatisticsCalls); Assert.Equal(1, metrics.Disposals);
        Assert.Throws<InvalidOperationException>(() => session.BoundSource);
        Assert.Throws<InvalidOperationException>(() => session.BoundContext);
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.Prepare(metrics, "B", "-"));
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.PrepareAndWriteOwnedAsync(metrics, "B", "-"));
        Assert.Equal("WroteOwnedOutput", session.SnapshotOwned().Name);
        var files = session.SnapshotOwnedFiles();
        Assert.True(files.OutputSpoolPosition > 0); Assert.True(files.ActualSpoolWrittenBytes > 0); Assert.Equal(0, files.AtomicStagePosition);
        await session.ReadOwnedOutputAsync();
        var readback = Assert.IsType<OfdDocumentPackage>(PrivateField(session, "ownedReadback"));
        Assert.Equal("A--C", TextReplacementTests.Text(readback).Text);
        Assert.Equal("CapturedOwnedReadbackGuard", session.SnapshotOwned().Name);
        Assert.Throws<InvalidOperationException>(() => session.ReserveOwnedAtomicMetadata(fixture.Stage, default));
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.ReadOwnedOutputAsync());
    }

    [Fact]
    public async Task PrivateRouteStillDetectsMutationAfterActualProviderCallback()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var source = session.BoundSource;
        var metrics = new TextReplacementTests.Metrics { Callback = phase => { if (phase == "measure") source.Options.Metadata.Title = "changed"; } };
        await Assert.ThrowsAsync<InvalidDataException>(() => session.PrepareAndWriteOwnedAsync(metrics, "B", "--"));
        Assert.Equal(1, metrics.Opens); Assert.Equal(1, metrics.Disposals);
        Assert.Throws<InvalidOperationException>(() => session.BoundSource);
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.ReadOwnedOutputAsync());
    }

    [Fact]
    public async Task NoOpStagesExactSnapshotThenReleasesPrivateInputBeforeCommit()
    {
        using var fixture = await Fixture.Create(); var original = await File.ReadAllBytesAsync(fixture.Input);
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        Assert.False((await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "B")).Changed);
        Assert.Null(PrivateField(session, "ownedSpool")); Assert.Null(PrivateField(session, "ownedWriter"));
        Assert.Equal(0, session.SnapshotOwnedFiles().OutputSpoolPosition);
        session.ReserveOwnedAtomicMetadata(fixture.Stage, default);
        using (var stage = new FileStream(fixture.Stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            await session.CopyOwnedNoOpToStageAsync(stage, default);
        await session.ClaimOwnedCommitAsync(default);
        Assert.Equal(original, await File.ReadAllBytesAsync(fixture.Stage));
        Assert.Equal(original, await File.ReadAllBytesAsync(fixture.Input));
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.ClaimOwnedCommitAsync(default));
    }

    [Fact]
    public async Task SnapshotCleanupFailureBeforeCommitPreservesTheExistingTarget()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "B"); session.ReserveOwnedAtomicMetadata(fixture.Stage, default);
        using (var stage = new FileStream(fixture.Stage, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            await session.CopyOwnedNoOpToStageAsync(stage, default);
        var field = typeof(OfdTextReplacementFileSession).GetField("snapshotPath", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var snapshot = (string)field.GetValue(session)!;
        if (!OperatingSystem.IsWindows())
        {
            File.Move(snapshot, fixture.Input + ".snapshot"); Directory.CreateDirectory(snapshot);
            try
            {
                var failure = await Record.ExceptionAsync(() => session.ClaimOwnedCommitAsync(default));
                Assert.True(failure is IOException or UnauthorizedAccessException);
            }
            finally { Directory.Delete(snapshot); }
            Assert.Equal("existing target sentinel", await File.ReadAllTextAsync(fixture.Target));
        }
    }

    [Fact]
    public async Task BusyClaimRejectsConcurrentWriteReadValidateAndDispose()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        using var entered = new ManualResetEventSlim(); using var release = new ManualResetEventSlim();
        var metrics = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "measure") return;
            entered.Set();
            if (!release.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Test did not release the private provider.");
        } };
        var writing = Task.Run(() => session.PrepareAndWriteOwnedAsync(metrics, "B", "--"));
        try
        {
            Assert.True(entered.Wait(TimeSpan.FromSeconds(10)));
            Assert.Throws<InvalidOperationException>(() => { _ = session.PrepareAndWriteOwnedAsync(metrics, "B", "--"); });
            await Assert.ThrowsAsync<InvalidOperationException>(() => session.ReadOwnedOutputAsync());
            await Assert.ThrowsAsync<InvalidOperationException>(() => session.AssertInputUnchangedAsync());
            Assert.Throws<InvalidOperationException>(() => session.Dispose());
        }
        finally { release.Set(); }
        Assert.True((await writing).Changed);
        await session.ReadOwnedOutputAsync();
    }

    [Fact]
    public async Task LaterFreshSourceGuardRejectsMutationAndCleansWrittenOwnerAndSpool()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var source = session.BoundSource;
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--");
        var spool = Assert.IsType<OfdTextReplacementSpool>(PrivateField(session, "ownedSpool"));
        var spoolPath = Assert.IsType<string>(PrivateField(spool, "path"));
        source.Options.Metadata.Title = "mutated after guarded ZIP disposal";
        await Assert.ThrowsAsync<InvalidDataException>(() => session.ReadOwnedOutputAsync());
        Assert.False(File.Exists(spoolPath));
        Assert.Null(PrivateField(session, "ownedWriter")); Assert.Null(PrivateField(session, "ownedSpool"));
        Assert.True(session.SnapshotOwnedFiles().ActualSpoolWrittenBytes > 0);
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"));
    }

    [Fact]
    public async Task HostileExceptionDataObservesTerminalContextAndKeepsPrimaryDuringSpoolCleanup()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var context = session.BoundContext; var inspections = 0; var terminalObserved = true; string? spoolPath = null;
        var primary = new HostileDataException(() =>
        {
            inspections++;
            var terminalContext = false;
            try { context.AssertActive(); } catch (InvalidOperationException) { terminalContext = true; }
            terminalObserved &= terminalContext && PrivateField(session, "state") is 5;
            throw new InvalidOperationException("secondary diagnostic failure");
        });
        var metrics = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "dispose") return;
            var spool = Assert.IsType<OfdTextReplacementSpool>(PrivateField(session, "ownedSpool"));
            spoolPath = Assert.IsType<string>(PrivateField(spool, "path"));
            throw primary;
        } };
        Assert.Same(primary, await Record.ExceptionAsync(() => session.PrepareAndWriteOwnedAsync(metrics, "B", "--")));
        Assert.True(inspections > 0); Assert.True(terminalObserved); Assert.Equal(1, metrics.Disposals);
        Assert.NotNull(spoolPath); Assert.False(File.Exists(spoolPath));
        Assert.Null(PrivateField(session, "ownedWriter")); Assert.Null(PrivateField(session, "ownedSpool"));
    }

    private sealed class HostileDataException(Action inspect) : Exception("specific primary provider failure")
    {
        public override IDictionary Data { get { inspect(); return base.Data; } }
    }

    private static object? PrivateField(object owner, string name)
        => owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(owner);

    private sealed class Fixture(string directory) : IDisposable
    {
        private readonly string directory = directory;
        internal string Input => Path.Combine(directory, "input.ofd");
        internal string Target => Path.Combine(directory, "target.ofd");
        internal string Stage => Path.Combine(directory, "stage.tmp");
        internal static async Task<Fixture> Create()
        {
            var fixture = new Fixture(Path.Combine(Path.GetTempPath(), "ofdrw-owned-session-tests-" + Guid.NewGuid().ToString("N")));
            Directory.CreateDirectory(fixture.directory); await File.WriteAllTextAsync(fixture.Target, "existing target sentinel");
            var package = await TextReplacementTests.Read();
            using var file = File.Create(fixture.Input); await new OfdPackageWriter().WriteAsync(package, file); return fixture;
        }
        public void Dispose() => Directory.Delete(directory, true);
    }
}
