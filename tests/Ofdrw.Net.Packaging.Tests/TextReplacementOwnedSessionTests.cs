using Ofdrw.Net.Layout.Editing.Internal;

namespace Ofdrw.Net.Packaging.Tests;

// Trusted friend ownership tests use the existing deterministic fake metrics.
// They do not qualify a CLI provider or prove real-font/production acceptance.
public sealed class TextReplacementOwnedSessionTests
{
    [Fact]
    public async Task PrivatePreparationReturnsOnlyCopiedValuesAndRejectsLegacyConsumption()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var metrics = new TextReplacementTests.Metrics(); var report = session.PrepareOwnedWriter(metrics, "B", "--");
        Assert.True(report.Changed); Assert.Single(report.Replacements); Assert.Single(report.Statistics.Analyses);
        Assert.Equal(1, metrics.Opens); Assert.Equal(2, metrics.Measures); Assert.Equal(1, metrics.StatisticsCalls); Assert.Equal(1, metrics.Disposals);
        Assert.Throws<InvalidOperationException>(() => session.BoundSource);
        Assert.Throws<InvalidOperationException>(() => session.BoundContext);
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.Prepare(metrics, "B", "-"));
        Assert.Throws<InvalidOperationException>(() => session.PrepareOwnedWriter(metrics, "B", "-"));
        Assert.Equal("PreparedOwnedWriter", session.SnapshotOwned().Name);
    }

    [Fact]
    public async Task PrivateRouteStillDetectsMutationAfterActualProviderCallback()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        var source = session.BoundSource;
        var metrics = new TextReplacementTests.Metrics { Callback = phase => { if (phase == "measure") source.Options.Metadata.Title = "changed"; } };
        Assert.Throws<InvalidDataException>(() => session.PrepareOwnedWriter(metrics, "B", "--"));
        Assert.Equal(1, metrics.Opens); Assert.Equal(1, metrics.Disposals);
        Assert.Throws<InvalidOperationException>(() => session.BoundSource);
        Assert.Throws<InvalidOperationException>(() => session.WriteOwnedAsync().GetAwaiter().GetResult());
    }

    [Fact]
    public async Task NoOpStagesExactSnapshotThenReleasesPrivateInputBeforeCommit()
    {
        using var fixture = await Fixture.Create(); var original = await File.ReadAllBytesAsync(fixture.Input);
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Input);
        Assert.False(session.PrepareOwnedWriter(new TextReplacementTests.Metrics(), "B", "B").Changed);
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
        session.PrepareOwnedWriter(new TextReplacementTests.Metrics(), "B", "B"); session.ReserveOwnedAtomicMetadata(fixture.Stage, default);
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
