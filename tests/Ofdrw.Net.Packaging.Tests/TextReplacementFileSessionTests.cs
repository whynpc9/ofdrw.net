using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementFileSessionTests
{
    [Fact]
    public async Task ActualFileReaderPreparationAndWriterShareExactlyOneContext()
    {
        using var fixture = await Fixture.Create();
        var original = await File.ReadAllBytesAsync(fixture.Path);
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        var initial = session.InitialReadStatistics;
        Assert.True(initial.StorageBytes > 0); Assert.True(initial.ValidationWorkBytes > 4 * original.Length);
        Assert.Equal(2L * original.Length, session.RawReadBytes);
        Assert.Equal(session.RawReadBytes, session.RawHashBytes);
        var provider = new TextReplacementTests.Metrics();
        var stats = session.Prepare(provider, "B", "--");
        Assert.Single(stats.Analyses); Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Disposals);
        Assert.All(stats.Phases, row => Assert.True(row.ValidationWorkBytes >= initial.ValidationWorkBytes));
        using var save = session.BeginSave(); using var use = save.UseOperation(); var operation = use.Operation;
        Assert.Same(session.BoundContext, operation.Context); Assert.Same(session.BoundSource, operation.Source);
        Assert.Equal("ABC", TextReplacementTests.Text(session.BoundSource).Text);
        Assert.Equal("A--C", TextReplacementTests.Text(operation.Candidate!.Package).Text);
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.Prepare(provider, "B", "-"));
        using var writer = OfdPackageWriter.PrepareOwnedWrite(operation.Candidate);
        using var output = new MemoryStream(); await writer.WriteAsync(output); output.Position = 0;
        var actual = await new OfdReader().ReadAsync(output);
        Assert.Equal("A--C", TextReplacementTests.Text(actual).Text);
        use.Dispose(); await save.AssertInputUnchangedAsync();
        Assert.Equal(original, await File.ReadAllBytesAsync(fixture.Path));
    }

    [Theory]
    [InlineData("source")]
    [InlineData("context")]
    [InlineData("token")]
    public async Task ForeignBindingFailsOnlyTheOwnedSessionAndCannotRetry(string mismatch)
    {
        using var fixture = await Fixture.Create();
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        using var foreign = new OfdControlledEditContext(new());
        using var cancellation = new CancellationTokenSource();
        Assert.Throws<InvalidOperationException>(() => OfdTextReplacementOperation.ForFileSession(session,
            mismatch == "source" ? new OfdDocumentPackage() : session.BoundSource,
            mismatch == "context" ? foreign : session.BoundContext,
            mismatch == "token" ? cancellation.Token : default));
        foreign.AssertActive();
        Assert.Throws<InvalidOperationException>(() => session.BoundContext.AssertActive());
        var provider = new TextReplacementTests.Metrics();
        Assert.Throws<InvalidOperationException>(() => session.Prepare(provider, "B", "-"));
        Assert.Equal(0, provider.Opens);
    }

    [Fact]
    public async Task OriginalCancellationRemainsBindingThroughSaveAndNeverRestarts()
    {
        using var fixture = await Fixture.Create(); using var cancellation = new CancellationTokenSource();
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path, token: cancellation.Token);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "-", token: cancellation.Token);
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.BoundContext.AssertActive());
    }

    [Fact]
    public async Task PreparationFailureIsTerminalAndProviderCleanupOccursExactlyOnce()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        var provider = new TextReplacementTests.Metrics { Callback = phase => { if (phase == "measure") throw new InvalidDataException("measurement failed"); } };
        Assert.Throws<InvalidDataException>(() => session.Prepare(provider, "B", "-"));
        Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Disposals);
        Assert.Throws<InvalidOperationException>(() => session.Prepare(provider, "B", "-"));
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
    }

    [Fact]
    public async Task PublicGapSourceMutationStillFailsExistingWriterGuard()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "-");
        session.BoundSource.Options.Metadata.Title = "mutated after preparation";
        using var save = session.BeginSave(); using var use = save.UseOperation(); var operation = use.Operation;
        Assert.Throws<InvalidDataException>(() => OfdPackageWriter.PrepareOwnedWrite(operation.Candidate!));
        Assert.Throws<InvalidOperationException>(() => session.BoundContext.AssertActive());
    }

    [Fact]
    public async Task ReplacedPathCannotPassByHashingOnlyTheRetainedOldHandle()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        File.Move(fixture.Path, fixture.Path + ".old");
        await File.WriteAllTextAsync(fixture.Path, "replacement at original pathname");
        await Assert.ThrowsAsync<InvalidDataException>(() => session.AssertInputUnchangedAsync());
        Assert.Throws<InvalidOperationException>(() => session.BoundContext.AssertActive());
        Assert.Equal("replacement at original pathname", await File.ReadAllTextAsync(fixture.Path));
    }

    [Fact]
    public async Task TinyQuotaRejectsBeforeOpeningMissingInputOrAllocatingOwnedBuffer()
    {
        var options = new OfdTextReplacementOptions(new(maxMaterializedBytes: 1));
        var absent = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".ofd");
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() => OfdTextReplacementFileSession.OpenAsync(absent, options));
        Assert.True(OfdTraversalBudget.IsLimit(failure)); Assert.Single(failure.Data); Assert.False(File.Exists(absent));
    }

    [Fact]
    public async Task ProviderReentryCannotValidateOrDisposeTheBusySession()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        var rejected = 0;
        var provider = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "measure") return;
            Assert.Throws<InvalidOperationException>(() => session.AssertInputUnchangedAsync().GetAwaiter().GetResult());
            Assert.Throws<InvalidOperationException>(() => session.Dispose());
            Assert.Throws<InvalidOperationException>(() => session.Prepare(new TextReplacementTests.Metrics(), "B", "-"));
            rejected++;
        } };
        session.Prepare(provider, "B", "-"); Assert.Equal(2, rejected);
        session.BoundContext.AssertActive(); await session.AssertInputUnchangedAsync();
    }

    [Fact]
    public async Task SimultaneousRawChecksEitherSerializeOrRejectWithoutCorruptingCounters()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        var before = session.RawReadBytes; var length = new FileInfo(fixture.Path).Length;
        var checking = session.AssertInputUnchangedAsync();
        var competing = await Record.ExceptionAsync(() => session.AssertInputUnchangedAsync());
        if (competing is not null) Assert.IsType<InvalidOperationException>(competing);
        await checking; session.BoundContext.AssertActive();
        Assert.Equal(before + (competing is null ? 2 : 1) * length, session.RawReadBytes);
        Assert.Equal(session.RawReadBytes, session.RawHashBytes);
    }

    [Fact]
    public async Task ActiveSaveLeaseRejectsDisposeAndRawReentryUntilConsumptionCompletes()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "-");
        using (var save = session.BeginSave())
        {
            Assert.Throws<InvalidOperationException>(() => session.Dispose());
            await Assert.ThrowsAsync<InvalidOperationException>(() => session.AssertInputUnchangedAsync());
            using (var use = save.UseOperation())
            {
                Assert.Throws<InvalidOperationException>(() => save.UseOperation());
                await Assert.ThrowsAsync<InvalidOperationException>(() => save.AssertInputUnchangedAsync());
                Assert.Throws<InvalidOperationException>(() => save.Dispose());
                Assert.Same(session.BoundContext, use.Operation.Context);
            }
            await save.AssertInputUnchangedAsync();
        }
        Assert.Throws<InvalidOperationException>(() => session.BeginSave());
        Assert.Throws<InvalidOperationException>(() => session.Prepare(new TextReplacementTests.Metrics(), "B", "-"));
    }

    [Fact]
    public async Task NoOpCopiesExactBoundSnapshotEvenWhenOriginalPathIsReplaced()
    {
        using var fixture = await Fixture.Create(); var original = await File.ReadAllBytesAsync(fixture.Path);
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "B");
        using var save = session.BeginSave();
        File.Move(fixture.Path, fixture.Path + ".old"); await File.WriteAllTextAsync(fixture.Path, "unbound replacement file");
        var staged = fixture.Path + ".staged";
        using (var destination = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            await save.CopyNoOpSnapshotAsync(destination);
        Assert.Equal(original, await File.ReadAllBytesAsync(staged));
        Assert.Equal(original.Length, session.NoOpSnapshotCopyBytes);
        Assert.Equal(2L * original.Length, session.BoundContext.Material.OutputUpperBoundBytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => save.AssertInputUnchangedAsync());
    }

    [Fact]
    public async Task ExternalAbaCounterexampleCannotChangeTheOwnedSnapshotUsedByNoOp()
    {
        using var fixture = await Fixture.Create();
        var sourceZip = await File.ReadAllBytesAsync(fixture.Path);
        static byte[] Repack(byte[] source, bool mutate)
        {
            using var original = new System.IO.Compression.ZipArchive(new MemoryStream(source), System.IO.Compression.ZipArchiveMode.Read);
            using var result = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(result, System.IO.Compression.ZipArchiveMode.Create, true))
            {
            foreach (var entry in original.Entries)
            {
                using var read = entry.Open(); using var data = new MemoryStream(); read.CopyTo(data);
                var bytes = data.ToArray();
                if (mutate && entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal))
                    bytes = System.Text.Encoding.UTF8.GetBytes(System.Text.Encoding.UTF8.GetString(bytes).Replace(">ABC<", ">AXC<", StringComparison.Ordinal));
                var added = zip.CreateEntry(entry.FullName, System.IO.Compression.CompressionLevel.NoCompression);
                added.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var write = added.Open(); write.Write(bytes);
            }
            // Keep the external file larger than FileStream's read cache so
            // seeking to zero for the final hash really rereads the file.
            var padding = zip.CreateEntry("Resources/aba-padding.bin", System.IO.Compression.CompressionLevel.NoCompression);
            padding.LastWriteTime = new DateTimeOffset(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);
            using var paddingStream = padding.Open(); paddingStream.Write(new byte[16_384]);
            }
            return result.ToArray();
        }
        var a = Repack(sourceZip, false); var b = Repack(sourceZip, true); Assert.Equal(a.Length, b.Length); Assert.NotEqual(a, b);
        await File.WriteAllBytesAsync(fixture.Path, a);
        // Unix native writes bypass the cooperative sharing lock used by the
        // rejected FileShare.Read algorithm. On Windows, where sharing is
        // enforced by the kernel, explicitly model an externally mutable file.
        var unix = !OperatingSystem.IsWindows();
        void Overwrite(byte[] bytes)
        {
            if (!unix) { File.WriteAllBytes(fixture.Path, bytes); return; }
            var fd = NativeOpen(fixture.Path, 1); Assert.True(fd >= 0);
            try { Assert.Equal(bytes.Length, NativeWrite(fd, bytes, (nuint)bytes.Length, 0).ToInt64()); }
            finally { Assert.Equal(0, NativeClose(fd)); }
        }
        using (var oldHandle = new FileStream(fixture.Path, FileMode.Open, FileAccess.Read, unix ? FileShare.Read : FileShare.ReadWrite))
        {
            var first = System.Security.Cryptography.SHA256.HashData(oldHandle);
            Overwrite(b); oldHandle.Position = 0;
            using var context = new OfdControlledEditContext(new());
            var wrong = await new OfdReader().ReadControlledAsync(oldHandle, new(), context);
            Assert.Equal("AXC", TextReplacementTests.Text(wrong).Text);
            Overwrite(a); oldHandle.Position = 0;
            Assert.Equal(first, System.Security.Cryptography.SHA256.HashData(oldHandle));
        }
        using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "B"); using var save = session.BeginSave();
        await File.WriteAllBytesAsync(fixture.Path, b);
        var staged = fixture.Path + ".aba-staged";
        using (var destination = new FileStream(staged, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            await save.CopyNoOpSnapshotAsync(destination);
        await File.WriteAllBytesAsync(fixture.Path, a); await save.AssertInputUnchangedAsync();
        Assert.Equal("ABC", TextReplacementTests.Text(session.BoundSource).Text);
        Assert.Equal(a, await File.ReadAllBytesAsync(staged));
    }

    [Fact]
    public async Task OperationUseAdmissionFailureIsTerminalAndCanStillReleaseItsSaveLease()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        session.Prepare(new TextReplacementTests.Metrics(), "B", "-"); using var save = session.BeginSave();
        var material = session.BoundContext.Material;
        material.ChargeCopy(session.Options.Limits.MaxMaterializedBytes - material.CloneCopyBytes, default);
        var failure = Assert.Throws<InvalidDataException>(() => save.UseOperation());
        Assert.True(OfdTraversalBudget.IsLimit(failure));
        Assert.Throws<InvalidOperationException>(() => session.BoundContext.AssertActive());
        Assert.Throws<InvalidOperationException>(() => save.UseOperation());
    }

    [Fact]
    public async Task FailedPreparationKeepsItsPrimaryErrorEvenWhenSnapshotDeletionFails()
    {
        using var fixture = await Fixture.Create(); var session = await OfdTextReplacementFileSession.OpenAsync(fixture.Path);
        var field = typeof(OfdTextReplacementFileSession).GetField("snapshotPath", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var path = (string)field.GetValue(session)!; var backup = fixture.Path + ".snapshot";
        var moved = false;
        try
        {
            // Unix namespace mutation of a held file is possible independently
            // of sharing flags. Inject a directory at the exact owned name to
            // make File.Delete fail, without changing any budget or source.
            if (!OperatingSystem.IsWindows())
            { File.Move(path, backup); Directory.CreateDirectory(path); moved = true; }
            var primary = new InvalidDataException("specific primary provider failure");
            var provider = new TextReplacementTests.Metrics { Callback = phase => { if (phase == "measure") throw primary; } };
            Assert.Same(primary, Assert.Throws<InvalidDataException>(() => session.Prepare(provider, "B", "-")));
            session.Dispose();
        }
        finally
        {
            session.Dispose();
            if (moved) Directory.Delete(path);
        }
    }

    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "open", SetLastError = true)]
    private static extern int NativeOpen(string path, int flags);
    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "pwrite", SetLastError = true)]
    private static extern IntPtr NativeWrite(int descriptor, byte[] buffer, nuint count, long offset);
    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "close", SetLastError = true)]
    private static extern int NativeClose(int descriptor);

    private sealed class Fixture(string directory) : IDisposable
    {
        internal string Path => System.IO.Path.Combine(directory, "source.ofd");
        internal static async Task<Fixture> Create()
        {
            var fixture = new Fixture(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ofdrw-file-session-" + Guid.NewGuid().ToString("N")));
            Directory.CreateDirectory(fixture.directory);
            var package = await TextReplacementTests.Read();
            using var output = File.Create(fixture.Path); await new OfdPackageWriter().WriteAsync(package, output);
            return fixture;
        }
        private readonly string directory = directory;
        public void Dispose() => Directory.Delete(directory, true);
    }
}
