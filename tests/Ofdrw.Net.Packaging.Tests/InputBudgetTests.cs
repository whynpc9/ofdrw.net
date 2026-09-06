using System.IO.Compression;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public class InputBudgetTests
{
    [Fact]
    public async Task Loader_ShouldRejectOversizedSeekableInputBeforeReading()
    {
        using var input = CreateZip();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(
            input, new OfdPackageLoadOptions { MaxInputBytes = input.Length - 1 }));
        Assert.Equal(0, input.Position);
        Assert.True(input.CanRead);
    }

    [Fact]
    public async Task Loader_ShouldStopNonSeekableInputAtBudgetPlusOneByte()
    {
        using var zip = CreateZip();
        using var input = new TrackingInput(zip);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(
            input, new OfdPackageLoadOptions { MaxInputBytes = 1024 }));
        Assert.Equal(1025, input.BytesRead);
    }

    [Fact]
    public async Task Reader_ShouldAcceptAnExplicitInputBudget()
    {
        using var zip = CreateZip();
        using var input = new TrackingInput(zip);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(
            input, new OfdPackageLoadOptions { MaxInputBytes = 1024 }));
        Assert.Equal(1025, input.BytesRead);
    }

    [Fact]
    public async Task Loader_ShouldHonorCancellationBeforeStaging()
    {
        using var zip = CreateZip();
        using var input = new TrackingInput(zip);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdPackageLoader().LoadAsync(
            input, cancellation.Token));
        Assert.Equal(0, input.BytesRead);
    }

    [Fact]
    public async Task Loader_ShouldAcceptEmptyEntriesAndExactInputLimit()
    {
        using var input = new MemoryStream();
        using (var zip = new ZipArchive(input, ZipArchiveMode.Create, leaveOpen: true))
        {
            zip.CreateEntry("empty.bin");
        }

        input.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(input,
            new OfdPackageLoadOptions { MaxInputBytes = input.Length });
        Assert.Empty(archive.GetBytes("empty.bin"));
        Assert.True(input.CanRead);
    }

    private static MemoryStream CreateZip()
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        using (var entry = zip.CreateEntry("payload.bin", CompressionLevel.NoCompression).Open())
        {
            entry.Write(new byte[8192]);
        }

        stream.Position = 0;
        return stream;
    }

    private sealed class TrackingInput(Stream source) : Stream
    {
        internal long BytesRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = source.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            var read = await source.ReadAsync(buffer, offset, count, cancellationToken);
            BytesRead += read;
            return read;
        }
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
