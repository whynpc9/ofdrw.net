using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class WriterFailClosedTests
{
    // SHA-256 over the sorted entry names and contents written for Package() by the writer
    // before the fail-closed change (computed at d34d0a3). ZIP timestamps are excluded.
    private const string ExpectedManifest = "8F07F17A0B52B28392857687F260D1BA42534B58F42ECD3FD1309413D3463595";

    [Fact]
    public async Task Write_FailureMidway_LeavesNoReadableZip()
    {
        var full = await Write(Package());
        using var stream = new FailOnceAfterBytesStream(full.Length / 2);

        await Assert.ThrowsAsync<IOException>(() => new OfdPackageWriter().WriteAsync(Package(), stream));

        Assert.True(stream.Failed);
        Assert.Throws<InvalidDataException>(() => new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read));
    }

    [Fact]
    public async Task Write_Cancelled_LeavesNoReadableZip()
    {
        using var source = new CancellationTokenSource();
        using var stream = new CancelOnFirstWriteStream(source);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdPackageWriter().WriteAsync(Package(), stream, source.Token));

        Assert.Throws<InvalidDataException>(() => new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read));
    }

    [Fact]
    public async Task Write_Success_KeepsTheSameEntries()
    {
        var bytes = await Write(Package());

        var manifest = Manifest(bytes);
        Assert.True(ExpectedManifest == manifest, $"manifest {manifest}");
        Assert.Equal(3, (await new OfdReader().ReadAsync(new MemoryStream(bytes))).Pages.Count);
    }

    [Fact]
    public async Task SignatureCleaner_FailureMidway_LeavesNoReadableZip()
    {
        var source = await Write(Package());
        using var full = new MemoryStream();
        await OfdPackageSignatureCleaner.CleanAsync(new MemoryStream(source), full);
        using var stream = new FailOnceAfterBytesStream(full.Length / 2);

        await Assert.ThrowsAsync<IOException>(() => OfdPackageSignatureCleaner.CleanAsync(new MemoryStream(source), stream));

        Assert.True(stream.Failed);
        Assert.Throws<InvalidDataException>(() => new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read));
    }

    internal static string Manifest(byte[] zipBytes)
    {
        using var zip = new ZipArchive(new MemoryStream(zipBytes), ZipArchiveMode.Read);
        var text = new StringBuilder();
        foreach (var entry in zip.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            // The writer indents XML with Environment.NewLine; compare content independent of the OS.
            var content = copy.ToArray();
            if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                content = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(content).Replace("\r\n", "\n"));
            text.Append(entry.FullName).Append(':').Append(Convert.ToHexString(SHA256.HashData(content))).Append('\n');
        }
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text.ToString())));
    }

    // Font-free content keeps the manifest independent of host fonts.
    private static OfdDocumentPackage Package()
    {
        var package = new OfdDocumentPackage();
        for (var page = 0; page < 3; page++)
        {
            var model = new OfdPage { Index = page, WidthMillimeters = 210, HeightMillimeters = 297 };
            for (var index = 0; index < 400; index++)
                model.Elements.Add(new OfdPathElement
                {
                    XMillimeters = 10 + index % 20 * 9, YMillimeters = 10 + index / 20 * 13, WidthMillimeters = 8, HeightMillimeters = 12,
                    AbbreviatedData = string.Create(CultureInfo.InvariantCulture, $"M 0 0 L 8 {index % 12} L {page + 1} 12 C 2 3 4 5 6 {index % 7} Z")
                });
            package.Pages.Add(model);
        }
        return package;
    }

    private static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    // Throws exactly once, on the first write that would pass the threshold; later writes succeed,
    // so an unguarded writer can still emit its central directory during disposal.
    private sealed class FailOnceAfterBytesStream(long threshold) : MemoryStream
    {
        public bool Failed { get; private set; }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Trip(count);
            base.Write(buffer, offset, count);
        }

        public override void Write(ReadOnlySpan<byte> buffer)
        {
            Trip(buffer.Length);
            base.Write(buffer);
        }

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            Trip(count);
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            Trip(buffer.Length);
            return base.WriteAsync(buffer, cancellationToken);
        }

        private void Trip(int count)
        {
            if (Failed || Length + count <= threshold) return;
            Failed = true;
            throw new IOException("Injected one-shot write failure.");
        }
    }

    private sealed class CancelOnFirstWriteStream(CancellationTokenSource source) : MemoryStream
    {
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            source.Cancel();
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default)
        {
            source.Cancel();
            return base.WriteAsync(buffer, cancellationToken);
        }
    }
}
