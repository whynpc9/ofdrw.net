using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Packaging.Archive;

public sealed class OfdPackageLoader
{
    public async Task<OfdPackageArchive> LoadAsync(Stream ofdStream, CancellationToken cancellationToken = default)
    {
        return await LoadAsync(ofdStream, new OfdPackageLoadOptions(), cancellationToken).ConfigureAwait(false);
    }

    public async Task<OfdPackageArchive> LoadAsync(
        Stream ofdStream,
        OfdPackageLoadOptions options,
        CancellationToken cancellationToken = default)
    {
        if (ofdStream is null)
        {
            throw new ArgumentNullException(nameof(ofdStream));
        }

        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }

        ValidateOptions(options);

        cancellationToken.ThrowIfCancellationRequested();
        using var buffer = ofdStream.CanSeek && ofdStream.Position == 0 ? null : new MemoryStream();
        Stream source = ofdStream;
        if (buffer is not null)
        {
            await BoundedStreamCopy.CopyAsync(
                ofdStream, buffer, options.MaxInputBytes, "OFD input", cancellationToken).ConfigureAwait(false);
            buffer.Position = 0;
            source = buffer;
        }
        else if (ofdStream.Length > options.MaxInputBytes)
        {
            throw new InvalidDataException("OFD input exceeds the configured compressed input size limit.");
        }

        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        using var zip = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        if (zip.Entries.Count > options.MaxEntryCount)
        {
            throw new InvalidDataException(
                $"OFD package contains {zip.Entries.Count} entries, exceeding the configured limit of {options.MaxEntryCount}.");
        }

        long totalUncompressedBytes = 0;
        foreach (var entry in zip.Entries)
        {
            if (string.IsNullOrWhiteSpace(entry.Name))
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            var normalizedName = NormalizeAndValidate(entry.FullName);
            if (result.ContainsKey(normalizedName))
            {
                throw new InvalidDataException($"OFD package contains a duplicate entry: {normalizedName}");
            }

            if (entry.Length > options.MaxEntryUncompressedBytes)
            {
                throw new InvalidDataException(
                    $"OFD package entry '{normalizedName}' exceeds the configured uncompressed size limit.");
            }

            totalUncompressedBytes = checked(totalUncompressedBytes + entry.Length);
            if (totalUncompressedBytes > options.MaxTotalUncompressedBytes)
            {
                throw new InvalidDataException("OFD package exceeds the configured total uncompressed size limit.");
            }

            if (entry.CompressedLength > 0 &&
                entry.Length / (double)entry.CompressedLength > options.MaxCompressionRatio)
            {
                throw new InvalidDataException(
                    $"OFD package entry '{normalizedName}' exceeds the configured compression ratio limit.");
            }

            using var entryStream = entry.Open();
            using var ms = new MemoryStream();
            var copied = await BoundedStreamCopy.CopyAsync(
                entryStream, ms, Math.Max(1, entry.Length), $"OFD entry '{normalizedName}'", cancellationToken)
                .ConfigureAwait(false);
            if (copied != entry.Length)
            {
                throw new InvalidDataException($"OFD entry '{normalizedName}' has an inconsistent expanded length.");
            }
            result[normalizedName] = ms.ToArray();
        }

        return new OfdPackageArchive(result);
    }

    private static string NormalizeAndValidate(string path)
    {
        var normalized = path.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal) ||
            normalized.Split('/').Any(x => x == ".."))
        {
            throw new InvalidDataException($"OFD package entry has an unsafe path: {path}");
        }

        return normalized.TrimStart('/');
    }

    private static void ValidateOptions(OfdPackageLoadOptions options)
    {
        if (options.MaxInputBytes <= 0 || options.MaxEntryCount <= 0 || options.MaxPageCount <= 0 ||
            options.MaxEntryUncompressedBytes <= 0 ||
            options.MaxTotalUncompressedBytes <= 0 ||
            options.MaxCompressionRatio <= 0 || double.IsNaN(options.MaxCompressionRatio) ||
            double.IsInfinity(options.MaxCompressionRatio))
        {
            throw new ArgumentOutOfRangeException(nameof(options), "All OFD package load limits must be positive.");
        }
    }
}
