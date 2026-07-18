using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

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

        using var buffer = new MemoryStream();
        await ofdStream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
        buffer.Position = 0;

        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        using var zip = new ZipArchive(buffer, ZipArchiveMode.Read, leaveOpen: true);
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
            await entryStream.CopyToAsync(ms, 81920, cancellationToken).ConfigureAwait(false);
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
        if (options.MaxEntryCount <= 0 ||
            options.MaxEntryUncompressedBytes <= 0 ||
            options.MaxTotalUncompressedBytes <= 0 ||
            options.MaxCompressionRatio <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "All OFD package load limits must be positive.");
        }
    }
}
