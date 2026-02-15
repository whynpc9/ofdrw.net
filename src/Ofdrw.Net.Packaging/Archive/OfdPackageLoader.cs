using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Packaging.Archive;

public sealed class OfdPackageLoader
{
    public async Task<OfdPackageArchive> LoadAsync(Stream ofdStream, CancellationToken cancellationToken = default)
    {
        if (ofdStream is null)
        {
            throw new ArgumentNullException(nameof(ofdStream));
        }

        using var buffer = new MemoryStream();
        await ofdStream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
        buffer.Position = 0;

        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        using var zip = new ZipArchive(buffer, ZipArchiveMode.Read, leaveOpen: true);
        foreach (var entry in zip.Entries)
        {
            if (string.IsNullOrWhiteSpace(entry.Name))
            {
                continue;
            }

            using var entryStream = entry.Open();
            using var ms = new MemoryStream();
            await entryStream.CopyToAsync(ms, 81920, cancellationToken).ConfigureAwait(false);
            result[Normalize(entry.FullName)] = ms.ToArray();
        }

        return new OfdPackageArchive(result);
    }

    private static string Normalize(string path)
    {
        return path.Replace('\\', '/').TrimStart('/');
    }
}
