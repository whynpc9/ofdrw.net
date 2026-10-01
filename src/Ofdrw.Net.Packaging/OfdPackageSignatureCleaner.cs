using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging;

/// <summary>Removes signature declarations across all DocBody entries without rewriting document content.</summary>
public static class OfdPackageSignatureCleaner
{
    /// <summary>Loads with explicit budgets, removes declarations and proven unreferenced signature payloads, and writes a new package.
    /// Stream callers must stage output if atomic publication is required. Input and output must be distinct streams.</summary>
    public static async Task<OfdPackageWriteResult> CleanAsync(Stream source, Stream destination,
        OfdPackageLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (destination is null) throw new ArgumentNullException(nameof(destination));
        if (ReferenceEquals(source, destination)) throw new ArgumentException("Input and output streams must differ.");
        var archive = await new OfdPackageLoader().LoadAsync(source, options ?? new OfdPackageLoadOptions(), cancellationToken).ConfigureAwait(false);
        var original = archive.EntryNames.ToDictionary(path => path, archive.GetBytes, StringComparer.OrdinalIgnoreCase);
        var entries = new Dictionary<string, byte[]>(original, StringComparer.OrdinalIgnoreCase);
        var result = new OfdPackageWriteResult();
        OfdPackagePruner.CleanSignatures(original, entries, result, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var entry in entries.OrderBy(pair => pair.Key, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var output = zip.CreateEntry(entry.Key, CompressionLevel.Optimal).Open();
            await output.WriteAsync(entry.Value, 0, entry.Value.Length, cancellationToken).ConfigureAwait(false);
        }
        return result;
    }
}
