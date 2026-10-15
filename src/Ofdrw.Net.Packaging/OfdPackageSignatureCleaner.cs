using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

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
        if (!original.TryGetValue("OFD.xml", out var rootBytes)) throw new InvalidDataException("OFD.xml is missing.");
        var material = new OfdModelMaterialization(new OfdModelLimits());
        var input = new OfdDocumentPackage(); foreach (var pair in original) input.PreservedEntries.Add(pair.Key, pair.Value);
        material.InspectStorage(input, cancellationToken); material.GuardPreservedXml(input, cancellationToken); material.EstimateWriter(input, cancellationToken);
        using (var rootStream = new MemoryStream(rootBytes, false))
        {
            var root = material.ReadXml("OFD.xml", rootBytes, cancellationToken).Root;
            if (root?.Name.LocalName != "OFD" ||
                (root.Name.NamespaceName != OfdConstants.Namespace && root.Name.NamespaceName != OfdConstants.StandardNamespace) ||
                !root.Elements(root.Name.Namespace + "DocBody").Any())
                throw new InvalidDataException("Input has no recognizable OFD root and DocBody.");
        }
        var entries = new Dictionary<string, byte[]>(original, StringComparer.OrdinalIgnoreCase);
        var result = new OfdPackageWriteResult();
        OfdPackagePruner.CleanSignatures(original, entries, result, cancellationToken, material: material);
        material.ValidateActualEntries(entries, cancellationToken);
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
