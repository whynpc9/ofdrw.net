using System.Collections.Generic;

namespace Ofdrw.Net.Packaging;

/// <summary>Describes physical cleanup performed while saving an edited package.</summary>
public sealed class OfdPackageWriteResult
{
    internal readonly List<string> Removed = new();
    internal readonly List<string> Warnings = new();
    internal readonly List<OfdFontEmbeddingResult> Fonts = new();

    public IReadOnlyList<string> RemovedEntries => Removed;
    public IReadOnlyList<string> Diagnostics => Warnings;
    /// <summary>Embedded font decisions grouped by source content identity.</summary>
    public IReadOnlyList<OfdFontEmbeddingResult> FontEmbedding => Fonts;
    public bool SignaturesInvalidated { get; internal set; }
}
