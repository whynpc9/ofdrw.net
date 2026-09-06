using System.Collections.Generic;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Contains a merged document and any explicitly permitted dropped-content diagnostics.</summary>
public sealed class OfdDocumentMergeResult
{
    internal OfdDocumentMergeResult(OfdDocumentPackage package, IReadOnlyList<string> diagnostics)
    {
        Package = package;
        Diagnostics = diagnostics;
    }

    public OfdDocumentPackage Package { get; }
    public IReadOnlyList<string> Diagnostics { get; }
}
