using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Explicitly removes signatures, preserving document bytes and all DocBody entries.</summary>
public static class OfdSignatureCleaner
{
    /// <summary>Removes declarations and unreferenced owned values/appearances. Unknown/shared payloads are retained with diagnostics.
    /// Use distinct streams and stage output for atomic publication. This is cleanup, not signing or signature validation.</summary>
    public static Task<OfdPackageWriteResult> CleanAsync(Stream source, Stream destination,
        OfdPackageLoadOptions? options = null, CancellationToken cancellationToken = default) =>
        OfdPackageSignatureCleaner.CleanAsync(source, destination, options, cancellationToken);
}
