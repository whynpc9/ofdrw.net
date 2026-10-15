using System;
using System.IO;
using System.Threading;
namespace Ofdrw.Net.Core.Fonts;

// Shared by all cmap parses in one embedding/export operation, including optional
// host fallback and nested signature fonts. Units are bytes copied/validated plus
// encoding, group, segment and variation entries visited; aliases pay only their
// encoding record. Check and charge before allocation or expansion.
internal sealed class FontParseBudget
{
    internal const long DefaultMaximumWork = 64L * 1024 * 1024;
    internal long Consumed { get; private set; }
    internal bool IsExhausted { get; private set; }
    internal CancellationToken CancellationToken { get; }
    private readonly long maximum;
    internal FontParseBudget(long maximum = DefaultMaximumWork, CancellationToken cancellationToken = default)
    {
        if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
        this.maximum = maximum; CancellationToken = cancellationToken;
    }
    internal void Spend(long work)
    {
        CancellationToken.ThrowIfCancellationRequested();
        if (work < 0 || IsExhausted || work > maximum - Consumed)
        {
            IsExhausted = true;
            throw new InvalidDataException("Font cmap parsing exceeds MaximumCmapParseWork.");
        }
        Consumed += work;
    }
}
