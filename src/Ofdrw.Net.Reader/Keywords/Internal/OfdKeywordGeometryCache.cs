using System.Collections.Generic;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// Operation-local immutable metrics only. Source hashes, ownership, anchors
// and callback guards are never cached. Standalone Find does not use this cache.
internal sealed class OfdKeywordGeometryCache
{
    private readonly OfdKeywordLocatorBudget budget;
    private readonly Dictionary<(OfdGlyphFontSnapshot Font, string Text, double Size, int Weight, bool Italic), OfdGlyphGeometry> values;
    internal OfdKeywordGeometryCache(OfdKeywordLocatorBudget budget)
    { this.budget = budget; budget.Copy(1024); values = new(); }
    internal bool TryGet(OfdGlyphFontSnapshot font, string text, double size, int weight, bool italic, out OfdGlyphGeometry geometry)
    {
        budget.Correlate(text.Length + font.ResourceId.Length);
        return values.TryGetValue((font, text, size, weight, italic), out geometry!);
    }
    internal void Add(OfdGlyphFontSnapshot font, string text, double size, int weight, bool italic, OfdGlyphGeometry geometry)
    {
        budget.Copy(256); budget.Correlate(text.Length + font.ResourceId.Length);
        values.Add((font, text, size, weight, italic), geometry);
    }
}
