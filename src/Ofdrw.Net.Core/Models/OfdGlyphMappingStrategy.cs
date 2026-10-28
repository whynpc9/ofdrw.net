using System;

namespace Ofdrw.Net.Core.Models;

/// <summary>Explicit interpretation of CGTransform source positions. No automatic fallback is performed.</summary>
public enum OfdGlyphMappingStrategy
{
    /// <summary>Preserves the historical bounded profile: positions index concatenated TextCode text.</summary>
    LegacyConcatenatedGlobal = 0,
    /// <summary>Binds preceding CGTransform siblings to the following TextCode, using local BMP positions.</summary>
    GroupedLocal = 1
}

internal static class OfdGlyphMappingStrategyGuard
{
    internal static void Validate(OfdGlyphMappingStrategy value)
    {
        if (value != OfdGlyphMappingStrategy.LegacyConcatenatedGlobal && value != OfdGlyphMappingStrategy.GroupedLocal)
            throw new ArgumentOutOfRangeException(nameof(value), value, "Unknown CGTransform mapping strategy.");
    }
}
