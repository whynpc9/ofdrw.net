namespace Ofdrw.Net.Core.Fonts;

/// <summary>Shared synthetic emphasis geometry for embedded regular faces.</summary>
internal static class FontEmphasisMetrics
{
    // Preserve the PDF export strategy: 2.5% of the em size, in the caller's
    // coordinate unit (OFD millimeters or PDF points). This is simulated weight,
    // not a substitute for the font designer's actual bold face.
    internal static double BoldStrokeWidth(double fontSize) => fontSize * 0.025;
}
