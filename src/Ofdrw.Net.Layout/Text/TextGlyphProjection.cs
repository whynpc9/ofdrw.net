namespace Ofdrw.Net.Layout.Text;

// Pure arithmetic only. Each caller retains its own admission, geometry checks and DTOs:
// a legacy writer result is not necessarily admissible to the strict positioning query.
internal readonly struct TextGlyphProjection
{
    private TextGlyphProjection(double x, double y, double advanceX, double advanceY,
        long nextPenX, long nextPenY)
    {
        X = x; Y = y; AdvanceX = advanceX; AdvanceY = advanceY;
        NextPenX = nextPenX; NextPenY = nextPenY;
    }

    internal double X { get; }
    internal double Y { get; }
    internal double AdvanceX { get; }
    internal double AdvanceY { get; }
    internal long NextPenX { get; }
    internal long NextPenY { get; }

    internal static TextGlyphProjection Create(TextShapedGlyph glyph, long penX, long penY,
        double scale, double originX, double baselineY) => new(
            originX + (penX + glyph.XOffset) * scale,
            baselineY - (penY + glyph.YOffset) * scale,
            glyph.XAdvance * scale, -((double)glyph.YAdvance) * scale,
            penX + glyph.XAdvance, penY + glyph.YAdvance);
}
