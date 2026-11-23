using System;
using System.Collections.Generic;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>One original-GID outline placed in the text object's local space (millimetres, y down).</summary>
internal sealed class OfdPlacedGlyph(int glyphIndex, int glyphId, TrueTypeGlyphOutline outline, IReadOnlyList<TrueTypePathCommand> commands)
{
    /// <summary>Slot in the explicit run (and in <see cref="OfdExplicitGlyphFont.Outlines"/>).</summary>
    internal int GlyphIndex { get; } = glyphIndex;
    internal int GlyphId { get; } = glyphId;
    internal TrueTypeGlyphOutline Outline { get; } = outline;
    /// <summary>Outline commands already scaled, flipped and offset to the glyph's anchor; quadratics stay quadratic.</summary>
    internal IReadOnlyList<TrueTypePathCommand> Commands { get; } = commands;
}

/// <summary>
/// Places explicit-glyph (CGTransform) outlines: each glyph's origin (complete source-span origins or the run's
/// DeltaX/DeltaY walk), the italic/emphasis anchor, horizontal-scale compensation, font-unit scaling and the y flip.
/// PDF and SVG/HTML were each doing this by hand; they now only turn placed commands into their own path syntax.
/// </summary>
internal static class OfdPlacedGlyphOutlines
{
    /// <param name="factor">The emphasis/italic drawing factor from <see cref="OfdTextEmphasis"/>.</param>
    /// <param name="hScale">Horizontal scale applied by the caller's transform; x is pre-divided to cancel it.</param>
    internal static IEnumerable<OfdPlacedGlyph> Place(OfdTextElement text, OfdExplicitGlyphRun run, OfdExplicitGlyphFont font,
        double[]? factor, double hScale, CancellationToken token)
    {
        var index = 0;
        foreach (var (x, y) in Origins(run, token))
        {
            token.ThrowIfCancellationRequested();
            var outline = font.Outlines[index];
            var anchor = OfdTextEmphasis.Anchor(x, y, factor);
            var scale = text.FontSizeMillimeters / outline.UnitsPerEm;
            double X(double value) => anchor.X / hScale + value * scale;
            double Y(double value) => anchor.Y - value * scale;
            var source = outline.ToPathCommands();
            var placed = new TrueTypePathCommand[source.Count];
            for (var command = 0; command < source.Count; command++)
            {
                var c = source[command];
                placed[command] = new TrueTypePathCommand(c.Verb, X(c.X), Y(c.Y),
                    X(c.ControlX), Y(c.ControlY), X(c.Control2X), Y(c.Control2Y));
            }
            yield return new OfdPlacedGlyph(index, run.GlyphIds[index], outline, placed);
            index++;
        }
    }

    /// <summary>
    /// Glyph origins: a complete source span uses its verified per-slot origins; otherwise each run starts at its own
    /// origin and advances by its expanded deltas between glyphs.
    /// </summary>
    internal static IEnumerable<(double X, double Y)> Origins(OfdExplicitGlyphRun run, CancellationToken token)
    {
        if (run.PathGlyphSourceSpan is not null)
        {
            foreach (var origin in run.PathGlyphOrigins(token)) yield return origin;
            yield break;
        }
        foreach (var textRun in run.Runs)
        {
            var glyphCount = run.SingleGlyphSourceSpan is not null ? 1 : textRun.Text.Length;
            var hasX = !string.IsNullOrWhiteSpace(textRun.DeltaX);
            var hasY = !string.IsNullOrWhiteSpace(textRun.DeltaY);
            var deltaX = hasX ? OfdTextGeometry.ExpandDeltas(textRun.DeltaX, glyphCount - 1, token) : [];
            var deltaY = hasY ? OfdTextGeometry.ExpandDeltas(textRun.DeltaY, glyphCount - 1, token) : [];
            var x = textRun.XMillimeters; var y = textRun.YMillimeters;
            for (var index = 0; index < glyphCount; index++)
            {
                yield return (x, y);
                if (index + 1 == glyphCount) continue;
                if (hasX) x += deltaX[index];
                if (hasY) y += deltaY[index];
            }
        }
    }
}
