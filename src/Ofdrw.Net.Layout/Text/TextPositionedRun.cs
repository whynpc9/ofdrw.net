using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Text;

/// <summary>An immutable millimeter projection of one exact shaped run, with Y increasing downward.</summary>
/// <remarks>This query does not shape, select a line, reorder runs, resolve styles or admit output
/// to a writer. Glyph order, source clusters, face and external context remain those of ShapedRun.
/// Coordinates are relative to the caller's chosen origin; no page or TextObject is created.</remarks>
public sealed class TextPositionedRun
{
    private TextPositionedRun(TextShapedRun shapedRun, double size, double originX, double baselineY,
        List<OfdPositionedGlyph> glyphs, double xAdvance, double yAdvance, TextLayoutBounds? ink)
    {
        ShapedRun = shapedRun; FontSizeMillimeters = size;
        OriginXMillimeters = originX; BaselineYMillimeters = baselineY;
        Glyphs = new ReadOnlyCollection<OfdPositionedGlyph>(glyphs);
        XAdvanceMillimeters = xAdvance; YAdvanceMillimeters = yAdvance; InkBoundsMillimeters = ink;
    }

    /// <summary>Projects an already shaped run using size/UPEM, pen plus offset and one Y-axis inversion.</summary>
    /// <remarks>Size must be finite and in (0, 1000] mm, with a representable positive scale.
    /// Input and resulting coordinates, advances and ink dimensions must be finite and within
    /// ±1,000,000 mm. An empty shaped run produces no glyphs or ink and zero advances.
    /// Cancellation or a budget failure returns no partial result. This managed query makes no native calls.</remarks>
    public static TextPositionedRun Create(TextShapedRun shapedRun, double fontSizeMillimeters,
        double originXMillimeters = 0, double baselineYMillimeters = 0,
        CancellationToken cancellationToken = default)
    {
        if (shapedRun is null) throw new ArgumentNullException(nameof(shapedRun));
        cancellationToken.ThrowIfCancellationRequested();
        var scale = fontSizeMillimeters / shapedRun.Request.Face.UnitsPerEm;
        if (!Finite(fontSizeMillimeters) || fontSizeMillimeters <= 0 || fontSizeMillimeters > 1000 || scale <= 0 ||
            !WithinBudget(originXMillimeters) || !WithinBudget(baselineYMillimeters))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid positioning size, scale or origin.");

        var glyphs = new List<OfdPositionedGlyph>(shapedRun.Glyphs.Count);
        long penX = 0, penY = 0;
        double? left = null;
        double top = 0, right = 0, bottom = 0;
        foreach (var glyph in shapedRun.Glyphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var projection = TextGlyphProjection.Create(glyph, penX, penY, scale,
                originXMillimeters, baselineYMillimeters);
            var x = Check(projection.X);
            var y = Check(projection.Y);
            var advanceX = Check(projection.AdvanceX);
            var advanceY = Check(projection.AdvanceY);
            glyphs.Add(new OfdPositionedGlyph(glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength,
                x, y, advanceX, advanceY));
            if (glyph.Ink is { } ink)
            {
                var inkLeft = Check(x + ink.X * scale);
                var inkTop = Check(y - (ink.Y + ink.Height) * scale);
                var inkRight = Check(inkLeft + ink.Width * scale);
                var inkBottom = Check(y - ink.Y * scale);
                if (!left.HasValue) { left = inkLeft; top = inkTop; right = inkRight; bottom = inkBottom; }
                else { left = Math.Min(left.Value, inkLeft); top = Math.Min(top, inkTop);
                    right = Math.Max(right, inkRight); bottom = Math.Max(bottom, inkBottom); }
            }
            penX = projection.NextPenX; penY = projection.NextPenY;
            Check(originXMillimeters + penX * scale);
            Check(baselineYMillimeters - penY * scale);
        }
        TextLayoutBounds? bounds = left.HasValue
            ? new TextLayoutBounds(left.Value, top, Check(right - left.Value), Check(bottom - top)) : null;
        var totalX = Check(shapedRun.XAdvance * scale);
        var totalY = Check(-((double)shapedRun.YAdvance) * scale);
        cancellationToken.ThrowIfCancellationRequested();
        return new TextPositionedRun(shapedRun, fontSizeMillimeters, originXMillimeters,
            baselineYMillimeters, glyphs, totalX, totalY, bounds);
    }

    /// <summary>The exact immutable shaped result, including its request, face and retained context.</summary>
    public TextShapedRun ShapedRun { get; }
    /// <summary>Explicit font em size used for this projection, in millimeters.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Caller-chosen initial pen X in millimeters.</summary>
    public double OriginXMillimeters { get; }
    /// <summary>Caller-chosen initial baseline Y in millimeters, increasing downward.</summary>
    public double BaselineYMillimeters { get; }
    /// <summary>Immutable origins and advances in the original shaper's order, with absolute UTF-16 clusters.</summary>
    public IReadOnlyList<OfdPositionedGlyph> Glyphs { get; }
    /// <summary>Final horizontal pen displacement, independent of offsets and origin deltas.</summary>
    public double XAdvanceMillimeters { get; }
    /// <summary>Final vertical pen displacement with Y downward, independent of offsets and origin deltas.</summary>
    public double YAdvanceMillimeters { get; }
    /// <summary>Union of translated glyph ink in caller coordinates, or null when no glyph has ink.</summary>
    public TextLayoutBounds? InkBoundsMillimeters { get; }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static bool WithinBudget(double value) => Finite(value) && Math.Abs(value) <= 1_000_000;
    private static double Check(double value) => WithinBudget(value) ? value :
        throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Positioned geometry exceeds the coordinate budget.");
}
