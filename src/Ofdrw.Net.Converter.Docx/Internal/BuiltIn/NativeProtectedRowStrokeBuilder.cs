using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

// One optional nonsemantic stroke per text object. The original text, font,
// requested weight and counter-constrained base stroke remain authoritative.
// This is an optical fallback; it does not create a physical weight700 face.
internal sealed class NativeProtectedRowStrokeBuilder(NativeBoldStrokePlanner planner, CancellationToken token)
{
    private const double PointsToMillimeters = 25.4 / 72;
    private const double CounterPaddingPoints = .5 + .0005;
    private int placements, commands, counters, characters;
    private bool exhausted;

    internal bool TryCreate(IReadOnlyList<OfdTextElement> texts, byte[] font, OfdPage page,
        out IReadOnlyList<OfdPathElement> result)
    {
        token.ThrowIfCancellationRequested();
        result = Array.Empty<OfdPathElement>();
        if (exhausted || texts.Count is 0 or > 64 || texts.Sum(t => (long)t.Text.Length) > 256 ||
            !Finite(page.WidthMillimeters) || !Finite(page.HeightMillimeters) ||
            page.WidthMillimeters is < .1 or > 10_000 || page.HeightMillimeters is < .1 or > 10_000 ||
            texts.Any(t => !Eligible(t))) return false;
        var paths = new List<OfdPathElement>();
        var groupHoles = new List<(double Left, double Top, double Right, double Bottom)>();
        try
        {
            foreach (var text in texts)
            {
                token.ThrowIfCancellationRequested();
                var size = Written(text.FontSizeMillimeters);
                var width = Math.Floor(size * .035 * 1000) / 1000;
                if (width <= text.SyntheticBoldStrokeWidthMillimeters!.Value) return false;
                var run = text.Runs[0];
                var advances = string.IsNullOrWhiteSpace(run.DeltaX)
                    ? Array.Empty<string>()
                    : run.DeltaX!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (advances.Length != text.Text.Length - 1) return false;
                var cursor = Written(text.XMillimeters) + Written(run.XMillimeters);
                var baseline = Written(text.YMillimeters) + Written(run.YMillimeters);
                if (!Finite(cursor) || !Finite(baseline)) return false;
                var data = new StringBuilder();
                for (var index = 0; index < text.Text.Length; index++)
                {
                    Charge(ref placements, 1, 2048);
                    var glyph = planner.ProtectedGlyph(font, size, text.Text[index]);
                    if (glyph is null) return false;
                    foreach (var curve in glyph.Curves)
                    {
                        Charge(ref commands, 1, 131072);
                        Append(data, curve.Command.ToString());
                        if (curve.Command == 'C') continue;
                        if (curve.Command == 'Q')
                            Pair(data, cursor + curve.ControlX * PointsToMillimeters,
                                baseline + (curve.ControlY + glyph.BaselineOffsetPoints) * PointsToMillimeters);
                        Pair(data, cursor + curve.X * PointsToMillimeters,
                            baseline + (curve.Y + glyph.BaselineOffsetPoints) * PointsToMillimeters);
                    }
                    foreach (var box in glyph.Counters)
                    {
                        Charge(ref counters, 1, 4096);
                        var left = cursor + (box.Left - CounterPaddingPoints) * PointsToMillimeters;
                        var top = baseline + (box.Top + glyph.BaselineOffsetPoints - CounterPaddingPoints) * PointsToMillimeters;
                        var right = cursor + (box.Right + CounterPaddingPoints) * PointsToMillimeters;
                        var bottom = baseline + (box.Bottom + glyph.BaselineOffsetPoints + CounterPaddingPoints) * PointsToMillimeters;
                        groupHoles.Add((Down(left), Down(top), Up(right), Up(bottom)));
                    }
                    if (index < advances.Length)
                    {
                        if (!double.TryParse(advances[index], NumberStyles.Float, CultureInfo.InvariantCulture, out var advance) ||
                            !Finite(advance)) return false;
                        cursor += advance;
                        if (!Finite(cursor) || Math.Abs(cursor) > 10_000) return false;
                    }
                }
                paths.Add(new OfdPathElement
                {
                    LayerType = text.LayerType, LayerId = text.LayerId,
                    WidthMillimeters = Written(page.WidthMillimeters),
                    HeightMillimeters = Written(page.HeightMillimeters),
                    Stroke = true, Fill = false, LineWidthMillimeters = width,
                    StrokeColor = text.FillColor, AbbreviatedData = data.ToString()
                });
            }
            // Adjacent equal-format runs can be separate TextObjects. Every
            // extra stroke must avoid all witnessed counters in the group,
            // including siblings whose glyph ink may overlap this object's ink.
            foreach (var path in paths)
            {
                var clips = new OfdClipCollection();
                foreach (var hole in groupHoles)
                {
                    Charge(ref counters, 1, 4096);
                    var clipData = new StringBuilder();
                    Rectangle(clipData, 0, 0, Written(page.WidthMillimeters), Written(page.HeightMillimeters));
                    // Independent complements intersect; overlapping exclusions
                    // cannot XOR their interiors back into another glyph.
                    Rectangle(clipData, hole.Left, hole.Top, hole.Right, hole.Bottom);
                    var clip = new OfdClip();
                    clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement
                    {
                        WidthMillimeters = Written(page.WidthMillimeters),
                        HeightMillimeters = Written(page.HeightMillimeters),
                        Stroke = false, Fill = true, FillRule = OfdFillRule.EvenOdd,
                        AbbreviatedData = clipData.ToString()
                    } });
                    clips.Items.Add(clip);
                }
                path.Clips = clips.Items.Count == 0 ? null : clips;
            }
        }
        catch (NotSupportedException) { return false; }
        // No path reaches the page until the whole equal-format group passed.
        result = paths;
        return paths.Count == texts.Count;
    }

    private static bool Eligible(OfdTextElement text) => text.EmitSyntheticBoldStroke &&
        text.Weight == OfdTextElement.BoldWeight && !text.Italic &&
        text.FillColor.Alpha == 255 && Finite(text.FontSizeMillimeters) && text.FontSizeMillimeters > 0 &&
        text.SyntheticBoldStrokeWidthMillimeters is > 0 &&
        string.IsNullOrEmpty(text.SourceXml) && string.IsNullOrEmpty(text.ClippingXml) &&
        text.Transform is null && text.Clips is null && text.Runs.Count == 1 &&
        text.Runs[0].Text == text.Text && string.IsNullOrEmpty(text.Runs[0].DeltaY) &&
        text.Runs[0].PositionedGlyphs is null && text.Text.Length > 0 && text.Text.All(IsHan);

    private static bool IsHan(char c) => c is >= '\u3400' and <= '\u4dbf' or >= '\u4e00' and <= '\u9fff' or >= '\uf900' and <= '\ufaff';
    private static double Written(double value) => double.Parse(value.ToString("0.###", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    private static double Down(double value) => Math.Floor(value * 1_000_000) / 1_000_000;
    private static double Up(double value) => Math.Ceiling(value * 1_000_000) / 1_000_000;
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private void Charge(ref int used, int count, int maximum)
    {
        token.ThrowIfCancellationRequested();
        used += count;
        if (exhausted || used > maximum)
        { exhausted = true; throw new NotSupportedException("Protected row emphasis output budget exhausted."); }
    }
    private void Append(StringBuilder data, string value)
    {
        Charge(ref characters, value.Length + 1, 4 * 1024 * 1024);
        if (data.Length + value.Length + 1 > 512 * 1024)
            throw new NotSupportedException("Protected row emphasis path is too large.");
        data.Append(value).Append(' ');
    }
    private void Pair(StringBuilder data, double x, double y)
    {
        if (!Finite(x) || !Finite(y) || Math.Abs(x) > 10_000 || Math.Abs(y) > 10_000)
            throw new NotSupportedException("Protected row emphasis placement is outside its finite budget.");
        Append(data, x.ToString("0.######", CultureInfo.InvariantCulture));
        Append(data, y.ToString("0.######", CultureInfo.InvariantCulture));
    }
    private void Rectangle(StringBuilder data, double left, double top, double right, double bottom)
    {
        Append(data, "M"); Pair(data, left, top);
        Append(data, "L"); Pair(data, right, top);
        Append(data, "L"); Pair(data, right, bottom);
        Append(data, "L"); Pair(data, left, bottom);
        Append(data, "C");
    }
}

internal sealed class NativeProtectedGlyph(int glyphId, NativeGlyphCurve[] curves,
    NativeCounterBounds[] counters, double baselineOffsetPoints)
{
    internal int GlyphId { get; } = glyphId;
    internal IReadOnlyList<NativeGlyphCurve> Curves { get; } = curves;
    internal IReadOnlyList<NativeCounterBounds> Counters { get; } = counters;
    internal double BaselineOffsetPoints { get; } = baselineOffsetPoints;
}

internal readonly struct NativeGlyphCurve(char command, double x = 0, double y = 0,
    double controlX = 0, double controlY = 0)
{
    internal char Command { get; } = command;
    internal double X { get; } = x;
    internal double Y { get; } = y;
    internal double ControlX { get; } = controlX;
    internal double ControlY { get; } = controlY;
    internal bool Finite => !double.IsNaN(X + Y + ControlX + ControlY) && !double.IsInfinity(X + Y + ControlX + ControlY);
}

internal readonly struct NativeCounterBounds(double left, double top, double right, double bottom)
{
    internal double Left { get; } = left;
    internal double Top { get; } = top;
    internal double Right { get; } = right;
    internal double Bottom { get; } = bottom;
}
