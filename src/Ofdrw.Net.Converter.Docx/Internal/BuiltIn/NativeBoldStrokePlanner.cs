using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Threading;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Converter.Pdf.Internal;
using PdfSharpCore.Drawing;
using SixLabors.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

// A bounded optical fallback, not a physical weight700 face. Only fresh native
// DOCX opts in. Preserve the old Han stroke unless original counter geometry
// permits an increase towards the existing general 2.5%-em emphasis.
internal sealed class NativeBoldStrokePlanner
{
    private readonly CancellationToken token;
    private readonly Dictionary<byte[], Profile?> profiles = new();
    private readonly WorkBudget budget;
    private long faceBytes;

    internal NativeBoldStrokePlanner(CancellationToken token)
    { this.token = token; budget = new WorkBudget(token); }

    internal double Width(byte[] bytes, bool physicalItalic, double size, string text)
    {
        token.ThrowIfCancellationRequested();
        var legacy = FontEmphasisMetrics.BoldStrokeWidth(size, text);
        var writtenSize = double.Parse(size.ToString("0.###", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        if (physicalItalic || writtenSize * 72 / 25.4 is < .1 or > 1000 ||
            legacy == FontEmphasisMetrics.BoldStrokeWidth(size) || budget.Exhausted) return legacy;
        if (!profiles.TryGetValue(bytes, out var profile))
        {
            profile = null;
            if (profiles.Count < 8 && bytes.Length <= 32 * 1024 * 1024 &&
                faceBytes + bytes.Length <= 64L * 1024 * 1024 && StaticTrueType(bytes))
            {
                faceBytes += bytes.Length;
                try { profile = new Profile(bytes, budget); }
                catch (FontException) { }
                catch (NotSupportedException) { }
            }
            // Keep the cache itself bounded when many distinct faces are used.
            if (profiles.Count < 8) profiles.Add(bytes, profile);
        }
        if (profile is null) return legacy;
        var fraction = .025;
        for (var i = 0; i < text.Length; i++)
        {
            token.ThrowIfCancellationRequested();
            var scalar = char.ConvertToUtf32(text, i);
            if (char.IsHighSurrogate(text[i])) i++;
            if (char.IsWhiteSpace(text, i)) continue;
            var measured = profile.Fraction(scalar, writtenSize * 72 / 25.4);
            if (measured is null) return legacy;
            fraction = Math.Min(fraction, measured.Value);
        }
        // This only bounds an increase around witnessed closed counters, not
        // every open gap or narrow neck in an arbitrary glyph.
        // Match the writer's serialized font size and round an increase down
        // to its 0.001mm grid, so formatting cannot exceed the checked width.
        return Math.Max(legacy, Math.Floor(writtenSize * fraction * 1000) / 1000);
    }

    private static bool StaticTrueType(byte[] bytes)
    {
        if (bytes.Length < 12 || bytes[0] != 0 || bytes[1] != 1 || bytes[2] != 0 || bytes[3] != 0) return false;
        var count = bytes[4] * 256 + bytes[5];
        if (count > 256 || 12 + count * 16 > bytes.Length) return false;
        var glyf = false;
        for (var i = 0; i < count; i++)
        {
            var tag = System.Text.Encoding.ASCII.GetString(bytes, 12 + i * 16, 4);
            if (tag == "glyf") glyf = true;
            if (tag is "fvar" or "gvar" or "CFF " or "CFF2" or "COLR" or "SVG " or "CBDT" or "EBDT" or "sbix") return false;
        }
        return glyf;
    }

    private sealed class WorkBudget(CancellationToken token)
    {
        private int work, points, glyphs;
        internal bool Exhausted { get; private set; }
        internal void Tick()
        {
            token.ThrowIfCancellationRequested();
            if (Exhausted || ++work > 50_000_000) Stop();
        }
        internal void Point() { Tick(); if (++points > 131072) Stop(); }
        internal void Glyph() { Tick(); if (++glyphs > 512) Stop(); }
        private void Stop() { Exhausted = true; throw new NotSupportedException("Native bold analysis budget exhausted."); }
    }

    private sealed class Profile
    {
        private readonly FontFamily family;
        private readonly WorkBudget budget;
        private readonly Dictionary<(int Scalar, double Size), double?> fractions = new();
        internal Profile(byte[] bytes, WorkBudget budget)
        {
            this.budget = budget;
            using var stream = new MemoryStream(bytes, writable: false);
            family = new FontCollection().Add(stream);
            if (!family.GetAvailableStyles().Contains(FontStyle.Regular))
                throw new NotSupportedException("No physical regular outline for native bold analysis.");
        }
        internal double? Fraction(int scalar, double size)
        {
            if (fractions.TryGetValue((scalar, size), out var cached)) return cached;
            double? result = null;
            try
            {
                budget.Glyph();
                var renderer = new CounterRenderer(budget);
                var font = family.CreateFont((float)size, FontStyle.Regular);
                TextRenderer.RenderTextTo(renderer, char.ConvertFromUtf32(scalar),
                    new TextOptions(font) { Dpi = 72, KerningMode = KerningMode.None });
                result = renderer.Fraction(font.Size);
            }
            catch (NotSupportedException) { }
            catch (FontException) { }
            if (fractions.Count < 512) fractions.Add((scalar, size), result);
            return result;
        }
    }

    private sealed class CounterRenderer(WorkBudget budget) : IGlyphRenderer
    {
        private readonly List<List<Point>> contours = new();
        private readonly PdfGlyphStrokeGeometry stroke = new(.001);
        private List<Point>? current;
        private Point position;
        private int glyphs, points;
        public void BeginText(in FontRectangle bounds) { }
        public void EndText() { }
        public bool BeginGlyph(in FontRectangle bounds, in GlyphRendererParameters parameters)
        {
            if (++glyphs != 1 || parameters.GlyphId == 0 || parameters.GlyphType != GlyphType.Standard)
                throw new NotSupportedException("Only one original monochrome glyph is analyzed.");
            return true;
        }
        public void EndGlyph() { }
        public TextDecorations EnabledDecorations() => TextDecorations.None;
        public void SetDecoration(TextDecorations decorations, Vector2 start, Vector2 end, float thickness) { }
        public void BeginFigure()
        {
            if (contours.Count >= 64) throw new NotSupportedException("Too many bold-analysis contours.");
            current = new List<Point>(); contours.Add(current); stroke.BeginFigure();
        }
        public void EndFigure() { current = null; stroke.EndFigure(); }
        public void MoveTo(Vector2 p) { position = new Point(p.X, p.Y); Add(position); stroke.MoveTo(p); }
        public void LineTo(Vector2 p) { position = new Point(p.X, p.Y); Add(position); stroke.LineTo(p); }
        public void CubicBezierTo(Vector2 a, Vector2 b, Vector2 c) =>
            throw new NotSupportedException("Native bold analysis requires quadratic TrueType outlines.");
        public void QuadraticBezierTo(Vector2 control, Vector2 target)
        {
            var end = new Point(target.X, target.Y);
            var start = new Vector2((float)position.X, (float)position.Y);
            stroke.CubicBezierTo(start, start + (control - start) * (2f / 3f),
                target + (control - target) * (2f / 3f), target);
            Flatten(position, new Point(control.X, control.Y), end, 0); position = end;
        }
        private void Add(Point p)
        {
            budget.Point();
            if (current is null || ++points > 4096 || !p.Finite)
                throw new NotSupportedException("Native bold outline exceeds its finite point budget.");
            current.Add(p);
        }
        private void Flatten(Point a, Point b, Point c, int depth)
        {
            budget.Tick();
            if (Distance(b, a, c) <= .00025) { Add(c); return; }
            if (depth >= 20) throw new NotSupportedException("Native bold curve exceeds its precision budget.");
            var ab = (a + b) * .5; var bc = (b + c) * .5; var mid = (ab + bc) * .5;
            Flatten(a, ab, mid, depth + 1); Flatten(mid, bc, c, depth + 1);
        }
        internal double Fraction(double size)
        {
            if (glyphs != 1 || contours.Count == 0) throw new NotSupportedException("No original outline to analyze.");
            var valid = contours.Where(c => c.Count > 2).ToArray();
            if (valid.Length == 0) throw new NotSupportedException("No closed original contour to analyze.");
            var largest = valid.OrderByDescending(c => Math.Abs(Area(c))).First();
            var sign = Math.Sign(Area(largest));
            if (sign == 0) throw new NotSupportedException("Ambiguous original contour orientation.");
            var holes = valid.Where(c => Math.Sign(Area(c)) == -sign).ToArray();
            var areas = holes.ToDictionary(c => c, c => Math.Abs(Area(c)));
            var witnesses = new List<(Point Center, double Radius)>();
            foreach (var contour in holes)
            {
                var left = contour.Min(p => p.X); var right = contour.Max(p => p.X);
                var top = contour.Min(p => p.Y); var bottom = contour.Max(p => p.Y);
                var best = 0d; var center = default(Point);
                for (var row = 1; row < 10; row++)
                for (var column = 1; column < 10; column++)
                {
                    var p = new Point(left + (right - left) * column / 10, top + (bottom - top) * row / 10);
                    if (Winding(p, contour) == 0 || valid.Sum(c => Winding(p, c)) != 0) continue;
                    // The innermost containing counter owns the witness. An
                    // outer counter may not borrow the inner hole in 回.
                    if (holes.Any(c => !ReferenceEquals(c, contour) && areas[c] <= areas[contour] && Winding(p, c) != 0)) continue;
                    var distance = double.PositiveInfinity;
                    foreach (var c in valid)
                    for (var i = 0; i < c.Count; i++)
                    { budget.Tick(); distance = Math.Min(distance, Distance(p, c[i], c[(i + 1) % c.Count])); }
                    // Conservative allowance for flattening and float callbacks.
                    var radius = distance - .0005;
                    if (radius > best) { best = radius; center = p; }
                }
                if (best <= 0) return .01;
                witnesses.Add((center, best));
            }
            bool Preserves(double width) => witnesses.All(w =>
                stroke.Clearance(new XPoint(w.Center.X, w.Center.Y), width, budget.Tick) >= w.Radius * .8 + .0005);
            var lower = size * .01; var upper = size * .025;
            if (witnesses.Count == 0 || Preserves(upper)) return .025;
            if (!Preserves(lower)) return .01;
            // Validate the actual miter10/bevel polygons at every candidate;
            // distance to original centerlines alone is insufficient at joins.
            for (var i = 0; i < 7; i++)
            {
                var middle = (lower + upper) * .5;
                if (Preserves(middle)) lower = middle; else upper = middle;
            }
            return lower / size;
        }
        private int Winding(Point p, List<Point> contour)
        {
            var winding = 0;
            for (var i = 0; i < contour.Count; i++)
            {
                budget.Tick(); var a = contour[i]; var b = contour[(i + 1) % contour.Count];
                var turn = (b.X - a.X) * (p.Y - a.Y) - (p.X - a.X) * (b.Y - a.Y);
                if (a.Y <= p.Y && p.Y < b.Y && turn > 0) winding++;
                else if (b.Y <= p.Y && p.Y < a.Y && turn < 0) winding--;
            }
            return winding;
        }
        private double Area(List<Point> contour)
        {
            var area = 0d;
            for (var i = 0; i < contour.Count; i++)
            { budget.Tick(); var a = contour[i]; var b = contour[(i + 1) % contour.Count]; area += a.X * b.Y - b.X * a.Y; }
            return area * .5;
        }
        private static double Distance(Point p, Point a, Point b)
        {
            var delta = b - a; var length = delta.X * delta.X + delta.Y * delta.Y;
            var t = length == 0 ? 0 : Math.Max(0, Math.Min(1, ((p.X - a.X) * delta.X + (p.Y - a.Y) * delta.Y) / length));
            var d = p - a - delta * t; return Math.Sqrt(d.X * d.X + d.Y * d.Y);
        }
        private readonly struct Point(double x, double y)
        {
            internal double X { get; } = x;
            internal double Y { get; } = y;
            internal bool Finite => !double.IsNaN(X) && !double.IsInfinity(X) && !double.IsNaN(Y) && !double.IsInfinity(Y);
            public static Point operator +(Point a, Point b) => new(a.X + b.X, a.Y + b.Y);
            public static Point operator -(Point a, Point b) => new(a.X - b.X, a.Y - b.Y);
            public static Point operator *(Point p, double scale) => new(p.X * scale, p.Y * scale);
        }
    }
}
