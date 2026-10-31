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
using Clipper2Lib;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

// A bounded optical fallback, not a physical weight700 face. Only fresh native
// DOCX opts in. Increases require both relative counter clearance and an
// absolute opening at the serialized point size. Legacy fallback does not
// claim to meet that optical minimum.
internal sealed class NativeBoldStrokePlanner
{
    private const double MaximumStrokeFraction = .035;
    private const double MinimumRetainedWitnessFraction = .6;
    private const double MinimumRetainedWitnessRadiusPoints = .4;
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
            legacy == FontEmphasisMetrics.BoldStrokeWidth(size)) return legacy;
        var profile = GetProfile(bytes);
        if (profile is null) return legacy;
        var fraction = MaximumStrokeFraction;
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

    internal NativeProtectedGlyph? ProtectedGlyph(byte[] bytes, double size, int scalar)
    {
        token.ThrowIfCancellationRequested();
        var writtenSize = double.Parse(size.ToString("0.###", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        if (writtenSize * 72 / 25.4 is < .1 or > 1000) return null;
        return GetProfile(bytes)?.ProtectedGlyph(scalar, writtenSize * 72 / 25.4);
    }

    private Profile? GetProfile(byte[] bytes)
    {
        if (!profiles.TryGetValue(bytes, out var profile))
        {
            if (budget.Exhausted) return null;
            profile = null;
            if (profiles.Count < 8 && bytes.Length <= 32 * 1024 * 1024 &&
                faceBytes + bytes.Length <= 64L * 1024 * 1024 && OpenTypeFace.HasStaticTrueTypeOutlines(bytes))
            {
                faceBytes += bytes.Length;
                try { profile = new Profile(bytes, budget); }
                catch (FontException) { }
                catch (NotSupportedException) { }
                catch (InvalidDataException) { }
            }
            // Keep the cache itself bounded when many distinct faces are used.
            if (profiles.Count < 8) profiles.Add(bytes, profile);
        }
        return profile;
    }

    private sealed class WorkBudget(CancellationToken token)
    {
        internal FontParseBudget Cmap { get; } = new(cancellationToken: token);
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
        private readonly OpenTypeCmap cmap;
        private readonly WorkBudget budget;
        private readonly bool forceIntegerPpem;
        private readonly Dictionary<(int Scalar, double Size), double?> fractions = new();
        private readonly Dictionary<(int Scalar, double Size), NativeProtectedGlyph?> protectedGlyphs = new();
        internal Profile(byte[] bytes, WorkBudget budget)
        {
            this.budget = budget;
            using var stream = new MemoryStream(bytes, writable: false);
            family = new FontCollection().Add(stream);
            if (!family.GetAvailableStyles().Contains(FontStyle.Regular))
                throw new NotSupportedException("No physical regular outline for native bold analysis.");
            var face = new OpenTypeFace(bytes);
            forceIntegerPpem = (OpenTypeFace.U16(face.Table("head"), 16) & 8) != 0;
            cmap = new OpenTypeCmap(face, budget.Cmap);
        }
        internal double? Fraction(int scalar, double size)
        {
            // Final pagination reuses the same planner. Later text can exhaust
            // its shared budget, but must not discard earlier completed proof.
            // Reading a cached result performs no new font/outline analysis.
            if (fractions.TryGetValue((scalar, size), out var cached)) return cached;
            if (budget.Exhausted) return null;
            double? result = null;
            NativeProtectedGlyph? outline = null;
            try
            {
                budget.Glyph();
                var originalGid = cmap.DirectGlyph(scalar);
                if (originalGid == 0) throw new NotSupportedException("No original cmap glyph for native bold analysis.");
                var renderer = new CounterRenderer(budget, originalGid);
                var font = family.CreateFont((float)size, FontStyle.Regular);
                TextRenderer.RenderTextTo(renderer, char.ConvertFromUtf32(scalar),
                    new TextOptions(font) { Dpi = 72, KerningMode = KerningMode.None });
                result = renderer.Fraction(font.Size);
                // The optional path must not change the uniform planner's result.
                // Failure to place an original curve only refuses extra paint.
                try { outline = renderer.ProtectedGlyph(font, scalar, forceIntegerPpem); }
                catch (NotSupportedException) { }
                catch (FontException) { }
                catch (InvalidDataException) { }
            }
            catch (NotSupportedException) { }
            catch (FontException) { }
            catch (InvalidDataException) { }
            if (fractions.Count < 512)
            {
                fractions.Add((scalar, size), result);
                protectedGlyphs.Add((scalar, size), outline);
            }
            return result;
        }
        internal NativeProtectedGlyph? ProtectedGlyph(int scalar, double size)
        {
            _ = Fraction(scalar, size);
            return protectedGlyphs.TryGetValue((scalar, size), out var outline) ? outline : null;
        }
    }

    private sealed class CounterRenderer(WorkBudget budget, int originalGid) : IGlyphRenderer
    {
        private readonly List<List<Point>> contours = new();
        private readonly List<NativeGlyphCurve> curves = new();
        private List<Point>[]? filledCounters;
        private readonly PdfGlyphStrokeGeometry stroke = new(.001);
        private List<Point>? current;
        private Point position;
        private int glyphs, points;
        public void BeginText(in FontRectangle bounds) { }
        public void EndText() { }
        public bool BeginGlyph(in FontRectangle bounds, in GlyphRendererParameters parameters)
        {
            if (++glyphs != 1 || parameters.GlyphId == 0 || parameters.GlyphId != originalGid ||
                parameters.GlyphType != GlyphType.Standard)
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
        public void EndFigure() { curves.Add(new NativeGlyphCurve('C')); current = null; stroke.EndFigure(); }
        public void MoveTo(Vector2 p) { curves.Add(new NativeGlyphCurve('M', p.X, p.Y)); position = new Point(p.X, p.Y); Add(position); stroke.MoveTo(p); }
        public void LineTo(Vector2 p) { curves.Add(new NativeGlyphCurve('L', p.X, p.Y)); position = new Point(p.X, p.Y); Add(position); stroke.LineTo(p); }
        public void CubicBezierTo(Vector2 a, Vector2 b, Vector2 c) =>
            throw new NotSupportedException("Native bold analysis requires quadratic TrueType outlines.");
        public void QuadraticBezierTo(Vector2 control, Vector2 target)
        {
            curves.Add(new NativeGlyphCurve('Q', target.X, target.Y, control.X, control.Y));
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
            var holes = FilledCounterContours(valid);
            filledCounters = holes;
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
                stroke.Clearance(new XPoint(w.Center.X, w.Center.Y), width, budget.Tick) >=
                Math.Max(w.Radius * MinimumRetainedWitnessFraction, MinimumRetainedWitnessRadiusPoints) + .0005);
            var lower = size * .01; var upper = size * MaximumStrokeFraction;
            if (witnesses.Count == 0 || Preserves(upper)) return MaximumStrokeFraction;
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
        internal NativeProtectedGlyph ProtectedGlyph(Font font, int scalar, bool forceIntegerPpem)
        {
            if (filledCounters is null || glyphs != 1 || curves.Count == 0 || curves.Count > 4352)
                throw new NotSupportedException("No complete original glyph for protected emphasis.");
            foreach (var command in curves)
            {
                budget.Tick();
                if (!command.Finite || Math.Abs(command.X) > 10_000 || Math.Abs(command.Y) > 10_000 ||
                    Math.Abs(command.ControlX) > 10_000 || Math.Abs(command.ControlY) > 10_000)
                    throw new NotSupportedException("Original protected curve exceeds its coordinate budget.");
            }
            // Undo Fonts1.0.1's line ascender/em centering and negative top
            // bearing exactly as the owning PDF outline route does. Analysis
            // remains translation invariant in its original callback coordinates.
            if (!font.TryGetGlyphs(new SixLabors.Fonts.Unicode.CodePoint(scalar),
                ColorFontSupport.None, out var metricsGlyphs))
                throw new NotSupportedException("No original glyph metrics for protected emphasis.");
            var count = 0; var topBearing = 0;
            foreach (var glyph in metricsGlyphs)
            {
                budget.Tick();
                if (++count != 1) throw new NotSupportedException("Multiple original glyph metrics.");
                topBearing = Math.Min(topBearing, glyph.GlyphMetrics.TopSideBearing);
            }
            if (count != 1) throw new NotSupportedException("Missing original glyph metrics.");
            var metrics = font.FontMetrics;
            var layoutAscender = metrics.HorizontalMetrics.Ascender - topBearing -
                (metrics.HorizontalMetrics.LineHeight - metrics.UnitsPerEm) * .5f;
            var offset = -layoutAscender * font.Size / metrics.UnitsPerEm;
            // Fonts1.0.1 rounds dpi*pointSize for head.ForcePPEMToInt even
            // without hinting. Preserve the font flag and uniform witness
            // analysis; only restore the optional vector replay and its
            // counter bounds to the declared size around the glyph baseline.
            var scaledSize = 72f * font.Size;
            var roundedSize = forceIntegerPpem ? MathF.Round(scaledSize) : scaledSize;
            if (roundedSize <= 0) throw new NotSupportedException("No finite protected glyph scale.");
            var correction = forceIntegerPpem ? (double)scaledSize / roundedSize : 1d;
            NativeGlyphCurve Restore(NativeGlyphCurve curve) => curve.Command == 'C' || correction == 1
                ? curve : new NativeGlyphCurve(curve.Command, curve.X * correction,
                    (curve.Y + offset) * correction - offset,
                    curve.Command == 'Q' ? curve.ControlX * correction : 0,
                    curve.Command == 'Q' ? (curve.ControlY + offset) * correction - offset : 0);
            var originalCurves = curves.Select(Restore).ToArray();
            var bounds = new List<NativeCounterBounds>();
            foreach (var hole in filledCounters)
            {
                budget.Tick();
                bounds.Add(new NativeCounterBounds(hole.Min(p => p.X) * correction,
                    (hole.Min(p => p.Y) + offset) * correction - offset,
                    hole.Max(p => p.X) * correction,
                    (hole.Max(p => p.Y) + offset) * correction - offset));
            }
            return new NativeProtectedGlyph(originalGid, originalCurves, bounds.ToArray(), offset);
        }

        private List<Point>[] FilledCounterContours(List<Point>[] input)
        {
            // Font contours describe nonzero fill, not a list of holes. For
            // example, overlapping clockwise bars can enclose a white region.
            // Use the polygon union only to enumerate counter candidates; all
            // witnesses and miter clearances still use the original contours.
            const double scale = 1_000_000;
            var paths = new Paths64();
            var edges = new List<(Point64 A, Point64 B)>();
            foreach (var contour in input)
            {
                var path = new Path64();
                foreach (var point in contour)
                {
                    budget.Tick();
                    // Keep integer magnitudes far below Clipper's precision
                    // degradation range. Invalid geometry retains legacy width.
                    if (Math.Abs(point.X) > 10_000 || Math.Abs(point.Y) > 10_000)
                        throw new NotSupportedException("Bold-analysis coordinate budget exceeded.");
                    var p = new Point64((long)Math.Round(point.X * scale), (long)Math.Round(point.Y * scale));
                    if (path.Count == 0 || p != path[path.Count - 1]) path.Add(p);
                }
                if (path.Count > 1 && path[0] == path[path.Count - 1]) path.RemoveAt(path.Count - 1);
                if (path.Count < 3) throw new NotSupportedException("Degenerate bold-analysis contour.");
                paths.Add(path);
                for (var i = 0; i < path.Count; i++) edges.Add((path[i], path[(i + 1) % path.Count]));
            }
            // Bound admitted intersection complexity before the synchronous
            // library call. Count touching/adjacent boxes too, conservatively.
            // These checks consume the existing document-wide visit budget.
            var potentialPairs = 0;
            for (var i = 0; i < edges.Count; i++)
            for (var j = i + 1; j < edges.Count; j++)
            {
                budget.Tick();
                var a = edges[i].A; var b = edges[i].B;
                var c = edges[j].A; var d = edges[j].B;
                if (Math.Max(Math.Min(a.X, b.X), Math.Min(c.X, d.X)) <= Math.Min(Math.Max(a.X, b.X), Math.Max(c.X, d.X)) &&
                    Math.Max(Math.Min(a.Y, b.Y), Math.Min(c.Y, d.Y)) <= Math.Min(Math.Max(a.Y, b.Y), Math.Max(c.Y, d.Y)) &&
                    ++potentialPairs > 8192)
                    throw new NotSupportedException("Bold-analysis intersection budget exceeded.");
            }
            var clipper = new Clipper64();
            var tree = new PolyTree64();
            budget.Tick();
            clipper.AddSubject(paths);
            var success = clipper.Execute(ClipType.Union, FillRule.NonZero, tree);
            // Execute has no cancellation hook. Always propagate a cancellation
            // observed after this bounded-input synchronous call, even on false.
            budget.Tick();
            if (!success || tree.Count == 0)
                throw new NotSupportedException("No valid filled bold-analysis topology.");
            var holes = new List<List<Point>>();
            var nodes = 0; var outputPoints = 0;
            void Visit(PolyPath64 parent, int depth)
            {
                if (depth > 64) throw new NotSupportedException("Bold-analysis nesting budget exceeded.");
                for (var i = 0; i < parent.Count; i++)
                {
                    budget.Tick();
                    var child = parent[i];
                    var polygon = child.Polygon!;
                    outputPoints += polygon.Count;
                    if (++nodes > 8192 || outputPoints > 65536)
                        throw new NotSupportedException("Bold-analysis output budget exceeded.");
                    if (child.IsHole)
                    {
                        if (holes.Count == 64) throw new NotSupportedException("Too many bold-analysis counters.");
                        var hole = new List<Point>(polygon.Count);
                        foreach (var p in polygon)
                        {
                            budget.Point();
                            hole.Add(new Point(p.X / scale, p.Y / scale));
                        }
                        holes.Add(hole);
                    }
                    Visit(child, depth + 1);
                }
            }
            Visit(tree, 0);
            return holes.ToArray();
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
