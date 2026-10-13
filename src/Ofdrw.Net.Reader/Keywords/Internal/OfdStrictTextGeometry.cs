using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

internal sealed class OfdKeywordTextSource
{
    internal int PageIndex, LayerOrdinal, ElementOrdinal;
    internal int[] OccurrencePath = Array.Empty<int>(), DefinitionPath = Array.Empty<int>();
    internal string? EntryPath;
    internal string Path = string.Empty, FontId = string.Empty;
    internal bool HasXml, Editable, Visible = true, Italic;
    internal int Weight = 400;
    internal double X, Y, Width, Height, Size, PageX, PageY, PageWidth, PageHeight;
    internal double[] Matrix = new double[] { 1, 0, 0, 1, 0, 0 };
    internal OfdGlyphFontSnapshot? Font;
    internal readonly List<OfdKeywordSourceRun> Runs = new();
}

internal sealed class OfdKeywordSourceRun(string text, int[] offsets, string path, double x, double y, string? dx, string? dy)
{
    internal readonly string Text = text, Path = path;
    internal readonly int[] SourceOffsets = offsets;
    internal readonly double X = x, Y = y;
    internal readonly string? DeltaX = dx, DeltaY = dy;
}

internal sealed class OfdPositionedCluster(OfdKeywordTextSource source, OfdKeywordSourceRun run,
    OfdGlyphClusterGeometry geometry, OfdKeywordPoint baseline, OfdKeywordPoint advance, OfdKeywordQuadrilateral? quad)
{
    internal readonly OfdKeywordTextSource Source = source;
    internal readonly OfdKeywordSourceRun Run = run;
    internal readonly OfdGlyphClusterGeometry Geometry = geometry;
    internal readonly OfdKeywordPoint Baseline = baseline, Advance = advance;
    internal readonly OfdKeywordQuadrilateral? Quad = quad;
}

internal static class OfdStrictTextGeometry
{
    internal static OfdKeywordTextSource Capture(OfdGraphicOccurrence occurrence, int layer, string path,
        OfdKeywordLocatorBudget budget, int reservedUnits)
    {
        var text = (OfdTextElement)occurrence.Element; var page = occurrence.Page;
        var result = new OfdKeywordTextSource
        {
            PageIndex = occurrence.OccurrencePath[0], LayerOrdinal = layer, ElementOrdinal = occurrence.OccurrencePath[3],
            OccurrencePath = occurrence.OccurrencePath.ToArray(), DefinitionPath = occurrence.DefinitionPath.ToArray(),
            EntryPath = occurrence.SourceEntryPath, Path = path, HasXml = text.SourceXml is not null,
            Editable = occurrence.IsEditableSource && !occurrence.IsInComposite,
            PageX = page.XMillimeters, PageY = page.YMillimeters, PageWidth = page.WidthMillimeters, PageHeight = page.HeightMillimeters
        };
        if (occurrence.ParentToPage.Count != 6 || !occurrence.ParentToPage.SequenceEqual(new double[] { 1, 0, 0, 1, 0, 0 }))
            throw new NotSupportedException("UnprovedParentTransform");
        if (!string.IsNullOrEmpty(text.ClippingXml)) throw new NotSupportedException("ExplicitClipping");
        if (text.SourceXml is { } source)
        {
            var root = budget.Parse(source);
            if (root.Name.LocalName != "TextObject" || !OfdKeywordSourceAdmission.KnownNamespace(root.Name.Namespace))
                throw new NotSupportedException("UnsupportedTextNamespace");
            Attributes(root, "ID", "Name", "Visible", "Boundary", "CTM", "Font", "Size", "Weight", "Italic",
                "Stroke", "Fill", "HScale", "ReadDirection", "CharDirection", "Alpha");
            result.Visible = Boolean(root.Attribute("Visible")?.Value, true);
            result.FontId = Required(root, "Font"); result.Size = Number(Required(root, "Size"));
            result.Weight = Integer(root.Attribute("Weight")?.Value ?? "400"); result.Italic = Boolean(root.Attribute("Italic")?.Value, false);
            if (text.ObjectId != root.Attribute("ID")?.Value || text.FontResourceId != result.FontId ||
                text.Weight != result.Weight || text.Italic != result.Italic)
                throw new NotSupportedException("SourceModelOverlayConflict");
            var boundary = Numbers(Required(root, "Boundary"), 4); SetBoundary(result, boundary);
            result.Matrix = root.Attribute("CTM") is { } matrix ? Numbers(matrix.Value, 6) : result.Matrix;
            if (Boolean(root.Attribute("Stroke")?.Value, false) || !Boolean(root.Attribute("Fill")?.Value, true) ||
                Number(root.Attribute("HScale")?.Value ?? "1") != 1 || Integer(root.Attribute("ReadDirection")?.Value ?? "0") != 0 ||
                Integer(root.Attribute("CharDirection")?.Value ?? "0") != 0 || Integer(root.Attribute("Alpha")?.Value ?? "255") != 255)
                throw new NotSupportedException("UnsupportedTextDrawingState");
            var ns = root.Name.Namespace;
            var codes = root.Elements(ns + "TextCode").ToArray();
            long rawUnits = 0;
            foreach (var code in codes)
                foreach (var node in code.Nodes().OfType<XText>()) { budget.Check(); rawUnits += node.Value.Length; }
            if (rawUnits > int.MaxValue) throw new InvalidDataException("Source text capacity exceeded.");
            if (rawUnits > reservedUnits) budget.Text((int)rawUnits - reservedUnits);
            foreach (var child in root.Elements())
            {
                budget.Check();
                if (child.Name == ns + "TextCode") continue;
                if (child.Name != ns + "FillColor") throw new NotSupportedException("UnsupportedTextChild");
                Attributes(child, "Value", "Alpha");
                if (child.HasElements || Integer(child.Attribute("Alpha")?.Value ?? "255") != 255 ||
                    Numbers(Required(child, "Value"), 3).Any(v => v < 0 || v > 255 || v != Math.Truncate(v)))
                    throw new NotSupportedException("UnsupportedFillColor");
            }
            if (root.Elements(ns + "FillColor").Count() > 1) throw new NotSupportedException("AmbiguousFillColor");
            PlainTextNodes(root, allowText: false);
            double? previousX = null, previousY = null;
            for (var i = 0; i < codes.Length; i++)
            {
                budget.Check(); var code = codes[i]; Attributes(code, "X", "Y", "DeltaX", "DeltaY");
                if (code.HasElements) throw new NotSupportedException("NestedTextCodeContent");
                PlainTextNodes(code, allowText: true);
                var x = code.Attribute("X") is { } xa ? Number(xa.Value) : previousX ?? throw new InvalidDataException("First TextCode requires X.");
                var y = code.Attribute("Y") is { } ya ? Number(ya.Value) : previousY ?? throw new InvalidDataException("First TextCode requires Y.");
                previousX = x; previousY = y; // Table46 inherits the previous START, never its last glyph.
                var raw = string.Concat(code.Nodes().OfType<XText>().Select(node => node.Value));
                var decoded = Decode(raw, budget);
                if (decoded.Text.Length == 0) throw new NotSupportedException("EmptyTextCode");
                result.Runs.Add(new OfdKeywordSourceRun(decoded.Text, decoded.Offsets,
                    path + "/" + code.Name + "[" + i.ToString(CultureInfo.InvariantCulture) + "]", x, y,
                    code.Attribute("DeltaX")?.Value, code.Attribute("DeltaY")?.Value));
            }
        }
        else
        {
            result.FontId = text.FontResourceId ?? throw new NotSupportedException("MissingExplicitFontIdentity");
            result.Size = text.FontSizeMillimeters; result.Weight = text.Weight; result.Italic = text.Italic;
            SetBoundary(result, new[] { text.XMillimeters, text.YMillimeters, text.WidthMillimeters, text.HeightMillimeters });
            if (text.Transform is not null) result.Matrix = (double[])text.Transform.Clone();
            if (text.FillColor is null || text.FillColor.Alpha != 255) throw new NotSupportedException("UnsupportedFillOpacity");
            for (var i = 0; i < text.Runs.Count; i++)
            {
                var run = text.Runs[i]; budget.Check();
                if (run.Text.Length == 0) throw new NotSupportedException("EmptyTextCode");
                var offsets = new int[run.Text.Length + 1];
                for (var p = 0; p < offsets.Length; p++) { if ((p & 255) == 0) budget.Check(); offsets[p] = p; }
                result.Runs.Add(new OfdKeywordSourceRun(run.Text, offsets, path + "/model:Runs[" + i + "]",
                    Finite(run.XMillimeters), Finite(run.YMillimeters), run.DeltaX, run.DeltaY));
            }
        }
        if (result.Runs.Count == 0) throw new NotSupportedException("NoExplicitTextRunAnchors");
        if (string.IsNullOrWhiteSpace(result.FontId) || result.Size <= 0 || !OfdGlyphBox.Finite(result.Size) ||
            result.Weight < 100 || result.Weight > 900 || result.Weight % 100 != 0)
            throw new InvalidDataException("Invalid exact text font or style.");
        ValidateMatrix(result.Matrix);
        foreach (var run in result.Runs) ValidateUnicode(run.Text, budget);
        return result;
    }

    private static void SetBoundary(OfdKeywordTextSource source, double[] values)
    {
        if (values.Any(v => !OfdGlyphBox.Finite(v)) || values[2] <= 0 || values[3] <= 0)
            throw new InvalidDataException("Invalid positive text boundary.");
        source.X = values[0]; source.Y = values[1]; source.Width = values[2]; source.Height = values[3];
    }
    internal static void ValidateMatrix(double[] matrix)
    {
        if (matrix.Length != 6 || matrix.Any(v => !OfdGlyphBox.Finite(v))) throw new InvalidDataException("Invalid finite six-value transform.");
        var determinant = matrix[0] * matrix[3] - matrix[1] * matrix[2];
        if (!OfdGlyphBox.Finite(determinant) || determinant == 0) throw new NotSupportedException("DegenerateTextTransform");
    }
    internal static List<OfdPositionedCluster> Position(OfdKeywordTextSource source, OfdKeywordSourceRun run,
        OfdGlyphGeometry geometry, OfdKeywordLocatorBudget budget, int remainingClusters)
    {
        if (geometry.Clusters.Count > remainingClusters) throw new InvalidDataException("Provider cluster capacity exceeded.");
        var cursor = 0;
        foreach (var cluster in geometry.Clusters)
        {
            budget.Check();
            if (cluster is null || cluster.Utf16Start != cursor || cursor >= run.Text.Length ||
                cluster.Utf16Length != (char.IsHighSurrogate(run.Text[cursor]) ? 2 : 1) || cluster.Utf16Length > run.Text.Length - cursor ||
                cluster.GlyphId == 0 || !OfdGlyphBox.Finite(cluster.AdvanceXMillimeters) || !OfdGlyphBox.Finite(cluster.AdvanceYMillimeters))
                throw new InvalidDataException("Provider must cover each source scalar exactly once in source order.");
            if (cluster.FontBoundsMillimeters is { } bounds &&
                (!OfdGlyphBox.Finite(bounds.Left) || !OfdGlyphBox.Finite(bounds.Top) || !OfdGlyphBox.Finite(bounds.Right) ||
                 !OfdGlyphBox.Finite(bounds.Bottom) || bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top))
                throw new InvalidDataException("Invalid provider font box.");
            cursor += cluster.Utf16Length;
        }
        if (cursor != run.Text.Length) throw new InvalidDataException("Provider mapping does not cover source text.");
        var count = geometry.Clusters.Count;
        var dx = Deltas(run.DeltaX, count - 1, budget); var dy = Deltas(run.DeltaY, count - 1, budget);
        // Validate every cumulative origin before clipping can classify an early
        // glyph as unsupported and otherwise hide a malformed later overflow.
        var xs = new double[count]; var ys = new double[count]; xs[0] = run.X; ys[0] = run.Y;
        for (var i = 1; i < count; i++)
        { budget.Check(); xs[i] = Finite(xs[i - 1] + dx[i - 1]); ys[i] = Finite(ys[i - 1] + dy[i - 1]); }
        var positions = new List<OfdPositionedCluster>(count);
        for (var i = 0; i < count; i++)
        {
            budget.Check(); var cluster = geometry.Clusters[i]; var x = xs[i]; var y = ys[i];
            var baseline = Transform(source, x, y, checkClip: false);
            var m = source.Matrix;
            var advance = new OfdKeywordPoint(m[0] * cluster.AdvanceXMillimeters + m[2] * cluster.AdvanceYMillimeters,
                m[1] * cluster.AdvanceXMillimeters + m[3] * cluster.AdvanceYMillimeters);
            OfdKeywordQuadrilateral? quad = null;
            if (cluster.FontBoundsMillimeters is { } box)
                quad = new OfdKeywordQuadrilateral(Transform(source, x + box.Left, y + box.Top, true),
                    Transform(source, x + box.Right, y + box.Top, true), Transform(source, x + box.Right, y + box.Bottom, true),
                    Transform(source, x + box.Left, y + box.Bottom, true));
            positions.Add(new OfdPositionedCluster(source, run, cluster, baseline, advance, quad));
        }
        return positions;
    }
    private static OfdKeywordPoint Transform(OfdKeywordTextSource source, double x, double y, bool checkClip)
    {
        var m = source.Matrix;
        var localX = Finite(m[0] * Finite(x) + m[2] * Finite(y) + m[4]);
        var localY = Finite(m[1] * x + m[3] * y + m[5]);
        if (checkClip && (localX < 0 || localY < 0 || localX > source.Width || localY > source.Height))
            throw new NotSupportedException("BoundaryClipping");
        var point = new OfdKeywordPoint(localX + source.X - source.PageX, localY + source.Y - source.PageY);
        if (checkClip && (point.X < 0 || point.Y < 0 || point.X > source.PageWidth || point.Y > source.PageHeight))
            throw new NotSupportedException("PhysicalBoxClipping");
        return point;
    }
    internal static bool Continuous(OfdPositionedCluster previous, OfdPositionedCluster next, bool sameRun)
    {
        var a = previous.Source; var b = next.Source;
        if (a.PageIndex != b.PageIndex || a.LayerOrdinal != b.LayerOrdinal) return false;
        if (!ReferenceEquals(a, b) && b.ElementOrdinal != a.ElementOrdinal + 1) return false;
        var am = a.Matrix; var bm = b.Matrix;
        var al = Math.Sqrt(am[0] * am[0] + am[1] * am[1]); var bl = Math.Sqrt(bm[0] * bm[0] + bm[1] * bm[1]);
        if (!OfdGlyphBox.Finite(al) || !OfdGlyphBox.Finite(bl) || al == 0 || bl == 0) return false;
        var ux = am[0] / al; var uy = am[1] / al;
        var dot = ux * bm[0] / bl + uy * bm[1] / bl;
        if (dot < Math.Cos(Math.PI / 1800)) return false; // 0.1 degrees.
        var dx = next.Baseline.X - previous.Baseline.X; var dy = next.Baseline.Y - previous.Baseline.Y;
        if (Math.Abs(dx * -uy + dy * ux) > 0.01) return false;
        if (sameRun) return true; // Explicit scalar order can contain negative/coincident X deltas.
        var advance = previous.Advance.X * ux + previous.Advance.Y * uy;
        if (advance <= 0 || dx * ux + dy * uy < 0) return false;
        return Math.Abs(dx - previous.Advance.X) <= 0.01 && Math.Abs(dy - previous.Advance.Y) <= 0.01;
    }

    internal static double[] Deltas(string? value, int expected, OfdKeywordLocatorBudget budget)
    {
        if (expected < 0) throw new InvalidDataException("Empty mapped run.");
        budget.Delta(expected); var result = new double[expected];
        if (value is null) return result; // Table46: zero movement, never font advance.
        var tokens = new Tokens(value, budget); var count = 0; var any = false;
        while (tokens.Next() is { } token)
        {
            any = true; var repeat = 1; double number;
            if (token == "g")
            {
                var n = tokens.Next() ?? throw new InvalidDataException("Missing delta repeat count.");
                if (!int.TryParse(n, NumberStyles.None, CultureInfo.InvariantCulture, out repeat) || repeat <= 0)
                    throw new InvalidDataException("Invalid positive delta repeat count.");
                number = Number(tokens.Next() ?? throw new InvalidDataException("Missing repeated delta value."));
            }
            else number = Number(token);
            if (repeat > expected - count) throw new InvalidDataException("Excess delta values.");
            for (var i = 0; i < repeat; i++) { if ((i & 255) == 0) budget.Check(); result[count++] = number; }
        }
        if (!any || count != expected) throw new InvalidDataException("Delta array must contain exactly one value per inter-glyph displacement.");
        return result;
    }
    private sealed class Tokens(string value, OfdKeywordLocatorBudget budget)
    {
        private int index;
        internal string? Next()
        {
            while (index < value.Length && char.IsWhiteSpace(value[index])) { if ((index & 255) == 0) budget.Check(); index++; }
            if (index == value.Length) return null;
            var start = index;
            while (index < value.Length && !char.IsWhiteSpace(value[index]))
            { if ((index & 255) == 0) budget.Check(); if (index - start >= 128) throw new InvalidDataException("Numeric token capacity exceeded."); index++; }
            budget.DeltaToken(); return value.Substring(start, index - start);
        }
    }
    internal static (string Text, int[] Offsets) Decode(string raw, OfdKeywordLocatorBudget budget)
    {
        var text = new StringBuilder(raw.Length); var offsets = new List<int>(raw.Length + 1);
        for (var i = 0; i < raw.Length;)
        {
            budget.Check(); offsets.Add(i); var c = raw[i++];
            if (c == '\\')
            {
                if (raw.Length - i < 4) throw new NotSupportedException("UnsupportedTextEncoding");
                var number = 0;
                for (var n = 0; n < 4; n++)
                {
                    var digit = raw[i++]; var hex = digit >= '0' && digit <= '9' ? digit - '0' :
                        digit >= 'A' && digit <= 'F' ? digit - 'A' + 10 : digit >= 'a' && digit <= 'f' ? digit - 'a' + 10 : -1;
                    if (hex < 0) throw new NotSupportedException("UnsupportedTextEncoding");
                    number = number * 16 + hex;
                }
                c = (char)number;
            }
            else if (char.IsWhiteSpace(c)) throw new NotSupportedException("UnescapedSourceWhitespace");
            if (c == '\u00a4' || c < 0x20) throw new NotSupportedException("UnsupportedTextEncoding");
            text.Append(c);
        }
        offsets.Add(raw.Length); return (text.ToString(), offsets.ToArray());
    }
    private static void ValidateUnicode(string text, OfdKeywordLocatorBudget budget)
    {
        for (var i = 0; i < text.Length; i++)
        {
            budget.Check();
            if (char.IsHighSurrogate(text[i])) { if (++i >= text.Length || !char.IsLowSurrogate(text[i])) throw new InvalidDataException("Unpaired source surrogate."); }
            else if (char.IsLowSurrogate(text[i])) throw new InvalidDataException("Unpaired source surrogate.");
        }
    }
    internal static void Attributes(XElement node, params string[] allowed)
    {
        foreach (var attribute in node.Attributes())
            if (!attribute.IsNamespaceDeclaration && (attribute.Name.Namespace != XNamespace.None || !allowed.Contains(attribute.Name.LocalName)))
                throw new NotSupportedException("UnsupportedSourceAttribute");
    }
    internal static void PlainTextNodes(XElement node, bool allowText)
    {
        foreach (var child in node.Nodes())
            if (child is not XElement and not XComment and not XText || child is XText text && !allowText && !string.IsNullOrWhiteSpace(text.Value))
                throw new NotSupportedException("UnsupportedSourceNode");
    }
    internal static string Required(XElement node, string attribute) => node.Attribute(attribute)?.Value
        ?? throw new InvalidDataException("Missing required source attribute.");
    internal static double Number(string text)
    {
        if (text.Length > 128 || !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            throw new InvalidDataException("Invalid finite source number.");
        return Finite(number);
    }
    internal static int Integer(string text)
    {
        if (!int.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var number)) throw new InvalidDataException("Invalid source integer.");
        return number;
    }
    internal static bool Boolean(string? value, bool fallback) => value switch
    { null => fallback, "true" or "1" => true, "false" or "0" => false, _ => throw new InvalidDataException("Invalid source boolean.") };
    internal static double[] Numbers(string value, int count)
    {
        // Fixed-arity attribute, bounded before Split; Delta arrays use the streaming tokenizer above.
        if (value.Length > count * 130) throw new InvalidDataException("Numeric attribute capacity exceeded.");
        var parts = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != count) throw new InvalidDataException("Invalid source numeric arity.");
        return parts.Select(Number).ToArray();
    }
    private static double Finite(double value) => OfdGlyphBox.Finite(value) ? value : throw new InvalidDataException("Nonfinite source geometry.");
}
