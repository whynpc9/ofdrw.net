using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Render-only values. Original mutable models are never cached or changed.
internal sealed class OfdPathStyle
{
    private static readonly ConditionalWeakTable<OfdPathElement, OfdPathStyle> Projected = new();
    internal double Alpha { get; }
    internal string? Cap { get; }
    internal string? Join { get; }
    internal double? MiterLimit { get; }
    internal IReadOnlyList<double>? DashPattern { get; }
    internal double DashOffset { get; }
    internal bool IsEvenOdd { get; }

    private OfdPathStyle(OfdPathElement path)
    {
        var styleXml = path.RenderSourceXml ?? path.SourceXml;
        var source = string.IsNullOrWhiteSpace(styleXml) ? null : XElement.Parse(styleXml!);
        string? Attribute(string name) => source?.Attribute(name)?.Value;
        var alpha = Read(Attribute("Alpha"), "Alpha", 255, 0);
        if (alpha > 255 || alpha != Math.Truncate(alpha)) throw Invalid("Alpha");
        Alpha = alpha / 255d;
        Cap = Choice(Attribute("Cap"), "Cap", "butt", "round", "square");
        Join = Choice(Attribute("Join"), "Join", "miter", "round", "bevel");
        var miter = Attribute("MiterLimit");
        if (miter is not null) MiterLimit = Read(miter, "MiterLimit", 10, 1);
        // Validate source lexemes too: a reader fallback must not hide corrupt input.
        Read(Attribute("LineWidth"), "LineWidth", path.LineWidthMillimeters, 0);
        Read(null, "LineWidth", path.LineWidthMillimeters, 0);
        DashOffset = Read(Attribute("DashOffset"), "DashOffset", 0, 0);
        var pattern = Attribute("DashPattern");
        if (pattern is not null)
        {
            var values = pattern.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => Read(value, "DashPattern", 0, 0)).ToArray();
            var sum = values.Sum();
            if (values.Length == 0 || sum <= 0 || !Finite(sum) || !Finite(sum * 72d / 25.4d)) throw Invalid("DashPattern");
            // Repeating an odd pattern is the SVG/PDF definition, whereas the
            // PDF backend appends an unrelated short gap to an odd XPen array.
            if (values.Length % 2 == 1) values = values.Concat(values).ToArray();
            DashPattern = Array.AsReadOnly(values);
        }
        IsEvenOdd = path.FillRule is { } rule ? rule == OfdFillRule.EvenOdd :
            string.Equals(Attribute("Rule"), "Even-Odd", StringComparison.OrdinalIgnoreCase);
    }

    internal static OfdPathStyle Read(OfdPathElement path) => Projected.TryGetValue(path, out var style) ? style : new OfdPathStyle(path);
    internal static void CaptureProjection(OfdPathElement original, OfdPathElement projected) => Projected.Add(projected, Read(original));

    private static string? Choice(string? value, string name, params string[] allowed)
    {
        if (value is null) return null;
        var normalized = value.ToLowerInvariant();
        if (!allowed.Contains(normalized)) throw Invalid(name);
        return normalized;
    }
    private static double Read(string? value, string name, double fallback, double minimum)
    {
        var number = fallback;
        if (value is not null && !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) throw Invalid(name);
        if (!Finite(number) || number < minimum || !Finite(number * 72d / 25.4d)) throw Invalid(name);
        return number;
    }
    private static bool Finite(double number) => !double.IsNaN(number) && !double.IsInfinity(number);
    private static InvalidDataException Invalid(string name) => new($"OFD path {name} has an invalid or nonrepresentable value.");
}
