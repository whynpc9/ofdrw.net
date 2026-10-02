using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal sealed class OfdClipRegion
{
    internal List<OfdPathElement> Paths { get; } = new();
    internal bool EvenOdd { get; set; }
}

internal static class OfdClipGeometry
{
    internal static IReadOnlyList<OfdClipRegion> Read(string? xml)
    {
        var regions = new List<OfdClipRegion>();
        if (string.IsNullOrWhiteSpace(xml)) return regions;
        var root = XElement.Parse(xml!);
        var ns = root.Name.Namespace;
        foreach (var clip in root.Elements(ns + "Clip"))
        {
            var region = new OfdClipRegion();
            if (clip.Elements().Any(area => area.Name.LocalName == "Area" && area.Name.Namespace != ns))
                throw new NotSupportedException("Foreign OFD clip areas cannot be exported safely.");
            foreach (var area in clip.Elements(ns + "Area"))
            {
                var pathsBeforeArea = region.Paths.Count;
                var areaTransform = Matrix(area.Attribute("CTM")?.Value);
                foreach (var shape in area.Elements())
                {
                    if (shape.Name.Namespace != ns) continue;
                    if (shape.Name.LocalName != "Path") throw new NotSupportedException("Only path-based OFD clipping areas are supported.");
                    var literal = string.Concat(shape.Element(shape.Name.Namespace + "AbbreviatedData")?.Nodes().OfType<XText>().Select(text => text.Value) ?? Enumerable.Empty<string>());
                    if (string.IsNullOrWhiteSpace(literal)) throw new NotSupportedException("OFD clip path has no supported literal geometry.");
                    var boundary = Numbers(shape.Attribute("Boundary")?.Value);
                    var transform = Matrix(shape.Attribute("CTM")?.Value);
                    if (boundary.Length >= 2)
                    {
                        transform[4] += boundary[0];
                        transform[5] += boundary[1];
                    }
                    region.Paths.Add(new OfdPathElement
                    {
                        AbbreviatedData = literal,
                        Transform = Multiply(areaTransform, transform), Stroke = false, Fill = true
                    });
                    region.EvenOdd |= string.Equals(shape.Attribute("Rule")?.Value, "Even-Odd", StringComparison.OrdinalIgnoreCase);
                }
                if (region.Paths.Count == pathsBeforeArea) throw new NotSupportedException("OFD clip area has no supported path geometry.");
            }
            if (region.Paths.Count == 0) throw new NotSupportedException("OFD clip has no supported path geometry.");
            regions.Add(region);
        }
        return regions;
    }

    private static double[] Matrix(string? value)
    {
        var values = Numbers(value);
        if (values.Length == 0) return new double[] { 1, 0, 0, 1, 0, 0 };
        if (values.Length != 6) throw new InvalidDataException("OFD clipping transform must have six values.");
        return values;
    }

    private static double[] Numbers(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return Array.Empty<double>();
        return value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(token =>
        {
            if (!double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ||
                double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidDataException("Invalid OFD clipping coordinate.");
            return number;
        }).ToArray();
    }

    private static double[] Multiply(double[] a, double[] b) => new[]
    {
        a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]
    };
}
