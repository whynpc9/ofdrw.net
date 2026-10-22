using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal sealed class OfdClipRegion
{
    internal List<OfdPathElement> Paths { get; } = new();
    internal bool EvenOdd { get; set; }
}

internal static class OfdClipGeometry
{
    internal static IReadOnlyList<OfdClipRegion> Read(OfdElement element, OfdModelMaterialization material, CancellationToken token)
    {
        element.ClipOwnership.Validate(element.ClippingXml, material, token);
        material.ChargeCopy(8192L + 64L * (element.ClippingXml?.Length ?? 0), token);
        var xml = OfdClipXml.Write(element, Ofdrw.Net.Core.Constants.OfdConstants.Namespace, material, token);
        return xml is null ? Array.Empty<OfdClipRegion>() : Read(xml, material, token, requireClosedProfile: element.Clips is not null);
    }
    internal static IReadOnlyList<OfdClipRegion> Read(string? xml)
    {
        var regions = new List<OfdClipRegion>();
        if (string.IsNullOrWhiteSpace(xml)) return regions;
        var root = XElement.Parse(xml!);
        return Read(root, null, default);
    }

    internal static IReadOnlyList<OfdClipRegion> Read(XElement root, OfdModelMaterialization? material, CancellationToken token, bool requireClosedProfile = false)
    {
        var regions = new List<OfdClipRegion>();
        var ns = root.Name.Namespace;
        if (root.Name.LocalName != "Clips" || (ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.Namespace && ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.StandardNamespace) ||
            requireClosedProfile && (!OfdGraphicXmlContract.HasKnownChildren(root) || OfdGraphicXmlContract.HasUnsupportedReferences(root) || root.DescendantsAndSelf().Attributes().Any(a => !OfdGraphicXmlContract.IsKnownAttribute(a))))
            throw new NotSupportedException("Clip extensions or references are outside the render profile.");
        foreach (var clip in root.Elements(ns + "Clip"))
        {
            token.ThrowIfCancellationRequested(); material?.ChargeCopy(256, token);
            var region = new OfdClipRegion();
            if (root.Attribute("TransFlag") is not null || clip.Elements(ns + "Area").Count() != 1)
                throw new NotSupportedException("Export supports exactly one Path Area per Clip.");
            if (clip.Elements().Any(area => area.Name.LocalName == "Area" && area.Name.Namespace != ns))
                throw new NotSupportedException("Foreign OFD clip areas cannot be exported safely.");
            foreach (var area in clip.Elements(ns + "Area"))
            {
                token.ThrowIfCancellationRequested(); material?.ChargeCopy(512L + 16L * (area.Attribute("CTM")?.Value.Length ?? 0), token);
                if (area.Attribute("DrawParam") is not null || area.Elements().Count(n => n.Name.Namespace == ns) != 1)
                    throw new NotSupportedException("Clip drawing parameters and text/union geometry are unsupported.");
                var pathsBeforeArea = region.Paths.Count;
                var areaTransform = Matrix(area.Attribute("CTM")?.Value);
                foreach (var shape in area.Elements())
                {
                    token.ThrowIfCancellationRequested();
                    if (shape.Name.Namespace != ns) continue;
                    material?.ChargeCopy(1024L + 16L * (shape.Attribute("CTM")?.Value.Length ?? 0) + 16L * (shape.Attribute("Boundary")?.Value.Length ?? 0), token);
                    if (shape.Name.LocalName != "Path") throw new NotSupportedException("Only path-based OFD clipping areas are supported.");
                    if (shape.Attribute("DrawParam") is not null || shape.Element(ns + "Clips") is not null) throw new NotSupportedException("Nested clip styling is unsupported.");
                    if (shape.Attribute("Stroke")?.Value is "true" or "1" || shape.Attribute("Fill")?.Value is "false" or "0")
                        throw new NotSupportedException("Stroked clip geometry is outside the render profile.");
                    var literal = string.Concat(shape.Element(shape.Name.Namespace + "AbbreviatedData")?.Nodes().OfType<XText>().Select(text => text.Value) ?? Enumerable.Empty<string>());
                    if (string.IsNullOrWhiteSpace(literal)) throw new NotSupportedException("OFD clip path has no supported literal geometry.");
                    if (!OfdPathDataSyntax.IsValidLiteral(literal, token) ||
                        requireClosedProfile && !OfdPathDataSyntax.HasSingleValidDrawing(shape, token))
                        throw new NotSupportedException("Malformed OFD clip path data cannot be exported.");
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
                    if (shape.Attribute("Rule")?.Value is { } rule && rule is not ("NonZero" or "Even-Odd")) throw new NotSupportedException("Unknown path fill rule.");
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
