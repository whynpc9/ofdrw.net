using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Reader.Readers;

// Import-time, read-only paint resolution. The source element is never changed;
// unsupported paint remains readable/writable and fails only on rendering.
internal sealed class OfdPaintResources
{
    private readonly Dictionary<string, XElement> spaces;
    private readonly Dictionary<string, XElement> parameters;
    private readonly HashSet<string> ambiguous;

    internal OfdPaintResources()
    {
        spaces = new(StringComparer.Ordinal); parameters = new(StringComparer.Ordinal);
        ambiguous = new(StringComparer.Ordinal);
    }

    internal OfdPaintResources(OfdPaintResources parent)
    {
        spaces = new(parent.spaces, StringComparer.Ordinal);
        parameters = new(parent.parameters, StringComparer.Ordinal);
        ambiguous = new(parent.ambiguous, StringComparer.Ordinal);
    }

    internal void Read(OfdPackageArchive archive, string path, OfdModelMaterialization material, CancellationToken token)
    {
        if (!archive.Contains(path)) return;
        var xml = material.ReadXml(path, archive.GetBytes(path), token);
        var ns = xml.Root?.Name.Namespace;
        if (ns is null || ns != XNamespace.None && ns != OfdConstants.Namespace && ns != OfdConstants.StandardNamespace ||
            xml.Root!.Name != ns + "Res") return;
        foreach (var container in xml.Root.Elements())
        {
            var table = container.Name.LocalName switch
            {
                "ColorSpaces" => spaces,
                "DrawParams" => parameters,
                _ => null
            };
            if (table is null) continue;
            foreach (var node in container.Elements())
            {
                token.ThrowIfCancellationRequested(); material.ReserveReadModel(256, token);
                if (node.Name != ns + (ReferenceEquals(table, spaces) ? "ColorSpace" : "DrawParam")) continue;
                var id = node.Attribute("ID")?.Value;
                if (string.IsNullOrWhiteSpace(id)) continue;
                if (ambiguous.Contains(id!) || table.ContainsKey(id!))
                { ambiguous.Add(id!); table.Remove(id!); }
                else table.Add(id!, node);
            }
        }
    }

    internal (XElement Node, string? Error) Normalize(XElement source, int sourceXmlCharacters,
        OfdModelMaterialization material, CancellationToken token)
    {
        try
        {
            // Reuse the caller's source-envelope length instead of creating a
            // second serialization string solely to account for this clone.
            token.ThrowIfCancellationRequested();
            material.ChargeCopy(checked(1024L + 16L * sourceXmlCharacters), token);
            var node = new XElement(source);
            if (node.Attribute("DrawParam") is { } draw)
            {
                var style = ResolveDrawParam(draw.Value, token);
                foreach (var attribute in style.Attributes())
                    if (attribute.Name.LocalName is not ("ID" or "Relative") && node.Attribute(attribute.Name) is null)
                        node.SetAttributeValue(attribute.Name, attribute.Value);
                foreach (var color in style.Elements())
                    if (color.Name.LocalName is "FillColor" or "StrokeColor" && node.Element(color.Name) is null)
                        node.Add(new XElement(color));
                draw.Remove();
            }
            foreach (var color in node.Elements().Where(n => n.Name.LocalName is "FillColor" or "StrokeColor").ToArray())
                NormalizeColor(color, token);
            return (node, null);
        }
        catch (Exception error) when ((error is NotSupportedException or InvalidDataException or FormatException or OverflowException) &&
            !OfdTraversalBudget.IsLimit(error))
        { return (source, "Unsupported OFD paint: " + error.Message); }
    }

    private XElement ResolveDrawParam(string id, CancellationToken token)
    {
        var chain = new List<XElement>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        for (string? current = id; ;)
        {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(current) || chain.Count >= 32 || !seen.Add(current))
                throw new NotSupportedException("DrawParam inheritance cycle or depth limit.");
            if (ambiguous.Contains(current) || !parameters.TryGetValue(current, out var item))
                throw new NotSupportedException($"Missing or ambiguous DrawParam '{current}'.");
            chain.Add(item);
            current = item.Attribute("Relative")?.Value;
            if (string.IsNullOrWhiteSpace(current)) break;
        }
        var result = new XElement(chain[0].Name);
        foreach (var item in chain.AsEnumerable().Reverse())
        {
            if (item.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not
                    ("ID" or "Relative" or "LineWidth" or "Join" or "Cap" or "DashOffset" or "DashPattern" or "MiterLimit"))))
                throw new NotSupportedException("DrawParam has an unknown attribute.");
            if (item.Elements().Any(e => e.Name.Namespace != item.Name.Namespace || e.Name.LocalName is not ("FillColor" or "StrokeColor")))
                throw new NotSupportedException("DrawParam has an unknown child.");
            foreach (var attribute in item.Attributes().Where(a => a.Name.LocalName is not ("ID" or "Relative")))
                result.SetAttributeValue(attribute.Name, attribute.Value);
            foreach (var color in item.Elements())
            { result.Element(color.Name)?.Remove(); result.Add(new XElement(color)); }
        }
        return result;
    }

    private void NormalizeColor(XElement color, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (color.HasElements || color.Attributes().Any(a => !a.IsNamespaceDeclaration &&
            (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("Value" or "Index" or "ColorSpace" or "Alpha"))))
            throw new NotSupportedException("Color extension, pattern or gradient.");
        var alpha = Integer(color.Attribute("Alpha")?.Value, 255);
        if (alpha is < 0 or > 255) throw new InvalidDataException("Color Alpha is outside 0..255.");
        var spaceId = color.Attribute("ColorSpace")?.Value;
        XElement? space = null;
        if (spaceId is not null)
        {
            if (ambiguous.Contains(spaceId) || !spaces.TryGetValue(spaceId, out space))
                throw new NotSupportedException($"Missing or ambiguous ColorSpace '{spaceId}'.");
        }
        var type = space?.Attribute("Type")?.Value ?? "RGB";
        if (space is not null)
        {
            if (space.Attribute("Profile") is not null ||
                Integer(space.Attribute("BitsPerComponent")?.Value, 8) != 8 ||
                space.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Type" or "BitsPerComponent" or "Profile"))) ||
                space.Elements().Any(e => e.Name != space.Name.Namespace + "Palette"))
                throw new NotSupportedException("ColorSpace profile, bit depth or extension.");
        }
        var count = type switch { "GRAY" => 1, "RGB" => 3, "CMYK" => 4,
            _ => throw new NotSupportedException($"ColorSpace type '{type}'.") };
        int[] components;
        if (color.Attribute("Value") is { } value)
        {
            var actualCount = space is null && value.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length == 1 ? 1 : count;
            components = Components(value.Value, actualCount);
        }
        else if (color.Attribute("Index") is { } index)
        {
            if (space is null) throw new NotSupportedException("Color Index needs a ColorSpace palette.");
            var palette = space.Element(space.Name.Namespace + "Palette");
            if (palette is null || palette.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                palette.Elements().Any(e => e.Name != space.Name.Namespace + "CV") ||
                palette.Nodes().Any(n => n is not XElement and not XText ||
                    n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Missing or unsupported palette.");
            var entries = palette.Elements().ToArray();
            if (entries.Length == 0 || entries.Length > 256) throw new NotSupportedException("Palette size outside supported range.");
            if (entries.Any(entry => entry.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                entry.Nodes().Any(n => n is not XText)))
                throw new NotSupportedException("Palette entry has unsupported content.");
            var ordinal = Integer(index.Value, -1);
            if (ordinal < 0 || ordinal >= entries.Length) throw new InvalidDataException("Palette Index is out of range.");
            components = Components(entries[ordinal].Value, count);
        }
        else components = new int[count];
        var rgb = components.Length switch
        {
            1 => new[] { components[0], components[0], components[0] },
            3 => components,
            4 => new[] { Ink(components[0], components[3]), Ink(components[1], components[3]), Ink(components[2], components[3]) },
            _ => throw new InvalidDataException("Invalid color component count.")
        };
        color.RemoveAttributes();
        color.SetAttributeValue("Value", string.Join(" ", rgb));
        if (alpha != 255) color.SetAttributeValue("Alpha", alpha.ToString(CultureInfo.InvariantCulture));
    }

    private static int Ink(int component, int black) => (255 - component) * (255 - black) / 255;
    private static int[] Components(string value, int count)
    {
        var parts = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != count) throw new InvalidDataException("Color component count does not match ColorSpace.");
        var result = new int[count];
        for (var i = 0; i < count; i++)
        {
            result[i] = Integer(parts[i], -1);
            if (result[i] is < 0 or > 255) throw new InvalidDataException("Color component is outside 0..255.");
        }
        return result;
    }

    private static int Integer(string? value, int fallback) => value is null ? fallback :
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number) ? number :
        throw new InvalidDataException("Invalid OFD paint integer.");
}
