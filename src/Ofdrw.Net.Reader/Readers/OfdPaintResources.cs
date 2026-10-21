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
    private readonly Dictionary<long, XElement> spaces;
    private readonly Dictionary<long, XElement> parameters;
    private readonly HashSet<long> ambiguous;

    internal OfdPaintResources()
    {
        spaces = new(); parameters = new();
        ambiguous = new();
    }

    internal OfdPaintResources(OfdPaintResources parent)
    {
        spaces = new(parent.spaces);
        parameters = new(parent.parameters);
        ambiguous = new(parent.ambiguous);
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
            if (container.Name.Namespace != ns) continue;
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
                var id = OfdSourceIdentity.ParseId(node.Attribute("ID")?.Value);
                if (id == 0) continue;
                if (ambiguous.Contains(id) || table.ContainsKey(id))
                { ambiguous.Add(id); table.Remove(id); }
                else table.Add(id, node);
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
            RejectDuplicateColors(node);
            if (node.Attribute("DrawParam") is { } draw)
            {
                var style = ResolveDrawParam(draw.Value, token);
                foreach (var attribute in style.Attributes())
                    if (attribute.Name.LocalName is not ("ID" or "Relative") && node.Attribute(attribute.Name) is null)
                        node.SetAttributeValue(attribute.Name, attribute.Value);
                foreach (var color in style.Elements())
                    if (color.Name.LocalName is "FillColor" or "StrokeColor" &&
                        !node.Elements().Any(existing => existing.Name.LocalName == color.Name.LocalName))
                    {
                        var inherited = new XElement(color);
                        inherited.Name = node.Name.Namespace + color.Name.LocalName;
                        node.Add(inherited);
                    }
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
        var chain = new List<XElement>(); var seen = new HashSet<long>();
        for (string? current = id; ;)
        {
            token.ThrowIfCancellationRequested();
            var numericId = OfdSourceIdentity.ParseId(current);
            if (numericId == 0 || chain.Count >= 32 || !seen.Add(numericId))
                throw new NotSupportedException("DrawParam inheritance cycle or depth limit.");
            if (ambiguous.Contains(numericId) || !parameters.TryGetValue(numericId, out var item))
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
            if (HasUnsupportedNodeContent(item))
                throw new NotSupportedException("DrawParam has unsupported content.");
            RejectDuplicateColors(item);
            // An overridden ancestor is still part of the referenced paint
            // chain. Validate its colors before a descendant can discard them.
            foreach (var color in item.Elements()) NormalizeColor(new XElement(color), token);
            foreach (var attribute in item.Attributes().Where(a => a.Name.LocalName is not ("ID" or "Relative")))
                result.SetAttributeValue(attribute.Name, attribute.Value);
            foreach (var color in item.Elements())
            {
                foreach (var prior in result.Elements().Where(existing => existing.Name.LocalName == color.Name.LocalName).ToArray())
                    prior.Remove();
                var inherited = new XElement(color);
                inherited.Name = result.Name.Namespace + color.Name.LocalName;
                result.Add(inherited);
            }
        }
        return result;
    }

    private static void RejectDuplicateColors(XElement node)
    {
        if (node.Elements().Count(child => child.Name.LocalName == "FillColor") > 1 ||
            node.Elements().Count(child => child.Name.LocalName == "StrokeColor") > 1)
            throw new NotSupportedException("Drawing has duplicate FillColor or StrokeColor children.");
    }

    private void NormalizeColor(XElement color, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (color.HasElements || HasUnsupportedNodeContent(color) || color.Attributes().Any(a => !a.IsNamespaceDeclaration &&
            (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("Value" or "Index" or "ColorSpace" or "Alpha"))))
            throw new NotSupportedException("Color extension, pattern or gradient.");
        var alpha = Integer(color.Attribute("Alpha")?.Value, 255);
        if (alpha is < 0 or > 255) throw new InvalidDataException("Color Alpha is outside 0..255.");
        var spaceId = color.Attribute("ColorSpace")?.Value;
        XElement? space = null;
        if (spaceId is not null)
        {
            var numericId = OfdSourceIdentity.ParseId(spaceId);
            if (numericId == 0 || ambiguous.Contains(numericId) || !spaces.TryGetValue(numericId, out space))
                throw new NotSupportedException($"Missing or ambiguous ColorSpace '{spaceId}'.");
        }
        var type = space?.Attribute("Type")?.Value ?? "RGB";
        if (space is not null)
        {
            if (space.Attribute("Profile") is not null || HasUnsupportedNodeContent(space) ||
                Integer(space.Attribute("BitsPerComponent")?.Value, 8) != 8 ||
                space.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Type" or "BitsPerComponent" or "Profile"))) ||
                space.Elements().Any(e => e.Name != space.Name.Namespace + "Palette") ||
                space.Elements().Skip(1).Any())
                throw new NotSupportedException("ColorSpace profile, bit depth or extension.");
        }
        var count = type switch { "GRAY" => 1, "RGB" => 3, "CMYK" => 4,
            _ => throw new NotSupportedException($"ColorSpace type '{type}'.") };
        var palette = space?.Element(space.Name.Namespace + "Palette");
        XElement[]? entries = null;
        if (palette is not null)
        {
            if (palette.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                palette.Elements().Any(e => e.Name != space!.Name.Namespace + "CV") ||
                palette.Nodes().Any(n => n is not XElement and not XText and not XComment ||
                    n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Unsupported palette.");
            entries = palette.Elements().ToArray();
            if (entries.Length == 0 || entries.Length > 256) throw new NotSupportedException("Palette size outside supported range.");
            foreach (var entry in entries)
            {
                if (entry.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                    entry.Nodes().Any(n => n is not XText and not XComment))
                    throw new NotSupportedException("Palette entry has unsupported content.");
                _ = Components(entry.Value, count);
            }
        }
        int[] components;
        if (color.Attribute("Value") is { } value)
        {
            var actualCount = space is null && value.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length == 1 ? 1 : count;
            components = Components(value.Value, actualCount);
        }
        else if (color.Attribute("Index") is { } index)
        {
            if (entries is null)
                throw new NotSupportedException("Missing or unsupported palette.");
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
    private static bool HasUnsupportedNodeContent(XElement element) => element.Nodes().Any(node => node switch
    {
        XElement => false,
        XComment => false,
        XText text => !string.IsNullOrWhiteSpace(text.Value),
        _ => true
    });
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
