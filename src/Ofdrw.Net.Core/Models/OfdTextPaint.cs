using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal sealed class OfdTextPaint
{
    internal bool Fill { get; }
    internal bool Stroke { get; }
    internal double HScale { get; }
    internal double LineWidthMillimeters { get; }
    internal OfdColor FillColor { get; }
    internal OfdColor StrokeColor { get; }

    private OfdTextPaint(OfdTextElement text)
    {
        var source = text.RenderSourceXml ?? text.SourceXml;
        var root = string.IsNullOrWhiteSpace(source) ? null : OfdXmlEnvelope.Parse(source!);
        Fill = Boolean(root?.Attribute("Fill")?.Value, true);
        Stroke = Boolean(root?.Attribute("Stroke")?.Value, false);
        HScale = Number(root?.Attribute("HScale")?.Value, 1);
        LineWidthMillimeters = Number(root?.Attribute("LineWidth")?.Value, 0.353);
        if (Stroke && (root?.Attribute("Cap") is not null || root?.Attribute("Join") is not null ||
            root?.Attribute("MiterLimit") is not null || root?.Attribute("DashPattern") is not null ||
            root?.Attribute("DashOffset") is not null))
            throw new NotSupportedException("OFD text stroke parameters are outside the render profile.");
        if (HScale <= 0 || HScale > 64 || LineWidthMillimeters < 0 || LineWidthMillimeters > 100)
            throw new NotSupportedException("OFD text scale or stroke width is outside the render profile.");
        FillColor = root is null ? text.FillColor : Color(root.Element(root.Name.Namespace + "FillColor"), text.FillColor);
        StrokeColor = root is null ? OfdColor.Black :
            Color(root.Element(root.Name.Namespace + "StrokeColor"), OfdColor.Black);
    }

    internal static OfdTextPaint Read(OfdTextElement text) => new(text);

    private static bool Boolean(string? value, bool fallback) => value is null ? fallback : value switch
    {
        "true" or "1" => true,
        "false" or "0" => false,
        _ => throw new InvalidDataException("Invalid OFD text paint boolean.")
    };

    private static double Number(string? value, double fallback)
    {
        if (value is null) return fallback;
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result) ||
            double.IsNaN(result) || double.IsInfinity(result))
            throw new InvalidDataException("Invalid OFD text paint number.");
        return result;
    }

    private static OfdColor Color(XElement? node, OfdColor fallback)
    {
        if (node is null) return fallback;
        // Annotation appearances can reach this parser without the Reader's
        // resource normalization. Admit only the unreferenced gray shorthand;
        // referenced and extended paints still need explicit resolution.
        if (node.HasElements || node.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("Value" or "Alpha"))))
            throw new InvalidDataException("OFD text paint needs resolved RGB components.");
        var values = node.Attribute("Value")?.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => int.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) && n is >= 0 and <= 255 ? n :
                throw new InvalidDataException("Invalid OFD text color component.")).ToArray();
        if (values is not { Length: 1 or 3 }) throw new InvalidDataException("OFD text paint needs resolved RGB components.");
        var alpha = 255;
        if (node.Attribute("Alpha") is { } attribute &&
            !int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out alpha))
            throw new InvalidDataException("Invalid OFD text paint alpha.");
        if (alpha is < 0 or > 255) throw new InvalidDataException("Invalid OFD text paint alpha.");
        return values.Length == 1 ? new OfdColor(values[0], values[0], values[0], alpha) :
            new OfdColor(values[0], values[1], values[2], alpha);
    }
}
