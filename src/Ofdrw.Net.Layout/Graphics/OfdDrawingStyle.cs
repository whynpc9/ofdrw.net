using System;
using System.Xml;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Graphics;

/// <summary>An immutable solid RGB/alpha brush. Gradients and shaders are not supported.</summary>
public sealed class OfdBrush
{
    /// <summary>Creates a solid brush.</summary>
    public OfdBrush(OfdColor color) => Color = color ?? throw new ArgumentNullException(nameof(color));
    /// <summary>Solid color.</summary>
    public OfdColor Color { get; }
}

/// <summary>An immutable solid pen with butt caps and miter joins; width is in user-space millimeters.</summary>
public sealed class OfdPen
{
    /// <summary>Creates a pen with positive width representable at OFD writer precision.</summary>
    public OfdPen(OfdColor color, double widthMillimeters = 0.353)
    {
        Color = color ?? throw new ArgumentNullException(nameof(color));
        GraphicsValidation.Positive(widthMillimeters);
        WidthMillimeters = widthMillimeters;
    }
    /// <summary>Stroke color.</summary>
    public OfdColor Color { get; }
    /// <summary>Stroke width in millimeters before the graphics transform.</summary>
    public double WidthMillimeters { get; }
}

/// <summary>Immutable text style and optional package-local font identity. Does not resolve, embed or subset fonts.</summary>
public sealed class OfdFont
{
    /// <summary>Describes a font. Resource ID, when provided, must identify exactly one existing resource in the destination package.</summary>
    public OfdFont(string name, double sizeMillimeters, int weight = 400, bool italic = false, string? resourceId = null)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Font name is required.", nameof(name));
        XmlConvert.VerifyXmlChars(name);
        if (resourceId is not null && string.IsNullOrWhiteSpace(resourceId)) throw new ArgumentException("Font resource ID cannot be empty.", nameof(resourceId));
        GraphicsValidation.Positive(sizeMillimeters);
        if (weight < 100 || weight > 900 || weight % 100 != 0) throw new ArgumentOutOfRangeException(nameof(weight));
        Name = name; SizeMillimeters = sizeMillimeters; Weight = weight; Italic = italic; ResourceId = resourceId;
    }
    /// <summary>Font family name, resolved by the existing writer/exporter when no ID is specified.</summary>
    public string Name { get; }
    /// <summary>Em size in millimeters.</summary>
    public double SizeMillimeters { get; }
    /// <summary>Per-object weight from 100 to 900 in steps of 100.</summary>
    public int Weight { get; }
    /// <summary>Per-object italic style.</summary>
    public bool Italic { get; }
    /// <summary>Optional identity in destination package.Fonts; never a source-package identity.</summary>
    public string? ResourceId { get; }
}
