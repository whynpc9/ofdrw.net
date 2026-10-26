using System;
using System.Collections.Generic;
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

/// <summary>OFD line ending style.</summary>
public enum OfdLineCap
{
    /// <summary>Stop at the segment endpoint.</summary>
    Butt,
    /// <summary>Extend by a semicircle.</summary>
    Round,
    /// <summary>Extend by a half-width square.</summary>
    Square
}

/// <summary>OFD join style between stroked segments.</summary>
public enum OfdLineJoin
{
    /// <summary>Extend adjacent edges to their intersection.</summary>
    Miter,
    /// <summary>Join with a circular arc.</summary>
    Round,
    /// <summary>Join with a straight bevel.</summary>
    Bevel
}

/// <summary>An immutable pen; width and dash lengths are user-space millimeters.</summary>
public sealed class OfdPen
{
    /// <summary>Creates the original solid butt/miter pen. This constructor retains the existing binary signature.</summary>
    public OfdPen(OfdColor color, double widthMillimeters = 0.353)
        : this(color, widthMillimeters, OfdLineCap.Butt) { }
    /// <summary>Creates a pen with a copied dash sequence and default butt/miter geometry.</summary>
    public OfdPen(OfdColor color, double widthMillimeters, IReadOnlyList<double> dashPattern, double dashOffset = 0)
        : this(color, widthMillimeters, OfdLineCap.Butt, dashPattern: dashPattern, dashOffset: dashOffset) { }
    /// <summary>Creates a pen with positive writer-representable width and optional line geometry.</summary>
    public OfdPen(OfdColor color, double widthMillimeters, OfdLineCap cap, OfdLineJoin join = OfdLineJoin.Miter,
        double miterLimit = 10, IReadOnlyList<double>? dashPattern = null, double dashOffset = 0)
    {
        Color = color ?? throw new ArgumentNullException(nameof(color));
        GraphicsValidation.Positive(widthMillimeters);
        if (!Enum.IsDefined(typeof(OfdLineCap), cap)) throw new ArgumentOutOfRangeException(nameof(cap));
        if (!Enum.IsDefined(typeof(OfdLineJoin), join)) throw new ArgumentOutOfRangeException(nameof(join));
        GraphicsValidation.Finite(miterLimit, dashOffset);
        if (miterLimit < 1 || miterLimit > 1000 || miterLimit != Math.Truncate(miterLimit))
            throw new ArgumentOutOfRangeException(nameof(miterLimit), "PDF requires an integer miter limit from 1 to 1000.");
        if (cap != OfdLineCap.Butt && miterLimit != 10)
            throw new NotSupportedException("The PDF backend cannot combine a nondefault miter limit with round or square caps.");
        if (dashOffset < 0 || GraphicsValidation.WriterValue(dashOffset) < 0)
            throw new ArgumentOutOfRangeException(nameof(dashOffset));
        if (dashPattern is not null)
        {
            if (dashPattern.Count == 0 || dashPattern.Count > 128)
                throw new ArgumentOutOfRangeException(nameof(dashPattern), "Dash pattern must have 1–128 entries.");
            var values = new double[dashPattern.Count];
            double sum = 0;
            for (var index = 0; index < values.Length; index++)
            {
                GraphicsValidation.Positive(dashPattern[index]);
                values[index] = dashPattern[index];
                if (double.IsInfinity(values[index] / widthMillimeters) || values[index] / widthMillimeters <= 0)
                    throw new ArgumentOutOfRangeException(nameof(dashPattern), "Dash lengths cannot be represented by the PDF pen width.");
                sum += values[index];
            }
            GraphicsValidation.Finite(sum, sum * 72d / 25.4d, dashOffset / widthMillimeters);
            DashPattern = Array.AsReadOnly(values);
        }
        else if (dashOffset != 0) throw new ArgumentException("Dash offset requires a dash pattern.", nameof(dashOffset));
        WidthMillimeters = widthMillimeters;
        Cap = cap; Join = join; MiterLimit = miterLimit; DashOffset = dashOffset;
    }
    /// <summary>Stroke color.</summary>
    public OfdColor Color { get; }
    /// <summary>Stroke width in millimeters before the graphics transform.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Line cap.</summary>
    public OfdLineCap Cap { get; }
    /// <summary>Line join.</summary>
    public OfdLineJoin Join { get; }
    /// <summary>Miter limit.</summary>
    public double MiterLimit { get; }
    /// <summary>Read-only copied dash sequence, or null for a solid line.</summary>
    public IReadOnlyList<double>? DashPattern { get; }
    /// <summary>Dash phase in millimeters.</summary>
    public double DashOffset { get; }
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
