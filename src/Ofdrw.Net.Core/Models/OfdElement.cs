namespace Ofdrw.Net.Core.Models;

public enum OfdElementType
{
    Text,
    Image,
    Path,
    Raw
}

public abstract class OfdElement
{
    public OfdElementType Type { get; protected set; }

    public string? ObjectId { get; set; }

    public string? LayerId { get; set; }

    public string LayerType { get; set; } = "Body";

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }
}

public sealed class OfdTextElement : OfdElement
{
    public OfdTextElement()
    {
        Type = OfdElementType.Text;
    }

    public string Text { get; set; } = string.Empty;

    public string FontName { get; set; } = "SimSun";

    public double FontSizeMillimeters { get; set; } = 4d;

    public string? FontResourceId { get; set; }

    public double[]? Transform { get; set; }

    public OfdColor FillColor { get; set; } = OfdColor.Black;

    public List<OfdTextRun> Runs { get; } = [];

    /// <summary>
    /// Complete source XML retained to preserve glyph positioning, transforms,
    /// colors and vendor attributes that are not strongly typed yet.
    /// </summary>
    public string? SourceXml { get; set; }
}

public sealed class OfdTextRun
{
    public string Text { get; set; } = string.Empty;

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public string? DeltaX { get; set; }

    public string? DeltaY { get; set; }
}

public sealed class OfdImageElement : OfdElement
{
    public OfdImageElement()
    {
        Type = OfdElementType.Image;
    }

    public string ResourceId { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;

    public string MediaType { get; set; } = "image/png";

    public byte[] Data { get; set; } = [];

    /// <summary>
    /// Complete source XML retained to preserve transforms, clipping, alpha and
    /// vendor attributes that are not strongly typed yet.
    /// </summary>
    public string? SourceXml { get; set; }
}

/// <summary>
/// An OFD path object. Coordinates in <see cref="AbbreviatedData"/> are local to
/// the object boundary and are transformed by <see cref="Transform"/> when present.
/// </summary>
public sealed class OfdPathElement : OfdElement
{
    public OfdPathElement()
    {
        Type = OfdElementType.Path;
    }

    /// <summary>
    /// Gets or sets the OFD abbreviated path command stream.
    /// </summary>
    public string AbbreviatedData { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the six-value OFD CTM matrix (a, b, c, d, e, f).
    /// </summary>
    public double[]? Transform { get; set; }

    public double LineWidthMillimeters { get; set; } = 0.353d;

    public bool Stroke { get; set; } = true;

    public bool Fill { get; set; }

    public OfdColor StrokeColor { get; set; } = OfdColor.Black;

    public OfdColor? FillColor { get; set; }

    /// <summary>
    /// Retains the complete source element so unsupported style attributes and
    /// extension children survive a read-write round trip.
    /// </summary>
    public string? SourceXml { get; set; }
}

public sealed class OfdColor
{
    public static OfdColor Black { get; } = new(0, 0, 0);

    public OfdColor(int red, int green, int blue, int alpha = 255)
    {
        Red = Clamp(red);
        Green = Clamp(green);
        Blue = Clamp(blue);
        Alpha = Clamp(alpha);
    }

    public int Red { get; }

    public int Green { get; }

    public int Blue { get; }

    public int Alpha { get; }

    private static int Clamp(int value)
    {
        return value < 0 ? 0 : value > 255 ? 255 : value;
    }
}

/// <summary>
/// A page object that is not yet represented by a strongly typed Ofdrw.Net model.
/// The original XML is retained so reading and writing a package does not silently
/// discard paths, composites, extensions, or future OFD object types.
/// </summary>
public sealed class OfdRawElement : OfdElement
{
    public OfdRawElement()
    {
        Type = OfdElementType.Raw;
    }

    public string LocalName { get; set; } = string.Empty;

    public string Xml { get; set; } = string.Empty;
}
