namespace Ofdrw.Net.Core.Models;

public enum OfdElementType
{
    Text,
    Image,
    Path,
    Raw,
    /// <summary>Placement of a shared composite vector resource.</summary>
    Composite
}

public abstract class OfdElement
{
    // Reader-resolved render view of an imported source fragment. The original
    // SourceXml/Raw XML remains authoritative for editing and package writing.
    internal string? RenderSourceXml { get; set; }
    internal string? RenderSourceAuthority { get; set; }
    internal string? RenderPaintError { get; set; }
    /// <summary>Invalidates Reader-resolved paint when a public source envelope is cleared.</summary>
    protected void InvalidateResolvedSourceIfCleared(string? sourceXml)
    {
        if (!string.IsNullOrWhiteSpace(sourceXml)) return;
        RenderSourceXml = null;
        RenderSourceAuthority = null;
    }
    internal OfdClipOwnership ClipOwnership { get; } = new();
    /// <summary>Typed path clips. Null preserves an unclaimed source; an empty collection clears known clips.</summary>
    public OfdClipCollection? Clips { get => ClipOwnership.Value; set => ClipOwnership.Set(value); }
    public OfdElementType Type { get; protected set; }

    public string? ObjectId { get; set; }

    public string? LayerId { get; set; }

    public string LayerType { get; set; } = "Body";

    /// <summary>OFD Clips XML in object coordinates, applied before the object's CTM. Text/path null preserves source XML; empty clears it. Image ClipsXml remains authoritative.</summary>
    public string? ClippingXml { get; set; }

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }
}

public sealed class OfdTextElement : OfdElement
{
    // Opt-in from freshly laid out native DOCX text only. Imported source XML
    // owns its painting attributes and never acquires this request implicitly.
    internal bool EmitSyntheticBoldStroke { get; set; }
    internal double? SyntheticBoldStrokeWidthMillimeters { get; set; }
    // Reader's PageBlock projection materializes Table 46 inherited starts;
    // ordinary imported typed runs retain their historical default values.
    internal bool SourceInheritedTextCodeAxes { get; set; }

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

    /// <summary>
    /// CT_Text <c>Weight</c> (100–900, default 400). Readers apply bold from this
    /// attribute; the <c>Bold</c> flag on the font resource only describes the font
    /// file and is ignored by most viewers for name-only fonts.
    /// </summary>
    public int Weight { get; set; } = DefaultWeight;

    /// <summary>CT_Text <c>Italic</c> (default false), applied by readers per text object.</summary>
    public bool Italic { get; set; }

    public const int DefaultWeight = 400;

    public const int BoldWeight = 700;

    public List<OfdTextRun> Runs { get; } = [];

    private string? sourceXml;
    /// <summary>
    /// Complete source XML retained to preserve glyph positioning, transforms,
    /// colors and vendor attributes that are not strongly typed yet.
    /// </summary>
    public string? SourceXml
    {
        get => sourceXml;
        set { InvalidateResolvedSourceIfCleared(value); sourceXml = value; }
    }
    internal void ClearSourceXmlForProjection() => sourceXml = null;
}

public sealed class OfdTextRun
{
    public string Text { get; set; } = string.Empty;

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public string? DeltaX { get; set; }

    public string? DeltaY { get; set; }

    /// <summary>Optional producer-owned glyph IDs, logical clusters and exact origins for this run.</summary>
    public OfdPositionedGlyphRun? PositionedGlyphs { get; set; }
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

    internal bool SourceMediaAmbiguous { get; set; }

    // Reader-resolved rendering view. The original resource reference and bytes
    // remain in the imported package; this data is never written in their place.
    internal string? MaskResourceId { get; set; }
    internal string? MaskMediaType { get; set; }
    internal byte[]? MaskData { get; set; }
    internal bool MaskSourceAmbiguous { get; set; }

    /// <summary>Maps the image's unit square into object coordinates. Null uses the boundary width/height.</summary>
    public double[]? Transform { get; set; }

    /// <summary>Object opacity, from zero (transparent) to 255 (opaque).</summary>
    public int Alpha { get; set; } = 255;

    /// <summary>Optional OFD Clips XML in object coordinates.</summary>
    public string? ClipsXml { get => ClippingXml; set => ClippingXml = value; }

    private string? sourceXml;
    /// <summary>
    /// Complete source XML retained to preserve transforms, clipping, alpha and
    /// vendor attributes that are not strongly typed yet.
    /// </summary>
    public string? SourceXml
    {
        get => sourceXml;
        set { InvalidateResolvedSourceIfCleared(value); sourceXml = value; }
    }
    internal void ClearSourceXmlForProjection() => sourceXml = null;
}

/// <summary>
/// An OFD path object. Coordinates in <see cref="AbbreviatedData"/> are local to
/// the object boundary and are transformed by <see cref="Transform"/> when present.
/// </summary>
public sealed class OfdPathElement : OfdElement
{
    // Set only on a transient projection of a reader-owned link annotation.
    // Ordinary page paths and public model construction cannot authorize AttachID.
    internal string? ReaderAttachmentActionId { get; set; }

    /// <summary>Null preserves a source Rule; newly constructed paths default to NonZero.</summary>
    public OfdFillRule? FillRule { get; set; }
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

    private string? sourceXml;
    /// <summary>
    /// Retains the complete source element so unsupported style attributes and
    /// extension children survive a read-write round trip.
    /// </summary>
    public string? SourceXml
    {
        get => sourceXml;
        set { InvalidateResolvedSourceIfCleared(value); sourceXml = value; }
    }
    internal void ClearSourceXmlForProjection() => sourceXml = null;
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
    // Reader-owned, read-only rendering view. Xml remains the authenticated
    // page root for package writeback; these children are never serialized.
    internal List<OfdElement>? PageBlockChildren { get; set; }
    internal string? PageBlockSourceXml { get; set; }
    // Set only by the Reader for annotation metadata and index envelopes outside the page drawing tree.
    // A caller changing Xml invalidates this render-only provenance.
    internal string? ReaderAnnotationMetadataXml { get; set; }
    public OfdRawElement()
    {
        Type = OfdElementType.Raw;
    }

    public string LocalName { get; set; } = string.Empty;

    public string Xml { get; set; } = string.Empty;
}
