using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ofdrw.Net.Layout.Text.Paragraphs;

/// <summary>How a requested bold/italic style was satisfied.</summary>
[Flags]
public enum TextSyntheticStyle
{
    /// <summary>The actual face provides the requested style; no synthetic emboldening or slant.</summary>
    None = 0,
    /// <summary>Weight is simulated with a glyph-outline stroke in the fill color
    /// (<see cref="TextResolvedStyle.SyntheticBoldStrokeMillimeters"/>).</summary>
    Bold = 1,
    /// <summary>Slant is simulated by shearing glyphs about their baseline
    /// (<see cref="TextResolvedStyle.ObliqueSkew"/>).</summary>
    Oblique = 2,
}

/// <summary>Immutable resolved style snapshot carried by every glyph run.</summary>
/// <remarks>Writers consume these values directly. They never re-query the request style,
/// reselect fonts, or reshape.</remarks>
public sealed class TextResolvedStyle
{
    /// <summary>Horizontal shear per unit of height above the baseline used for synthetic oblique.</summary>
    public const double SyntheticObliqueSkew = 0.2;

    internal TextResolvedStyle(TextStyle requested, TextFontFamily family, int fallbackIndex, TextFontFace face,
        TextSyntheticStyle synthetic = TextSyntheticStyle.None, double boldStroke = 0)
    {
        Requested = requested; FamilyName = family.Name; FallbackIndex = fallbackIndex;
        FontSizeMillimeters = requested.FontSizeMillimeters;
        RequestedBold = requested.Bold; RequestedItalic = requested.Italic;
        FaceBold = face.Bold; FaceItalic = face.Italic; Language = requested.Language;
        Synthetic = synthetic;
        SyntheticBoldStrokeMillimeters = (synthetic & TextSyntheticStyle.Bold) != 0 ? boldStroke : 0;
        ObliqueSkew = (synthetic & TextSyntheticStyle.Oblique) != 0 ? SyntheticObliqueSkew : 0;
        var paint = requested.Paint;
        // Synthetic bold paints the outline with the fill color; it never overrides a requested stroke.
        Paint = SyntheticBoldStrokeMillimeters > 0
            ? new TextPaint(paint.FillColor, paint.FillColor, SyntheticBoldStrokeMillimeters)
            : paint;
        Decorations = requested.Decorations;
        DecorationColor = requested.Decorations == TextDecorationLines.None ? null :
            requested.DecorationColor ?? paint.FillColor ?? paint.StrokeColor;
        HighlightColor = requested.HighlightColor;
        LetterSpacingMillimeters = requested.LetterSpacingMillimeters;
    }
    /// <summary>The immutable requested style (a source identity, not a paint authority).</summary>
    public TextStyle Requested { get; }
    /// <summary>Name of the family that supplied the face.</summary>
    public string FamilyName { get; }
    /// <summary>Zero-based index of that family in the style's fallback chain.</summary>
    public int FallbackIndex { get; }
    /// <summary>Resolved fill/stroke paint, including a synthetic bold stroke.</summary>
    public TextPaint Paint { get; }
    /// <summary>Resolved em size in millimeters.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Requested bold.</summary>
    public bool RequestedBold { get; }
    /// <summary>Requested italic.</summary>
    public bool RequestedItalic { get; }
    /// <summary>Bold declared by the actual face.</summary>
    public bool FaceBold { get; }
    /// <summary>Italic declared by the actual face.</summary>
    public bool FaceItalic { get; }
    /// <summary>Synthetic styling applied to satisfy the request.</summary>
    public TextSyntheticStyle Synthetic { get; }
    /// <summary>Width of the synthetic bold outline stroke in millimeters; zero without synthetic bold.</summary>
    public double SyntheticBoldStrokeMillimeters { get; }
    /// <summary>Horizontal shear (x += skew × height above the baseline); zero without synthetic oblique.</summary>
    public double ObliqueSkew { get; }
    /// <summary>Resolved BCP 47 language.</summary>
    public string Language { get; }
    /// <summary>Decoration lines drawn for this run.</summary>
    public TextDecorationLines Decorations { get; }
    /// <summary>Resolved decoration color, or null without decorations.</summary>
    public TextColor? DecorationColor { get; }
    /// <summary>Highlight color behind the run, or null.</summary>
    public TextColor? HighlightColor { get; }
    /// <summary>Extra advance after each grapheme cluster, in millimeters.</summary>
    public double LetterSpacingMillimeters { get; }
}

/// <summary>Kind of a run decoration rectangle.</summary>
public enum ParagraphLayoutDecorationKind
{
    /// <summary>A filled background rectangle drawn before the run's glyphs.</summary>
    Highlight,
    /// <summary>A filled underline rectangle drawn after the run's glyphs.</summary>
    Underline,
    /// <summary>A filled strikethrough rectangle drawn after the run's glyphs.</summary>
    Strikethrough,
}

/// <summary>A final, filled decoration rectangle in page-content coordinates (millimeters, Y down).</summary>
public sealed class ParagraphLayoutDecoration
{
    internal ParagraphLayoutDecoration(ParagraphLayoutDecorationKind kind, TextLayoutBounds bounds, TextColor color)
    { Kind = kind; BoundsMillimeters = bounds; Color = color; }
    /// <summary>Highlight, underline or strikethrough.</summary>
    public ParagraphLayoutDecorationKind Kind { get; }
    /// <summary>The filled rectangle. Hanging trailing spaces of a line are excluded.</summary>
    public TextLayoutBounds BoundsMillimeters { get; }
    /// <summary>Fill color of the rectangle.</summary>
    public TextColor Color { get; }
}

/// <summary>One final glyph occurrence in page-content coordinates (millimeters, Y down).</summary>
public sealed class ParagraphLayoutGlyph
{
    internal ParagraphLayoutGlyph(uint glyphId, int clusterStart, int clusterLength, double x, double y,
        double advanceX, double advanceY, double justification, TextLayoutBounds? ink)
    {
        GlyphId = glyphId; ClusterStart = clusterStart; ClusterLength = clusterLength;
        XMillimeters = x; YMillimeters = y; AdvanceXMillimeters = advanceX; AdvanceYMillimeters = advanceY;
        JustificationMillimeters = justification; InkBoundsMillimeters = ink;
    }
    /// <summary>Original GID in the run's face.</summary>
    public uint GlyphId { get; }
    /// <summary>Absolute UTF-16 start of the source cluster. Several glyphs may share a cluster.</summary>
    public int ClusterStart { get; }
    /// <summary>UTF-16 length of the source cluster.</summary>
    public int ClusterLength { get; }
    /// <summary>Glyph origin X (pen plus offset).</summary>
    public double XMillimeters { get; }
    /// <summary>Glyph origin Y on the baseline (pen plus offset).</summary>
    public double YMillimeters { get; }
    /// <summary>Final horizontal advance, including justification.</summary>
    public double AdvanceXMillimeters { get; }
    /// <summary>Final vertical advance (Y down); zero for horizontal text.</summary>
    public double AdvanceYMillimeters { get; }
    /// <summary>Extra advance added by justification (already included in AdvanceX).</summary>
    public double JustificationMillimeters { get; }
    /// <summary>Ink rectangle, or null for a glyph without outline. It includes a synthetic bold
    /// stroke and the sheared extent of synthetic oblique.</summary>
    public TextLayoutBounds? InkBoundsMillimeters { get; }
}

/// <summary>A single-face, single-direction, single-style run of final glyphs in visual order.</summary>
public sealed class ParagraphLayoutRun
{
    internal ParagraphLayoutRun(int sourceStart, string text, TextShapedRun shapedRun, int bidiLevel,
        TextResolvedStyle style, double originX, double baselineY, double advance,
        IReadOnlyList<ParagraphLayoutGlyph> glyphs, TextLayoutBounds? ink,
        IReadOnlyList<ParagraphLayoutDecoration>? decorations = null, bool isTab = false)
    {
        SourceStart = sourceStart; Text = text; ShapedRun = shapedRun; BidiLevel = bidiLevel; Style = style;
        OriginXMillimeters = originX; BaselineYMillimeters = baselineY; AdvanceMillimeters = advance;
        Glyphs = new ReadOnlyCollection<ParagraphLayoutGlyph>(new List<ParagraphLayoutGlyph>(glyphs));
        InkBoundsMillimeters = ink;
        Decorations = new ReadOnlyCollection<ParagraphLayoutDecoration>(
            new List<ParagraphLayoutDecoration>(decorations ?? Array.Empty<ParagraphLayoutDecoration>()));
        IsTab = isTab;
    }
    /// <summary>Highlight, underline and strikethrough rectangles of this run, in drawing order.</summary>
    public IReadOnlyList<ParagraphLayoutDecoration> Decorations { get; }
    /// <summary>Whether this run is a positioned tab (one blank glyph whose advance reaches the tab stop).
    /// Its ShapedRun was produced by the layout engine, not by the shaper.</summary>
    public bool IsTab { get; }
    /// <summary>Absolute UTF-16 start of this run's logical source.</summary>
    public int SourceStart { get; }
    /// <summary>UTF-16 length of the logical source.</summary>
    public int SourceLength => Text.Length;
    /// <summary>Unmodified logical source of the run.</summary>
    public string Text { get; }
    /// <summary>The exact shaped result whose glyphs are positioned here.</summary>
    public TextShapedRun ShapedRun { get; }
    /// <summary>Selected immutable face (byte identity <see cref="TextFontFace.FaceSha256"/>).</summary>
    public TextFontFace Face => ShapedRun.Request.Face;
    /// <summary>ISO 15924 script passed to shaping.</summary>
    public string Script => ShapedRun.Request.Script;
    /// <summary>Shaping direction.</summary>
    public TextDirection Direction => ShapedRun.Request.Direction;
    /// <summary>Resolved UAX #9 embedding level after L1.</summary>
    public int BidiLevel { get; }
    /// <summary>Resolved style snapshot.</summary>
    public TextResolvedStyle Style { get; }
    /// <summary>Em size in millimeters.</summary>
    public double FontSizeMillimeters => Style.FontSizeMillimeters;
    /// <summary>Pen X at the run's visual left edge.</summary>
    public double OriginXMillimeters { get; }
    /// <summary>Baseline Y.</summary>
    public double BaselineYMillimeters { get; }
    /// <summary>Total horizontal advance including justification.</summary>
    public double AdvanceMillimeters { get; }
    /// <summary>Final glyphs in visual left-to-right order.</summary>
    public IReadOnlyList<ParagraphLayoutGlyph> Glyphs { get; }
    /// <summary>Union of glyph ink, or null when no glyph has ink.</summary>
    public TextLayoutBounds? InkBoundsMillimeters { get; }
}

/// <summary>One final line. Coordinates are relative to its page's content box.</summary>
public sealed class ParagraphLayoutLine
{
    internal ParagraphLayoutLine(int sourceStart, int sourceLength, int terminatorLength, int paragraphIndex,
        int pageIndex, int baseLevel, bool endsParagraph, bool hardBreak, double top, double baseline, double height,
        double ascent, double descent, double contentLeft, double contentWidth,
        IReadOnlyList<ParagraphLayoutRun> runs, TextLayoutBounds? ink)
    {
        SourceStart = sourceStart; SourceLength = sourceLength; TerminatorLength = terminatorLength;
        ParagraphIndex = paragraphIndex; PageIndex = pageIndex; BaseLevel = baseLevel;
        EndsParagraph = endsParagraph; EndsWithHardBreak = hardBreak; TopMillimeters = top;
        BaselineMillimeters = baseline; HeightMillimeters = height; AscentMillimeters = ascent;
        DescentMillimeters = descent; ContentLeftMillimeters = contentLeft; ContentWidthMillimeters = contentWidth;
        Runs = new ReadOnlyCollection<ParagraphLayoutRun>(new List<ParagraphLayoutRun>(runs));
        InkBoundsMillimeters = ink;
    }
    /// <summary>Absolute UTF-16 start of the line.</summary>
    public int SourceStart { get; }
    /// <summary>UTF-16 length including trailing spaces and the terminator.</summary>
    public int SourceLength { get; }
    /// <summary>UTF-16 length of a hard line/paragraph terminator ending this line (not drawn).</summary>
    public int TerminatorLength { get; }
    /// <summary>Zero-based paragraph index.</summary>
    public int ParagraphIndex { get; }
    /// <summary>Zero-based page index.</summary>
    public int PageIndex { get; }
    /// <summary>Paragraph embedding level (0 LTR, 1 RTL).</summary>
    public int BaseLevel { get; }
    /// <summary>Whether this is the last line of its paragraph.</summary>
    public bool EndsParagraph { get; }
    /// <summary>Whether the line ends at a mandatory break (terminator or end of text).</summary>
    public bool EndsWithHardBreak { get; }
    /// <summary>Top of the line box.</summary>
    public double TopMillimeters { get; }
    /// <summary>Common baseline of every run.</summary>
    public double BaselineMillimeters { get; }
    /// <summary>Line box height.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Maximum run ascent above the baseline.</summary>
    public double AscentMillimeters { get; }
    /// <summary>Maximum run descent below the baseline.</summary>
    public double DescentMillimeters { get; }
    /// <summary>Left edge of aligned content, excluding hanging trailing spaces.</summary>
    public double ContentLeftMillimeters { get; }
    /// <summary>Content advance excluding hanging trailing spaces.</summary>
    public double ContentWidthMillimeters { get; }
    /// <summary>Runs in visual left-to-right order.</summary>
    public IReadOnlyList<ParagraphLayoutRun> Runs { get; }
    /// <summary>Union of glyph ink, or null.</summary>
    public TextLayoutBounds? InkBoundsMillimeters { get; }
}

/// <summary>One page of final lines.</summary>
public sealed class ParagraphLayoutPage
{
    internal ParagraphLayoutPage(int index, IReadOnlyList<ParagraphLayoutLine> lines, double usedHeight)
    {
        Index = index; Lines = new ReadOnlyCollection<ParagraphLayoutLine>(new List<ParagraphLayoutLine>(lines));
        UsedHeightMillimeters = usedHeight;
    }
    /// <summary>Zero-based page index.</summary>
    public int Index { get; }
    /// <summary>Lines on the page in reading order.</summary>
    public IReadOnlyList<ParagraphLayoutLine> Lines { get; }
    /// <summary>Bottom of the last line box.</summary>
    public double UsedHeightMillimeters { get; }
}

/// <summary>The complete, immutable final layout. No partial result is ever returned.</summary>
public sealed class ParagraphLayoutResult
{
    internal ParagraphLayoutResult(ParagraphLayoutRequest request, IReadOnlyList<ParagraphLayoutPage> pages,
        int shapeCalls, string engineVersion)
    {
        Request = request; Pages = new ReadOnlyCollection<ParagraphLayoutPage>(new List<ParagraphLayoutPage>(pages));
        ShapeCalls = shapeCalls; EngineVersion = engineVersion;
        var lines = new List<ParagraphLayoutLine>();
        foreach (var page in pages) lines.AddRange(page.Lines);
        Lines = new ReadOnlyCollection<ParagraphLayoutLine>(lines);
    }
    /// <summary>The request, including its immutable styles.</summary>
    public ParagraphLayoutRequest Request { get; }
    /// <summary>Complete logical source.</summary>
    public string Text => Request.Text;
    /// <summary>Frame width.</summary>
    public double WidthMillimeters => Request.Options.WidthMillimeters;
    /// <summary>Pages in order.</summary>
    public IReadOnlyList<ParagraphLayoutPage> Pages { get; }
    /// <summary>Every line in reading order.</summary>
    public IReadOnlyList<ParagraphLayoutLine> Lines { get; }
    /// <summary>Unicode data version of the bidi, line break and grapheme algorithms.</summary>
    public string UnicodeVersion => Unicode.UnicodeData.Version;
    /// <summary>Shaper identity reported by the run shaper (empty when nothing was shaped).</summary>
    public string EngineVersion { get; }
    /// <summary>Number of run shaping calls performed.</summary>
    public int ShapeCalls { get; }
}
