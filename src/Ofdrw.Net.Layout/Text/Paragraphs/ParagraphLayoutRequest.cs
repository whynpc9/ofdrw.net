using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ofdrw.Net.Layout.Text.Paragraphs;

/// <summary>An immutable sRGB (DeviceRGB) color with 8-bit alpha.</summary>
public sealed class TextColor : IEquatable<TextColor>
{
    /// <summary>Creates a color from 0–255 channel values; alpha 255 is opaque.</summary>
    public TextColor(int red, int green, int blue, int alpha = 255)
    {
        if (red is < 0 or > 255 || green is < 0 or > 255 || blue is < 0 or > 255 || alpha is < 0 or > 255)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Color channels must be within 0–255.");
        Red = red; Green = green; Blue = blue; Alpha = alpha;
    }
    /// <summary>Opaque black.</summary>
    public static TextColor Black { get; } = new(0, 0, 0);
    /// <summary>Red channel, 0–255.</summary>
    public int Red { get; }
    /// <summary>Green channel, 0–255.</summary>
    public int Green { get; }
    /// <summary>Blue channel, 0–255.</summary>
    public int Blue { get; }
    /// <summary>Alpha, 0 (transparent) – 255 (opaque).</summary>
    public int Alpha { get; }
    /// <summary>Color space of the channels; always DeviceRGB in this profile.</summary>
    public string ColorSpace => "DeviceRGB";
    /// <inheritdoc />
    public bool Equals(TextColor? other) => other is not null && Red == other.Red && Green == other.Green &&
        Blue == other.Blue && Alpha == other.Alpha;
    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as TextColor);
    /// <inheritdoc />
    public override int GetHashCode() => (Red << 24) ^ (Green << 16) ^ (Blue << 8) ^ Alpha;
}

/// <summary>Immutable glyph paint: fill and optional stroke, each with its own color and alpha.</summary>
public sealed class TextPaint : IEquatable<TextPaint>
{
    /// <summary>Creates a paint. A null fill disables filling; a null stroke disables stroking.</summary>
    public TextPaint(TextColor? fillColor, TextColor? strokeColor = null, double strokeWidthMillimeters = 0)
    {
        if (fillColor is null && strokeColor is null)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "A text paint needs a fill or a stroke.");
        if (double.IsNaN(strokeWidthMillimeters) || double.IsInfinity(strokeWidthMillimeters) ||
            strokeWidthMillimeters < 0 || strokeWidthMillimeters > 100 ||
            strokeColor is not null && strokeWidthMillimeters <= 0 ||
            strokeColor is null && strokeWidthMillimeters != 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "A stroke needs a positive finite width (at most 100 mm); no stroke needs width zero.");
        FillColor = fillColor; StrokeColor = strokeColor; StrokeWidthMillimeters = strokeWidthMillimeters;
    }
    /// <summary>Opaque black fill without stroke: the first verified writer profile.</summary>
    public static TextPaint OpaqueBlack { get; } = new(TextColor.Black);
    /// <summary>Whether glyphs are filled.</summary>
    public bool Fill => FillColor is not null;
    /// <summary>Fill color, or null when not filled.</summary>
    public TextColor? FillColor { get; }
    /// <summary>Whether glyph outlines are stroked.</summary>
    public bool Stroke => StrokeColor is not null;
    /// <summary>Stroke color, or null when not stroked.</summary>
    public TextColor? StrokeColor { get; }
    /// <summary>Stroke width in millimeters; zero without stroke.</summary>
    public double StrokeWidthMillimeters { get; }
    /// <inheritdoc />
    public bool Equals(TextPaint? other) => other is not null && Equals(FillColor, other.FillColor) &&
        Equals(StrokeColor, other.StrokeColor) && StrokeWidthMillimeters == other.StrokeWidthMillimeters;
    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as TextPaint);
    /// <inheritdoc />
    public override int GetHashCode() => (FillColor?.GetHashCode() ?? 0) ^ ((StrokeColor?.GetHashCode() ?? 0) * 31) ^
        StrokeWidthMillimeters.GetHashCode();
}

/// <summary>A named family of explicitly supplied faces. No host font discovery is performed.</summary>
public sealed class TextFontFamily
{
    /// <summary>Creates a family from up to four faces with distinct (bold, italic) declarations.</summary>
    public TextFontFamily(string name, IReadOnlyList<TextFontFace> faces)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A family name is required.", nameof(name));
        if (faces is null || faces.Count is < 1 or > 4 || faces.Any(face => face is null))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "A family needs one to four faces.");
        if (faces.Select(face => (face.Bold, face.Italic)).Distinct().Count() != faces.Count)
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "Faces in one family must declare distinct bold/italic styles.");
        Name = name;
        Faces = new ReadOnlyCollection<TextFontFace>(faces.ToArray());
    }
    /// <summary>Creates a single-face family.</summary>
    public TextFontFamily(string name, TextFontFace face) : this(name, new[] { face }) { }
    /// <summary>Caller supplied family name.</summary>
    public string Name { get; }
    /// <summary>Immutable faces of this family.</summary>
    public IReadOnlyList<TextFontFace> Faces { get; }
    /// <summary>Returns the face whose declared style exactly matches, or null. No synthesis.</summary>
    public TextFontFace? Select(bool bold, bool italic) =>
        Faces.FirstOrDefault(face => face.Bold == bold && face.Italic == italic);
}

/// <summary>An OpenType feature applied to every run of a style, for example ("liga", 0).</summary>
public sealed class TextFontFeature : IEquatable<TextFontFeature>
{
    /// <summary>Creates a feature setting with a four-character ASCII tag.</summary>
    public TextFontFeature(string tag, uint value)
    {
        _ = new TextFeature(tag, value, 0, 0); // validates the tag
        Tag = tag; Value = value;
    }
    /// <summary>Four-character ASCII OpenType feature tag.</summary>
    public string Tag { get; }
    /// <summary>Feature value, usually 0 (off) or 1 (on).</summary>
    public uint Value { get; }
    /// <inheritdoc />
    public bool Equals(TextFontFeature? other) => other is not null && Tag == other.Tag && Value == other.Value;
    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as TextFontFeature);
    /// <inheritdoc />
    public override int GetHashCode() => Tag.GetHashCode() ^ (int)Value;
}

/// <summary>An immutable requested character style. Layout snapshots it into the result.</summary>
public sealed class TextStyle
{
    /// <summary>Creates a style with an ordered font fallback chain.</summary>
    /// <param name="fontFallback">One to sixteen families, tried in order per grapheme cluster.</param>
    /// <param name="fontSizeMillimeters">Em size in millimeters, in (0, 1000].</param>
    /// <param name="paint">Glyph paint; defaults to opaque black fill.</param>
    /// <param name="bold">Request a face declared bold. No synthetic bold is produced.</param>
    /// <param name="italic">Request a face declared italic. No synthetic italic is produced.</param>
    /// <param name="language">BCP 47 language tag passed to shaping.</param>
    /// <param name="features">OpenType feature settings for every run of this style.</param>
    public TextStyle(IReadOnlyList<TextFontFamily> fontFallback, double fontSizeMillimeters,
        TextPaint? paint = null, bool bold = false, bool italic = false, string language = "und",
        IReadOnlyList<TextFontFeature>? features = null)
    {
        if (fontFallback is null || fontFallback.Count is < 1 or > 16 || fontFallback.Any(family => family is null))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Supply one to sixteen font families.");
        if (double.IsNaN(fontSizeMillimeters) || double.IsInfinity(fontSizeMillimeters) ||
            fontSizeMillimeters <= 0 || fontSizeMillimeters > 1000)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Font size must be within (0, 1000] mm.");
        if (string.IsNullOrWhiteSpace(language) || language.Length > 32)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "A short BCP 47 language tag is required.");
        if (features is not null && (features.Count > 64 || features.Any(feature => feature is null)))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Supply at most 64 feature settings.");
        FontFallback = new ReadOnlyCollection<TextFontFamily>(fontFallback.ToArray());
        FontSizeMillimeters = fontSizeMillimeters;
        Paint = paint ?? TextPaint.OpaqueBlack;
        Bold = bold; Italic = italic; Language = language;
        Features = new ReadOnlyCollection<TextFontFeature>(features?.ToArray() ?? Array.Empty<TextFontFeature>());
    }
    /// <summary>Convenience constructor for one family.</summary>
    public TextStyle(TextFontFamily family, double fontSizeMillimeters, TextPaint? paint = null,
        bool bold = false, bool italic = false, string language = "und")
        : this(new[] { family }, fontSizeMillimeters, paint, bold, italic, language) { }
    /// <summary>Ordered font families; each grapheme cluster uses the first that covers it.</summary>
    public IReadOnlyList<TextFontFamily> FontFallback { get; }
    /// <summary>Em size in millimeters.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Requested glyph paint.</summary>
    public TextPaint Paint { get; }
    /// <summary>Whether a face declared bold is requested.</summary>
    public bool Bold { get; }
    /// <summary>Whether a face declared italic is requested.</summary>
    public bool Italic { get; }
    /// <summary>BCP 47 language tag.</summary>
    public string Language { get; }
    /// <summary>Feature settings applied to each shaped run of this style.</summary>
    public IReadOnlyList<TextFontFeature> Features { get; }

    internal bool ShapesLike(TextStyle other) => ReferenceEquals(this, other) ||
        FontSizeMillimeters == other.FontSizeMillimeters && Language == other.Language &&
        Features.SequenceEqual(other.Features);
}

/// <summary>Applies a style to a half-open UTF-16 interval of the request text.</summary>
public sealed class TextStyleSpan
{
    /// <summary>Creates a style span; spans of one request must tile the text exactly.</summary>
    public TextStyleSpan(int start, int length, TextStyle style)
    {
        if (start < 0 || length < 0 || (long)start + length > int.MaxValue)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid style span interval.");
        Start = start; Length = length; Style = style ?? throw new ArgumentNullException(nameof(style));
    }
    /// <summary>Absolute UTF-16 start.</summary>
    public int Start { get; }
    /// <summary>UTF-16 length.</summary>
    public int Length { get; }
    /// <summary>The requested style.</summary>
    public TextStyle Style { get; }
}

/// <summary>Paragraph base direction (UAX #9 P2/P3 for Auto).</summary>
public enum TextParagraphDirection
{
    /// <summary>First strong character, defaulting to left-to-right.</summary>
    Auto,
    /// <summary>Paragraph embedding level 0.</summary>
    LeftToRight,
    /// <summary>Paragraph embedding level 1.</summary>
    RightToLeft,
}

/// <summary>Horizontal alignment of each line's content within the available width.</summary>
public enum TextAlignment
{
    /// <summary>Left for LTR paragraphs, right for RTL paragraphs.</summary>
    Start,
    /// <summary>Right for LTR paragraphs, left for RTL paragraphs.</summary>
    End,
    /// <summary>Physical left.</summary>
    Left,
    /// <summary>Physical right.</summary>
    Right,
    /// <summary>Centered.</summary>
    Center,
    /// <summary>Start-aligned, with lines that do not end a paragraph stretched to the full width.</summary>
    Justify,
}

/// <summary>How a segment wider than the line, without a legal break, is handled.</summary>
public enum TextOverflowPolicy
{
    /// <summary>Fail with <see cref="TextLayoutError.LayoutOverflow"/>; never break outside UAX #14.</summary>
    Reject,
    /// <summary>Emergency break at extended grapheme cluster boundaries (never inside a cluster).</summary>
    BreakGraphemes,
}

/// <summary>Line height rule.</summary>
public sealed class TextLineHeight
{
    private TextLineHeight(string kind, double value) { Kind = kind; Value = value; }
    /// <summary>Font ascender + descender + line gap of the line's runs.</summary>
    public static TextLineHeight Normal { get; } = new("Normal", 1);
    /// <summary>A multiple (0.5–10) of the normal line height, extra space split above/below.</summary>
    public static TextLineHeight Multiple(double factor) =>
        factor is >= 0.5 and <= 10 ? new TextLineHeight("Multiple", factor) :
        throw new TextLayoutException(TextLayoutError.InvalidRequest, "Line height factor must be within [0.5, 10].");
    /// <summary>An exact line height in millimeters, (0, 1000].</summary>
    public static TextLineHeight Exact(double millimeters) =>
        millimeters > 0 && millimeters <= 1000 ? new TextLineHeight("Exact", millimeters) :
        throw new TextLayoutException(TextLayoutError.InvalidRequest, "Exact line height must be within (0, 1000] mm.");
    /// <summary>"Normal", "Multiple" or "Exact".</summary>
    public string Kind { get; }
    /// <summary>Factor for Multiple, millimeters for Exact, 1 for Normal.</summary>
    public double Value { get; }
}

/// <summary>Frame, paragraph and budget options. All lengths are millimeters.</summary>
public sealed class ParagraphLayoutOptions
{
    /// <summary>Creates options for a frame of the given width.</summary>
    public ParagraphLayoutOptions(double widthMillimeters, double? pageHeightMillimeters = null,
        TextParagraphDirection direction = TextParagraphDirection.Auto,
        TextAlignment alignment = TextAlignment.Start, TextLineHeight? lineHeight = null,
        double paragraphSpacingMillimeters = 0, double firstLineIndentMillimeters = 0,
        TextOverflowPolicy overflow = TextOverflowPolicy.Reject,
        int maximumLines = 10000, int maximumShapeCalls = 100000, int maximumTextLength = 262144)
    {
        if (!Finite(widthMillimeters) || widthMillimeters <= 0 || widthMillimeters > 1_000_000 ||
            pageHeightMillimeters is { } height && (!Finite(height) || height <= 0 || height > 1_000_000) ||
            !Finite(paragraphSpacingMillimeters) || paragraphSpacingMillimeters < 0 || paragraphSpacingMillimeters > 1000 ||
            !Finite(firstLineIndentMillimeters) || firstLineIndentMillimeters < 0 ||
            firstLineIndentMillimeters >= widthMillimeters)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid frame, spacing or indent geometry.");
        if (!Enum.IsDefined(typeof(TextParagraphDirection), direction) || !Enum.IsDefined(typeof(TextAlignment), alignment) ||
            !Enum.IsDefined(typeof(TextOverflowPolicy), overflow))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Unknown direction, alignment or overflow policy.");
        if (maximumLines is < 1 or > 1_000_000 || maximumShapeCalls is < 1 or > 10_000_000 ||
            maximumTextLength is < 1 or > 16_777_216)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid line, shaping-call or text budget.");
        WidthMillimeters = widthMillimeters; PageHeightMillimeters = pageHeightMillimeters;
        Direction = direction; Alignment = alignment; LineHeight = lineHeight ?? TextLineHeight.Normal;
        ParagraphSpacingMillimeters = paragraphSpacingMillimeters; FirstLineIndentMillimeters = firstLineIndentMillimeters;
        Overflow = overflow; MaximumLines = maximumLines; MaximumShapeCalls = maximumShapeCalls;
        MaximumTextLength = maximumTextLength;
    }
    /// <summary>Available line width.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Content height of each page; null places every line on one unbounded page.</summary>
    public double? PageHeightMillimeters { get; }
    /// <summary>Base direction of every paragraph.</summary>
    public TextParagraphDirection Direction { get; }
    /// <summary>Line alignment.</summary>
    public TextAlignment Alignment { get; }
    /// <summary>Line height rule.</summary>
    public TextLineHeight LineHeight { get; }
    /// <summary>Space after each paragraph except the last; dropped at a page break.</summary>
    public double ParagraphSpacingMillimeters { get; }
    /// <summary>Indent of each paragraph's first line on its start side.</summary>
    public double FirstLineIndentMillimeters { get; }
    /// <summary>Handling of segments wider than a line.</summary>
    public TextOverflowPolicy Overflow { get; }
    /// <summary>Maximum number of lines before failing.</summary>
    public int MaximumLines { get; }
    /// <summary>Maximum number of run shaping calls before failing.</summary>
    public int MaximumShapeCalls { get; }
    /// <summary>Maximum UTF-16 text length.</summary>
    public int MaximumTextLength { get; }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>Complete logical text, its style spans and layout options.</summary>
public sealed class ParagraphLayoutRequest
{
    /// <summary>Creates a request whose spans tile the text in order, without gaps or overlap.</summary>
    public ParagraphLayoutRequest(string text, IReadOnlyList<TextStyleSpan> spans, ParagraphLayoutOptions options)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Options = options ?? throw new ArgumentNullException(nameof(options));
        if (spans is null || spans.Any(span => span is null))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Style spans are required.");
        if (text.Length > options.MaximumTextLength)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "The text exceeds MaximumTextLength.");
        var cursor = 0;
        foreach (var span in spans)
        {
            if (span.Start != cursor)
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "Style spans must tile the text in order without gaps or overlap.", sourceStart: span.Start,
                    sourceLength: span.Length);
            cursor += span.Length;
        }
        if (cursor != text.Length || text.Length > 0 && spans.Count == 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Style spans must cover the complete text.");
        if (text.Length == 0 && spans.Count == 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "An empty text still needs one empty style span for its line metrics.");
        Spans = new ReadOnlyCollection<TextStyleSpan>(spans.ToArray());
    }
    /// <summary>Creates a request with one style for the complete text.</summary>
    public ParagraphLayoutRequest(string text, TextStyle style, ParagraphLayoutOptions options)
        : this(text, new[] { new TextStyleSpan(0, text?.Length ?? 0, style) }, options) { }
    /// <summary>Unmodified logical Unicode source.</summary>
    public string Text { get; }
    /// <summary>Ordered style spans tiling the text.</summary>
    public IReadOnlyList<TextStyleSpan> Spans { get; }
    /// <summary>Layout options.</summary>
    public ParagraphLayoutOptions Options { get; }
}
