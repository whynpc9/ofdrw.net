using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>Structured failure categories for new text queries and their limited output path.</summary>
public enum TextLayoutError
{
    /// <summary>The supplied font bytes or metrics are malformed.</summary>
    InvalidFont,
    /// <summary>The input needs a font or text profile that has not been verified.</summary>
    UnsupportedFontProfile,
    /// <summary>The logical Unicode input is invalid for the selected run.</summary>
    InvalidText,
    /// <summary>A request or replacement shaper result violates the contract.</summary>
    InvalidRequest,
    /// <summary>No supplied face can represent a required glyph.</summary>
    MissingGlyph,
    /// <summary>The pinned native shaping library could not be loaded or called.</summary>
    NativeUnavailable,
    /// <summary>A finite font, text, glyph, line, or shaping budget was exceeded.</summary>
    BudgetExceeded,
    /// <summary>The completed query cannot be represented by the limited output model.</summary>
    UnsupportedOutput,
}

/// <summary>A structured failure from the new text query path. No partial result is returned.</summary>
public sealed class TextLayoutException : Exception
{
    /// <summary>Creates a structured failure with an optional absolute UTF-16 source interval.</summary>
    public TextLayoutException(TextLayoutError error, string message, Exception? inner = null,
        int sourceStart = -1, int sourceLength = 0) : base(message, inner)
    { Error = error; SourceStart = sourceStart; SourceLength = sourceLength; }
    /// <summary>Failure category suitable for caller decisions and diagnostics.</summary>
    public TextLayoutError Error { get; }
    /// <summary>Absolute UTF-16 offset when a specific source interval caused the failure; otherwise -1.</summary>
    public int SourceStart { get; }
    /// <summary>Length in UTF-16 code units of the offending interval, or zero if none is known.</summary>
    public int SourceLength { get; }
}

/// <summary>Logical direction for one already analyzed shaping run.</summary>
public enum TextDirection
{
    /// <summary>Glyph advances progress left to right.</summary>
    LeftToRight,
    /// <summary>Glyph advances progress right to left; paragraph bidi remains a separate responsibility.</summary>
    RightToLeft,
}

/// <summary>An OpenType feature over a half-open UTF-16 interval relative to the shaped run.</summary>
public sealed class TextFeature
{
    /// <summary>Sets one OpenType feature over a run-relative half-open UTF-16 interval.</summary>
    public TextFeature(string tag, uint value, int start, int length)
    {
        if (tag is null || tag.Length != 4) throw new ArgumentException("Feature tag must contain four characters.", nameof(tag));
        for (var i = 0; i < 4; i++) if (tag[i] < 0x20 || tag[i] > 0x7e) throw new ArgumentException("Feature tag must be ASCII.", nameof(tag));
        if (start < 0 || length < 0) throw new ArgumentOutOfRangeException(nameof(start));
        Tag = tag; Value = value; Start = start; Length = length;
    }
    /// <summary>Four-character ASCII feature tag.</summary>
    public string Tag { get; }
    /// <summary>OpenType feature value, often zero or one.</summary>
    public uint Value { get; }
    /// <summary>Run-relative UTF-16 start offset.</summary>
    public int Start { get; }
    /// <summary>Run-relative UTF-16 length.</summary>
    public int Length { get; }
}

/// <summary>A single logical run with one selected font, script, language, and direction.</summary>
public sealed class TextShapeRequest
{
    /// <summary>Creates a single-font logical run without selecting a host font or performing bidi analysis.</summary>
    public TextShapeRequest(string text, int sourceStart, TextFontFace face, TextDirection direction = TextDirection.LeftToRight,
        string script = "Latn", string language = "en", IReadOnlyList<TextFeature>? features = null)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Face = face ?? throw new ArgumentNullException(nameof(face));
        if (direction != TextDirection.LeftToRight && direction != TextDirection.RightToLeft)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Unsupported text direction.");
        if (sourceStart < 0 || (long)sourceStart + text.Length > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(sourceStart));
        if (script is null || script.Length != 4) throw new ArgumentException("Script must be a four-letter OpenType tag.", nameof(script));
        foreach (var c in script) if (c < 'A' || c > 'z' || c is >= '[' and <= '`') throw new ArgumentException("Script tag must use ASCII letters.", nameof(script));
        if (string.IsNullOrWhiteSpace(language) || language.Length > 32) throw new ArgumentException("A short language tag is required.", nameof(language));
        if (text.Length > 4096) throw new TextLayoutException(TextLayoutError.BudgetExceeded, "A shaping run exceeds 4096 UTF-16 code units.");
        for (var i = 0; i < text.Length; i++)
        {
            var scalar = text[i];
            if (char.IsHighSurrogate(text[i]))
            {
                if (i + 1 >= text.Length || !char.IsLowSurrogate(text[i + 1]))
                    throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                        sourceStart: sourceStart + i, sourceLength: 1);
                var codepoint = char.ConvertToUtf32(text, i);
                if (codepoint is >= 0xe0100 and <= 0xe01ef)
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "Variation selectors need a separately verified text profile.",
                        sourceStart: sourceStart + i, sourceLength: 2);
                i++;
            }
            else if (char.IsLowSurrogate(text[i]))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                    sourceStart: sourceStart + i, sourceLength: 1);
            else if (text[i] == '\r' || text[i] == '\n' || text[i] == '\t' ||
                IsUnsupportedLineSeparator(text[i]))
                throw new TextLayoutException(TextLayoutError.InvalidText,
                    "A shaping run cannot contain line or tab controls.",
                    sourceStart: sourceStart + i, sourceLength: 1);
            else if (char.GetUnicodeCategory(scalar) == System.Globalization.UnicodeCategory.Format ||
                scalar is >= '\ufe00' and <= '\ufe0f')
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Format controls and variation selectors need a separately verified text profile.",
                    sourceStart: sourceStart + i, sourceLength: 1);
        }
        var featureCopy = new List<TextFeature>();
        if (features is not null) foreach (var feature in features)
        {
            if (feature is null || feature.Start > text.Length || feature.Length > text.Length - feature.Start)
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "A feature interval is outside the source run.");
            if ((feature.Start > 0 && feature.Start < text.Length && char.IsLowSurrogate(text[feature.Start])) ||
                (feature.Start + feature.Length > 0 && feature.Start + feature.Length < text.Length &&
                char.IsLowSurrogate(text[feature.Start + feature.Length])))
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "A feature interval splits a surrogate pair.");
            featureCopy.Add(feature);
        }
        SourceStart = sourceStart; Direction = direction; Script = script; Language = language;
        Features = new ReadOnlyCollection<TextFeature>(featureCopy);
    }
    internal static bool IsUnsupportedLineSeparator(char value) =>
        value is '\u000b' or '\u000c' or '\u0085' or '\u2028' or '\u2029';
    /// <summary>Unmodified logical Unicode source for this run.</summary>
    public string Text { get; }
    /// <summary>Absolute UTF-16 start offset of this run in the caller's source.</summary>
    public int SourceStart { get; }
    /// <summary>The immutable selected face used by shaping and measurement.</summary>
    public TextFontFace Face { get; }
    /// <summary>Direction already chosen by the caller's run analysis.</summary>
    public TextDirection Direction { get; }
    /// <summary>Four-letter ISO 15924 script tag for this run.</summary>
    public string Script { get; }
    /// <summary>BCP 47 language tag for this run.</summary>
    public string Language { get; }
    /// <summary>Immutable feature intervals relative to the run.</summary>
    public IReadOnlyList<TextFeature> Features { get; }
}

/// <summary>Glyph ink rectangle in font design units with Y pointing upward.</summary>
public sealed class TextInkBounds
{
    /// <summary>Creates a glyph ink rectangle in font design units with Y upward.</summary>
    public TextInkBounds(double x, double y, double width, double height)
    {
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y) ||
            double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height) ||
            width < 0 || height < 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid glyph ink bounds.");
        X = x; Y = y; Width = width; Height = height;
    }
    /// <summary>Leftmost ink coordinate relative to the glyph origin.</summary>
    public double X { get; }
    /// <summary>Bottom ink coordinate relative to the glyph origin.</summary>
    public double Y { get; }
    /// <summary>Ink width, independent of advance.</summary>
    public double Width { get; }
    /// <summary>Ink height, independent of line height.</summary>
    public double Height { get; }
}

/// <summary>One shaped GID and placement. Cluster intervals are source UTF-16 offsets, not caret stops.</summary>
public sealed class TextShapedGlyph
{
    /// <summary>Creates a positioned glyph with its logical source cluster and ink bounds.</summary>
    public TextShapedGlyph(uint glyphId, int clusterStart, int clusterLength, int xAdvance, int yAdvance,
        int xOffset, int yOffset, TextInkBounds? ink, bool unsafeToBreak)
    {
        GlyphId = glyphId; ClusterStart = clusterStart; ClusterLength = clusterLength;
        XAdvance = xAdvance; YAdvance = yAdvance; XOffset = xOffset; YOffset = yOffset;
        Ink = ink; UnsafeToBreak = unsafeToBreak;
    }
    /// <summary>Actual glyph ID in the selected face.</summary>
    public uint GlyphId { get; }
    /// <summary>Absolute UTF-16 start of the source cluster.</summary>
    public int ClusterStart { get; }
    /// <summary>UTF-16 source length associated with this cluster.</summary>
    public int ClusterLength { get; }
    /// <summary>Horizontal advance in font design units.</summary>
    public int XAdvance { get; }
    /// <summary>Vertical advance in font design units.</summary>
    public int YAdvance { get; }
    /// <summary>Horizontal origin offset in font design units.</summary>
    public int XOffset { get; }
    /// <summary>Vertical origin offset in font design units.</summary>
    public int YOffset { get; }
    /// <summary>Glyph ink rectangle, or null for no visible outline.</summary>
    public TextInkBounds? Ink { get; }
    /// <summary>Whether the shaper marks a break immediately before this glyph as unsafe.</summary>
    public bool UnsafeToBreak { get; }
}

/// <summary>A complete immutable run result; all positions are in the selected face's design units.</summary>
public sealed class TextShapedRun
{
    /// <summary>Freezes and validates one complete run result from a replaceable shaper.</summary>
    public TextShapedRun(TextShapeRequest request, IReadOnlyList<TextShapedGlyph> glyphs, string engineVersion)
    {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        Glyphs = new ReadOnlyCollection<TextShapedGlyph>(new List<TextShapedGlyph>(glyphs ?? throw new ArgumentNullException(nameof(glyphs))));
        EngineVersion = engineVersion ?? throw new ArgumentNullException(nameof(engineVersion));
        if (request.Text.Length > 0 && Glyphs.Count == 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Shaper discarded a nonempty source run.");
        if (Glyphs.Count > request.Text.Length * 4L + 16)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Shaped glyph count exceeds the expansion budget.");
        long x = 0, y = 0;
        var intervals = new List<(int Start, int End)>(Glyphs.Count);
        foreach (var glyph in Glyphs)
        {
            if (glyph is null || glyph.GlyphId == 0 || glyph.GlyphId >= request.Face.GlyphCount ||
                glyph.ClusterStart < request.SourceStart ||
                glyph.ClusterLength <= 0 || (long)glyph.ClusterStart + glyph.ClusterLength >
                (long)request.SourceStart + request.Text.Length)
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "Shaper returned an invalid glyph or source interval.");
            var relativeStart = glyph.ClusterStart - request.SourceStart;
            var relativeEnd = relativeStart + glyph.ClusterLength;
            if (relativeStart > 0 && char.IsLowSurrogate(request.Text[relativeStart]) ||
                relativeEnd < request.Text.Length && char.IsLowSurrogate(request.Text[relativeEnd]))
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "Shaper split a UTF-16 surrogate pair at a cluster boundary.");
            intervals.Add((glyph.ClusterStart, glyph.ClusterStart + glyph.ClusterLength));
            x += glyph.XAdvance; y += glyph.YAdvance;
        }
        intervals.Sort((left, right) => left.Start.CompareTo(right.Start));
        var coveredTo = request.SourceStart;
        foreach (var interval in intervals)
        {
            if (interval.Start > coveredTo)
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "Shaper omitted a source interval from the completed run.");
            coveredTo = Math.Max(coveredTo, interval.End);
        }
        if (coveredTo != request.SourceStart + request.Text.Length)
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "Shaper omitted the end of the source run.");
        if (Math.Abs(x) > int.MaxValue || Math.Abs(y) > int.MaxValue)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Shaped advance exceeds the coordinate budget.");
        XAdvance = (int)x; YAdvance = (int)y;
    }
    /// <summary>The exact input request returned by the shaper.</summary>
    public TextShapeRequest Request { get; }
    /// <summary>Immutable glyphs in the shaper's output order.</summary>
    public IReadOnlyList<TextShapedGlyph> Glyphs { get; }
    /// <summary>Shaping adapter and pinned package version used for this result.</summary>
    public string EngineVersion { get; }
    /// <summary>Total horizontal advance in font design units.</summary>
    public int XAdvance { get; }
    /// <summary>Total vertical advance in font design units.</summary>
    public int YAdvance { get; }
}

/// <summary>Shapes one preselected font run. Bidi, fallback, line breaking, and output remain caller responsibilities.</summary>
public interface ITextShaper
{
    /// <summary>Shapes one preselected logical run; cancellation is checked around the native call.</summary>
    TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default);
}
