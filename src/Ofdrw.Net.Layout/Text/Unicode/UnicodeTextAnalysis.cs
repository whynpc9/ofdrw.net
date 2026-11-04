using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ofdrw.Net.Layout.Text.Unicode;

/// <summary>A line break opportunity before a UTF-16 offset.</summary>
public readonly struct TextLineBreakOpportunity
{
    internal TextLineBreakOpportunity(int offset, bool mandatory) { Offset = offset; Mandatory = mandatory; }
    /// <summary>UTF-16 offset before which a line may (or must) break.</summary>
    public int Offset { get; }
    /// <summary>Whether the break is a hard (mandatory) break.</summary>
    public bool Mandatory { get; }
}

/// <summary>Resolved bidi levels of one paragraph, per UTF-16 code unit.</summary>
public sealed class TextBidiParagraph
{
    private readonly BidiAlgorithm _bidi;
    private readonly CodePointText _text;

    internal TextBidiParagraph(BidiAlgorithm bidi, CodePointText text) { _bidi = bidi; _text = text; }

    /// <summary>Paragraph embedding level: 0 (LTR) or 1 (RTL).</summary>
    public int ParagraphLevel => _bidi.ParagraphLevel;

    /// <summary>Applies L1 to one line [start, end) (UTF-16) and returns the level of each UTF-16
    /// code unit. Both offsets must be on code point boundaries.</summary>
    public IReadOnlyList<int> LineLevels(int start, int end)
    {
        var first = _text.IndexOfOffset(start);
        var last = _text.IndexOfOffset(end);
        if (first > last) throw new ArgumentOutOfRangeException(nameof(start));
        var levels = _bidi.LineLevels(first, last);
        var result = new int[end - start];
        for (var index = first; index < last; index++)
            for (var unit = _text.Offsets[index]; unit < _text.Offsets[index + 1]; unit++)
                result[unit - start] = levels[index - first];
        return new ReadOnlyCollection<int>(result);
    }

    /// <summary>L2: logical UTF-16 offsets of code points in visual left-to-right order for one line.</summary>
    public IReadOnlyList<int> VisualOrder(int start, int end)
    {
        var first = _text.IndexOfOffset(start);
        var last = _text.IndexOfOffset(end);
        var order = BidiAlgorithm.VisualOrder(_bidi.LineLevels(first, last));
        var result = new int[order.Length];
        for (var index = 0; index < order.Length; index++) result[index] = _text.Offsets[first + order[index]];
        return new ReadOnlyCollection<int>(result);
    }
}

/// <summary>Public entry points to the frozen Unicode 17.0.0 segmentation and bidi algorithms used
/// by paragraph layout (UAX #9 rev 51, UAX #14 rev 55, UAX #29 rev 47; untailored).</summary>
public static class UnicodeTextAnalysis
{
    /// <summary>Unicode Character Database version of every property table.</summary>
    public static string UnicodeVersion => UnicodeData.Version;

    /// <summary>Extended grapheme cluster boundaries as UTF-16 offsets, including 0 and the length.</summary>
    public static IReadOnlyList<int> GraphemeBoundaries(string text)
    {
        var points = new CodePointText(text ?? throw new ArgumentNullException(nameof(text)));
        var boundaries = GraphemeBreaker.Boundaries(points.CodePoints);
        var result = new int[boundaries.Length];
        for (var index = 0; index < boundaries.Length; index++) result[index] = points.Offsets[boundaries[index]];
        return new ReadOnlyCollection<int>(result);
    }

    /// <summary>Default UAX #14 break opportunities, including the mandatory break at the end.</summary>
    public static IReadOnlyList<TextLineBreakOpportunity> LineBreakOpportunities(string text)
    {
        var points = new CodePointText(text ?? throw new ArgumentNullException(nameof(text)));
        var kinds = LineBreaker.Analyze(points.CodePoints);
        var result = new List<TextLineBreakOpportunity>();
        for (var index = 1; index < kinds.Length; index++)
            if (kinds[index] != LineBreakKind.Prohibited)
                result.Add(new TextLineBreakOpportunity(points.Offsets[index], kinds[index] == LineBreakKind.Mandatory));
        return new ReadOnlyCollection<TextLineBreakOpportunity>(result);
    }

    /// <summary>Resolves one paragraph (no paragraph separators inside) with UAX #9.</summary>
    /// <param name="paragraph">Paragraph text.</param>
    /// <param name="direction">Null for P2/P3 auto-LTR, otherwise the base direction.</param>
    public static TextBidiParagraph ResolveBidi(string paragraph, TextDirection? direction = null)
    {
        var points = new CodePointText(paragraph ?? throw new ArgumentNullException(nameof(paragraph)));
        var types = new BidiClass[points.Length];
        for (var index = 0; index < types.Length; index++)
        {
            types[index] = UnicodeProperties.Bidi(points.CodePoints[index]);
            if (types[index] == BidiClass.B && index != types.Length - 1)
                throw new TextLayoutException(TextLayoutError.InvalidText,
                    "Split paragraphs at paragraph separators before bidi resolution.",
                    sourceStart: points.Offsets[index], sourceLength: points.Offsets[index + 1] - points.Offsets[index]);
        }
        var level = direction switch { TextDirection.LeftToRight => 0, TextDirection.RightToLeft => 1, _ => -1 };
        return new TextBidiParagraph(BidiAlgorithm.Resolve(types, points.CodePoints, level), points);
    }
}
