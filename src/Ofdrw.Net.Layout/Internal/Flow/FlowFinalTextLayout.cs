using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;

namespace Ofdrw.Net.Layout.Internal.Flow;

/// <summary>Lays out Flow/DOCX inlines with the final paragraph layout engine and returns Flow lines
/// carrying the final lines, so the existing pagination, tables and decorations place them unchanged.</summary>
/// <remarks>Line controls keep the compatibility semantics: CR/CRLF/LF end a line inside the paragraph
/// (the first-line indent applies once), trailing line controls reserve terminal line heights, and form
/// feed is a page control. Each line-delimited chunk is one engine request.</remarks>
internal sealed class FlowFinalTextLayout
{
    private readonly ParagraphLayoutEngine _engine;
    private readonly Func<FlowTextStyle, TextStyle> _map;
    private readonly TextSyntheticStylePolicy _synthetic;
    private readonly Dictionary<FlowTextStyle, TextStyle> _styles = new(ReferenceComparer.Instance);
    private readonly Dictionary<TextStyle, double> _emptyHeights = new(ReferenceComparer<TextStyle>.Instance);

    internal FlowFinalTextLayout(ParagraphLayoutEngine engine, Func<FlowTextStyle, TextStyle> map,
        TextSyntheticStylePolicy synthetic)
    { _engine = engine; _map = map; _synthetic = synthetic; }

    internal static FlowFinalTextLayout? Create(FlowTextEngine? engine, Func<FlowTextStyle, TextColor> color)
    {
        if (engine is null) return null;
        return new FlowFinalTextLayout(engine.Engine, style =>
        {
            var text = new TextStyle(engine.Chain(style.FontFamily), style.FontSizeMillimeters, new TextPaint(color(style)),
                style.Bold, style.Italic, engine.Language);
            if (style.Underline || style.Strikethrough)
                text = text.WithDecorations((style.Underline ? TextDecorationLines.Underline : 0) |
                    (style.Strikethrough ? TextDecorationLines.Strikethrough : 0));
            if (style.HighlightColor is { } highlight) text = text.WithHighlight(highlight);
            if (style.LetterSpacingMillimeters != 0) text = text.WithLetterSpacing(style.LetterSpacingMillimeters);
            return text;
        }, engine.SyntheticStyles);
    }

    internal static TextColor Color(OfdColor color) => new(color.Red, color.Green, color.Blue, color.Alpha);

    private TextStyle Style(FlowTextStyle style)
    {
        if (!_styles.TryGetValue(style, out var value)) _styles.Add(style, value = _map(style));
        return value;
    }

    /// <summary>Returns null when the inlines need the compatibility layout (inline images).</summary>
    internal IReadOnlyList<FlowLine>? Layout(IReadOnlyList<FlowInline> inlines, FlowParagraphFormat format, double width,
        FlowTextStyle fallback, CancellationToken token)
    {
        if (inlines.Any(inline => inline.Image is not null)) return null;
        // Concatenate first, then normalize: CRLF may straddle two inlines and belongs to the leading one.
        var raw = new StringBuilder();
        var rawStyles = new List<FlowTextStyle>();
        foreach (var inline in inlines)
            for (var index = 0; index < inline.Text.Length; index++)
            {
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                raw.Append(inline.Text[index]);
                rawStyles.Add(inline.Style);
            }
        var text = new StringBuilder(raw.Length);
        var styles = new List<FlowTextStyle>(raw.Length);
        for (var index = 0; index < raw.Length; index++)
        {
            var value = raw[index];
            var style = rawStyles[index];
            if (value == '\r')
            {
                if (index + 1 < raw.Length && raw[index + 1] == '\n') index++;
                value = '\n';
            }
            text.Append(value);
            styles.Add(style);
        }
        var source = text.ToString();
        var result = new List<FlowLine>();
        var lastContent = -1;
        for (var index = source.Length - 1; index >= 0; index--)
            if (source[index] != '\n' && source[index] != '\f') { lastContent = index; break; }
        var terminalNewline = false;
        var emptyStyle = fallback;
        var firstLine = true;
        var chunkStart = -1;
        void Empty(bool terminal, FlowTextStyle style)
        {
            var height = EmptyHeight(Style(style), format);
            result.Add(new FlowLine(new List<FlowGlyph>(), height, 0, FlowAlignment.Left, terminalNewline: terminal));
            firstLine = false;
        }
        void Chunk(int end)
        {
            Lines(source.Substring(chunkStart, end - chunkStart), chunkStart, styles, format, width, firstLine, result, token);
            firstLine = false;
            chunkStart = -1;
        }
        for (var index = 0; index < source.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            var value = source[index];
            if (value == '\n' || value == '\f')
            {
                if (value == '\n' && (!terminalNewline || index <= lastContent || lastContent < 0)) emptyStyle = styles[index];
                if (value == '\f' && terminalNewline) Empty(index > lastContent, emptyStyle);
                if (chunkStart >= 0) Chunk(index);
                else if (value == '\n') Empty(index > lastContent, emptyStyle);
                if (value == '\f') result.Add(new FlowLine(new List<FlowGlyph>(), 0, 0, FlowAlignment.Left, pageBreak: true));
                if (value == '\n') emptyStyle = styles[index];
                terminalNewline = value == '\n';
                continue;
            }
            terminalNewline = false;
            if (chunkStart < 0) chunkStart = index;
        }
        if (chunkStart >= 0) Chunk(source.Length);
        else if (result.Count == 0 || terminalNewline) Empty(terminalNewline, result.Count == 0 && source.Length == 0 ? fallback : emptyStyle);
        return result;
    }

    private ParagraphLayoutOptions Options(FlowParagraphFormat format, double width, double indent, int length)
    {
        var alignment = format.Alignment switch
        {
            FlowAlignment.Center => TextAlignment.Center,
            FlowAlignment.Right => TextAlignment.Right,
            FlowAlignment.Justify => TextAlignment.Justify,
            _ => TextAlignment.Left,
        };
        var options = new ParagraphLayoutOptions(width, null, TextParagraphDirection.Auto, alignment,
            format.MinimumLineHeightMillimeters > 0 ? TextLineHeight.AtLeast(Math.Min(1000, format.MinimumLineHeightMillimeters)) : TextLineHeight.Normal,
            0, indent, TextOverflowPolicy.BreakGraphemes, maximumLines: 1_000_000, maximumShapeCalls: 10_000_000,
            maximumTextLength: Math.Max(1, Math.Min(16_777_216, length)));
        options = options.WithSyntheticStyles(_synthetic);
        return format.TabStops.Count > 0 || format.DefaultTabIntervalMillimeters is not null
            ? options.WithTabStops(format.TabStops, format.DefaultTabIntervalMillimeters ?? 12.7, format.DefaultTabOriginMillimeters)
            : options;
    }

    private void Lines(string text, int offset, List<FlowTextStyle> styles, FlowParagraphFormat format, double width,
        bool first, List<FlowLine> result, CancellationToken token)
    {
        var spans = new List<TextStyleSpan>();
        for (var index = 0; index < text.Length;)
        {
            var style = styles[offset + index];
            var end = index + 1;
            // A span boundary may not split a surrogate pair; the pair belongs to its leading style.
            while (end < text.Length && (ReferenceEquals(styles[offset + end], style) || char.IsLowSurrogate(text[end]))) end++;
            spans.Add(new TextStyleSpan(index, end - index, Style(style)));
            index = end;
        }
        // A style change inside an extended grapheme cluster: that cluster takes its leading style, as in
        // the compatibility layout; the following text keeps its own style.
        var layout = Run(text, ClusterAlignedSpans(text, spans),
            Options(format, width, first ? format.FirstLineIndentMillimeters : 0, text.Length), token);
        foreach (var line in layout.Lines)
            result.Add(new FlowLine(new List<FlowGlyph>(), line.HeightMillimeters, 0, FlowAlignment.Left) { Final = line });
    }

    private ParagraphLayoutResult Run(string text, IReadOnlyList<TextStyleSpan> spans, ParagraphLayoutOptions options,
        CancellationToken token)
    {
        try { return _engine.Layout(new ParagraphLayoutRequest(text, spans, options), token); }
        catch (TextLayoutException exception)
        {
            throw new InvalidOperationException("Final text layout failed (" + exception.Error + "): " + exception.Message, exception);
        }
    }

    // Re-tiles the spans on UAX #29 extended grapheme cluster boundaries (the engine's own breaker): every
    // cluster takes the style of its first code unit; spans are otherwise unchanged.
    internal static List<TextStyleSpan> ClusterAlignedSpans(string text, IReadOnlyList<TextStyleSpan> spans)
    {
        if (spans.Count <= 1) return spans.ToList();
        var codePoints = new Ofdrw.Net.Layout.Text.Unicode.CodePointText(text);
        var boundaries = new bool[text.Length + 1];
        foreach (var boundary in Ofdrw.Net.Layout.Text.Unicode.GraphemeBreaker.Boundaries(codePoints.CodePoints))
            boundaries[codePoints.Offsets[boundary]] = true;
        boundaries[0] = boundaries[text.Length] = true;
        var styles = new TextStyle[text.Length];
        foreach (var span in spans)
            for (var index = span.Start; index < span.Start + span.Length; index++) styles[index] = span.Style;
        for (var index = 1; index < text.Length; index++)
            if (!boundaries[index]) styles[index] = styles[index - 1];
        var result = new List<TextStyleSpan>();
        for (var index = 0; index < text.Length;)
        {
            var end = index + 1;
            while (end < text.Length && ReferenceEquals(styles[end], styles[index])) end++;
            result.Add(new TextStyleSpan(index, end - index, styles[index]));
            index = end;
        }
        return result;
    }

    private double EmptyHeight(TextStyle style, FlowParagraphFormat format)
    {
        if (_emptyHeights.TryGetValue(style, out var height) && format.MinimumLineHeightMillimeters <= 0) return height;
        var layout = Run(string.Empty, new[] { new TextStyleSpan(0, 0, style) }, Options(format, 1000, 0, 1), CancellationToken.None);
        height = layout.Lines[0].HeightMillimeters;
        if (format.MinimumLineHeightMillimeters <= 0) _emptyHeights[style] = height;
        return height;
    }

    private sealed class ReferenceComparer : IEqualityComparer<FlowTextStyle>
    {
        internal static readonly ReferenceComparer Instance = new();
        public bool Equals(FlowTextStyle? x, FlowTextStyle? y) => ReferenceEquals(x, y);
        public int GetHashCode(FlowTextStyle obj) => RuntimeHelpers.GetHashCode(obj);
    }

    private sealed class ReferenceComparer<T> : IEqualityComparer<T> where T : class
    {
        internal static readonly ReferenceComparer<T> Instance = new();
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}

/// <summary>Writes final lines carried by Flow lines and owns their embedded font resources.</summary>
internal sealed class FlowFinalTextWriter
{
    private readonly OfdDocumentPackage _package;
    private readonly Dictionary<string, OfdFontResource> _fonts = new(StringComparer.Ordinal);

    internal FlowFinalTextWriter(OfdDocumentPackage package) => _package = package;

    /// <summary>Appends a final line whose line box top is placed at <paramref name="top"/>; returns the TextObject count.</summary>
    internal int Draw(OfdPage page, ParagraphLayoutLine line, double left, double top, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        foreach (var run in line.Runs) ParagraphLayoutOfdWriter.PreflightRun(run);
        ParagraphLayoutOfdWriter.AppendLine(page.Elements, line, left, top - line.TopMillimeters, Font);
        return line.Runs.Count;
    }

    private string Font(TextFontFace face)
    {
        if (_fonts.TryGetValue(face.FaceSha256, out var resource)) return resource.Id;
        var id = "F" + (_package.Fonts.Count + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
        while (_package.Fonts.Any(font => font.Id == id)) id += "x";
        resource = new OfdFontResource
        {
            Id = id, FontName = face.Name, FamilyName = face.Name, Charset = "unicode",
            Bold = face.Bold, Italic = face.Italic, Data = face.GetFaceBytes(), CollectionFaceIndex = 0,
        };
        _fonts.Add(face.FaceSha256, resource);
        _package.Fonts.Add(resource);
        return id;
    }
}
