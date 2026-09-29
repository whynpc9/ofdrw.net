using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Ofdrw.Net.Layout.Internal.Flow;

internal enum FlowAlignment { Left, Center, Right }

internal sealed class FlowTextStyle
{
    internal string FontFamily { get; set; } = "SimSun";
    internal double FontSizeMillimeters { get; set; } = 3.704167d;
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal object? Source { get; set; }
}

internal sealed class FlowInline
{
    internal string Text { get; set; } = string.Empty;
    internal FlowTextStyle Style { get; set; } = new();
    internal object? Image { get; set; }
    internal double ImageWidthMillimeters { get; set; }
    internal double ImageHeightMillimeters { get; set; }
}

internal sealed class FlowParagraphFormat
{
    internal FlowAlignment Alignment { get; set; }
    internal double FirstLineIndentMillimeters { get; set; }
    internal double MinimumLineHeightMillimeters { get; set; }
}

internal interface IFlowFontMetrics
{
    double AdvanceMillimeters(string grapheme, FlowTextStyle style);
}

internal sealed class FlowGlyph
{
    internal FlowGlyph(string text, FlowTextStyle style, double width, object? image = null, double imageHeight = 0)
    { Text = text; Style = style; Width = width; Image = image; ImageHeight = imageHeight; }
    internal string Text { get; }
    internal FlowTextStyle Style { get; }
    internal double Width { get; }
    internal object? Image { get; }
    internal double ImageHeight { get; }
}

internal sealed class FlowLine
{
    internal FlowLine(List<FlowGlyph> glyphs, double height, double indent, FlowAlignment alignment, bool pageBreak = false)
    { Glyphs = glyphs; Height = height; Indent = indent; Alignment = alignment; PageBreak = pageBreak; }
    internal List<FlowGlyph> Glyphs { get; }
    internal double Height { get; }
    internal double Indent { get; }
    internal FlowAlignment Alignment { get; }
    internal bool PageBreak { get; }
    internal double BaselineMillimeters => Glyphs.Count == 0 ? 0 : Glyphs.Max(g => g.Style.FontSizeMillimeters);
}

/// <summary>The same grapheme, Latin word, line-height and alignment calculations serve public flow and DOCX Native.</summary>
internal static class FlowParagraphLayout
{
    internal static IReadOnlyList<FlowLine> Layout(
        IReadOnlyList<FlowInline> inlines, FlowParagraphFormat format, double width,
        IFlowFontMetrics metrics, FlowTextStyle fallback, CancellationToken cancellationToken,
        bool allowOversizeGlyph = false)
    {
        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        var glyphs = new List<FlowGlyph>();
        var rawText = new StringBuilder();
        var rawStyles = new List<FlowTextStyle>();
        void FlushText()
        {
            if (rawText.Length == 0) return;
            var normalized = new StringBuilder(rawText.Length);
            var styles = new List<FlowTextStyle>(rawText.Length);
            for (var index = 0; index < rawText.Length; index++)
            {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                var value = rawText[index];
                var style = rawStyles[index];
                if (value == '\r')
                {
                    // CRLF may straddle two spans; the control belongs to its leading span.
                    if (index + 1 < rawText.Length && rawText[index + 1] == '\n') index++;
                    value = '\n';
                }
                normalized.Append(value);
                styles.Add(style);
            }
            var enumerator = StringInfo.GetTextElementEnumerator(normalized.ToString());
            while (enumerator.MoveNext())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var element = enumerator.GetTextElement();
                // A grapheme crossing a span boundary takes the leading scalar's style.
                var style = styles[enumerator.ElementIndex];
                glyphs.Add(new FlowGlyph(element, style,
                    element == "\n" || element == "\f" ? 0 : metrics.AdvanceMillimeters(element, style)));
            }
            rawText.Clear();
            rawStyles.Clear();
        }
        foreach (var inline in inlines)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (inline.Image is not null)
            {
                FlushText();
                if (double.IsNaN(inline.ImageWidthMillimeters) || double.IsInfinity(inline.ImageWidthMillimeters) ||
                    inline.ImageWidthMillimeters <= 0 || double.IsNaN(inline.ImageHeightMillimeters) ||
                    double.IsInfinity(inline.ImageHeightMillimeters) || inline.ImageHeightMillimeters <= 0)
                    throw new InvalidDataException("An inline image has invalid dimensions.");
                var imageWidth = Math.Min(inline.ImageWidthMillimeters, width);
                var scale = imageWidth / inline.ImageWidthMillimeters;
                glyphs.Add(new FlowGlyph(string.Empty, inline.Style, imageWidth, inline.Image,
                    inline.ImageHeightMillimeters * scale));
                continue;
            }
            for (var index = 0; index < inline.Text.Length; index++)
            {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                rawText.Append(inline.Text[index]);
                rawStyles.Add(inline.Style);
            }
        }
        FlushText();

        var result = new List<FlowLine>();
        var current = new List<FlowGlyph>();
        var currentWidth = 0d;
        var indent = format.FirstLineIndentMillimeters;
        void Flush()
        {
            var size = current.Count == 0 ? fallback.FontSizeMillimeters : current.Max(g => g.Style.FontSizeMillimeters);
            var imageHeight = current.Count == 0 ? 0 : current.Max(g => g.ImageHeight);
            result.Add(new FlowLine(current, Math.Max(imageHeight, Math.Max(size * 1.3, format.MinimumLineHeightMillimeters)),
                indent, format.Alignment));
            current = new List<FlowGlyph>();
            currentWidth = 0;
            indent = 0;
        }
        static bool Word(FlowGlyph glyph) => glyph.Text.Length == 1 && glyph.Text[0] < 128 &&
            char.IsLetterOrDigit(glyph.Text[0]);
        for (var i = 0; i < glyphs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var glyph = glyphs[i];
            if (glyph.Text == "\f" || glyph.Text == "\n")
            {
                if (current.Count > 0 || glyph.Text == "\n") Flush();
                if (glyph.Text == "\f") result.Add(new FlowLine(new List<FlowGlyph>(), 0, 0, format.Alignment, true));
                continue;
            }
            if (Word(glyph) && (i == 0 || !Word(glyphs[i - 1])))
            {
                var wordWidth = 0d;
                for (var j = i; j < glyphs.Count && Word(glyphs[j]); j++) wordWidth += glyphs[j].Width;
                if (current.Count > 0 && wordWidth <= width && currentWidth + wordWidth > width - indent) Flush();
            }
            if (current.Count > 0 && currentWidth + glyph.Width > width - indent) Flush();
            if (glyph.Image is not null && width > indent && glyph.Width > width - indent && current.Count == 0)
            {
                var scale = (width - indent) / glyph.Width;
                glyph = new FlowGlyph(glyph.Text, glyph.Style, width - indent,
                    glyph.Image, glyph.ImageHeight * scale);
            }
            if (glyph.Width > width - indent && current.Count == 0 && !allowOversizeGlyph)
                throw new InvalidOperationException("A flow glyph is wider than the usable line width.");
            current.Add(glyph);
            currentWidth += glyph.Width;
        }
        if (current.Count > 0 || result.Count == 0) Flush();
        return result;
    }

    internal static double Align(double left, double width, double contentWidth, FlowAlignment alignment) =>
        contentWidth >= width ? left : alignment switch
        {
            FlowAlignment.Center => left + (width - contentWidth) / 2,
            FlowAlignment.Right => left + width - contentWidth,
            _ => left
        };
}

internal static class FlowPagination
{
    /// <summary>Returns true when content needs another page, and fails if it cannot fit on an empty page.</summary>
    internal static bool NeedsNewPage(double y, double height, double contentTop, double contentBottom)
    {
        if (double.IsNaN(height) || double.IsInfinity(height) || height < 0 || contentBottom <= contentTop)
            throw new ArgumentOutOfRangeException(nameof(height));
        if (height > contentBottom - contentTop + 0.000001d)
            throw new InvalidOperationException("A flow line is taller than the usable page area.");
        return y + height > contentBottom + 0.000001d;
    }
}

internal static class FlowTextMetrics
{
    internal static bool IsCjkTypographicUnit(string grapheme)
    {
        if (string.IsNullOrEmpty(grapheme)) return false;
        var scalar = char.IsHighSurrogate(grapheme[0]) && grapheme.Length > 1 &&
            char.IsLowSurrogate(grapheme[1])
            ? char.ConvertToUtf32(grapheme, 0)
            : grapheme[0];
        return scalar >= 0x1100 && scalar <= 0x11FF ||
               scalar >= 0x2E80 && scalar <= 0x9FFF ||
               scalar >= 0xA960 && scalar <= 0xA97F ||
               scalar >= 0xAC00 && scalar <= 0xD7AF ||
               scalar >= 0xD7B0 && scalar <= 0xD7FF ||
               scalar >= 0xF900 && scalar <= 0xFAFF ||
               scalar >= 0xFF00 && scalar <= 0xFFEF ||
               scalar >= 0x1B000 && scalar <= 0x1B16F ||
               scalar >= 0x20000 && scalar <= 0x3FFFF;
    }
}
