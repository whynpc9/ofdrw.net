using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// OFD keeps logical TextObject order. PDF readers infer bidi from character
// positions and content order, so a disjoint RTL phrase is painted left to right.
// Do not reorder arbitrary objects, clipped/transformed text or overlapping ink.
internal static class PdfRtlPaintOrder
{
    internal static IReadOnlyList<OfdElement> Arrange(IReadOnlyList<OfdElement> source,
        DocumentFontContext fonts, CancellationToken token)
    {
        OfdElement[]? ordered = null;
        for (var index = 0; index + 2 < source.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            if (source[index] is not OfdTextElement first || Ink(first, fonts) is not { } previousInk) continue;
            var end = index;
            var previousX = first.XMillimeters;
            while (end + 2 < source.Count && source[end + 1] is OfdTextElement space &&
                   source[end + 2] is OfdTextElement next && Plain(space) && Plain(next) &&
                   space.Text == " " && !EmbeddedFontCoverage.HasExplicitGlyphReferences(space) &&
                   SameLine(first, space) && SameLine(first, next) &&
                   Ink(next, fonts) is { } nextInk && fonts.IsEmptyRtlSeparator(space) &&
                   nextInk.Right <= previousInk.Left &&
                   next.XMillimeters < space.XMillimeters && space.XMillimeters < previousX)
            {
                end += 2;
                previousInk = nextInk;
                previousX = next.XMillimeters;
            }
            if (end == index) continue;
            ordered ??= source.ToArray();
            for (var offset = 0; offset <= end - index; offset++)
                ordered[index + offset] = source[end - offset];
            index = end;
        }
        return ordered ?? source;
    }

    private static bool Plain(OfdTextElement text) => text.Runs.Count == 1 &&
        text.Runs[0].Text == text.Text && text.Transform is null && text.Clips is null &&
        string.IsNullOrWhiteSpace(text.ClippingXml) && OfdGraphicXmlContract.IsVisible(text);

    private static (double Left, double Right)? Ink(OfdTextElement text, DocumentFontContext fonts) =>
        Plain(text) && OfdPositionedGlyphProfile.IsHebrewWord(text.Text) ? fonts.VisualHebrewInk(text) : null;

    private static bool SameLine(OfdTextElement left, OfdTextElement right) =>
        left.FontResourceId == right.FontResourceId && left.FontSizeMillimeters == right.FontSizeMillimeters &&
        Math.Abs(left.YMillimeters + left.Runs[0].YMillimeters -
                 right.YMillimeters - right.Runs[0].YMillimeters) <= 0.001;
}
