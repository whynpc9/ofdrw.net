using System;
using Ofdrw.Net.Core.Fonts;
using UglyToad.PdfPig.Content;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// Geometric word order does not establish logical bidi order. Unless the caller
// opts into the bounded, verified visual-line profile (PdfRtlVisualLineImport),
// do not publish it as semantic text.
internal static class PdfImportReadingOrder
{
    internal const string Code = "PDF_TEXT_READING_ORDER_UNVERIFIED";

    internal static void Validate(Page page)
    {
        if (ContainsRightToLeft(page))
        {
            throw new NotSupportedException(
                $"{Code}: PDF page {page.Number} contains right-to-left text or bidi controls. " +
                "The default PDF importer cannot establish its logical reading order. " +
                "Use TextLayerMode=None with RequireRasterization=true for an image-only conversion, " +
                "or opt into RightToLeftTextImport=VisualLineProfile for the bounded verified import.");
        }
    }

    internal static bool ContainsRightToLeft(Page page)
    {
        foreach (var letter in page.Letters)
            if (ContainsRightToLeft(letter.Value)) return true;
        return false;
    }

    internal static bool ContainsRightToLeft(string text)
    {
        for (var index = 0; index < text.Length; index++)
        {
            var scalar = (int)text[index];
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length &&
                char.IsLowSurrogate(text[index + 1]))
            {
                scalar = char.ConvertToUtf32(text[index], text[++index]);
            }
            // Pinned Unicode 17 R/AL/bidi-control ranges: a rejection
            // boundary, not an inverse-bidi implementation.
            if (UnicodeFontSubsetProfile.RequiresBidiMirroring(scalar)) return true;
        }
        return false;
    }

    // A rejection by the opt-in profile keeps the same stable code and the
    // explicit image-only alternative.
    internal static NotSupportedException ProfileRejection(int pageNumber, string reason) => new(
        $"{Code}: PDF page {pageNumber}: {reason} The bounded RightToLeftTextImport=VisualLineProfile does not admit it. " +
        "Use TextLayerMode=None with RequireRasterization=true for an image-only conversion.");
}
