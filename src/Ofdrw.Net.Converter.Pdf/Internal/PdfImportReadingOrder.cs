using System;
using Ofdrw.Net.Core.Fonts;
using UglyToad.PdfPig.Content;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// Geometric word order does not establish logical bidi order. Until import
// consumes a verified logical structure, do not publish it as semantic text.
internal static class PdfImportReadingOrder
{
    internal static void Validate(Page page)
    {
        foreach (var letter in page.Letters)
        {
            var text = letter.Value;
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
                if (UnicodeFontSubsetProfile.RequiresBidiMirroring(scalar))
                {
                    throw new NotSupportedException(
                        $"PDF_TEXT_READING_ORDER_UNVERIFIED: PDF page {page.Number} contains right-to-left text or bidi controls. " +
                        "The default PDF importer cannot establish its logical reading order. " +
                        "Use TextLayerMode=None with RequireRasterization=true for an image-only conversion.");
                }
            }
        }
    }
}
