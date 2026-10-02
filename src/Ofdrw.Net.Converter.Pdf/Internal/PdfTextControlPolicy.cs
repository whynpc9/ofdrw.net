using System;
using System.Text;
using Ofdrw.Net.Core.Fonts;
namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class PdfTextControlPolicy
{
    internal static void Validate(string text)
    {
        foreach (var scalar in OpenTypeFace.Scalars(text))
        {
            if (OpenTypeCmap.IsVariationSelector(scalar))
                throw new NotSupportedException("PDFsharp cannot preserve Unicode variation sequences; use OFD/SVG or a glyph-aware PDF renderer.");
            if (UnicodeFontSubsetProfile.IsNonRenderingControl(scalar) && UnicodeFontSubsetProfile.RequiresBidiMirroring(scalar))
                throw new NotSupportedException("PDFsharp cannot preserve bidi formatting control semantics; the original text remains supported in OFD/SVG.");
        }
    }
    // Drawing-only filtering. The OFD model/Unicode text is never rewritten.
    // Positioned runs retain their original grapheme/delta slots and only skip
    // painting the zero-width control; implicit runs draw the visible string.
    internal static string VisibleText(string text)
    {
        var output = new StringBuilder(text.Length);
        foreach (var scalar in OpenTypeFace.Scalars(text))
            if (!UnicodeFontSubsetProfile.IsNonRenderingControl(scalar)) output.Append(char.ConvertFromUtf32(scalar));
        return output.ToString();
    }
}
