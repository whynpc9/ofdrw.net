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
            if (scalar is 0x200C or 0x200D)
                throw new NotSupportedException("PDFsharp cannot preserve Unicode join-control shaping semantics; use OFD/SVG or a shaping-capable PDF renderer.");
            if (OpenTypeCmap.IsVariationSelector(scalar) || scalar is >= 0x180B and <= 0x180D or 0x180F)
                throw new NotSupportedException("PDFsharp cannot preserve Unicode variation sequences; use OFD/SVG or a glyph-aware PDF renderer.");
            if (scalar is >= 0x202A and <= 0x202E or >= 0x2066 and <= 0x2069 or 0x200F or 0x061C)
                throw new NotSupportedException("PDFsharp cannot preserve bidi formatting control semantics; the original text remains supported in OFD/SVG.");
        }
    }
    // Drawing-only filtering. The OFD model/Unicode text is never rewritten.
    // Positioned runs retain their original grapheme/delta slots and only skip
    // painting the zero-width control; implicit runs draw the visible string.
    private static bool IsZeroWidthFormat(int scalar) => scalar is
        0x00AD or 0x180E or >= 0x200B and <= 0x200F or >= 0x2060 and <= 0x2064 or >= 0x206A and <= 0x206F or 0xFEFF or 0x061C;

    internal static string VisibleText(string text, OpenTypeCmap? coverage = null)
    {
        Validate(text);
        var output = new StringBuilder(text.Length);
        foreach (var scalar in OpenTypeFace.Scalars(text))
            if (!IsZeroWidthFormat(scalar) &&
                (!UnicodeFontSubsetProfile.IsNonRenderingControl(scalar) || coverage is null || coverage.Glyph(scalar) != 0))
                output.Append(char.ConvertFromUtf32(scalar));
        return output.ToString();
    }
}
