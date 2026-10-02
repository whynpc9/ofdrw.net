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
            if (OpenTypeCmap.IsVariationSelector(scalar) || scalar is >= 0x180B and <= 0x180D or 0x180F)
                throw new NotSupportedException("PDFsharp cannot preserve Unicode variation sequences; use OFD/SVG or a glyph-aware PDF renderer.");
            if (scalar is >= 0x202A and <= 0x202E or >= 0x2066 and <= 0x2069 or 0x200E or 0x200F or 0x061C)
                throw new NotSupportedException("PDFsharp cannot preserve bidi formatting control semantics; the original text remains supported in OFD/SVG.");
            if (UnicodeFontSubsetProfile.IsNonRenderingControl(scalar) && !IsHangulFiller(scalar))
                throw new NotSupportedException("PDFsharp cannot preserve semantic default-ignorable controls; use OFD/SVG or a renderer that retains their Unicode semantics.");
        }
    }
    // These mapped spacing letters have dedicated advance/Delta regressions.
    // Unicode permits default-ignorables with no glyph to be omitted; this
    // exception is not a general shaping or semantic-preservation claim.
    private static bool IsHangulFiller(int scalar) => scalar is 0x115F or 0x1160 or 0x3164 or 0xFFA0;

    internal static string VisibleText(string text, OpenTypeCmap? coverage = null)
    {
        Validate(text);
        var output = new StringBuilder(text.Length);
        foreach (var scalar in OpenTypeFace.Scalars(text))
            if (!IsHangulFiller(scalar) || coverage is null || coverage.Glyph(scalar) != 0)
                output.Append(char.ConvertFromUtf32(scalar));
        return output.ToString();
    }
}
