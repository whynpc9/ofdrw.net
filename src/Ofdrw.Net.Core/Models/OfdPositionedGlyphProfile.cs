using System.Globalization;

namespace Ofdrw.Net.Core.Models;

// Shared narrow source profiles for typed production and explicit-GID readers.
internal static class OfdPositionedGlyphProfile
{
    internal static bool IsLatinCombiningPair(string text) => text.Length == 2 &&
        (text[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z') &&
        text[1] is >= '\u0300' and <= '\u036F' &&
        CharUnicodeInfo.GetUnicodeCategory(text[1]) == UnicodeCategory.NonSpacingMark;

    // Full-word RTL source with one nominal glyph occurrence per source letter. This is
    // deliberately narrower than Hebrew script or paragraph bidi in general.
    internal static bool IsHebrewWord(string text)
    {
        if (text.Length is < 2 or > 8) return false;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] is < '\u05d0' or > '\u05ea') return false;
        }
        return true;
    }
}
