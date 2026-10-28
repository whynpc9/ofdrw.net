using System.Globalization;

namespace Ofdrw.Net.Core.Models;

// Shared narrow source profile for typed production and explicit-GID readers.
internal static class OfdPositionedGlyphProfile
{
    internal static bool IsLatinCombiningPair(string text) => text.Length == 2 &&
        (text[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z') &&
        text[1] is >= '\u0300' and <= '\u036F' &&
        CharUnicodeInfo.GetUnicodeCategory(text[1]) == UnicodeCategory.NonSpacingMark;
}
