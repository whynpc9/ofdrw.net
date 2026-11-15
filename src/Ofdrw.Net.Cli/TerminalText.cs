using System.Globalization;
using System.Text;

/// <summary>Makes text that originates in an input file safe to print on one terminal line.</summary>
internal static class TerminalText
{
    /// <summary>Replaces control and bidirectional-format characters with visible \uXXXX escapes.
    /// Other text, including CJK, is returned unchanged.</summary>
    internal static string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        StringBuilder? builder = null;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (!NeedsEscape(character))
            {
                builder?.Append(character);
                continue;
            }

            builder ??= new StringBuilder(value, 0, index, value.Length + 8);
            builder.Append("\\u").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture));
        }
        return builder?.ToString() ?? value;
    }

    private static bool NeedsEscape(char c) =>
        c < 0x20 || c is >= (char)0x7F and <= (char)0x9F ||      // C0, DEL and C1
        c == 0x061C || c == 0x200E || c == 0x200F ||             // bidi marks
        c is >= (char)0x202A and <= (char)0x202E ||              // bidi embeddings and overrides
        c is >= (char)0x2066 and <= (char)0x2069 ||              // bidi isolates
        c == 0x2028 || c == 0x2029;                             // line and paragraph separators
}
