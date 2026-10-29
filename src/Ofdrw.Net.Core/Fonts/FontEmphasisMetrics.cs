namespace Ofdrw.Net.Core.Fonts;

/// <summary>Shared synthetic emphasis geometry for embedded regular faces.</summary>
internal static class FontEmphasisMetrics
{
    // The units belong to the caller (OFD millimeters or PDF points). Dense Han
    // counters close under the Latin 2.5%-em overlay at ordinary document sizes.
    // This is simulated weight, not a substitute for a physical bold face.
    internal static double BoldStrokeWidth(double fontSize) => fontSize * 0.025;

    internal static double BoldStrokeWidth(double fontSize, string text)
    {
        for (var index = 0; index < text.Length; index++)
        {
            int codePoint = text[index];
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]))
            {
                codePoint = char.ConvertToUtf32(text[index], text[index + 1]);
                index++;
            }
            if (codePoint is >= 0x3400 and <= 0x9FFF or >= 0xF900 and <= 0xFAFF or >= 0x20000 and <= 0x323AF)
                return fontSize * 0.01;
        }
        return BoldStrokeWidth(fontSize);
    }
}
