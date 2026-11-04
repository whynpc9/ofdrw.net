using System;
using System.Collections.Generic;

namespace Ofdrw.Net.Layout.Text.Unicode;

internal enum BidiClass : byte
{
    L, R, AL, EN, ES, ET, AN, CS, NSM, BN, B, S, WS, ON,
    LRE, LRO, RLE, RLO, PDF, LRI, RLI, FSI, PDI,
}

// Order matches UnicodeData.LineBreakClassNames (generated).
internal enum LineBreakClass : byte
{
    XX, BK, CR, LF, CM, NL, SG, WJ, ZW, GL, SP, ZWJ, B2, BA, BB, HY, HH, CB, CL, CP, EX, IN, NS,
    OP, QU, IS, NU, PO, PR, SY, AI, AK, AL, AP, AS, CJ, EB, EM, H2, H3, HL, ID, JL, JV, JT, RI,
    SA, VF, VI,
}

// Order matches UnicodeData.GraphemeClassNames (generated).
internal enum GraphemeClass : byte
{
    Other, CR, LF, Control, Extend, ZWJ, RegionalIndicator, Prepend, SpacingMark, L, V, T, LV, LVT,
}

internal enum IndicConjunctBreak : byte { None, Consonant, Linker, Extend }

[Flags]
internal enum LineBreakFlags : ushort
{
    None = 0,
    InitialPunctuation = 64,
    FinalPunctuation = 128,
    EastAsian = 256,
    UnassignedPictographic = 512,
    Mark = 1024,
}

/// <summary>Range lookups into the frozen, generated Unicode 17.0.0 tables.</summary>
internal static class UnicodeProperties
{
    private static readonly Dictionary<int, (int Pair, bool Opening)> Brackets = CreateBrackets();

    static UnicodeProperties()
    {
        // Fail at type initialization rather than silently mis-classify text when
        // a generated enum ordering drifts from the data tables.
        Verify(UnicodeData.BidiClassNames, typeof(BidiClass));
        Verify(UnicodeData.LineBreakClassNames, typeof(LineBreakClass));
        if (UnicodeData.GraphemeClassNames.Length != Enum.GetValues(typeof(GraphemeClass)).Length)
            throw new InvalidOperationException("Generated grapheme classes do not match the enum.");
    }

    internal static BidiClass Bidi(int codePoint) =>
        (BidiClass)UnicodeData.BidiValues[Find(UnicodeData.BidiStarts, codePoint)];

    internal static LineBreakClass LineBreak(int codePoint, out LineBreakFlags flags)
    {
        var value = UnicodeData.LineBreakValues[Find(UnicodeData.LineBreakStarts, codePoint)];
        flags = (LineBreakFlags)(value & ~63);
        return (LineBreakClass)(value & 63);
    }

    internal static GraphemeClass Grapheme(int codePoint, out IndicConjunctBreak incb, out bool extendedPictographic)
    {
        var value = UnicodeData.GraphemeValues[Find(UnicodeData.GraphemeStarts, codePoint)];
        incb = (IndicConjunctBreak)((value >> 4) & 3);
        extendedPictographic = (value & 64) != 0;
        return (GraphemeClass)(value & 15);
    }

    /// <summary>ISO 15924 script code (Zyyy common, Zinh inherited, Zzzz unknown).</summary>
    internal static string Script(int codePoint) =>
        UnicodeData.ScriptCodes[UnicodeData.ScriptValues[Find(UnicodeData.ScriptStarts, codePoint)]];

    internal static bool TryGetBracket(int codePoint, out int pair, out bool opening)
    {
        if (Brackets.TryGetValue(codePoint, out var value)) { pair = value.Pair; opening = value.Opening; return true; }
        pair = 0; opening = false;
        return false;
    }

    /// <summary>BD16 compares brackets by canonical equivalence; only these two pairs differ.</summary>
    internal static int CanonicalBracket(int codePoint) => codePoint switch
    {
        0x2329 => 0x3008,
        0x232A => 0x3009,
        _ => codePoint,
    };

    private static Dictionary<int, (int, bool)> CreateBrackets()
    {
        var result = new Dictionary<int, (int, bool)>(UnicodeData.BracketCodes.Length);
        for (var index = 0; index < UnicodeData.BracketCodes.Length; index++)
            result[UnicodeData.BracketCodes[index]] = (UnicodeData.BracketPairs[index], UnicodeData.BracketOpening[index] != 0);
        return result;
    }

    private static int Find(int[] starts, int codePoint)
    {
        if (codePoint < 0 || codePoint > 0x10ffff) throw new ArgumentOutOfRangeException(nameof(codePoint));
        int low = 0, high = starts.Length - 1;
        while (low < high)
        {
            var middle = (low + high + 1) >> 1;
            if (starts[middle] <= codePoint) low = middle; else high = middle - 1;
        }
        return low;
    }

    private static void Verify(string[] names, Type type)
    {
        var values = Enum.GetNames(type);
        if (values.Length != names.Length)
            throw new InvalidOperationException("Generated Unicode classes do not match " + type.Name + ".");
        for (var index = 0; index < names.Length; index++)
            if (values[index] != names[index])
                throw new InvalidOperationException("Generated Unicode class order differs for " + type.Name + ".");
    }
}

/// <summary>A UTF-16 string decoded to scalar values with their UTF-16 offsets.</summary>
internal sealed class CodePointText
{
    internal CodePointText(string text)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        var points = new List<int>(text.Length);
        var offsets = new List<int>(text.Length + 1);
        for (var index = 0; index < text.Length; index++)
        {
            offsets.Add(index);
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]))
            {
                points.Add(char.ConvertToUtf32(text[index], text[index + 1]));
                index++;
            }
            else if (char.IsSurrogate(text[index]))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                    sourceStart: index, sourceLength: 1);
            else points.Add(text[index]);
        }
        offsets.Add(text.Length);
        CodePoints = points.ToArray();
        Offsets = offsets.ToArray();
    }

    internal CodePointText(int[] codePoints)
    {
        CodePoints = codePoints;
        Offsets = new int[codePoints.Length + 1];
        var builder = new System.Text.StringBuilder(codePoints.Length);
        for (var index = 0; index < codePoints.Length; index++)
        {
            Offsets[index] = builder.Length;
            builder.Append(char.ConvertFromUtf32(codePoints[index]));
        }
        Offsets[codePoints.Length] = builder.Length;
        Text = builder.ToString();
    }

    internal string Text { get; }
    internal int[] CodePoints { get; }
    /// <summary>UTF-16 offset of each code point, plus the text length at the end.</summary>
    internal int[] Offsets { get; }
    internal int Length => CodePoints.Length;

    internal int IndexOfOffset(int utf16Offset)
    {
        var index = Array.BinarySearch(Offsets, utf16Offset);
        if (index < 0) throw new ArgumentException("Offset splits a surrogate pair.", nameof(utf16Offset));
        return index;
    }
}
