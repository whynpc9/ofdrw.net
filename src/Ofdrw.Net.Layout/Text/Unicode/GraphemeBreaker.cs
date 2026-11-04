using System.Collections.Generic;

namespace Ofdrw.Net.Layout.Text.Unicode;

/// <summary>UAX #29 (revision 47, Unicode 17.0.0) default extended grapheme cluster boundaries.</summary>
/// <remarks>No tailoring. Rules GB1–GB13 and GB999, including GB9c (Indic_Conjunct_Break)
/// and GB11 (emoji ZWJ sequences), are implemented on scalar values.</remarks>
internal static class GraphemeBreaker
{
    /// <summary>Returns boundary indexes in code points: 0, every interior boundary and Length.
    /// An empty input returns no boundaries (GB1/GB2 do not apply to empty text).</summary>
    internal static int[] Boundaries(int[] codePoints)
    {
        var result = new List<int>();
        if (codePoints.Length == 0) return result.ToArray();
        result.Add(0);
        var previous = UnicodeProperties.Grapheme(codePoints[0], out var previousIncb, out var previousPictographic);
        // GB9c: a consonant followed by any Extend/Linker sequence with at least one Linker.
        var conjunctConsonant = previousIncb == IndicConjunctBreak.Consonant;
        var conjunctLinker = false;
        // GB11: Extended_Pictographic Extend* ZWJ.
        var pictographicRun = previousPictographic;
        var pictographicZwj = false;
        var regionalIndicators = previous == GraphemeClass.RegionalIndicator ? 1 : 0;
        for (var index = 1; index < codePoints.Length; index++)
        {
            var current = UnicodeProperties.Grapheme(codePoints[index], out var incb, out var pictographic);
            if (!NoBreak(previous, current, incb, pictographic, conjunctConsonant && conjunctLinker,
                    pictographicZwj, regionalIndicators))
                result.Add(index);

            // Update GB9c state.
            if (incb == IndicConjunctBreak.Consonant) { conjunctConsonant = true; conjunctLinker = false; }
            else if (conjunctConsonant && incb == IndicConjunctBreak.Linker) conjunctLinker = true;
            else if (conjunctConsonant && incb == IndicConjunctBreak.Extend) { }
            else { conjunctConsonant = false; conjunctLinker = false; }
            // Update GB11 state.
            pictographicZwj = pictographicRun && current == GraphemeClass.ZWJ;
            if (pictographic) pictographicRun = true;
            else if (current != GraphemeClass.Extend) pictographicRun = false;
            // GB12/GB13 state.
            regionalIndicators = current == GraphemeClass.RegionalIndicator ? regionalIndicators + 1 : 0;
            previous = current;
        }
        result.Add(codePoints.Length);
        return result.ToArray();
    }

    private static bool NoBreak(GraphemeClass previous, GraphemeClass current, IndicConjunctBreak currentIncb,
        bool currentPictographic, bool conjunctLinked, bool pictographicZwj, int regionalIndicators)
    {
        if (previous == GraphemeClass.CR && current == GraphemeClass.LF) return true; // GB3
        if (previous is GraphemeClass.Control or GraphemeClass.CR or GraphemeClass.LF) return false; // GB4
        if (current is GraphemeClass.Control or GraphemeClass.CR or GraphemeClass.LF) return false; // GB5
        if (previous == GraphemeClass.L && current is GraphemeClass.L or GraphemeClass.V or GraphemeClass.LV or
                GraphemeClass.LVT) return true; // GB6
        if (previous is GraphemeClass.LV or GraphemeClass.V && current is GraphemeClass.V or GraphemeClass.T)
            return true; // GB7
        if (previous is GraphemeClass.LVT or GraphemeClass.T && current == GraphemeClass.T) return true; // GB8
        if (current is GraphemeClass.Extend or GraphemeClass.ZWJ) return true; // GB9
        if (current == GraphemeClass.SpacingMark) return true; // GB9a
        if (previous == GraphemeClass.Prepend) return true; // GB9b
        if (conjunctLinked && currentIncb == IndicConjunctBreak.Consonant) return true; // GB9c
        if (pictographicZwj && currentPictographic) return true; // GB11
        if (previous == GraphemeClass.RegionalIndicator && current == GraphemeClass.RegionalIndicator)
            return regionalIndicators % 2 == 1; // GB12, GB13
        return false; // GB999
    }
}
