using System.Collections.Generic;

namespace Ofdrw.Net.Layout.Text.Unicode;

internal enum LineBreakKind : byte
{
    /// <summary>No line break is allowed before this position.</summary>
    Prohibited,
    /// <summary>A line break opportunity.</summary>
    Allowed,
    /// <summary>A hard break: the line must end here.</summary>
    Mandatory,
}

/// <summary>UAX #14 (revision 55, Unicode 17.0.0) default line breaking, without tailoring.</summary>
/// <remarks>LB1 uses the documented default resolution (AI/SG/XX→AL, SA→CM for Mn/Mc else AL,
/// CJ→NS). No dictionary segmentation, hyphenation or language tailoring is applied.</remarks>
internal static class LineBreaker
{
    private struct Element
    {
        internal int Index;          // code point index of the base character
        internal LineBreakClass Class;
        internal LineBreakFlags Flags;
        internal int CodePoint;
    }

    /// <summary>Returns the break kind before each code point index 0..Length. Index 0 is always
    /// Prohibited (LB2) and index Length is Mandatory (LB3) for nonempty text.</summary>
    internal static LineBreakKind[] Analyze(int[] codePoints)
    {
        var count = codePoints.Length;
        var result = new LineBreakKind[count + 1];
        if (count == 0) return result;
        var raw = new LineBreakClass[count];
        var flags = new LineBreakFlags[count];
        for (var index = 0; index < count; index++)
        {
            var value = UnicodeProperties.LineBreak(codePoints[index], out var flag);
            // LB1 default resolution.
            value = value switch
            {
                LineBreakClass.AI or LineBreakClass.SG or LineBreakClass.XX => LineBreakClass.AL,
                LineBreakClass.SA => (flag & LineBreakFlags.Mark) != 0 ? LineBreakClass.CM : LineBreakClass.AL,
                LineBreakClass.CJ => LineBreakClass.NS,
                _ => value,
            };
            raw[index] = value; flags[index] = flag;
        }

        // LB9/LB10: fold combining sequences into their base. Absorbed characters never
        // start a new element, so no break is possible before them.
        var elements = new List<Element>(count);
        var absorbed = new bool[count];
        for (var index = 0; index < count; index++)
        {
            var value = raw[index];
            if (value is LineBreakClass.CM or LineBreakClass.ZWJ)
            {
                if (elements.Count > 0 && elements[elements.Count - 1].Class is not (LineBreakClass.BK or
                        LineBreakClass.CR or LineBreakClass.LF or LineBreakClass.NL or LineBreakClass.SP or
                        LineBreakClass.ZW))
                {
                    absorbed[index] = true;
                    continue;
                }
                // LB10: as U+0041 (AL, Lu, ea=Na, not pictographic).
                elements.Add(new Element { Index = index, Class = LineBreakClass.AL, Flags = LineBreakFlags.None,
                    CodePoint = codePoints[index] });
                continue;
            }
            elements.Add(new Element { Index = index, Class = value, Flags = flags[index], CodePoint = codePoints[index] });
        }

        for (var index = 1; index < count; index++) result[index] = LineBreakKind.Prohibited;
        result[0] = LineBreakKind.Prohibited; // LB2
        result[count] = LineBreakKind.Mandatory; // LB3
        // Linear-time lookbehind: for each element, the last element at or before it that is not SP,
        // not SY/IS, and the length of the regional-indicator run ending there. Every rule that
        // scans back over SP*, (SY|IS)* or RI runs reads these instead of rescanning the prefix.
        var count2 = elements.Count;
        var context = new Lookbehind(new int[count2], new int[count2], new int[count2]);
        for (var k = 0; k < count2; k++)
        {
            var value = elements[k].Class;
            context.NotSpace[k] = value == LineBreakClass.SP ? k > 0 ? context.NotSpace[k - 1] : -1 : k;
            context.NotSyIs[k] = value is LineBreakClass.SY or LineBreakClass.IS ? k > 0 ? context.NotSyIs[k - 1] : -1 : k;
            context.RegionalRun[k] = value == LineBreakClass.RI ? (k > 0 ? context.RegionalRun[k - 1] : 0) + 1 : 0;
        }
        for (var k = 1; k < elements.Count; k++)
        {
            var kind = Decide(elements, k, raw, context);
            result[elements[k].Index] = kind;
        }
        // Absorbed positions remain Prohibited (LB9), except that LB4/LB5 may still
        // force a break after a hard break (which never absorbs, see above).
        return result;
    }

    private static LineBreakClass At(List<Element> elements, int index) =>
        index >= 0 && index < elements.Count ? elements[index].Class : (LineBreakClass)255;

    private static bool IsAksara(Element element) =>
        element.Class is LineBreakClass.AK or LineBreakClass.AS || element.CodePoint == 0x25CC;

    private sealed class Lookbehind
    {
        internal Lookbehind(int[] notSpace, int[] notSyIs, int[] regionalRun)
        { NotSpace = notSpace; NotSyIs = notSyIs; RegionalRun = regionalRun; }
        internal int[] NotSpace { get; }
        internal int[] NotSyIs { get; }
        internal int[] RegionalRun { get; }
    }

    private static LineBreakKind Decide(List<Element> e, int k, LineBreakClass[] raw, Lookbehind context)
    {
        var left = e[k - 1];
        var right = e[k];
        var l = left.Class;
        var r = right.Class;
        var hasNext = k + 1 < e.Count;
        var next = hasNext ? e[k + 1] : default;
        var n = hasNext ? next.Class : (LineBreakClass)255;
        const LineBreakKind no = LineBreakKind.Prohibited, yes = LineBreakKind.Allowed, must = LineBreakKind.Mandatory;

        if (l == LineBreakClass.BK) return must; // LB4
        if (l == LineBreakClass.CR && r == LineBreakClass.LF) return no; // LB5
        if (l is LineBreakClass.CR or LineBreakClass.LF or LineBreakClass.NL) return must;
        if (r is LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or LineBreakClass.NL) return no; // LB6
        if (r is LineBreakClass.SP or LineBreakClass.ZW) return no; // LB7
        // LB8: ZW SP* ÷
        {
            var j = context.NotSpace[k - 1];
            if (j >= 0 && e[j].Class == LineBreakClass.ZW) return yes;
        }
        if (raw[right.Index - 1] == LineBreakClass.ZWJ) return no; // LB8a
        // LB9 and LB10 are applied while building elements.
        if (l == LineBreakClass.WJ || r == LineBreakClass.WJ) return no; // LB11
        if (l == LineBreakClass.GL) return no; // LB12
        if (r == LineBreakClass.GL && l is not (LineBreakClass.SP or LineBreakClass.BA or LineBreakClass.HY or
                LineBreakClass.HH)) return no; // LB12a
        if (r is LineBreakClass.CL or LineBreakClass.CP or LineBreakClass.EX or LineBreakClass.SY) return no; // LB13
        // LB14: OP SP* ×
        var beforeSpaces = context.NotSpace[k - 1];
        if (beforeSpaces >= 0 && e[beforeSpaces].Class == LineBreakClass.OP) return no;
        // LB15a: (sot | BK | CR | LF | NL | OP | QU | GL | SP | ZW) [\p{Pi}&QU] SP* ×
        if (beforeSpaces >= 0 && e[beforeSpaces].Class == LineBreakClass.QU &&
            (e[beforeSpaces].Flags & LineBreakFlags.InitialPunctuation) != 0)
        {
            var before = beforeSpaces - 1;
            if (before < 0 || e[before].Class is LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or
                    LineBreakClass.NL or LineBreakClass.OP or LineBreakClass.QU or LineBreakClass.GL or
                    LineBreakClass.SP or LineBreakClass.ZW)
                return no;
        }
        // LB15b: × [\p{Pf}&QU] ( SP | GL | WJ | CL | QU | CP | EX | IS | SY | BK | CR | LF | NL | ZW | eot)
        if (r == LineBreakClass.QU && (right.Flags & LineBreakFlags.FinalPunctuation) != 0 &&
            (!hasNext || n is LineBreakClass.SP or LineBreakClass.GL or LineBreakClass.WJ or LineBreakClass.CL or
                LineBreakClass.QU or LineBreakClass.CP or LineBreakClass.EX or LineBreakClass.IS or
                LineBreakClass.SY or LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or
                LineBreakClass.NL or LineBreakClass.ZW))
            return no;
        if (l == LineBreakClass.SP && r == LineBreakClass.IS && n == LineBreakClass.NU) return yes; // LB15c
        if (r == LineBreakClass.IS) return no; // LB15d
        // LB16: (CL | CP) SP* × NS
        if (r == LineBreakClass.NS && beforeSpaces >= 0 &&
            e[beforeSpaces].Class is LineBreakClass.CL or LineBreakClass.CP) return no;
        // LB17: B2 SP* × B2
        if (r == LineBreakClass.B2 && beforeSpaces >= 0 && e[beforeSpaces].Class == LineBreakClass.B2) return no;
        if (l == LineBreakClass.SP) return yes; // LB18
        // LB19
        if (r == LineBreakClass.QU && (right.Flags & LineBreakFlags.InitialPunctuation) == 0) return no;
        if (l == LineBreakClass.QU && (left.Flags & LineBreakFlags.FinalPunctuation) == 0) return no;
        // LB19a
        if (r == LineBreakClass.QU)
        {
            if ((left.Flags & LineBreakFlags.EastAsian) == 0) return no;
            if (!hasNext || (next.Flags & LineBreakFlags.EastAsian) == 0) return no;
        }
        if (l == LineBreakClass.QU)
        {
            if ((right.Flags & LineBreakFlags.EastAsian) == 0) return no;
            if (k - 2 < 0 || (e[k - 2].Flags & LineBreakFlags.EastAsian) == 0) return no;
        }
        if (l == LineBreakClass.CB || r == LineBreakClass.CB) return yes; // LB20
        // LB20a: ( sot | BK | CR | LF | NL | SP | ZW | CB | GL ) ( HY | HH ) × ( AL | HL )
        if (l is LineBreakClass.HY or LineBreakClass.HH && r is LineBreakClass.AL or LineBreakClass.HL)
        {
            var before = At(e, k - 2);
            if (k - 2 < 0 || before is LineBreakClass.BK or LineBreakClass.CR or LineBreakClass.LF or
                    LineBreakClass.NL or LineBreakClass.SP or LineBreakClass.ZW or LineBreakClass.CB or
                    LineBreakClass.GL)
                return no;
        }
        if (r is LineBreakClass.BA or LineBreakClass.HH or LineBreakClass.HY or LineBreakClass.NS) return no; // LB21
        if (l == LineBreakClass.BB) return no;
        if (l is LineBreakClass.HY or LineBreakClass.HH && At(e, k - 2) == LineBreakClass.HL &&
            r != LineBreakClass.HL) return no; // LB21a
        if (l == LineBreakClass.SY && r == LineBreakClass.HL) return no; // LB21b
        if (r == LineBreakClass.IN) return no; // LB22
        // LB23
        if (l is LineBreakClass.AL or LineBreakClass.HL && r == LineBreakClass.NU) return no;
        if (l == LineBreakClass.NU && r is LineBreakClass.AL or LineBreakClass.HL) return no;
        // LB23a
        if (l == LineBreakClass.PR && r is LineBreakClass.ID or LineBreakClass.EB or LineBreakClass.EM) return no;
        if (l is LineBreakClass.ID or LineBreakClass.EB or LineBreakClass.EM && r == LineBreakClass.PO) return no;
        // LB24
        if (l is LineBreakClass.PR or LineBreakClass.PO && r is LineBreakClass.AL or LineBreakClass.HL) return no;
        if (l is LineBreakClass.AL or LineBreakClass.HL && r is LineBreakClass.PR or LineBreakClass.PO) return no;
        // LB25
        if (r is LineBreakClass.PO or LineBreakClass.PR)
        {
            var j = k - 1;
            if (j >= 0 && e[j].Class is LineBreakClass.CL or LineBreakClass.CP) j--;
            if (j >= 0) j = context.NotSyIs[j];
            if (j >= 0 && e[j].Class == LineBreakClass.NU) return no;
        }
        if (l is LineBreakClass.PO or LineBreakClass.PR)
        {
            if (r == LineBreakClass.NU) return no;
            if (r == LineBreakClass.OP && n == LineBreakClass.NU) return no;
            if (r == LineBreakClass.OP && n == LineBreakClass.IS && At(e, k + 2) == LineBreakClass.NU) return no;
        }
        if (l is LineBreakClass.HY or LineBreakClass.IS && r == LineBreakClass.NU) return no;
        if (r == LineBreakClass.NU)
        {
            var j = k - 1;
            j = context.NotSyIs[j];
            if (j >= 0 && e[j].Class == LineBreakClass.NU) return no;
        }
        // LB26
        if (l == LineBreakClass.JL && r is LineBreakClass.JL or LineBreakClass.JV or LineBreakClass.H2 or
                LineBreakClass.H3) return no;
        if (l is LineBreakClass.JV or LineBreakClass.H2 && r is LineBreakClass.JV or LineBreakClass.JT) return no;
        if (l is LineBreakClass.JT or LineBreakClass.H3 && r == LineBreakClass.JT) return no;
        // LB27
        if (l is LineBreakClass.JL or LineBreakClass.JV or LineBreakClass.JT or LineBreakClass.H2 or
                LineBreakClass.H3 && r == LineBreakClass.PO) return no;
        if (l == LineBreakClass.PR && r is LineBreakClass.JL or LineBreakClass.JV or LineBreakClass.JT or
                LineBreakClass.H2 or LineBreakClass.H3) return no;
        // LB28
        if (l is LineBreakClass.AL or LineBreakClass.HL && r is LineBreakClass.AL or LineBreakClass.HL) return no;
        // LB28a
        if (l == LineBreakClass.AP && IsAksara(right)) return no;
        if (IsAksara(left) && r is LineBreakClass.VF or LineBreakClass.VI) return no;
        if (l == LineBreakClass.VI && k - 2 >= 0 && IsAksara(e[k - 2]) &&
            (r == LineBreakClass.AK || right.CodePoint == 0x25CC)) return no;
        if (IsAksara(left) && IsAksara(right) && n == LineBreakClass.VF) return no;
        if (l == LineBreakClass.IS && r is LineBreakClass.AL or LineBreakClass.HL) return no; // LB29
        // LB30
        if (l is LineBreakClass.AL or LineBreakClass.HL or LineBreakClass.NU && r == LineBreakClass.OP &&
            (right.Flags & LineBreakFlags.EastAsian) == 0) return no;
        if (l == LineBreakClass.CP && (left.Flags & LineBreakFlags.EastAsian) == 0 &&
            r is LineBreakClass.AL or LineBreakClass.HL or LineBreakClass.NU) return no;
        // LB30a
        if (l == LineBreakClass.RI && r == LineBreakClass.RI)
        {
            if (context.RegionalRun[k - 1] % 2 == 1) return no;
        }
        // LB30b
        if (r == LineBreakClass.EM && (l == LineBreakClass.EB ||
                (left.Flags & LineBreakFlags.UnassignedPictographic) != 0)) return no;
        return yes; // LB31
    }
}
