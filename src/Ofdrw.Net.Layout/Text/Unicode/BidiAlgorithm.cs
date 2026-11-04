using System;
using System.Collections.Generic;

namespace Ofdrw.Net.Layout.Text.Unicode;

/// <summary>UAX #9 (revision 51, Unicode 17.0.0) Unicode Bidirectional Algorithm through rule L2.</summary>
/// <remarks>Implements P2–P3, X1–X10, W1–W7, N0–N2, I1–I2, L1 and L2 for one paragraph. P1 paragraph
/// splitting is the caller's responsibility. L3/L4 (combining-mark reordering, mirroring) are not
/// performed: mirroring is left to the run shaper. Characters removed by X9 report level
/// <see cref="Removed"/> from <see cref="Resolve"/>.</remarks>
internal sealed class BidiAlgorithm
{
    internal const byte Removed = 255;
    private const int MaxDepth = 125;

    private readonly BidiClass[] _initial;
    private readonly int[]? _codePoints;
    private readonly BidiClass[] _types;
    private readonly byte[] _levels;
    private readonly int[] _matchingPdi;
    private readonly int[] _matchingIsolateInitiator;
    private readonly int _length;
    private byte[] _explicitLevels = Array.Empty<byte>();

    private BidiAlgorithm(BidiClass[] types, int[]? codePoints)
    {
        _initial = (BidiClass[])types.Clone();
        _codePoints = codePoints;
        _types = (BidiClass[])types.Clone();
        _length = types.Length;
        _levels = new byte[_length];
        _matchingPdi = new int[_length];
        _matchingIsolateInitiator = new int[_length];
    }

    internal byte ParagraphLevel { get; private set; }

    /// <summary>Resolves embedding levels for one paragraph.</summary>
    /// <param name="types">Bidi classes of the paragraph's code points.</param>
    /// <param name="codePoints">Code points used for bracket pairs (N0), or null to treat no
    /// character as a paired bracket.</param>
    /// <param name="paragraphLevel">0 (LTR), 1 (RTL) or -1 for P2/P3 auto-LTR detection.</param>
    internal static BidiAlgorithm Resolve(BidiClass[] types, int[]? codePoints, int paragraphLevel)
    {
        if (types is null) throw new ArgumentNullException(nameof(types));
        if (codePoints is not null && codePoints.Length != types.Length)
            throw new ArgumentException("Code points and classes differ in length.", nameof(codePoints));
        if (paragraphLevel is < -1 or > 1) throw new ArgumentOutOfRangeException(nameof(paragraphLevel));
        var algorithm = new BidiAlgorithm(types, codePoints);
        algorithm.Run(paragraphLevel);
        return algorithm;
    }

    /// <summary>Resolved levels before line-level rules; X9-removed characters are <see cref="Removed"/>.</summary>
    internal byte[] Levels => (byte[])_levels.Clone();

    /// <summary>Applies L1 to [start, end) as one line and returns its levels, giving every
    /// X9-removed character the level of the preceding character (or the paragraph level).</summary>
    internal byte[] LineLevels(int start, int end, bool keepRemovedMarker = false)
    {
        if (start < 0 || end > _length || start > end) throw new ArgumentOutOfRangeException(nameof(start));
        var result = new byte[end - start];
        for (var index = start; index < end; index++) result[index - start] = _levels[index];
        // L1: segment/paragraph separators, and whitespace/isolate formatting before them or the line end.
        var trailing = true;
        for (var index = end - 1; index >= start; index--)
        {
            var type = _initial[index];
            if (type is BidiClass.B or BidiClass.S)
            {
                result[index - start] = ParagraphLevel;
                trailing = true;
            }
            else if (type is BidiClass.WS or BidiClass.FSI or BidiClass.LRI or BidiClass.RLI or BidiClass.PDI ||
                     IsRemovedByX9(type))
            {
                if (trailing) result[index - start] = ParagraphLevel;
            }
            else trailing = false;
        }
        if (!keepRemovedMarker)
        {
            var prior = ParagraphLevel;
            for (var index = start; index < end; index++)
            {
                if (IsRemovedByX9(_initial[index]))
                {
                    if (result[index - start] == Removed) result[index - start] = prior;
                }
                else prior = result[index - start];
            }
        }
        else
        {
            for (var index = start; index < end; index++)
                if (IsRemovedByX9(_initial[index])) result[index - start] = Removed;
        }
        return result;
    }

    /// <summary>L2: returns logical indexes in visual left-to-right order. Entries whose level is
    /// <see cref="Removed"/> are omitted.</summary>
    internal static int[] VisualOrder(byte[] levels)
    {
        var indexes = new List<int>(levels.Length);
        var kept = new List<byte>(levels.Length);
        for (var index = 0; index < levels.Length; index++)
        {
            if (levels[index] == Removed) continue;
            indexes.Add(index); kept.Add(levels[index]);
        }
        var order = indexes.ToArray();
        var lineLevels = kept.ToArray();
        byte highest = 0, lowestOdd = 255;
        foreach (var level in lineLevels)
        {
            if (level > highest) highest = level;
            if ((level & 1) != 0 && level < lowestOdd) lowestOdd = level;
        }
        for (int level = highest; level >= lowestOdd && level > 0; level--)
        {
            for (var index = 0; index < order.Length;)
            {
                if (lineLevels[index] < level) { index++; continue; }
                var end = index;
                while (end < order.Length && lineLevels[end] >= level) end++;
                Array.Reverse(order, index, end - index);
                Array.Reverse(lineLevels, index, end - index);
                index = end;
            }
        }
        return order;
    }

    internal static bool IsRemovedByX9(BidiClass type) => type is BidiClass.RLE or BidiClass.LRE or
        BidiClass.RLO or BidiClass.LRO or BidiClass.PDF or BidiClass.BN;

    private static bool IsIsolateInitiator(BidiClass type) => type is BidiClass.LRI or BidiClass.RLI or BidiClass.FSI;

    private void Run(int paragraphLevel)
    {
        DetermineMatchingIsolates();
        ParagraphLevel = paragraphLevel >= 0 ? (byte)paragraphLevel : (byte)FirstStrongLevel(0, _length, 0);
        DetermineExplicitLevels();
        // Every sequence's sos/eos depends on explicit levels, not on levels already resolved
        // by I1/I2 in another sequence.
        _explicitLevels = (byte[])_levels.Clone();
        foreach (var sequence in IsolatingRunSequences())
            ResolveSequence(sequence);
        for (var index = 0; index < _length; index++)
            if (IsRemovedByX9(_initial[index])) _levels[index] = Removed;
    }

    private void DetermineMatchingIsolates()
    {
        for (var index = 0; index < _length; index++) { _matchingPdi[index] = -1; _matchingIsolateInitiator[index] = -1; }
        var stack = new Stack<int>();
        for (var index = 0; index < _length; index++)
        {
            var type = _types[index];
            if (IsIsolateInitiator(type)) stack.Push(index);
            else if (type == BidiClass.PDI && stack.Count > 0)
            {
                var opener = stack.Pop();
                _matchingPdi[opener] = index;
                _matchingIsolateInitiator[index] = opener;
            }
            else if (type == BidiClass.B) stack.Clear();
        }
    }

    // P2/P3 over [start, end); returns 0/1, or fallback if no strong character is found.
    private int FirstStrongLevel(int start, int end, int fallback)
    {
        for (var index = start; index < end; index++)
        {
            var type = _types[index];
            if (type == BidiClass.L) return 0;
            if (type is BidiClass.R or BidiClass.AL) return 1;
            if (IsIsolateInitiator(type))
            {
                if (_matchingPdi[index] < 0) return fallback;
                index = _matchingPdi[index];
            }
            else if (type == BidiClass.B) return fallback;
        }
        return fallback;
    }

    private struct Status
    {
        internal byte Level;
        internal BidiClass Override; // ON = neutral
        internal bool Isolate;
    }

    private void DetermineExplicitLevels()
    {
        var stack = new Status[MaxDepth + 2];
        var depth = 0;
        stack[depth++] = new Status { Level = ParagraphLevel, Override = BidiClass.ON, Isolate = false };
        int overflowIsolates = 0, overflowEmbeddings = 0, validIsolates = 0;
        for (var index = 0; index < _length; index++)
        {
            var type = _types[index];
            var top = stack[depth - 1];
            switch (type)
            {
                case BidiClass.RLE:
                case BidiClass.LRE:
                case BidiClass.RLO:
                case BidiClass.LRO:
                {
                    var rtl = type is BidiClass.RLE or BidiClass.RLO;
                    var level = rtl ? (top.Level + 1) | 1 : (top.Level + 2) & ~1;
                    _levels[index] = top.Level;
                    if (level <= MaxDepth && overflowIsolates == 0 && overflowEmbeddings == 0)
                        stack[depth++] = new Status { Level = (byte)level, Isolate = false,
                            Override = type == BidiClass.RLO ? BidiClass.R : type == BidiClass.LRO ? BidiClass.L : BidiClass.ON };
                    else if (overflowIsolates == 0) overflowEmbeddings++;
                    break;
                }
                case BidiClass.RLI:
                case BidiClass.LRI:
                case BidiClass.FSI:
                {
                    _levels[index] = top.Level;
                    if (top.Override != BidiClass.ON) _types[index] = top.Override;
                    var rtl = type == BidiClass.RLI;
                    if (type == BidiClass.FSI)
                    {
                        var end = _matchingPdi[index] < 0 ? _length : _matchingPdi[index];
                        rtl = FirstStrongLevel(index + 1, end, 0) == 1;
                    }
                    var level = rtl ? (top.Level + 1) | 1 : (top.Level + 2) & ~1;
                    if (level <= MaxDepth && overflowIsolates == 0 && overflowEmbeddings == 0)
                    {
                        validIsolates++;
                        stack[depth++] = new Status { Level = (byte)level, Override = BidiClass.ON, Isolate = true };
                    }
                    else overflowIsolates++;
                    break;
                }
                case BidiClass.PDI:
                {
                    if (overflowIsolates > 0) overflowIsolates--;
                    else if (validIsolates > 0)
                    {
                        overflowEmbeddings = 0;
                        while (!stack[depth - 1].Isolate) depth--;
                        depth--;
                        validIsolates--;
                    }
                    top = stack[depth - 1];
                    _levels[index] = top.Level;
                    if (top.Override != BidiClass.ON) _types[index] = top.Override;
                    break;
                }
                case BidiClass.PDF:
                {
                    _levels[index] = top.Level;
                    if (overflowIsolates > 0) { }
                    else if (overflowEmbeddings > 0) overflowEmbeddings--;
                    else if (!top.Isolate && depth >= 2) depth--;
                    break;
                }
                case BidiClass.B:
                    _levels[index] = ParagraphLevel;
                    break;
                case BidiClass.BN:
                    _levels[index] = top.Level;
                    break;
                default:
                    _levels[index] = top.Level;
                    if (top.Override != BidiClass.ON) _types[index] = top.Override;
                    break;
            }
        }
    }

    private List<int[]> IsolatingRunSequences()
    {
        // BD13 level runs over characters not removed by X9.
        var runs = new List<List<int>>();
        List<int>? current = null;
        var currentLevel = -1;
        for (var index = 0; index < _length; index++)
        {
            if (IsRemovedByX9(_initial[index])) continue;
            if (current is null || _levels[index] != currentLevel)
            {
                current = new List<int>();
                runs.Add(current);
                currentLevel = _levels[index];
            }
            current.Add(index);
        }
        var runOfCharacter = new Dictionary<int, int>();
        for (var run = 0; run < runs.Count; run++) runOfCharacter[runs[run][0]] = run;
        var sequences = new List<int[]>();
        for (var run = 0; run < runs.Count; run++)
        {
            var first = runs[run][0];
            // A run starting with a PDI that matches an initiator continues an earlier sequence.
            if (_initial[first] == BidiClass.PDI && _matchingIsolateInitiator[first] >= 0) continue;
            var sequence = new List<int>();
            var cursor = run;
            while (true)
            {
                sequence.AddRange(runs[cursor]);
                var last = runs[cursor][runs[cursor].Count - 1];
                if (!IsIsolateInitiator(_initial[last]) || _matchingPdi[last] < 0) break;
                var pdi = _matchingPdi[last];
                if (!runOfCharacter.TryGetValue(pdi, out cursor)) break;
            }
            sequences.Add(sequence.ToArray());
        }
        return sequences;
    }

    private int PrecedingLevel(int index)
    {
        for (var prior = index - 1; prior >= 0; prior--)
            if (!IsRemovedByX9(_initial[prior])) return _explicitLevels[prior];
        return ParagraphLevel;
    }

    private int FollowingLevel(int index)
    {
        for (var next = index + 1; next < _length; next++)
            if (!IsRemovedByX9(_initial[next])) return _explicitLevels[next];
        return ParagraphLevel;
    }

    private static bool IsStrongForNeutral(BidiClass type) => type is BidiClass.L or BidiClass.R or
        BidiClass.EN or BidiClass.AN;

    private static BidiClass AsStrong(BidiClass type) => type == BidiClass.L ? BidiClass.L : BidiClass.R;

    private static bool IsNeutralOrIsolate(BidiClass type) => type is BidiClass.B or BidiClass.S or
        BidiClass.WS or BidiClass.ON or BidiClass.LRI or BidiClass.RLI or BidiClass.FSI or BidiClass.PDI;

    private void ResolveSequence(int[] sequence)
    {
        var count = sequence.Length;
        var level = _explicitLevels[sequence[0]];
        var sos = Math.Max(PrecedingLevel(sequence[0]), level) % 2 == 1 ? BidiClass.R : BidiClass.L;
        var last = sequence[count - 1];
        // A sequence can only end in an isolate initiator when it has no matching PDI here.
        var eosLevel = IsIsolateInitiator(_initial[last]) ? ParagraphLevel : FollowingLevel(last);
        var eos = Math.Max(eosLevel, level) % 2 == 1 ? BidiClass.R : BidiClass.L;
        var types = new BidiClass[count];
        for (var index = 0; index < count; index++) types[index] = _types[sequence[index]];

        // W1
        for (var index = 0; index < count; index++)
        {
            if (types[index] != BidiClass.NSM) continue;
            if (index == 0) types[index] = sos;
            else types[index] = IsIsolateInitiator(types[index - 1]) || types[index - 1] == BidiClass.PDI
                ? BidiClass.ON : types[index - 1];
        }
        // W2 and W3
        var lastStrong = sos;
        for (var index = 0; index < count; index++)
        {
            var type = types[index];
            if (type is BidiClass.L or BidiClass.R or BidiClass.AL) lastStrong = type;
            else if (type == BidiClass.EN && lastStrong == BidiClass.AL) types[index] = BidiClass.AN;
        }
        for (var index = 0; index < count; index++) if (types[index] == BidiClass.AL) types[index] = BidiClass.R;
        // W4
        for (var index = 1; index < count - 1; index++)
        {
            var type = types[index];
            if (type == BidiClass.ES && types[index - 1] == BidiClass.EN && types[index + 1] == BidiClass.EN)
                types[index] = BidiClass.EN;
            else if (type == BidiClass.CS && types[index - 1] == BidiClass.EN && types[index + 1] == BidiClass.EN)
                types[index] = BidiClass.EN;
            else if (type == BidiClass.CS && types[index - 1] == BidiClass.AN && types[index + 1] == BidiClass.AN)
                types[index] = BidiClass.AN;
        }
        // W5
        for (var index = 0; index < count; index++)
        {
            if (types[index] != BidiClass.ET) continue;
            var end = index;
            while (end < count && types[end] == BidiClass.ET) end++;
            var adjacentEn = index > 0 && types[index - 1] == BidiClass.EN || end < count && types[end] == BidiClass.EN;
            if (adjacentEn) for (var j = index; j < end; j++) types[j] = BidiClass.EN;
            index = end - 1;
        }
        // W6
        for (var index = 0; index < count; index++)
            if (types[index] is BidiClass.ES or BidiClass.ET or BidiClass.CS) types[index] = BidiClass.ON;
        // W7
        lastStrong = sos;
        for (var index = 0; index < count; index++)
        {
            var type = types[index];
            if (type is BidiClass.L or BidiClass.R) lastStrong = type;
            else if (type == BidiClass.EN && lastStrong == BidiClass.L) types[index] = BidiClass.L;
        }

        var embedding = level % 2 == 1 ? BidiClass.R : BidiClass.L;
        ResolvePairedBrackets(sequence, types, sos, embedding);

        // N1 and N2
        for (var index = 0; index < count; index++)
        {
            if (!IsNeutralOrIsolate(types[index])) continue;
            var end = index;
            while (end < count && IsNeutralOrIsolate(types[end])) end++;
            var before = index == 0 ? sos : AsStrong(types[index - 1]);
            var after = end == count ? eos : AsStrong(types[end]);
            var resolved = before == after ? before : embedding;
            for (var j = index; j < end; j++) types[j] = resolved;
            index = end - 1;
        }

        // I1 and I2
        for (var index = 0; index < count; index++)
        {
            var type = types[index];
            var target = sequence[index];
            if (level % 2 == 0)
            {
                if (type == BidiClass.R) _levels[target] = (byte)(level + 1);
                else if (type is BidiClass.AN or BidiClass.EN) _levels[target] = (byte)(level + 2);
            }
            else if (type is BidiClass.L or BidiClass.EN or BidiClass.AN) _levels[target] = (byte)(level + 1);
        }
    }

    private void ResolvePairedBrackets(int[] sequence, BidiClass[] types, BidiClass sos, BidiClass embedding)
    {
        if (_codePoints is null) return;
        // BD16
        var pairs = new List<(int Open, int Close)>();
        var stack = new List<(int Bracket, int Position)>();
        for (var index = 0; index < sequence.Length; index++)
        {
            if (types[index] != BidiClass.ON) continue;
            var codePoint = _codePoints[sequence[index]];
            if (!UnicodeProperties.TryGetBracket(codePoint, out var pair, out var opening)) continue;
            if (opening)
            {
                if (stack.Count == 63) break;
                stack.Add((UnicodeProperties.CanonicalBracket(pair), index));
            }
            else
            {
                var canonical = UnicodeProperties.CanonicalBracket(codePoint);
                for (var depth = stack.Count - 1; depth >= 0; depth--)
                {
                    if (stack[depth].Bracket != canonical) continue;
                    pairs.Add((stack[depth].Position, index));
                    stack.RemoveRange(depth, stack.Count - depth);
                    break;
                }
            }
        }
        if (pairs.Count == 0) return;
        pairs.Sort((left, right) => left.Open.CompareTo(right.Open));
        var opposite = embedding == BidiClass.L ? BidiClass.R : BidiClass.L;
        foreach (var (open, close) in pairs)
        {
            var foundEmbedding = false;
            var foundOpposite = false;
            for (var index = open + 1; index < close; index++)
            {
                var type = types[index];
                if (!IsStrongForNeutral(type)) continue;
                var strong = AsStrong(type);
                if (strong == embedding) { foundEmbedding = true; break; }
                foundOpposite = true;
            }
            BidiClass? resolved = null;
            if (foundEmbedding) resolved = embedding;
            else if (foundOpposite)
            {
                var context = sos;
                for (var index = open - 1; index >= 0; index--)
                {
                    if (!IsStrongForNeutral(types[index])) continue;
                    context = AsStrong(types[index]);
                    break;
                }
                resolved = context == opposite ? opposite : embedding;
            }
            if (resolved is not { } value) continue;
            types[open] = value; types[close] = value;
            foreach (var bracket in new[] { open, close })
                for (var index = bracket + 1; index < sequence.Length && _initial[sequence[index]] == BidiClass.NSM; index++)
                    types[index] = value;
        }
    }
}
