using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Unicode;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// F06 bounded visual-line RTL import (Issue #109). The evidence for each
// accepted line comes from three independent sources:
//   1. character mapping: every glyph's decoded Unicode value (ToUnicode or
//      font encoding) must exist, be BMP and belong to the admitted classes;
//   2. geometry: glyph origins and advances on one horizontal baseline define
//      the visual left-to-right order, independent of content-stream order;
//   3. logical order: a candidate logical text is derived from the visual
//      order and accepted only when the conformance-tested UAX #9 forward
//      reordering (Unicode 17, UnicodeTextAnalysis) reproduces exactly the
//      observed glyph permutation.
// GetWords() strings are never reordered: word segmentation is rebuilt from
// glyph gaps. Everything outside the profile is rejected, not guessed.
internal static class PdfRtlVisualLineImport
{
    // Tolerances are fractions of the larger glyph size involved (ems).
    private const double BaselineTolerance = 0.05;
    private const double NeighbourLineDistance = 0.6;
    private const double SpaceGap = 0.15;
    private const double MaximumGap = 3.0;
    private const double MinimumOriginStep = 0.5;
    private const double TrailingSpaceAdvance = 0.25;

    internal sealed class ImportedWord
    {
        internal ImportedWord(string text, double[] origins, double baseline, double left, double right,
            double bottom, double top, string? fontName, double pointSize)
        {
            Text = text; Origins = origins; Baseline = baseline; Left = left; Right = right;
            Bottom = bottom; Top = top; FontName = fontName; PointSize = pointSize;
        }

        /// <summary>Logical word followed by one space, like the legacy word layer.</summary>
        internal string Text { get; }
        /// <summary>PDF user-space X origin of each UTF-16 unit of <see cref="Text"/>.</summary>
        internal double[] Origins { get; }
        internal double Baseline { get; }
        internal double Left { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal double Top { get; }
        internal string? FontName { get; }
        internal double PointSize { get; }
    }

    internal sealed class ImportedPage
    {
        internal ImportedPage(IReadOnlyList<IReadOnlyList<ImportedWord>> lines, ISet<Letter> consumed)
        { Lines = lines; Consumed = consumed; }
        /// <summary>Right-to-left lines, top to bottom, each in logical word order.</summary>
        internal IReadOnlyList<IReadOnlyList<ImportedWord>> Lines { get; }
        /// <summary>Letters owned by those lines; the legacy word layer must not repeat them.</summary>
        internal ISet<Letter> Consumed { get; }
    }

    private sealed class Unit
    {
        internal Letter? Letter;
        internal string Text = string.Empty;
        internal bool Rtl;
        internal double Start, End, Em;
    }

    private sealed class ReferenceComparer : IEqualityComparer<Letter>
    {
        internal static readonly ReferenceComparer Instance = new();
        public bool Equals(Letter? x, Letter? y) => ReferenceEquals(x, y);
        public int GetHashCode(Letter obj) => RuntimeHelpers.GetHashCode(obj);
    }

    internal static ImportedPage Import(PdfDocument document, Page page, PdfTextBaseDirection direction, int maximumWords,
        CancellationToken cancellationToken)
    {
        var number = page.Number;
        if (page.Rotation.Value != 0)
            throw PdfImportReadingOrder.ProfileRejection(number, $"the page is rotated by {page.Rotation.Value} degrees.");
        RejectSubstitutedText(page.GetMarkedContents(), number, cancellationToken);

        var horizontal = new List<(Letter Letter, int Order)>();
        var order = 0;
        foreach (var letter in page.Letters)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var value = letter.Value ?? string.Empty;
            if (value.Length > 0 && value.All(char.IsWhiteSpace)) { order++; continue; }
            if (letter.TextOrientation != TextOrientation.Horizontal)
            {
                if (PdfImportReadingOrder.ContainsRightToLeft(value))
                    throw PdfImportReadingOrder.ProfileRejection(number, "right-to-left text is rotated or vertical.");
                order++;
                continue;
            }
            horizontal.Add((letter, order++));
        }

        // Baseline clusters, top to bottom (PDF Y grows upwards).
        var sorted = horizontal.OrderByDescending(item => item.Letter.StartBaseLine.Y).ThenBy(item => item.Order).ToList();
        var clusters = new List<List<(Letter Letter, int Order)>>();
        var clusterY = new List<double>();
        var clusterEm = new List<double>();
        foreach (var item in sorted)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var em = Em(item.Letter);
            var y = item.Letter.StartBaseLine.Y;
            var last = clusters.Count - 1;
            if (last >= 0 && Math.Abs(clusterY[last] - y) <= BaselineTolerance * Math.Max(clusterEm[last], em))
            {
                clusters[last].Add(item);
                clusterEm[last] = Math.Max(clusterEm[last], em);
            }
            else
            {
                clusters.Add(new List<(Letter, int)> { item });
                clusterY.Add(y);
                clusterEm.Add(em);
            }
        }

        PdfPaintedTextEvidence? painted = null;
        var consumed = new HashSet<Letter>(ReferenceComparer.Instance);
        var lines = new List<IReadOnlyList<ImportedWord>>();
        var words = 0;
        for (var index = 0; index < clusters.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!clusters[index].Any(item => PdfImportReadingOrder.ContainsRightToLeft(item.Letter.Value ?? string.Empty)))
                continue;
            for (var other = 0; other < clusters.Count; other++)
            {
                if (other == index) continue;
                if (Math.Abs(clusterY[other] - clusterY[index]) < NeighbourLineDistance * Math.Max(clusterEm[other], clusterEm[index]))
                    throw PdfImportReadingOrder.ProfileRejection(number,
                        "right-to-left text shares its line band with glyphs on a different baseline (superscript, overlap or ruby).");
            }
            // Only directly painted, unclipped glyphs are geometry evidence.
            painted ??= PdfPaintedTextEvidence.Scan(document, page, cancellationToken);
            foreach (var item in clusters[index])
                if (painted.Reject(item.Letter) is { } reason)
                    throw PdfImportReadingOrder.ProfileRejection(number, $"right-to-left line {lines.Count + 1}: {reason}");
            var line = ImportLine(clusters[index], clusterY[index], direction, number, lines.Count + 1, cancellationToken);
            words += line.Count;
            if (words > maximumWords)
                throw new System.IO.InvalidDataException(
                    $"PDF page {number} contains more than {maximumWords} text objects, which exceeds the configured limit.");
            lines.Add(line);
            foreach (var item in clusters[index]) consumed.Add(item.Letter);
        }
        return new ImportedPage(lines, consumed);
    }

    private static void RejectSubstitutedText(IEnumerable<MarkedContentElement> contents, int number,
        CancellationToken cancellationToken)
    {
        var pending = new Stack<MarkedContentElement>(contents);
        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var content = pending.Pop();
            if (content.ActualText is not null)
                throw PdfImportReadingOrder.ProfileRejection(number,
                    "the page carries /ActualText marked content, whose substituted logical text the profile does not verify.");
            if (string.Equals(content.Tag, "ReversedChars", StringComparison.Ordinal))
                throw PdfImportReadingOrder.ProfileRejection(number, "the page carries /ReversedChars marked content.");
            if (string.Equals(content.Tag, "OC", StringComparison.Ordinal))
                throw PdfImportReadingOrder.ProfileRejection(number,
                    "the page carries optional content, whose visibility the profile does not verify.");
            foreach (var child in content.Children) pending.Push(child);
        }
    }

    private static IReadOnlyList<ImportedWord> ImportLine(List<(Letter Letter, int Order)> cluster, double baseline,
        PdfTextBaseDirection direction, int number, int lineNumber, CancellationToken cancellationToken)
    {
        string Where() => $"right-to-left line {lineNumber}";
        var glyphs = cluster.OrderBy(item => item.Letter.StartBaseLine.X).ThenBy(item => item.Order).ToList();
        var units = new List<Unit>();
        Unit? previous = null;
        foreach (var (letter, _) in glyphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var value = letter.Value ?? string.Empty;
            if (value.Length == 0)
                throw PdfImportReadingOrder.ProfileRejection(number, $"{Where()} has a glyph without a Unicode mapping.");
            var rtl = Classify(value, number, Where());
            var em = Em(letter);
            var start = letter.StartBaseLine.X;
            var end = letter.EndBaseLine.X;
            if (!(end - start > 0) || Math.Abs(letter.EndBaseLine.Y - letter.StartBaseLine.Y) > BaselineTolerance * em)
                throw PdfImportReadingOrder.ProfileRejection(number, $"{Where()} has a glyph without a positive horizontal advance.");
            if (previous is not null)
            {
                var step = start - previous.Start;
                if (step < MinimumOriginStep * (previous.End - previous.Start))
                    throw PdfImportReadingOrder.ProfileRejection(number, $"{Where()} has overlapping glyphs, so its visual order is not unique.");
                var gap = start - previous.End;
                var gapEm = Math.Max(previous.Em, em);
                if (gap > MaximumGap * gapEm)
                    throw PdfImportReadingOrder.ProfileRejection(number,
                        $"{Where()} has a gap wider than {MaximumGap.ToString(CultureInfo.InvariantCulture)} em; columns cannot be told apart.");
                if (gap >= SpaceGap * gapEm)
                    units.Add(new Unit { Text = " ", Start = previous.End, End = start, Em = gapEm });
            }
            previous = new Unit { Letter = letter, Text = value, Rtl = rtl, Start = start, End = end, Em = em };
            units.Add(previous);
        }

        // Visual code units. A multi-character RTL glyph (for example a lam-alef
        // ligature) maps to logical characters; its visual contribution is reversed.
        var visual = new List<(char Char, int Unit)>();
        for (var index = 0; index < units.Count; index++)
        {
            var text = units[index].Text;
            if (units[index].Rtl)
                for (var offset = text.Length - 1; offset >= 0; offset--) visual.Add((text[offset], index));
            else
                foreach (var character in text) visual.Add((character, index));
        }

        var candidates = direction switch
        {
            PdfTextBaseDirection.LeftToRight => new[] { 0 },
            PdfTextBaseDirection.RightToLeft => new[] { 1 },
            _ => new[] { 0, 1 },
        };
        int[]? accepted = null;
        string? acceptedText = null;
        foreach (var level in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var logical = CandidateLogicalOrder(visual, units, level);
            var text = new string(logical.Select(position => visual[position].Char).ToArray());
            if (!Verify(text, logical, level, direction == PdfTextBaseDirection.Auto)) continue;
            if (acceptedText is not null && acceptedText != text)
                throw PdfImportReadingOrder.ProfileRejection(number,
                    $"{Where()} is consistent with both base directions; set RightToLeftBaseDirection explicitly.");
            accepted ??= logical;
            acceptedText ??= text;
        }
        if (accepted is null)
            throw PdfImportReadingOrder.ProfileRejection(number, direction == PdfTextBaseDirection.Auto
                ? $"{Where()} has no logical order whose own UAX #9 base direction reproduces the glyph order; set RightToLeftBaseDirection explicitly."
                : $"{Where()} has no logical order that UAX #9 reorders into the observed glyph order for the requested base direction.");

        // Logical words: split at space units, each followed by one space.
        var result = new List<ImportedWord>();
        var current = new List<int>();
        void Flush(int? spaceUnit)
        {
            if (current.Count == 0) return;
            var text = new char[current.Count + 1];
            var origins = new double[current.Count + 1];
            double left = double.MaxValue, right = double.MinValue, bottom = double.MaxValue, top = double.MinValue;
            double size = 0;
            string? font = null;
            for (var index = 0; index < current.Count; index++)
            {
                var (character, unitIndex) = visual[current[index]];
                var unit = units[unitIndex];
                text[index] = character;
                origins[index] = unit.Start;
                var letter = unit.Letter!;
                font ??= letter.FontName;
                size = Math.Max(size, unit.Em);
                var rect = letter.BoundingBox;
                left = Math.Min(left, Math.Min(unit.Start, rect.Left));
                right = Math.Max(right, Math.Max(unit.End, rect.Right));
                bottom = Math.Min(bottom, Math.Min(baseline - 0.2 * unit.Em, rect.Bottom));
                top = Math.Max(top, Math.Max(baseline + 0.8 * unit.Em, rect.Top));
            }
            var lastUnit = units[visual[current[current.Count - 1]].Unit];
            var spaceOrigin = spaceUnit is { } space ? units[space].Start
                : lastUnit.Rtl ? lastUnit.Start - TrailingSpaceAdvance * lastUnit.Em : lastUnit.End;
            text[current.Count] = ' ';
            origins[current.Count] = spaceOrigin;
            left = Math.Min(left, spaceOrigin);
            right = Math.Max(right, spaceOrigin);
            result.Add(new ImportedWord(new string(text), origins, baseline, left, right, bottom, top, font, size));
            current.Clear();
        }
        foreach (var position in accepted)
        {
            var unitIndex = visual[position].Unit;
            if (units[unitIndex].Letter is null) Flush(unitIndex);
            else current.Add(position);
        }
        Flush(null);
        return result;
    }

    // Returns whether a mapped glyph value is right-to-left; rejects anything
    // outside the admitted L / R / AL letters.
    private static bool Classify(string value, int number, string where)
    {
        bool? rtl = null;
        foreach (var character in value)
        {
            if (char.IsSurrogate(character))
                throw PdfImportReadingOrder.ProfileRejection(number, $"{where} contains a supplementary-plane character.");
            if (character is >= 'יִ' and <= '﷿' or >= 'ﹰ' and <= '﻿')
                throw PdfImportReadingOrder.ProfileRejection(number,
                    $"{where} maps a glyph to presentation form U+{(int)character:X4}; only base letters are admitted.");
            var type = UnicodeProperties.Bidi(character);
            bool characterRtl;
            if (type == BidiClass.L) characterRtl = false;
            else if (type is BidiClass.R or BidiClass.AL) characterRtl = true;
            else
                throw PdfImportReadingOrder.ProfileRejection(number,
                    $"{where} contains U+{(int)character:X4} (bidi class {type}); the profile admits only L, R and AL letters separated by spaces.");
            if (rtl is { } known && known != characterRtl)
                throw PdfImportReadingOrder.ProfileRejection(number, $"{where} has one glyph mapped to characters of both directions.");
            rtl = characterRtl;
        }
        return rtl == true;
    }

    // Resolved levels in visual order (only strong letters and spaces between
    // glyphs: N1/N2 depend only on the nearest strong neighbours, which the
    // reordering preserves), then L2. L2 is an involution for these levels.
    private static int[] CandidateLogicalOrder(List<(char Char, int Unit)> visual, List<Unit> units, int paragraphLevel)
    {
        var count = visual.Count;
        var rtl = new bool?[count];
        for (var index = 0; index < count; index++)
            if (units[visual[index].Unit].Letter is not null) rtl[index] = units[visual[index].Unit].Rtl;
        var levels = new int[count];
        for (var index = 0; index < count; index++)
        {
            bool isRtl;
            if (rtl[index] is { } strong) isRtl = strong;
            else
            {
                bool? before = null, after = null;
                for (var probe = index - 1; probe >= 0 && before is null; probe--) before = rtl[probe];
                for (var probe = index + 1; probe < count && after is null; probe++) after = rtl[probe];
                isRtl = before is { } b && after is { } a && b == a ? b : paragraphLevel == 1;
            }
            levels[index] = paragraphLevel == 0 ? (isRtl ? 1 : 0) : (isRtl ? 1 : 2);
        }
        var positions = Enumerable.Range(0, count).ToArray();
        var highest = levels.Length == 0 ? 0 : levels.Max();
        var lowestOdd = levels.Length == 0 ? 1 : levels.Where(level => level % 2 == 1).DefaultIfEmpty(highest + 1).Min();
        for (var level = highest; level >= lowestOdd; level--)
        {
            for (var start = 0; start < count;)
            {
                if (levels[positions[start]] < level) { start++; continue; }
                var end = start;
                while (end < count && levels[positions[end]] >= level) end++;
                Array.Reverse(positions, start, end - start);
                start = end;
            }
        }
        // positions[logical] = visual index after applying the involution to the visual sequence.
        return positions;
    }

    private static bool Verify(string logical, int[] visualOfLogical, int paragraphLevel, bool auto)
    {
        var bidi = UnicodeTextAnalysis.ResolveBidi(logical,
            auto ? null : paragraphLevel == 0 ? TextDirection.LeftToRight : TextDirection.RightToLeft);
        if (bidi.ParagraphLevel != paragraphLevel) return false;
        var order = bidi.VisualOrder(0, logical.Length);
        if (order.Count != visualOfLogical.Length) return false;
        // order[visual] = logical; it must invert visualOfLogical exactly.
        for (var visual = 0; visual < order.Count; visual++)
            if (visualOfLogical[order[visual]] != visual) return false;
        return true;
    }

    private static double Em(Letter letter)
    {
        var size = letter.PointSize;
        if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0) size = letter.FontSize;
        if (double.IsNaN(size) || double.IsInfinity(size) || size <= 0)
            size = Math.Max(letter.BoundingBox.Height, 1);
        return size;
    }
}
