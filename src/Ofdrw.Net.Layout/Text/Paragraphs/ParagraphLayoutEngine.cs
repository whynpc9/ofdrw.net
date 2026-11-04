using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Layout.Text.Unicode;

namespace Ofdrw.Net.Layout.Text.Paragraphs;

/// <summary>General paragraph layout: Unicode 17.0.0 bidi (UAX #9), line breaking (UAX #14) and
/// grapheme clusters (UAX #29), font fallback, run shaping, line metrics, alignment and pagination.</summary>
/// <remarks>
/// <para>The result is final: writers consume its faces, GIDs, positions and resolved styles without
/// reshaping or re-measuring. Existing Flow layout is unaffected.</para>
/// <para>Admission profile: tabs, explicit bidi formatting characters, other format controls
/// (ZWJ/ZWNJ, soft hyphen, word joiner, BOM), variation selectors and C0/C1 controls other than
/// line/paragraph separators are rejected before shaping. Synthetic bold/italic is never produced.</para>
/// </remarks>
public sealed class ParagraphLayoutEngine
{
    private const double Epsilon = 1e-9;
    private const int MaximumItemUtf16 = 4000;
    private const int MaximumFitAttempts = 32;
    private readonly ITextShaper _shaper;

    /// <summary>Creates an engine around a run shaper (for example HarfBuzz).</summary>
    public ParagraphLayoutEngine(ITextShaper shaper) => _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));

    /// <summary>Lays out the complete request. Failure or cancellation returns no partial result.</summary>
    public ParagraphLayoutResult Layout(ParagraphLayoutRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        Validate(request.Text);
        var session = new Session(this, request, cancellationToken);
        return session.Run();
    }

    private static void Validate(string text)
    {
        for (var index = 0; index < text.Length; index++)
        {
            var value = text[index];
            if (char.IsHighSurrogate(value))
            {
                if (index + 1 >= text.Length || !char.IsLowSurrogate(text[index + 1]))
                    throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                        sourceStart: index, sourceLength: 1);
                var scalar = char.ConvertToUtf32(value, text[index + 1]);
                if (scalar is >= 0xe0100 and <= 0xe01ef || CharUnicodeInfo.GetUnicodeCategory(text, index) == UnicodeCategory.Format)
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "Format controls and variation selectors need a separately verified text profile.",
                        sourceStart: index, sourceLength: 2);
                index++;
                continue;
            }
            if (char.IsLowSurrogate(value))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                    sourceStart: index, sourceLength: 1);
            if (value == '\t')
                throw new TextLayoutException(TextLayoutError.InvalidText,
                    "Tab stops are outside the paragraph layout profile.", sourceStart: index, sourceLength: 1);
            if (char.IsControl(value) && !IsHardBreak(value) && !IsParagraphSeparator(value))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unsupported control character.",
                    sourceStart: index, sourceLength: 1);
            if (CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.Format || value is >= '︀' and <= '️')
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Explicit bidi formatting, joiner and other format controls or variation selectors " +
                    "need a separately verified text profile.", sourceStart: index, sourceLength: 1);
        }
    }

    // UAX #9 paragraph separators (Bidi_Class B) admitted by this profile.
    private static bool IsParagraphSeparator(int value) => value is '\n' or '\r' or '\u0085' or 0x2029;
    // Mandatory line breaks inside a paragraph (UAX #14 BK that are not paragraph separators).
    private static bool IsHardBreak(int value) => value is '\u000b' or '\u000c' or 0x2028;

    private static bool IsHangingSpace(int codePoint)
    {
        if (codePoint == 0x20) return true;
        if (codePoint > 0xffff) return false;
        if (CharUnicodeInfo.GetUnicodeCategory((char)codePoint) != UnicodeCategory.SpaceSeparator) return false;
        var lineBreak = UnicodeProperties.LineBreak(codePoint, out _);
        return lineBreak is LineBreakClass.SP or LineBreakClass.BA;
    }

    private static bool IsJustificationSpace(int codePoint) =>
        codePoint <= 0xffff && CharUnicodeInfo.GetUnicodeCategory((char)codePoint) == UnicodeCategory.SpaceSeparator &&
        UnicodeProperties.LineBreak(codePoint, out _) is not LineBreakClass.GL;

    private static bool IsInterCharacterScript(string script) => script is "Hani" or "Hira" or "Kana" or "Hang" or "Bopo";

    private sealed class Item
    {
        internal int Start;   // paragraph code point index
        internal int End;
        internal byte Level;
        internal string Script = "Zyyy";
        internal TextFontFace Face = null!;
        internal TextStyle Style = null!;
        internal TextResolvedStyle Resolved = null!;
        internal TextShapedRun? Shaped;
    }

    private sealed class Paragraph
    {
        internal int Index;
        internal int Start;            // absolute code point index of the first code point
        internal int End;              // exclusive, before the terminator
        internal int TerminatorLength; // UTF-16
        internal int[] CodePoints = Array.Empty<int>();
        internal BidiAlgorithm Bidi = null!;
        internal LineBreakKind[] Breaks = Array.Empty<LineBreakKind>();
        internal bool[] GraphemeStart = Array.Empty<bool>();
        internal byte[] ClusterLevel = Array.Empty<byte>();
        internal int[] ItemOf = Array.Empty<int>();
        internal TextStyle[] StyleOf = Array.Empty<TextStyle>();
        internal string[] ScriptOf = Array.Empty<string>();
        internal List<Item> Items = new();
        internal double[] Prefix = Array.Empty<double>(); // measured advance prefix, mm
        internal int Length => End - Start;
    }

    private sealed class DraftRun
    {
        internal int Start, End; // paragraph code point indexes
        internal byte Level;
        internal Item Item = null!;
        internal TextShapedRun Shaped = null!;
        internal double Scale;
        internal double[] Justification = Array.Empty<double>();
    }

    private sealed class DraftLine
    {
        internal Paragraph Paragraph = null!;
        internal int Start, End, ContentEnd, TrimEnd; // paragraph code point indexes
        internal int TerminatorLength;
        internal bool FirstOfParagraph, EndsParagraph, HardBreak;
        internal List<DraftRun> Visual = new();
        internal double Ascent, Descent, Height, BaselineOffset;
        internal double LineStartX, ContentLeft, ContentWidth;
    }

    private sealed class Session
    {
        private readonly ParagraphLayoutEngine _engine;
        private readonly ParagraphLayoutRequest _request;
        private readonly ParagraphLayoutOptions _options;
        private readonly CancellationToken _token;
        private readonly CodePointText _text;
        private readonly TextStyle[] _styleByCodePoint;
        private int _shapeCalls;
        private string _engineVersion = string.Empty;
        private int _lineCount;

        internal Session(ParagraphLayoutEngine engine, ParagraphLayoutRequest request, CancellationToken token)
        {
            _engine = engine; _request = request; _options = request.Options; _token = token;
            _text = new CodePointText(request.Text);
            _styleByCodePoint = new TextStyle[_text.Length];
            var span = 0;
            for (var index = 0; index < _text.Length; index++)
            {
                var offset = _text.Offsets[index];
                while (offset >= request.Spans[span].Start + request.Spans[span].Length) span++;
                _styleByCodePoint[index] = request.Spans[span].Style;
                // A span boundary inside a surrogate pair cannot be represented.
                if (_text.Offsets[index + 1] - offset == 2 && offset + 1 == request.Spans[span].Start + request.Spans[span].Length)
                    throw new TextLayoutException(TextLayoutError.InvalidRequest, "A style span splits a surrogate pair.",
                        sourceStart: offset, sourceLength: 2);
            }
        }

        internal ParagraphLayoutResult Run()
        {
            var drafts = new List<DraftLine>();
            foreach (var paragraph in SplitParagraphs())
            {
                _token.ThrowIfCancellationRequested();
                Analyze(paragraph);
                drafts.AddRange(BreakLines(paragraph));
            }
            var pages = Paginate(drafts);
            return new ParagraphLayoutResult(_request, pages, _shapeCalls, _engineVersion);
        }

        private IEnumerable<Paragraph> SplitParagraphs()
        {
            var points = _text.CodePoints;
            var start = 0;
            var index = 0;
            var paragraphIndex = 0;
            while (index < points.Length)
            {
                var value = points[index];
                if (IsParagraphSeparator(value))
                {
                    var terminator = value == '\r' && index + 1 < points.Length && points[index + 1] == '\n' ? 2 : 1;
                    yield return Create(paragraphIndex++, start, index, terminator);
                    index += terminator;
                    start = index;
                    continue;
                }
                index++;
            }
            if (start < points.Length) yield return Create(paragraphIndex, start, points.Length, 0);
        }

        private Paragraph Create(int index, int start, int end, int terminatorCodePoints)
        {
            var paragraph = new Paragraph { Index = index, Start = start, End = end };
            paragraph.TerminatorLength = _text.Offsets[end + terminatorCodePoints] - _text.Offsets[end];
            paragraph.CodePoints = new int[end - start];
            Array.Copy(_text.CodePoints, start, paragraph.CodePoints, 0, end - start);
            return paragraph;
        }

        private int Offset(Paragraph paragraph, int index) => _text.Offsets[paragraph.Start + index];

        private void Analyze(Paragraph paragraph)
        {
            var points = paragraph.CodePoints;
            var count = points.Length;
            var types = new BidiClass[count];
            for (var index = 0; index < count; index++) types[index] = UnicodeProperties.Bidi(points[index]);
            var direction = _options.Direction switch
            {
                TextParagraphDirection.LeftToRight => 0,
                TextParagraphDirection.RightToLeft => 1,
                _ => -1,
            };
            paragraph.Bidi = BidiAlgorithm.Resolve(types, points, direction);
            paragraph.Breaks = LineBreaker.Analyze(points);
            paragraph.GraphemeStart = new bool[count + 1];
            foreach (var boundary in GraphemeBreaker.Boundaries(points)) paragraph.GraphemeStart[boundary] = true;
            paragraph.GraphemeStart[count] = true;
            paragraph.GraphemeStart[0] = true;
            paragraph.StyleOf = new TextStyle[count];
            Array.Copy(_styleByCodePoint, paragraph.Start, paragraph.StyleOf, 0, count);

            // Per-cluster bidi level (a cluster never straddles runs) and style check.
            var levels = paragraph.Bidi.Levels;
            paragraph.ClusterLevel = new byte[count];
            for (var index = 0; index < count;)
            {
                var end = NextCluster(paragraph, index);
                for (var j = index; j < end; j++)
                {
                    paragraph.ClusterLevel[j] = levels[index] == BidiAlgorithm.Removed ? paragraph.Bidi.ParagraphLevel : levels[index];
                    if (!ReferenceEquals(paragraph.StyleOf[j], paragraph.StyleOf[index]))
                        throw new TextLayoutException(TextLayoutError.InvalidRequest,
                            "A style span boundary splits an extended grapheme cluster.",
                            sourceStart: Offset(paragraph, index), sourceLength: Offset(paragraph, end) - Offset(paragraph, index));
                }
                index = end;
            }
            ResolveScripts(paragraph);
            Itemize(paragraph);
            Measure(paragraph);
        }

        private static int NextCluster(Paragraph paragraph, int index)
        {
            var end = index + 1;
            while (end < paragraph.Length && !paragraph.GraphemeStart[end]) end++;
            return end;
        }

        private static void ResolveScripts(Paragraph paragraph)
        {
            var count = paragraph.Length;
            var scripts = new string[count];
            // Cluster script: its first non-common code point.
            for (var index = 0; index < count;)
            {
                var end = NextCluster(paragraph, index);
                var script = "Zyyy";
                for (var j = index; j < end; j++)
                {
                    var value = UnicodeProperties.Script(paragraph.CodePoints[j]);
                    if (value is "Zyyy" or "Zinh" or "Zzzz") continue;
                    script = value;
                    break;
                }
                for (var j = index; j < end; j++) scripts[j] = script;
                index = end;
            }
            // Common clusters take the preceding real script, else the following one.
            string? previous = null;
            for (var index = 0; index < count; index++)
            {
                if (scripts[index] != "Zyyy") previous = scripts[index];
                else if (previous is not null) scripts[index] = previous;
            }
            string? next = null;
            for (var index = count - 1; index >= 0; index--)
            {
                if (scripts[index] != "Zyyy") next = scripts[index];
                else if (next is not null) scripts[index] = next;
            }
            paragraph.ScriptOf = scripts;
        }

        private void Itemize(Paragraph paragraph)
        {
            var count = paragraph.Length;
            paragraph.ItemOf = new int[count];
            Item? current = null;
            TextFontFace? previousFace = null;
            TextStyle? previousStyle = null;
            for (var index = 0; index < count;)
            {
                _token.ThrowIfCancellationRequested();
                var end = NextCluster(paragraph, index);
                if (IsHardBreak(paragraph.CodePoints[index]))
                {
                    for (var j = index; j < end; j++) paragraph.ItemOf[j] = -1;
                    current = null;
                    index = end;
                    continue;
                }
                var style = paragraph.StyleOf[index];
                var (face, family, fallbackIndex) = SelectFace(paragraph, index, end, style,
                    ReferenceEquals(previousStyle, style) ? previousFace : null);
                var level = paragraph.ClusterLevel[index];
                var script = paragraph.ScriptOf[index];
                var utf16 = Offset(paragraph, end) - Offset(paragraph, index);
                if (current is null || current.Level != level || current.Script != script ||
                    !ReferenceEquals(current.Face, face) || !ReferenceEquals(current.Style, style) ||
                    Offset(paragraph, current.End) - Offset(paragraph, current.Start) + utf16 > MaximumItemUtf16 &&
                    paragraph.Breaks[index] != LineBreakKind.Prohibited ||
                    Offset(paragraph, current.End) - Offset(paragraph, current.Start) + utf16 > 4096)
                {
                    current = new Item
                    {
                        Start = index, End = index, Level = level, Script = script, Face = face, Style = style,
                        Resolved = new TextResolvedStyle(style, family, fallbackIndex, face),
                    };
                    paragraph.Items.Add(current);
                }
                current.End = end;
                for (var j = index; j < end; j++) paragraph.ItemOf[j] = paragraph.Items.Count - 1;
                previousFace = face; previousStyle = style;
                index = end;
            }
        }

        private (TextFontFace Face, TextFontFamily Family, int FallbackIndex) SelectFace(Paragraph paragraph,
            int start, int end, TextStyle style, TextFontFace? previous)
        {
            bool Covers(TextFontFace face)
            {
                for (var j = start; j < end; j++)
                    if (!face.Covers(paragraph.CodePoints[j])) return false;
                return true;
            }
            // Common/inherited clusters (spaces, punctuation, digits) stay in the previous face when
            // it covers them, so punctuation does not jump to a fallback font mid-word.
            var common = UnicodeProperties.Script(paragraph.CodePoints[start]) is "Zyyy" or "Zinh";
            var anyStyledFace = false;
            for (var index = 0; index < style.FontFallback.Count; index++)
            {
                var family = style.FontFallback[index];
                var face = family.Select(style.Bold, style.Italic);
                if (face is null) continue;
                anyStyledFace = true;
                if (common && previous is not null && !ReferenceEquals(previous, face) && Covers(previous))
                {
                    for (var j = 0; j < style.FontFallback.Count; j++)
                        if (ReferenceEquals(style.FontFallback[j].Select(style.Bold, style.Italic), previous))
                            return (previous, style.FontFallback[j], j);
                }
                if (Covers(face)) return (face, family, index);
            }
            var offset = Offset(paragraph, start);
            var length = Offset(paragraph, end) - offset;
            if (!anyStyledFace)
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "No family supplies a face declaring the requested bold/italic style; synthetic styles are not produced.",
                    sourceStart: offset, sourceLength: length);
            throw new TextLayoutException(TextLayoutError.MissingGlyph,
                "No face in the style's fallback chain covers the grapheme cluster at UTF-16 offset " + offset + ".",
                sourceStart: offset, sourceLength: length);
        }

        private TextShapedRun Shape(Paragraph paragraph, Item item, int start, int end, byte level)
        {
            _token.ThrowIfCancellationRequested();
            if (++_shapeCalls > _options.MaximumShapeCalls)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded MaximumShapeCalls.");
            var offset = Offset(paragraph, start);
            var text = _request.Text.Substring(offset, Offset(paragraph, end) - offset);
            var features = item.Style.Features.Select(feature => new TextFeature(feature.Tag, feature.Value, 0, text.Length)).ToArray();
            var shapeRequest = new TextShapeRequest(text, offset, item.Face,
                level % 2 == 1 ? TextDirection.RightToLeft : TextDirection.LeftToRight,
                item.Script, item.Style.Language, features);
            var shaped = _engine._shaper.Shape(shapeRequest, _token);
            if (!ReferenceEquals(shaped.Request, shapeRequest))
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "The shaper returned a result for another request.");
            if (_engineVersion.Length == 0) _engineVersion = shaped.EngineVersion;
            return shaped;
        }

        private void Measure(Paragraph paragraph)
        {
            var count = paragraph.Length;
            var advance = new double[count];
            foreach (var item in paragraph.Items)
            {
                item.Shaped = Shape(paragraph, item, item.Start, item.End, item.Level);
                var scale = item.Style.FontSizeMillimeters / item.Face.UnitsPerEm;
                foreach (var glyph in item.Shaped.Glyphs)
                {
                    var index = _text.IndexOfOffset(glyph.ClusterStart) - paragraph.Start;
                    advance[index] += glyph.XAdvance * scale;
                }
            }
            paragraph.Prefix = new double[count + 1];
            for (var index = 0; index < count; index++) paragraph.Prefix[index + 1] = paragraph.Prefix[index] + advance[index];
        }

        private int TrimEnd(Paragraph paragraph, int start, int end)
        {
            while (end > start && (IsHardBreak(paragraph.CodePoints[end - 1]) || IsHangingSpace(paragraph.CodePoints[end - 1])))
                end--;
            return end;
        }

        private double Estimate(Paragraph paragraph, int start, int end) =>
            paragraph.Prefix[TrimEnd(paragraph, start, end)] - paragraph.Prefix[start];

        private IEnumerable<DraftLine> BreakLines(Paragraph paragraph)
        {
            var count = paragraph.Length;
            if (count == 0)
            {
                yield return BuildLine(paragraph, 0, 0, firstOfParagraph: true, endsParagraph: true);
                yield break;
            }
            var candidates = new List<int>();
            for (var index = 1; index <= count; index++)
                if (paragraph.Breaks[index] != LineBreakKind.Prohibited && paragraph.GraphemeStart[index])
                    candidates.Add(index);
            var start = 0;
            var first = true;
            var candidate = 0;
            while (start < count)
            {
                _token.ThrowIfCancellationRequested();
                var available = _options.WidthMillimeters - (first ? _options.FirstLineIndentMillimeters : 0);
                while (candidate < candidates.Count && candidates[candidate] <= start) candidate++;
                // Candidate breaks in order, ending at the first mandatory break.
                var fitting = new List<int>();
                var firstOver = -1;
                for (var c = candidate; c < candidates.Count; c++)
                {
                    var position = candidates[c];
                    if (Estimate(paragraph, start, position) <= available + Epsilon) fitting.Add(position);
                    else { firstOver = position; break; }
                    if (paragraph.Breaks[position] == LineBreakKind.Mandatory) break;
                }
                DraftLine? line = null;
                // Verify by building (reshaping cut runs); back off on overflow.
                for (var attempt = fitting.Count - 1; attempt >= 0 && fitting.Count - attempt <= MaximumFitAttempts; attempt--)
                {
                    var draft = BuildLine(paragraph, start, fitting[attempt], first, fitting[attempt] == count);
                    if (draft.ContentWidth <= available + 1e-7) { line = draft; break; }
                }
                if (line is null)
                {
                    var limit = firstOver >= 0 ? firstOver : fitting.Count > 0 ? fitting[0] : count;
                    if (fitting.Count > 0 && fitting.Count > MaximumFitAttempts)
                        throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                            "Line fitting after reshaping did not converge.", sourceStart: Offset(paragraph, start),
                            sourceLength: Offset(paragraph, limit) - Offset(paragraph, start));
                    if (_options.Overflow == TextOverflowPolicy.Reject)
                        throw new TextLayoutException(TextLayoutError.LayoutOverflow,
                            "A segment without a legal line break exceeds the line width.",
                            sourceStart: Offset(paragraph, start), sourceLength: Offset(paragraph, limit) - Offset(paragraph, start));
                    line = EmergencyBreak(paragraph, start, limit, available, first);
                }
                yield return line;
                first = false;
                start = line.End;
            }
        }

        private DraftLine EmergencyBreak(Paragraph paragraph, int start, int limit, double available, bool first)
        {
            var boundaries = new List<int>();
            for (var index = start + 1; index <= limit; index++)
                if (paragraph.GraphemeStart[index]) boundaries.Add(index);
            // Largest grapheme boundary whose rebuilt line fits; binary search on the estimate first.
            var low = 0;
            var high = boundaries.Count - 1;
            var best = -1;
            while (low <= high)
            {
                var middle = (low + high) / 2;
                if (Estimate(paragraph, start, boundaries[middle]) <= available + Epsilon) { best = middle; low = middle + 1; }
                else high = middle - 1;
            }
            for (var attempt = best; attempt >= 0 && best - attempt < MaximumFitAttempts; attempt--)
            {
                var draft = BuildLine(paragraph, start, boundaries[attempt], first, boundaries[attempt] == paragraph.Length);
                if (draft.ContentWidth <= available + 1e-7) return draft;
            }
            throw new TextLayoutException(TextLayoutError.LayoutOverflow,
                "A single extended grapheme cluster exceeds the line width.",
                sourceStart: Offset(paragraph, start),
                sourceLength: Offset(paragraph, boundaries.Count > 0 ? boundaries[0] : limit) - Offset(paragraph, start));
        }

        private DraftLine BuildLine(Paragraph paragraph, int start, int end, bool firstOfParagraph, bool endsParagraph)
        {
            _token.ThrowIfCancellationRequested();
            if (++_lineCount > _options.MaximumLines * 4L + 64)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded its line budget.");
            var line = new DraftLine
            {
                Paragraph = paragraph, Start = start, End = end, FirstOfParagraph = firstOfParagraph,
                EndsParagraph = endsParagraph,
            };
            var contentEnd = end;
            if (contentEnd > start && IsHardBreak(paragraph.CodePoints[contentEnd - 1]))
            {
                contentEnd--;
                line.HardBreak = true;
                line.TerminatorLength = Offset(paragraph, end) - Offset(paragraph, contentEnd);
            }
            if (endsParagraph)
            {
                line.HardBreak = true;
                line.TerminatorLength += paragraph.TerminatorLength;
            }
            line.ContentEnd = contentEnd;
            line.TrimEnd = TrimEnd(paragraph, start, contentEnd);

            // L1 line levels, uniform per grapheme cluster.
            var lineLevels = paragraph.Bidi.LineLevels(start, contentEnd);
            var levels = new byte[contentEnd - start];
            for (var index = start; index < contentEnd;)
            {
                var clusterEnd = Math.Min(NextCluster(paragraph, index), contentEnd);
                for (var j = index; j < clusterEnd; j++) levels[j - start] = lineLevels[index - start];
                index = clusterEnd;
            }
            var logical = new List<DraftRun>();
            for (var index = start; index < contentEnd;)
            {
                var itemIndex = paragraph.ItemOf[index];
                var level = levels[index - start];
                var runEnd = index + 1;
                while (runEnd < contentEnd && paragraph.ItemOf[runEnd] == itemIndex && levels[runEnd - start] == level) runEnd++;
                var item = paragraph.Items[itemIndex];
                var shaped = index == item.Start && runEnd == item.End && level == item.Level
                    ? item.Shaped! : Shape(paragraph, item, index, runEnd, level);
                logical.Add(new DraftRun
                {
                    Start = index, End = runEnd, Level = level, Item = item, Shaped = shaped,
                    Scale = item.Style.FontSizeMillimeters / item.Face.UnitsPerEm,
                    Justification = new double[shaped.Glyphs.Count],
                });
                index = runEnd;
            }
            line.Visual = Reorder(logical);
            Metrics(line, paragraph, start);
            Align(line, paragraph);
            return line;
        }

        private static List<DraftRun> Reorder(List<DraftRun> logical)
        {
            var runs = logical.ToArray();
            if (runs.Length == 0) return new List<DraftRun>();
            var highest = runs.Max(run => run.Level);
            var lowestOdd = runs.Where(run => run.Level % 2 == 1).Select(run => (int)run.Level).DefaultIfEmpty(255).Min();
            for (int level = highest; level >= lowestOdd && level > 0; level--)
            {
                for (var index = 0; index < runs.Length;)
                {
                    if (runs[index].Level < level) { index++; continue; }
                    var end = index;
                    while (end < runs.Length && runs[end].Level >= level) end++;
                    Array.Reverse(runs, index, end - index);
                    index = end;
                }
            }
            return runs.ToList();
        }

        private void Metrics(DraftLine line, Paragraph paragraph, int start)
        {
            double ascent = 0, descent = 0, gapAscent = 0, gapDescent = 0;
            void Add(TextFontFace face, double size)
            {
                var scale = size / face.UnitsPerEm;
                var a = face.Ascender * scale;
                var d = -face.Descender * scale;
                var g = Math.Max(0, face.LineGap) * scale;
                ascent = Math.Max(ascent, a); descent = Math.Max(descent, d);
                gapAscent = Math.Max(gapAscent, a + g / 2); gapDescent = Math.Max(gapDescent, d + g / 2);
            }
            if (line.Visual.Count == 0)
            {
                // An empty line takes the metrics of its style's first available face.
                var style = paragraph.Length == 0
                    ? _styleByCodePoint.Length == 0 ? _request.Spans[0].Style
                        : _styleByCodePoint[Math.Min(Math.Max(paragraph.Start - 1, 0), _styleByCodePoint.Length - 1)]
                    : paragraph.StyleOf[Math.Min(start, paragraph.Length - 1)];
                var face = style.FontFallback.Select(family => family.Select(style.Bold, style.Italic))
                    .FirstOrDefault(value => value is not null) ?? style.FontFallback[0].Faces[0];
                Add(face, style.FontSizeMillimeters);
            }
            foreach (var run in line.Visual) Add(run.Item.Face, run.Item.Style.FontSizeMillimeters);
            var normal = gapAscent + gapDescent;
            var rule = _options.LineHeight;
            switch (rule.Kind)
            {
                case "Exact":
                    line.Height = rule.Value;
                    line.BaselineOffset = (rule.Value - (ascent + descent)) / 2 + ascent;
                    break;
                case "Multiple":
                    line.Height = normal * rule.Value;
                    line.BaselineOffset = gapAscent + (line.Height - normal) / 2;
                    break;
                default:
                    line.Height = normal;
                    line.BaselineOffset = gapAscent;
                    break;
            }
            line.Ascent = ascent; line.Descent = descent;
        }

        private void Align(DraftLine line, Paragraph paragraph)
        {
            var trimOffset = Offset(paragraph, line.TrimEnd);
            double total = 0, hanging = 0;
            foreach (var run in line.Visual)
                foreach (var glyph in run.Shaped.Glyphs)
                {
                    var advance = glyph.XAdvance * run.Scale;
                    total += advance;
                    if (glyph.ClusterStart >= trimOffset) hanging += advance;
                }
            var content = total - hanging;
            var rtl = paragraph.Bidi.ParagraphLevel == 1;
            var indent = line.FirstOfParagraph ? _options.FirstLineIndentMillimeters : 0;
            var width = _options.WidthMillimeters;
            var available = width - indent;
            if (_options.Alignment == TextAlignment.Justify && !line.EndsParagraph && !line.HardBreak && content < available)
                content += Justify(line, paragraph, available - content, trimOffset);
            var alignment = _options.Alignment switch
            {
                TextAlignment.Start or TextAlignment.Justify => rtl ? TextAlignment.Right : TextAlignment.Left,
                TextAlignment.End => rtl ? TextAlignment.Left : TextAlignment.Right,
                _ => _options.Alignment,
            };
            double contentLeft = alignment switch
            {
                TextAlignment.Left => rtl ? 0 : indent,
                TextAlignment.Right => (rtl ? width - indent : width) - content,
                _ => (rtl ? 0 : indent) + (available - content) / 2,
            };
            line.ContentLeft = contentLeft;
            line.ContentWidth = content;
            line.LineStartX = rtl ? contentLeft - hanging : contentLeft;
        }

        private double Justify(DraftLine line, Paragraph paragraph, double extra, int trimOffset)
        {
            // Opportunities: word spaces, plus gaps between graphemes next to scripts set without
            // word spaces (Han, kana, Hangul, Bopomofo), except before closing punctuation.
            var expandAfter = new HashSet<int>(); // paragraph code point index of a grapheme start
            for (var index = line.Start; index < line.TrimEnd;)
            {
                var next = NextCluster(paragraph, index);
                if (next >= line.TrimEnd) break;
                var space = IsJustificationSpace(paragraph.CodePoints[index]);
                var nextSpace = IsJustificationSpace(paragraph.CodePoints[next]);
                var closing = UnicodeProperties.LineBreak(paragraph.CodePoints[next], out _) is LineBreakClass.CL or
                    LineBreakClass.CP or LineBreakClass.EX or LineBreakClass.IS or LineBreakClass.NS;
                if (space || !nextSpace && !closing && (IsInterCharacterScript(paragraph.ScriptOf[index]) ||
                        IsInterCharacterScript(paragraph.ScriptOf[next])))
                    expandAfter.Add(index);
                index = next;
            }
            var opportunities = new List<(DraftRun Run, int Glyph)>();
            foreach (var run in line.Visual)
            {
                // The last glyph (in shaper order) of each selected grapheme receives the expansion.
                var last = new Dictionary<int, int>();
                for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                {
                    var glyph = run.Shaped.Glyphs[index];
                    if (glyph.ClusterStart >= trimOffset) continue;
                    var grapheme = _text.IndexOfOffset(glyph.ClusterStart) - paragraph.Start;
                    while (grapheme > 0 && !paragraph.GraphemeStart[grapheme]) grapheme--;
                    if (expandAfter.Contains(grapheme)) last[grapheme] = index;
                }
                foreach (var pair in last) opportunities.Add((run, pair.Value));
            }
            if (opportunities.Count == 0) return 0;
            var each = extra / opportunities.Count;
            foreach (var (run, glyph) in opportunities) run.Justification[glyph] += each;
            return each * opportunities.Count;
        }

        private List<ParagraphLayoutPage> Paginate(List<DraftLine> drafts)
        {
            var pages = new List<ParagraphLayoutPage>();
            var current = new List<ParagraphLayoutLine>();
            var height = _options.PageHeightMillimeters;
            double cursor = 0;
            var pageIndex = 0;
            var lines = 0;
            foreach (var draft in drafts)
            {
                _token.ThrowIfCancellationRequested();
                if (++lines > _options.MaximumLines)
                    throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded MaximumLines.");
                if (draft.FirstOfParagraph && draft.Paragraph.Index > 0 && current.Count > 0)
                    cursor += _options.ParagraphSpacingMillimeters;
                if (height is { } limit)
                {
                    if (draft.Height > limit + Epsilon)
                        throw new TextLayoutException(TextLayoutError.LayoutOverflow, "A line is taller than the page.",
                            sourceStart: Offset(draft.Paragraph, draft.Start),
                            sourceLength: Offset(draft.Paragraph, draft.End) - Offset(draft.Paragraph, draft.Start));
                    if (cursor + draft.Height > limit + Epsilon && current.Count > 0)
                    {
                        pages.Add(new ParagraphLayoutPage(pageIndex++, current, current[current.Count - 1].TopMillimeters +
                            current[current.Count - 1].HeightMillimeters));
                        current = new List<ParagraphLayoutLine>();
                        cursor = 0;
                    }
                }
                current.Add(Finalize(draft, pageIndex, cursor));
                cursor += draft.Height;
            }
            pages.Add(new ParagraphLayoutPage(pageIndex, current, current.Count == 0 ? 0 :
                current[current.Count - 1].TopMillimeters + current[current.Count - 1].HeightMillimeters));
            return pages;
        }

        private ParagraphLayoutLine Finalize(DraftLine draft, int pageIndex, double top)
        {
            var paragraph = draft.Paragraph;
            var baseline = top + draft.BaselineOffset;
            var pen = draft.LineStartX;
            var runs = new List<ParagraphLayoutRun>(draft.Visual.Count);
            Bounds? lineInk = null;
            foreach (var run in draft.Visual)
            {
                var origin = pen;
                var glyphs = new List<ParagraphLayoutGlyph>(run.Shaped.Glyphs.Count);
                Bounds? runInk = null;
                double penY = 0;
                for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                {
                    var glyph = run.Shaped.Glyphs[index];
                    var x = Check(pen + glyph.XOffset * run.Scale);
                    var y = Check(baseline - (penY + glyph.YOffset) * run.Scale);
                    var advance = glyph.XAdvance * run.Scale + run.Justification[index];
                    TextLayoutBounds? ink = null;
                    if (glyph.Ink is { } value)
                    {
                        ink = new TextLayoutBounds(Check(x + value.X * run.Scale),
                            Check(y - (value.Y + value.Height) * run.Scale), value.Width * run.Scale, value.Height * run.Scale);
                        runInk = Bounds.Union(runInk, ink);
                    }
                    glyphs.Add(new ParagraphLayoutGlyph(glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength, x, y,
                        Check(advance), -glyph.YAdvance * run.Scale, run.Justification[index], ink));
                    pen = Check(pen + advance);
                    penY += glyph.YAdvance;
                }
                var offset = Offset(paragraph, run.Start);
                var text = _request.Text.Substring(offset, Offset(paragraph, run.End) - offset);
                var runBounds = runInk?.ToBounds();
                if (runBounds is not null) lineInk = Bounds.Union(lineInk, runBounds);
                runs.Add(new ParagraphLayoutRun(offset, text, run.Shaped, run.Level, run.Item.Resolved, origin, baseline,
                    pen - origin, glyphs, runBounds));
            }
            var lineStart = Offset(paragraph, draft.Start);
            var lineEnd = Offset(paragraph, draft.End) + (draft.EndsParagraph ? paragraph.TerminatorLength : 0);
            return new ParagraphLayoutLine(lineStart, lineEnd - lineStart, draft.TerminatorLength, paragraph.Index,
                pageIndex, paragraph.Bidi.ParagraphLevel, draft.EndsParagraph, draft.HardBreak, top, baseline,
                draft.Height, draft.Ascent, draft.Descent, draft.ContentLeft, draft.ContentWidth, runs, lineInk?.ToBounds());
        }

        private static double Check(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value) && Math.Abs(value) <= 1_000_000 ? value :
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Layout geometry exceeds the coordinate budget.");
    }

    private sealed class Bounds
    {
        private double _left, _top, _right, _bottom;
        internal static Bounds Union(Bounds? bounds, TextLayoutBounds value)
        {
            if (bounds is null) return new Bounds { _left = value.X, _top = value.Y, _right = value.X + value.Width, _bottom = value.Y + value.Height };
            bounds._left = Math.Min(bounds._left, value.X); bounds._top = Math.Min(bounds._top, value.Y);
            bounds._right = Math.Max(bounds._right, value.X + value.Width);
            bounds._bottom = Math.Max(bounds._bottom, value.Y + value.Height);
            return bounds;
        }
        internal TextLayoutBounds ToBounds() => new(_left, _top, _right - _left, _bottom - _top);
    }
}
