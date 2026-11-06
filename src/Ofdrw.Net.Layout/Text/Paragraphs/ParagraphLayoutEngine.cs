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
/// <para>Admission profile: explicit bidi formatting characters, format controls other than ZWJ/ZWNJ
/// (soft hyphen, word joiner, BOM, ...) and C0/C1 controls other than tab and line/paragraph separators
/// are rejected before shaping. ZWJ, ZWNJ and variation selectors are shaped with their grapheme
/// cluster. Tabs advance to the options' tab stops. Synthetic bold/oblique is produced only when
/// <see cref="ParagraphLayoutOptions.SyntheticStyles"/> allows it.</para>
/// </remarks>
public sealed class ParagraphLayoutEngine
{
    private const double Epsilon = 1e-9;
    private const int MaximumItemUtf16 = 4000;
    private const int MaximumFitAttempts = 32;
    private const int MaximumEmergencyMisses = 4;
    private const int MaximumNearMisses = 3;
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
                if (scalar is < 0xe0100 or > 0xe01ef && CharUnicodeInfo.GetUnicodeCategory(text, index) == UnicodeCategory.Format)
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "Format controls and variation selectors need a separately verified text profile.",
                        sourceStart: index, sourceLength: 2);
                index++;
                continue;
            }
            if (char.IsLowSurrogate(value))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.",
                    sourceStart: index, sourceLength: 1);
            if (char.IsControl(value) && value != '\t' && !IsHardBreak(value) && !IsParagraphSeparator(value))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unsupported control character.",
                    sourceStart: index, sourceLength: 1);
            if (CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.Format && !IsJoinerOrSelector(value))
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Explicit bidi formatting and format controls other than ZWJ/ZWNJ " +
                    "need a separately verified text profile.", sourceStart: index, sourceLength: 1);
        }
    }

    // UAX #9 paragraph separators (Bidi_Class B) admitted by this profile.
    private static bool IsParagraphSeparator(int value) => value is '\n' or '\r' or '\u0085' or 0x2029;
    // Mandatory line breaks inside a paragraph (UAX #14 BK that are not paragraph separators).
    private static bool IsHardBreak(int value) => value is '\u000b' or '\u000c' or 0x2028;

    // ZWNJ, ZWJ and variation selectors: shaped inside their grapheme cluster, never cmap-required.
    private static bool IsJoinerOrSelector(int value) =>
        value is 0x200c or 0x200d or >= 0xfe00 and <= 0xfe0f or >= 0xe0100 and <= 0xe01ef;

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
        internal int ContextStart, ContextEnd; // shaping context of Shaped
        internal bool IsTab;
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
        internal List<int> Tabs = new();                   // paragraph code point indexes of tab characters
        internal double[] Prefix = Array.Empty<double>(); // measured advance prefix, mm
        internal double MaximumEm;                         // largest style size in the paragraph, mm
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
        internal double[] Extra = Array.Empty<double>(); // tab advance and letter spacing, mm
        internal bool IsTab => Item.IsTab;
        internal double Advance(int glyph) => Shaped.Glyphs[glyph].XAdvance * Scale + Extra[glyph] + Justification[glyph];
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
            // An empty request still has one empty, styled line.
            if (start < points.Length || points.Length == 0) yield return Create(paragraphIndex, start, points.Length, 0);
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
                var (face, family, fallbackIndex, synthetic) = SelectFace(paragraph, index, end, style,
                    ReferenceEquals(previousStyle, style) ? previousFace : null);
                var level = paragraph.ClusterLevel[index];
                var script = paragraph.ScriptOf[index];
                var utf16 = Offset(paragraph, end) - Offset(paragraph, index);
                var tab = paragraph.CodePoints[index] == '\t';
                if (tab)
                {
                    // A tab is its own item: the engine positions it at a tab stop and never shapes it.
                    paragraph.Tabs.Add(index);
                    var tabItem = new Item
                    {
                        Start = index, End = end, Level = level, Script = script, Face = face, Style = style, IsTab = true,
                        Resolved = new TextResolvedStyle(style, family, fallbackIndex, face, synthetic,
                            BoldStroke(style, script)),
                    };
                    paragraph.Items.Add(tabItem);
                    for (var j = index; j < end; j++) paragraph.ItemOf[j] = paragraph.Items.Count - 1;
                    current = null;
                    previousFace = face; previousStyle = style;
                    index = end;
                    continue;
                }
                if (current is null || current.Level != level || current.Script != script ||
                    !ReferenceEquals(current.Face, face) || !ReferenceEquals(current.Style, style) ||
                    Offset(paragraph, current.End) - Offset(paragraph, current.Start) + utf16 > MaximumItemUtf16 &&
                    paragraph.Breaks[index] != LineBreakKind.Prohibited ||
                    Offset(paragraph, current.End) - Offset(paragraph, current.Start) + utf16 > 4096)
                {
                    current = new Item
                    {
                        Start = index, End = index, Level = level, Script = script, Face = face, Style = style,
                        Resolved = new TextResolvedStyle(style, family, fallbackIndex, face, synthetic,
                            BoldStroke(style, script)),
                    };
                    paragraph.Items.Add(current);
                }
                current.End = end;
                for (var j = index; j < end; j++) paragraph.ItemOf[j] = paragraph.Items.Count - 1;
                previousFace = face; previousStyle = style;
                index = end;
            }
        }

        // Synthetic bold uses the shared emphasis geometry of the existing native OFD output: a thinner
        // stroke for Han, whose dense counters close under the Latin width.
        private static double BoldStroke(TextStyle style, string script) =>
            script == "Hani" ? Ofdrw.Net.Core.Fonts.FontEmphasisMetrics.HanBoldStrokeWidth(style.FontSizeMillimeters)
                : Ofdrw.Net.Core.Fonts.FontEmphasisMetrics.BoldStrokeWidth(style.FontSizeMillimeters);

        // The declared face for the requested style or, when synthesis is allowed, the closest declared
        // face plus the synthetic styles that make up the difference. Never "un-bolds" or "un-slants".
        private (TextFontFace? Face, TextSyntheticStyle Synthetic) Candidate(TextFontFamily family, TextStyle style)
        {
            var exact = family.Select(style.Bold, style.Italic);
            if (exact is not null || _options.SyntheticStyles != TextSyntheticStylePolicy.Synthesize) return (exact, TextSyntheticStyle.None);
            if (style.Italic && family.Select(style.Bold, false) is { } upright) return (upright, TextSyntheticStyle.Oblique);
            if (style.Bold && family.Select(false, style.Italic) is { } regular) return (regular, TextSyntheticStyle.Bold);
            if (style.Bold && style.Italic && family.Select(false, false) is { } plain)
                return (plain, TextSyntheticStyle.Bold | TextSyntheticStyle.Oblique);
            return (null, TextSyntheticStyle.None);
        }

        private (TextFontFace Face, TextFontFamily Family, int FallbackIndex, TextSyntheticStyle Synthetic) SelectFace(
            Paragraph paragraph, int start, int end, TextStyle style, TextFontFace? previous)
        {
            bool Covers(TextFontFace face)
            {
                for (var j = start; j < end; j++)
                {
                    var value = paragraph.CodePoints[j];
                    // Tabs are positioned by the engine; joiners and selectors are resolved by the shaper.
                    if (value == '\t' || IsJoinerOrSelector(value)) continue;
                    if (!face.Covers(value)) return false;
                }
                return true;
            }
            var offset = Offset(paragraph, start);
            var length = Offset(paragraph, end) - offset;
            (TextFontFace, TextFontFamily, int, TextSyntheticStyle) Accept(TextFontFace face, TextFontFamily family, int index,
                TextSyntheticStyle synthetic)
            {
                if ((synthetic & TextSyntheticStyle.Bold) != 0 && (style.Paint.Stroke || !style.Paint.Fill))
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "Synthetic bold paints a fill-colored outline stroke and cannot combine with a requested stroke or an unfilled paint.",
                        sourceStart: offset, sourceLength: length);
                return (face, family, index, synthetic);
            }
            // Common/inherited clusters (spaces, punctuation, digits) stay in the previous face when
            // it covers them, so punctuation does not jump to a fallback font mid-word.
            var common = UnicodeProperties.Script(paragraph.CodePoints[start]) is "Zyyy" or "Zinh" ||
                paragraph.CodePoints[start] == '\t';
            var anyStyledFace = false;
            for (var index = 0; index < style.FontFallback.Count; index++)
            {
                var family = style.FontFallback[index];
                var (face, synthetic) = Candidate(family, style);
                if (face is null) continue;
                anyStyledFace = true;
                if (common && previous is not null && !ReferenceEquals(previous, face) && Covers(previous))
                {
                    for (var j = 0; j < style.FontFallback.Count; j++)
                    {
                        var (candidate, candidateSynthetic) = Candidate(style.FontFallback[j], style);
                        if (ReferenceEquals(candidate, previous))
                            return Accept(previous, style.FontFallback[j], j, candidateSynthetic);
                    }
                }
                if (Covers(face)) return Accept(face, family, index, synthetic);
            }
            if (!anyStyledFace)
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    _options.SyntheticStyles == TextSyntheticStylePolicy.Synthesize
                        ? "No family supplies a face that can satisfy the requested bold/italic style."
                        : "No family supplies a face declaring the requested bold/italic style; synthetic styles are not produced.",
                    sourceStart: offset, sourceLength: length);
            throw new TextLayoutException(TextLayoutError.MissingGlyph,
                "No face in the style's fallback chain covers the grapheme cluster at UTF-16 offset " + offset + ".",
                sourceStart: offset, sourceLength: length);
        }

        // Adjacent text shaped the same way (face, script, language, features) at the same level and
        // inside [limitStart, limitEnd) is shaping context, so a paint- or size-only style change
        // inside a word keeps joining/kerning. Line edges and level changes end the context.
        private (int Start, int End) ShapingContext(Paragraph paragraph, Item item, int start, int end, byte level,
            int limitStart, int limitEnd, Func<int, byte> levelAt)
        {
            if (_engine._shaper is not IContextualTextShaper) return (start, end);
            bool Compatible(int index)
            {
                var other = paragraph.ItemOf[index];
                if (other < 0 || levelAt(index) != level) return false;
                var candidate = paragraph.Items[other];
                if (candidate.IsTab) return false;
                return ReferenceEquals(candidate, item) || ReferenceEquals(candidate.Face, item.Face) &&
                    candidate.Script == item.Script && candidate.Style.Language == item.Style.Language &&
                    candidate.Style.ShapingFeatures.SequenceEqual(item.Style.ShapingFeatures);
            }
            var contextStart = start;
            var contextEnd = end;
            while (contextStart > limitStart && Compatible(contextStart - 1) &&
                   Offset(paragraph, contextEnd) - Offset(paragraph, contextStart - 1) <= 4096) contextStart--;
            while (contextEnd < limitEnd && Compatible(contextEnd) &&
                   Offset(paragraph, contextEnd + 1) - Offset(paragraph, contextStart) <= 4096) contextEnd++;
            return (contextStart, contextEnd);
        }

        private TextShapedRun Shape(Paragraph paragraph, Item item, int start, int end, byte level,
            (int Start, int End) context)
        {
            _token.ThrowIfCancellationRequested();
            if (++_shapeCalls > _options.MaximumShapeCalls)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Paragraph layout exceeded MaximumShapeCalls.");
            var offset = Offset(paragraph, start);
            var text = _request.Text.Substring(offset, Offset(paragraph, end) - offset);
            var direction = level % 2 == 1 ? TextDirection.RightToLeft : TextDirection.LeftToRight;
            TextShapeRequest shapeRequest;
            TextShapedRun shaped;
            if ((context.Start < start || context.End > end) && _engine._shaper is IContextualTextShaper contextual)
            {
                var contextOffset = Offset(paragraph, context.Start);
                var contextText = _request.Text.Substring(contextOffset, Offset(paragraph, context.End) - contextOffset);
                var contextFeatures = item.Style.ShapingFeatures.Select(feature =>
                    new TextFeature(feature.Tag, feature.Value, 0, contextText.Length)).ToArray();
                var whole = new TextShapeRequest(contextText, contextOffset, item.Face, direction, item.Script,
                    item.Style.Language, contextFeatures, paragraphProfile: true);
                TextShapeRequest? contextItem = null;
                // CreateContextualItem checks .NET text-element boundaries; if the runtime disagrees
                // with a UAX #29 item boundary, shape without context rather than fail.
                try { contextItem = TextShapeRequest.CreateContextualItem(whole, new TextSourceSpan(offset, text.Length)); }
                catch (TextLayoutException error) when (error.Error == TextLayoutError.InvalidRequest) { }
                if (contextItem is not null)
                {
                    shapeRequest = contextItem;
                    shaped = contextual.ShapeContextualItem(shapeRequest, _token);
                }
                else
                {
                    var features = item.Style.ShapingFeatures.Select(feature => new TextFeature(feature.Tag, feature.Value, 0, text.Length)).ToArray();
                    shapeRequest = new TextShapeRequest(text, offset, item.Face, direction,
                        item.Script, item.Style.Language, features, paragraphProfile: true);
                    shaped = _engine._shaper.Shape(shapeRequest, _token);
                }
            }
            else
            {
                var features = item.Style.ShapingFeatures.Select(feature => new TextFeature(feature.Tag, feature.Value, 0, text.Length)).ToArray();
                shapeRequest = new TextShapeRequest(text, offset, item.Face, direction,
                    item.Script, item.Style.Language, features, paragraphProfile: true);
                shaped = _engine._shaper.Shape(shapeRequest, _token);
            }
            if (!ReferenceEquals(shaped.Request, shapeRequest))
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "The shaper returned a result for another request.");
            if (_engineVersion.Length == 0) _engineVersion = shaped.EngineVersion;
            return shaped;
        }

        // A tab is one blank glyph (the face's space) with no advance of its own; the line builder adds
        // the distance to the tab stop. The logical text keeps U+0009.
        private TextShapedRun TabRun(Paragraph paragraph, Item item, byte level)
        {
            var space = item.Face.Covers(0x20) ? item.Face.NominalGlyph(0x20) : 0;
            var offset = Offset(paragraph, item.Start);
            if (space <= 0)
                throw new TextLayoutException(TextLayoutError.MissingGlyph,
                    "A tab needs a space glyph in its selected face.", sourceStart: offset, sourceLength: 1);
            var request = new TextShapeRequest("\t", offset, item.Face,
                level % 2 == 1 ? TextDirection.RightToLeft : TextDirection.LeftToRight, item.Script, item.Style.Language,
                null, paragraphProfile: true);
            return new TextShapedRun(request, new[] { new TextShapedGlyph((uint)space, offset, 1, 0, 0, 0, 0, null, false) },
                "Ofdrw.Net.Layout/tab");
        }

        private void Measure(Paragraph paragraph)
        {
            var count = paragraph.Length;
            var advance = new double[count];
            foreach (var item in paragraph.Items)
            {
                if (item.IsTab)
                {
                    item.Shaped = TabRun(paragraph, item, item.Level);
                    (item.ContextStart, item.ContextEnd) = (item.Start, item.End);
                    continue;
                }
                var context = ShapingContext(paragraph, item, item.Start, item.End, item.Level, 0, count,
                    index => paragraph.ClusterLevel[index]);
                item.Shaped = Shape(paragraph, item, item.Start, item.End, item.Level, context);
                (item.ContextStart, item.ContextEnd) = context;
                var scale = item.Style.FontSizeMillimeters / item.Face.UnitsPerEm;
                foreach (var glyph in item.Shaped.Glyphs)
                {
                    var index = _text.IndexOfOffset(glyph.ClusterStart) - paragraph.Start;
                    advance[index] += glyph.XAdvance * scale;
                }
                if (item.Style.LetterSpacingMillimeters != 0)
                    for (var index = item.Start; index < item.End; index = NextCluster(paragraph, index))
                        if (HasSpacingAfter(paragraph, index, count)) advance[index] += item.Style.LetterSpacingMillimeters;
            }
            paragraph.MaximumEm = paragraph.Items.Count == 0 ? 0 : paragraph.Items.Max(item => item.Style.FontSizeMillimeters);
            paragraph.Prefix = new double[count + 1];
            for (var index = 0; index < count; index++) paragraph.Prefix[index + 1] = paragraph.Prefix[index] + advance[index];
        }

        private int TrimEnd(Paragraph paragraph, int start, int end)
        {
            while (end > start && (IsHardBreak(paragraph.CodePoints[end - 1]) || IsHangingSpace(paragraph.CodePoints[end - 1])))
                end--;
            return end;
        }

        // Letter spacing follows a grapheme cluster when the next cluster, before `limit`, has the same style.
        private static bool HasSpacingAfter(Paragraph paragraph, int index, int limit)
        {
            var next = NextCluster(paragraph, index);
            return next < limit && paragraph.CodePoints[index] != '\t' && paragraph.CodePoints[next] != '\t' &&
                ReferenceEquals(paragraph.StyleOf[index], paragraph.StyleOf[next]) &&
                paragraph.StyleOf[index].LetterSpacingMillimeters != 0;
        }

        // Advance of a tab at start-relative position `x`, followed by a segment `width` wide whose part
        // before its first decimal separator is `beforeDecimal` wide. Stops past the frame clamp to it.
        // A start-aligned or implicit stop must leave at least the minimum gap (Issue #115).
        private double TabAdvance(double x, double width, double beforeDecimal, double fontSize)
        {
            double target = double.NaN;
            var minimum = _options.MinimumTabAdvanceEm * fontSize;
            foreach (var stop in _options.TabStops)
            {
                if (stop.PositionMillimeters <= x + Epsilon) continue;
                if (stop.Alignment == TextTabAlignment.Start && stop.PositionMillimeters < x + minimum - Epsilon) continue;
                target = stop.Alignment switch
                {
                    TextTabAlignment.Center => stop.PositionMillimeters - width / 2,
                    TextTabAlignment.End => stop.PositionMillimeters - width,
                    TextTabAlignment.Decimal => stop.PositionMillimeters - beforeDecimal,
                    _ => stop.PositionMillimeters,
                };
                break;
            }
            if (double.IsNaN(target))
            {
                var interval = _options.DefaultTabIntervalMillimeters;
                var origin = _options.DefaultTabOriginMillimeters;
                // Zero minimum: the first stop strictly past x. Otherwise the first stop at least
                // `minimum` past x, accepting one exactly at that distance like an explicit stop.
                target = minimum <= Epsilon
                    ? origin + (Math.Floor((x - origin) / interval + Epsilon) + 1) * interval
                    : origin + Math.Ceiling((x + minimum - origin) / interval - Epsilon) * interval;
            }
            var advance = Math.Max(0, target - x);
            return Math.Max(0, Math.Min(advance, _options.WidthMillimeters - x));
        }

        private int FirstDecimal(Paragraph paragraph, int start, int end)
        {
            for (var index = start; index < end; index++)
                if (paragraph.CodePoints[index] == '.') return index;
            return end;
        }

        // Content width of [start, end) excluding hanging trailing spaces, with tabs advanced to their stops.
        private double Estimate(Paragraph paragraph, int start, int end, double indent)
        {
            var trim = TrimEnd(paragraph, start, end);
            var tabs = paragraph.Tabs;
            var tab = tabs.BinarySearch(start);
            if (tab < 0) tab = ~tab;
            if (tab >= tabs.Count || tabs[tab] >= trim) return paragraph.Prefix[trim] - paragraph.Prefix[start];
            var x = indent + paragraph.Prefix[tabs[tab]] - paragraph.Prefix[start];
            for (; tab < tabs.Count && tabs[tab] < trim; tab++)
            {
                var segmentStart = tabs[tab] + 1;
                var segmentEnd = tab + 1 < tabs.Count && tabs[tab + 1] < trim ? tabs[tab + 1] : trim;
                var width = paragraph.Prefix[segmentEnd] - paragraph.Prefix[segmentStart];
                var before = paragraph.Prefix[FirstDecimal(paragraph, segmentStart, segmentEnd)] - paragraph.Prefix[segmentStart];
                x += TabAdvance(x, width, before, paragraph.StyleOf[tabs[tab]].FontSizeMillimeters) + width;
            }
            return x - indent;
        }

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
                var c = candidate;
                for (; c < candidates.Count; c++)
                {
                    var position = candidates[c];
                    if (Estimate(paragraph, start, position, first ? _options.FirstLineIndentMillimeters : 0) <=
                        available + Epsilon) fitting.Add(position);
                    else { firstOver = position; break; }
                    if (paragraph.Breaks[position] == LineBreakKind.Mandatory) break;
                }
                // The estimate charges cross-break kerning or contextual forms to the left side, so a
                // legal break estimated just over the frame can still fit once its line is rebuilt.
                // Such shaping changes are local to the break: rebuild up to MaximumNearMisses further
                // legal candidates whose estimate exceeds the frame by at most half the paragraph's
                // largest em, before trusting the estimate.
                var nearMisses = new List<int>();
                for (; firstOver >= 0 && c < candidates.Count && nearMisses.Count < MaximumNearMisses; c++)
                {
                    var position = candidates[c];
                    if (Estimate(paragraph, start, position, first ? _options.FirstLineIndentMillimeters : 0) >
                        available + paragraph.MaximumEm / 2) break;
                    nearMisses.Add(position);
                    if (paragraph.Breaks[position] == LineBreakKind.Mandatory) break;
                }
                DraftLine? line = null;
                // Verify by building (reshaping cut runs), largest candidate first; back off on overflow.
                for (var attempt = nearMisses.Count - 1; attempt >= 0 && line is null; attempt--)
                {
                    var draft = BuildLine(paragraph, start, nearMisses[attempt], first, nearMisses[attempt] == count);
                    if (draft.ContentWidth <= available + 1e-7) line = draft;
                }
                for (var attempt = fitting.Count - 1; line is null && attempt >= 0 &&
                         fitting.Count - attempt <= MaximumFitAttempts; attempt--)
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
            // Largest grapheme boundary whose rebuilt (reshaped) line fits. The measured prefix is only a
            // heuristic (a ligature's whole advance sits on its first code point), and rebuilt widths are
            // not monotonic either: an Arabic lam-alef ligature makes "لا" narrower than "ل". So a binary
            // search on rebuilt widths only proposes a start; a bounded upward scan then tests later cuts
            // and stops after MaximumEmergencyMisses consecutive overflows (or the build budget).
            var low = 0;
            var high = boundaries.Count - 1;
            DraftLine? best = null;
            var bestIndex = -1;
            var builds = 0;
            while (low <= high && builds < MaximumFitAttempts)
            {
                builds++;
                var middle = (low + high) / 2;
                var draft = BuildLine(paragraph, start, boundaries[middle], first, boundaries[middle] == paragraph.Length);
                if (draft.ContentWidth <= available + 1e-7) { best = draft; bestIndex = middle; low = middle + 1; }
                else high = middle - 1;
            }
            var misses = 0;
            for (var index = bestIndex + 1; index < boundaries.Count && misses < MaximumEmergencyMisses &&
                     builds < 2 * MaximumFitAttempts; index++)
            {
                builds++;
                var draft = BuildLine(paragraph, start, boundaries[index], first, boundaries[index] == paragraph.Length);
                if (draft.ContentWidth <= available + 1e-7) { best = draft; misses = 0; }
                else misses++;
            }
            if (best is not null) return best;
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
                var context = ShapingContext(paragraph, item, index, runEnd, level, start, contentEnd,
                    position => levels[position - start]);
                var shaped = item.IsTab ? (level == item.Level ? item.Shaped! : TabRun(paragraph, item, level)) :
                    index == item.Start && runEnd == item.End && level == item.Level &&
                    context.Start == item.ContextStart && context.End == item.ContextEnd
                    ? item.Shaped! : Shape(paragraph, item, index, runEnd, level, context);
                logical.Add(new DraftRun
                {
                    Start = index, End = runEnd, Level = level, Item = item, Shaped = shaped,
                    Scale = item.Style.FontSizeMillimeters / item.Face.UnitsPerEm,
                    Justification = new double[shaped.Glyphs.Count],
                    Extra = new double[shaped.Glyphs.Count],
                });
                index = runEnd;
            }
            ApplyLetterSpacing(logical, paragraph, line.TrimEnd);
            ApplyTabs(logical, paragraph, line, firstOfParagraph ? _options.FirstLineIndentMillimeters : 0);
            line.Visual = Reorder(logical);
            Metrics(line, paragraph, start);
            Align(line, paragraph);
            return line;
        }

        // Letter spacing is charged per grapheme with a same-style successor on the line (as in Measure).
        // Each shaped cluster receives the gaps of every grapheme it covers, so a ligature or other
        // many-grapheme cluster keeps the total. In LTR runs a gap follows its grapheme's cluster; in RTL
        // runs (glyphs in visual order) it is placed on the cluster of the logical successor, which is the
        // visual left neighbor, so tracking widens the gap between the two graphemes rather than the run edge.
        private void ApplyLetterSpacing(List<DraftRun> runs, Paragraph paragraph, int trimEnd)
        {
            foreach (var run in runs)
            {
                var spacing = run.Item.Style.LetterSpacingMillimeters;
                if (spacing == 0 || run.IsTab) continue;
                var rtl = run.Level % 2 == 1;
                // Shaped clusters of the run: paragraph code point interval and last glyph in shaper order.
                var clusters = new Dictionary<int, (int End, int LastGlyph)>();
                for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                {
                    var glyph = run.Shaped.Glyphs[index];
                    var start = _text.IndexOfOffset(glyph.ClusterStart) - paragraph.Start;
                    var end = _text.IndexOfOffset(glyph.ClusterStart + glyph.ClusterLength) - paragraph.Start;
                    clusters[start] = (end, index);
                }
                int? ClusterOf(int codePoint)
                {
                    foreach (var pair in clusters)
                        if (codePoint >= pair.Key && codePoint < pair.Value.End) return pair.Key;
                    return null;
                }
                foreach (var pair in clusters)
                {
                    for (var grapheme = pair.Key; grapheme < pair.Value.End; grapheme = NextCluster(paragraph, grapheme))
                    {
                        if (!paragraph.GraphemeStart[grapheme] || !HasSpacingAfter(paragraph, grapheme, trimEnd)) continue;
                        var target = pair.Value.LastGlyph;
                        if (rtl && NextCluster(paragraph, grapheme) is var successor && successor >= pair.Value.End &&
                            ClusterOf(successor) is { } next)
                            target = clusters[next].LastGlyph;
                        run.Extra[target] += spacing;
                    }
                }
            }
        }

        private double LogicalAdvance(DraftRun run, Paragraph paragraph, int limit)
        {
            double total = 0;
            for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                if (_text.IndexOfOffset(run.Shaped.Glyphs[index].ClusterStart) - paragraph.Start < limit)
                    total += run.Shaped.Glyphs[index].XAdvance * run.Scale + run.Extra[index];
            return total;
        }

        // Tab advances in logical order from the line's start edge. Tabs resolve to the paragraph level
        // (UAX #9 L1), so the visual reordering keeps every segment between its tabs.
        private void ApplyTabs(List<DraftRun> runs, Paragraph paragraph, DraftLine line, double indent)
        {
            if (!runs.Any(run => run.IsTab)) return;
            var x = indent;
            for (var index = 0; index < runs.Count; index++)
            {
                var run = runs[index];
                if (!run.IsTab) { x += LogicalAdvance(run, paragraph, line.TrimEnd); continue; }
                double width = 0, before = 0;
                var decimalSeen = false;
                for (var next = index + 1; next < runs.Count && !runs[next].IsTab; next++)
                {
                    var segment = runs[next];
                    var advance = LogicalAdvance(segment, paragraph, line.TrimEnd);
                    width += advance;
                    if (decimalSeen) continue;
                    var dot = FirstDecimal(paragraph, segment.Start, segment.End);
                    if (dot < segment.End)
                    {
                        before += LogicalAdvance(segment, paragraph, dot);
                        decimalSeen = true;
                    }
                    else before += advance;
                }
                run.Extra[0] = TabAdvance(x, width, before, run.Item.Style.FontSizeMillimeters);
                x += run.Extra[0];
            }
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
                // An empty paragraph uses the style of its own terminator (at paragraph.Start); only a
                // completely empty request falls back to its single empty span.
                var style = paragraph.Length > 0 ? paragraph.StyleOf[Math.Min(start, paragraph.Length - 1)]
                    : paragraph.Start < _styleByCodePoint.Length ? _styleByCodePoint[paragraph.Start]
                    : _styleByCodePoint.Length > 0 ? _styleByCodePoint[_styleByCodePoint.Length - 1]
                    : _request.Spans[0].Style;
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
                case "AtLeast":
                    // Extra space goes below the text, as in the compatibility layout and word processors'
                    // proportional ("multiple") line spacing, so a taller line never pushes ink downward.
                    line.Height = Math.Max(normal, rule.Value);
                    line.BaselineOffset = gapAscent;
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
                for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                {
                    var advance = run.Advance(index);
                    total += advance;
                    if (run.Shaped.Glyphs[index].ClusterStart >= trimOffset) hanging += advance;
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
            // Text before the line's last tab is aligned by its tab stops, so it is never stretched.
            var lastTab = paragraph.Tabs.LastOrDefault(tab => tab >= line.Start && tab < line.TrimEnd);
            var justifyFrom = paragraph.Tabs.Any(tab => tab >= line.Start && tab < line.TrimEnd) ? lastTab + 1 : line.Start;
            for (var index = justifyFrom; index < line.TrimEnd;)
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
            var trimOffset = Offset(paragraph, draft.TrimEnd);
            var runs = new List<ParagraphLayoutRun>(draft.Visual.Count);
            Bounds? lineInk = null;
            foreach (var run in draft.Visual)
            {
                var origin = pen;
                var glyphs = new List<ParagraphLayoutGlyph>(run.Shaped.Glyphs.Count);
                var resolved = run.Item.Resolved;
                var stroke = resolved.SyntheticBoldStrokeMillimeters / 2;
                var skew = resolved.ObliqueSkew;
                Bounds? runInk = null;
                double? contentLeft = null, contentRight = null;
                double penY = 0;
                for (var index = 0; index < run.Shaped.Glyphs.Count; index++)
                {
                    var glyph = run.Shaped.Glyphs[index];
                    var x = Check(pen + glyph.XOffset * run.Scale);
                    var y = Check(baseline - (penY + glyph.YOffset) * run.Scale);
                    var advance = run.Advance(index);
                    TextLayoutBounds? ink = null;
                    if (glyph.Ink is { } value)
                    {
                        var left = x + value.X * run.Scale - stroke;
                        var right = x + (value.X + value.Width) * run.Scale + stroke;
                        var inkTop = y - (value.Y + value.Height) * run.Scale - stroke;
                        var inkBottom = y - value.Y * run.Scale + stroke;
                        if (skew != 0)
                        {
                            // Oblique shears about the baseline: x += skew × (baseline − y).
                            var shifts = new[] { skew * (baseline - inkTop), skew * (baseline - inkBottom) };
                            left += shifts.Min(); right += shifts.Max();
                        }
                        ink = new TextLayoutBounds(Check(left), Check(inkTop), right - left, inkBottom - inkTop);
                        runInk = Bounds.Union(runInk, ink);
                    }
                    if (glyph.ClusterStart < trimOffset)
                    {
                        contentLeft = Math.Min(contentLeft ?? pen, pen);
                        contentRight = Math.Max(contentRight ?? pen + advance, pen + advance);
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
                var decorations = contentLeft is { } decorationLeft && contentRight is { } decorationRight &&
                    decorationRight > decorationLeft
                    ? Decorate(run, resolved, decorationLeft, decorationRight - decorationLeft, baseline)
                    : null;
                runs.Add(new ParagraphLayoutRun(offset, text, run.Shaped, run.Level, resolved, origin, baseline,
                    pen - origin, glyphs, runBounds, decorations, run.IsTab));
            }
            var lineStart = Offset(paragraph, draft.Start);
            var lineEnd = Offset(paragraph, draft.End) + (draft.EndsParagraph ? paragraph.TerminatorLength : 0);
            return new ParagraphLayoutLine(lineStart, lineEnd - lineStart, draft.TerminatorLength, paragraph.Index,
                pageIndex, paragraph.Bidi.ParagraphLevel, draft.EndsParagraph, draft.HardBreak, top, baseline,
                draft.Height, draft.Ascent, draft.Descent, draft.ContentLeft, draft.ContentWidth, runs, lineInk?.ToBounds());
        }

        // Highlight spans the ascent and descent, underline and strikethrough use the position and thickness,
        // of the style's primary face (its first family), so a decoration stays level across font fallback
        // inside one style. All are filled rectangles over the run's non-hanging advance.
        private static List<ParagraphLayoutDecoration>? Decorate(DraftRun run, TextResolvedStyle resolved, double left,
            double width, double baseline)
        {
            if (resolved.HighlightColor is null && resolved.Decorations == TextDecorationLines.None) return null;
            var style = run.Item.Style;
            var primary = style.FontFallback[0];
            var face = primary.Select(style.Bold, style.Italic) ?? primary.Select(false, false) ?? primary.Faces[0];
            var scale = style.FontSizeMillimeters / face.UnitsPerEm;
            var result = new List<ParagraphLayoutDecoration>(3);
            if (resolved.HighlightColor is { } highlight)
                result.Add(new ParagraphLayoutDecoration(ParagraphLayoutDecorationKind.Highlight,
                    new TextLayoutBounds(Check(left), Check(baseline - face.Ascender * scale), width,
                        (face.Ascender - face.Descender) * scale), highlight));
            var color = resolved.DecorationColor;
            if (color is not null && (resolved.Decorations & TextDecorationLines.Underline) != 0)
                result.Add(new ParagraphLayoutDecoration(ParagraphLayoutDecorationKind.Underline,
                    new TextLayoutBounds(Check(left), Check(baseline - face.UnderlinePosition * scale), width,
                        face.UnderlineThickness * scale), color));
            if (color is not null && (resolved.Decorations & TextDecorationLines.Strikethrough) != 0)
                result.Add(new ParagraphLayoutDecoration(ParagraphLayoutDecorationKind.Strikethrough,
                    new TextLayoutBounds(Check(left), Check(baseline - face.StrikeoutPosition * scale), width,
                        face.StrikeoutThickness * scale), color));
            return result;
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
