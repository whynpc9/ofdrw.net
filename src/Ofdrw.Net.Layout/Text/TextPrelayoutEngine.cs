using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>A bounded new-text query. The first profile supports one static face per logical run and horizontal LTR line layout.</summary>
public sealed class TextPrelayoutRequest
{
    /// <summary>Creates a bounded line query with an ordered list of explicit whole-run font choices.</summary>
    public TextPrelayoutRequest(string text, IReadOnlyList<TextFontFace> faces, double fontSizeMillimeters,
        double widthMillimeters, double? maximumHeightMillimeters = null, string script = "Latn", string language = "en",
        IReadOnlyList<TextFeature>? features = null, int maximumLines = 256, int maximumShapeCalls = 4096)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        if (faces is null || faces.Count == 0 || faces.Count > 16 || faces.Any(face => face is null))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Supply one to sixteen explicit font faces.");
        if (!FinitePositive(fontSizeMillimeters) || fontSizeMillimeters > 1000 ||
            !FinitePositive(widthMillimeters) || widthMillimeters > 1_000_000 ||
            maximumHeightMillimeters.HasValue && (!FinitePositive(maximumHeightMillimeters.Value) ||
                maximumHeightMillimeters.Value > 1_000_000))
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Font size, width, and optional height must be finite positive millimeters.");
        if (maximumLines < 1 || maximumLines > 256 || maximumShapeCalls < 1 || maximumShapeCalls > 8192)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid line or shaping-call budget.");
        if (text.Length > 4096) throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Prelayout exceeds 4096 UTF-16 code units.");
        if (script != "Latn" && script != "Hani")
            throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                "The initial paragraph query supports Latin or Han runs; other scripts require a separate run profile.");
        Faces = new ReadOnlyCollection<TextFontFace>(faces.ToArray());
        FontSizeMillimeters = fontSizeMillimeters; WidthMillimeters = widthMillimeters;
        MaximumHeightMillimeters = maximumHeightMillimeters; MaximumLines = maximumLines; MaximumShapeCalls = maximumShapeCalls;
        Script = script; Language = language;
        Features = new ReadOnlyCollection<TextFeature>(features?.ToArray() ?? Array.Empty<TextFeature>());
        // Check script, language, and the original UTF-16 before any native call.
        _ = new TextShapeRequest(string.Empty, 0, Faces[0], script: script, language: language);
        for (var i = 0; i < text.Length; i++)
        {
            if (char.IsHighSurrogate(text[i]))
            {
                if (i + 1 >= text.Length || !char.IsLowSurrogate(text[i + 1]))
                    throw new TextLayoutException(TextLayoutError.InvalidText, "Unpaired UTF-16 surrogate.");
                if (char.ConvertToUtf32(text, i) is >= 0xe0100 and <= 0xe01ef)
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "Variation selectors need a separately verified text profile.",
                        sourceStart: i, sourceLength: 2);
                i++;
            }
            else if (char.IsLowSurrogate(text[i]) || text[i] == '\t' || TextShapeRequest.IsUnsupportedLineSeparator(text[i]) || text[i] == '\r' &&
                (i + 1 >= text.Length || text[i + 1] != '\n'))
                throw new TextLayoutException(TextLayoutError.InvalidText, "Unsupported control or unpaired UTF-16 surrogate.",
                    sourceStart: i, sourceLength: 1);
            else if (char.GetUnicodeCategory(text[i]) == UnicodeCategory.Format ||
                text[i] is >= '\ufe00' and <= '\ufe0f')
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Format controls and variation selectors need a separately verified text profile.",
                    sourceStart: i, sourceLength: 1);
        }
        foreach (var feature in Features)
            if (feature is null || feature.Start > text.Length || feature.Length > text.Length - feature.Start)
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "Feature interval exceeds the original text.");
    }
    /// <summary>Unmodified logical Unicode source, including LF or CRLF line separators.</summary>
    public string Text { get; }
    /// <summary>Ordered immutable font choices; fallback selects a single face for the complete run.</summary>
    public IReadOnlyList<TextFontFace> Faces { get; }
    /// <summary>Requested em size in millimeters.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Available advance width in millimeters.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Optional available height for overflow reporting.</summary>
    public double? MaximumHeightMillimeters { get; }
    /// <summary>Four-letter script tag for the uniform run profile.</summary>
    public string Script { get; }
    /// <summary>BCP 47 language tag for the uniform run profile.</summary>
    public string Language { get; }
    /// <summary>Immutable feature intervals relative to the full source text.</summary>
    public IReadOnlyList<TextFeature> Features { get; }
    /// <summary>Maximum number of final lines before failure.</summary>
    public int MaximumLines { get; }
    /// <summary>Maximum number of full-run native shaping calls before failure.</summary>
    public int MaximumShapeCalls { get; }
    private static bool FinitePositive(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>A rectangle in millimeters with Y increasing down the page.</summary>
public sealed class TextLayoutBounds
{
    /// <summary>Creates a page-coordinate ink rectangle in millimeters.</summary>
    public TextLayoutBounds(double x, double y, double width, double height)
    {
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y) ||
            double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height) ||
            width < 0 || height < 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid page ink bounds.");
        X = x; Y = y; Width = width; Height = height;
    }
    /// <summary>Left edge relative to the laid-out line.</summary>
    public double X { get; }
    /// <summary>Top edge with Y increasing down the page.</summary>
    public double Y { get; }
    /// <summary>Visible ink width, independent of the advance.</summary>
    public double Width { get; }
    /// <summary>Visible ink height, independent of line height.</summary>
    public double Height { get; }
}

/// <summary>A nonfatal, source-addressed decision made during prelayout.</summary>
public sealed class TextLayoutDiagnostic
{
    /// <summary>Creates a diagnostic with an absolute UTF-16 source interval.</summary>
    public TextLayoutDiagnostic(string code, int sourceStart, int sourceLength, string detail)
    { Code = code; SourceStart = sourceStart; SourceLength = sourceLength; Detail = detail; }
    /// <summary>Stable machine-readable diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Absolute UTF-16 source start.</summary>
    public int SourceStart { get; }
    /// <summary>UTF-16 source interval length.</summary>
    public int SourceLength { get; }
    /// <summary>Human-readable detail, including selected face identity when relevant.</summary>
    public string Detail { get; }
}

/// <summary>Final line geometry and the exact shaped run used to measure it.</summary>
public sealed class TextLayoutLine
{
    internal TextLayoutLine(int sourceStart, int sourceLength, int terminatorLength, double top, double baseline,
        double height, double advance, TextLayoutBounds? ink, TextShapedRun run, bool overflow)
    {
        SourceStart = sourceStart; SourceLength = sourceLength; TerminatorLength = terminatorLength;
        TopMillimeters = top; BaselineMillimeters = baseline; HeightMillimeters = height;
        AdvanceMillimeters = advance; InkBoundsMillimeters = ink; ShapedRun = run; OverflowsWidth = overflow;
    }
    /// <summary>Absolute UTF-16 start of visible line source.</summary>
    public int SourceStart { get; }
    /// <summary>UTF-16 length of visible line source, excluding its terminator.</summary>
    public int SourceLength { get; }
    /// <summary>UTF-16 length of a following LF or CRLF, if present.</summary>
    public int TerminatorLength { get; }
    /// <summary>Line top in millimeters relative to the query origin.</summary>
    public double TopMillimeters { get; }
    /// <summary>Baseline in millimeters relative to the query origin.</summary>
    public double BaselineMillimeters { get; }
    /// <summary>Reserved line height in millimeters.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Horizontal advance width in millimeters.</summary>
    public double AdvanceMillimeters { get; }
    /// <summary>Visible ink rectangle, or null when the line has no ink.</summary>
    public TextLayoutBounds? InkBoundsMillimeters { get; }
    /// <summary>The exact final shaped run used for this line's measurement.</summary>
    public TextShapedRun ShapedRun { get; }
    /// <summary>Whether the final visible advance, excluding trailing breakable spaces, exceeds the available width.</summary>
    public bool OverflowsWidth { get; }
}

/// <summary>Immutable prelayout lines, selected face, dimensions, overflow, and diagnostics.</summary>
public sealed class TextPrelayoutResult
{
    internal TextPrelayoutResult(TextPrelayoutRequest request, TextFontFace face, List<TextLayoutLine> lines,
        List<TextLayoutDiagnostic> diagnostics)
    {
        Request = request; Face = face;
        Lines = new ReadOnlyCollection<TextLayoutLine>(lines);
        Diagnostics = new ReadOnlyCollection<TextLayoutDiagnostic>(diagnostics);
        WidthMillimeters = lines.Count == 0 ? 0 : lines.Max(line => line.AdvanceMillimeters);
        HeightMillimeters = lines.Sum(line => line.HeightMillimeters);
        Overflows = lines.Any(line => line.OverflowsWidth) ||
            request.MaximumHeightMillimeters.HasValue && HeightMillimeters > request.MaximumHeightMillimeters.Value;
    }
    /// <summary>The logical source and available geometry that produced this result.</summary>
    public TextPrelayoutRequest Request { get; }
    /// <summary>The actual immutable face selected for the whole run.</summary>
    public TextFontFace Face { get; }
    /// <summary>Final lines in logical source order.</summary>
    public IReadOnlyList<TextLayoutLine> Lines { get; }
    /// <summary>Nonfatal fallback and profile diagnostics.</summary>
    public IReadOnlyList<TextLayoutDiagnostic> Diagnostics { get; }
    /// <summary>Maximum final line advance in millimeters.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Sum of final reserved line heights in millimeters.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Whether any final line exceeds width or total height exceeds the optional limit.</summary>
    public bool Overflows { get; }
}

/// <summary>Queries final line positions from one replaceable run shaper. It never silently uses legacy approximate metrics.</summary>
public sealed class TextPrelayoutEngine
{
    private readonly ITextShaper _shaper;
    /// <summary>Creates a query engine with an explicit, replaceable single-run shaper.</summary>
    public TextPrelayoutEngine(ITextShaper shaper) => _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));

    /// <summary>Selects one supplied face, shapes complete candidate lines, and returns only final results.</summary>
    public TextPrelayoutResult Layout(TextPrelayoutRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        var scalars = Scalars(request.Text, cancellationToken).ToArray();
        var uniqueScalars = scalars.Select(value => value.Scalar).Distinct().ToArray();
        var coverage = new Dictionary<(TextFontFace Face, int Scalar), bool>();
        const long maximumCoverageWork = 8_000_000;
        long coverageWork = 0;
        bool Covers(TextFontFace candidate, int scalar)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (candidate, scalar);
            if (coverage.TryGetValue(key, out var cached)) return cached;
            var cost = candidate.CoverageLookupWorkUpperBound;
            if (coverageWork > maximumCoverageWork - cost)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Prelayout font coverage lookup budget exceeded.");
            coverageWork += cost;
            var covered = candidate.Covers(scalar);
            coverage.Add(key, covered);
            return covered;
        }
        TextFontFace? face = null;
        foreach (var candidate in request.Faces)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!uniqueScalars.All(scalar => Covers(candidate, scalar))) continue;
            face = candidate;
            break;
        }
        if (face is null)
        {
            foreach (var value in scalars)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (request.Faces.Any(candidate => Covers(candidate, value.Scalar))) continue;
                throw new TextLayoutException(TextLayoutError.MissingGlyph,
                    "No supplied face covers the scalar at UTF-16 offset " + value.Offset + ".",
                    sourceStart: value.Offset, sourceLength: value.Scalar > 0xffff ? 2 : 1);
            }
            throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                "The text needs cluster-level mixed-face fallback, which this whole-run profile does not provide.");
        }
        var diagnostics = new List<TextLayoutDiagnostic>();
        if (!ReferenceEquals(face, request.Faces[0]))
            diagnostics.Add(new TextLayoutDiagnostic("TEXT_FACE_FALLBACK", 0, request.Text.Length,
                "Whole-run fallback selected face " + face.FaceSha256 + "."));
        var scale = request.FontSizeMillimeters / face.UnitsPerEm;
        var height = Math.Max(face.UnitsPerEm, face.Ascender - face.Descender + face.LineGap) * scale;
        var baselineOffset = face.Ascender * scale;
        var lines = new List<TextLayoutLine>();
        var calls = 0;
        const long maximumCumulativeShapeWork = 512L * 1024 * 1024;
        long cumulativeShapeWork = 0;
        TextShapedRun Shape(int start, int length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++calls > request.MaximumShapeCalls)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Prelayout shaping-call budget exceeded.");
            // The pinned adapter copies each nonempty face into managed and native
            // storage per call. Charge that traffic and weighted text work before
            // invoking any replacement shaper, including reshaping for final lines.
            var work = (length == 0 ? 0L : 2L * face.FaceByteLength) + 64L * length;
            if (cumulativeShapeWork > maximumCumulativeShapeWork - work)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Prelayout cumulative shaping-work budget exceeded.");
            cumulativeShapeWork += work;
            var features = request.Features.Where(feature => feature.Start < start + length &&
                feature.Start + feature.Length > start).Select(feature =>
                new TextFeature(feature.Tag, feature.Value, Math.Max(0, feature.Start - start),
                    Math.Min(length, feature.Start + feature.Length - start) - Math.Max(0, feature.Start - start))).ToArray();
            var shapeRequest = new TextShapeRequest(request.Text.Substring(start, length), start, face,
                script: request.Script, language: request.Language, features: features);
            var shaped = _shaper.Shape(shapeRequest, cancellationToken);
            if (shaped is null || !ReferenceEquals(shaped.Request, shapeRequest) || shaped.XAdvance < 0)
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "Shaper returned a result for a different run or invalid direction.");
            return shaped;
        }
        double FitAdvance(int start, int end, TextShapedRun run)
        {
            if (run.XAdvance * scale <= request.WidthMillimeters + 1e-8) return run.XAdvance;
            var visibleEnd = end;
            while (visibleEnd > start && IsBreakableWhitespace(request.Text[visibleEnd - 1])) visibleEnd--;
            return visibleEnd == end ? run.XAdvance : Shape(start, visibleEnd - start).XAdvance;
        }
        void AddLine(int start, int length, int terminatorLength, TextShapedRun run, double fitAdvance)
        {
            if (lines.Count >= request.MaximumLines)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Prelayout line budget exceeded.");
            var top = lines.Count * height;
            var baseline = top + baselineOffset;
            var advance = run.XAdvance * scale;
            double? minX = null, maxX = null, minY = null, maxY = null;
            double penX = 0, penY = 0;
            foreach (var glyph in run.Glyphs)
            {
                if (glyph.Ink is not null)
                {
                    var left = (penX + glyph.XOffset + glyph.Ink.X) * scale;
                    var right = left + glyph.Ink.Width * scale;
                    var bottom = baseline - (penY + glyph.YOffset + glyph.Ink.Y) * scale;
                    var inkTop = bottom - glyph.Ink.Height * scale;
                    minX = Math.Min(minX ?? left, left); maxX = Math.Max(maxX ?? right, right);
                    minY = Math.Min(minY ?? inkTop, inkTop); maxY = Math.Max(maxY ?? bottom, bottom);
                }
                penX += glyph.XAdvance; penY += glyph.YAdvance;
            }
            var ink = minX.HasValue ? new TextLayoutBounds(minX.Value, minY!.Value,
                maxX!.Value - minX.Value, maxY!.Value - minY.Value) : null;
            lines.Add(new TextLayoutLine(start, length, terminatorLength, top, baseline, height,
                advance, ink, run, fitAdvance * scale > request.WidthMillimeters + 1e-8));
        }
        var paragraphStart = 0;
        while (paragraphStart <= request.Text.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var newline = request.Text.IndexOf('\n', paragraphStart);
            var paragraphEnd = newline < 0 ? request.Text.Length : newline > paragraphStart &&
                request.Text[newline - 1] == '\r' ? newline - 1 : newline;
            var terminator = newline < 0 ? 0 : newline - paragraphEnd + 1;
            var cursor = paragraphStart;
            if (cursor == paragraphEnd) AddLine(cursor, 0, terminator, Shape(cursor, 0), 0);
            while (cursor < paragraphEnd)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var opportunities = BreakOpportunities(request.Text, cursor, paragraphEnd);
                var fittingEnd = -1;
                var selectedEnd = paragraphEnd;
                TextShapedRun? selected = null;
                double selectedFitAdvance = 0;
                foreach (var end in opportunities)
                {
                    var candidate = Shape(cursor, end - cursor);
                    var fitAdvance = FitAdvance(cursor, end, candidate);
                    if (fitAdvance * scale <= request.WidthMillimeters + 1e-8)
                    { fittingEnd = end; selected = candidate; selectedFitAdvance = fitAdvance; continue; }
                    selectedEnd = fittingEnd > cursor ? fittingEnd : end;
                    if (fittingEnd <= cursor) { selected = candidate; selectedFitAdvance = fitAdvance; }
                    break;
                }
                if (selected is null) throw new TextLayoutException(TextLayoutError.InvalidRequest, "No line break candidate was available.");
                if (selectedEnd == paragraphEnd && fittingEnd == paragraphEnd) { /* complete paragraph */ }
                else if (fittingEnd > cursor) selectedEnd = fittingEnd;
                if (selected.Request.Text.Length != selectedEnd - cursor)
                {
                    selected = Shape(cursor, selectedEnd - cursor);
                    selectedFitAdvance = FitAdvance(cursor, selectedEnd, selected);
                }
                AddLine(cursor, selectedEnd - cursor, selectedEnd == paragraphEnd ? terminator : 0,
                    selected, selectedFitAdvance);
                cursor = selectedEnd;
            }
            if (newline < 0) break;
            paragraphStart = newline + 1;
        }
        return new TextPrelayoutResult(request, face, lines, diagnostics);
    }

    private static IEnumerable<int> BreakOpportunities(string text, int start, int end)
    {
        var part = text.Substring(start, end - start);
        var enumerator = StringInfo.GetTextElementEnumerator(part);
        var elements = new List<(int After, int Scalar, char Last)>();
        while (enumerator.MoveNext())
        {
            var element = enumerator.GetTextElement();
            var after = start + enumerator.ElementIndex + element.Length;
            var scalar = char.IsHighSurrogate(element[0]) && element.Length > 1 ? char.ConvertToUtf32(element, 0) : element[0];
            elements.Add((after, scalar, element[element.Length - 1]));
        }
        for (var i = 0; i < elements.Count; i++)
        {
            var current = elements[i];
            if (current.After == end) { yield return end; continue; }
            var next = elements[i + 1].Scalar;
            if (IsNoBreakSpace(current.Scalar) || IsNoBreakSpace(next) ||
                IsNoStartPunctuation(next) || IsOpeningPunctuation(current.Scalar)) continue;
            if (IsBreakableWhitespace(current.Last) || IsCjk(current.Scalar) || IsHan(next))
                yield return current.After;
        }
    }

    private static bool IsBreakableWhitespace(char value) => char.IsWhiteSpace(value) &&
        !IsNoBreakSpace(value);

    private static bool IsNoBreakSpace(int scalar) => scalar is 0x00a0 or 0x2007 or 0x202f;

    private static bool IsNoStartPunctuation(int scalar) =>
        scalar is '、' or '。' or '，' or '．' or '：' or '；' or '！' or '？' or
            '）' or '］' or '｝' or '》' or '」' or '』' or '】' or '〗' or '〙' or '〛' or
            ')' or ']' or '}' or ',' or '.' or ':' or ';' or '!' or '?';

    private static bool IsOpeningPunctuation(int scalar) =>
        scalar is '（' or '［' or '｛' or '《' or '「' or '『' or '【' or '〖' or '〘' or '〚' or
            '(' or '[' or '{';

    private static bool IsCjk(int scalar) => scalar is >= 0x3000 and <= 0x33ff or
        >= 0xac00 and <= 0xd7af || IsHan(scalar);

    private static bool IsHan(int scalar) => scalar is >= 0x3400 and <= 0x9fff or
        >= 0xf900 and <= 0xfaff or
        >= 0x20000 and <= 0x2a6df or >= 0x2a700 and <= 0x2ee5f or
        >= 0x2f800 and <= 0x2fa1f or >= 0x30000 and <= 0x3347f;

    private static IEnumerable<(int Scalar, int Offset)> Scalars(string text, CancellationToken cancellationToken)
    {
        for (var i = 0; i < text.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (text[i] == '\n' || text[i] == '\r') continue;
            if (char.IsHighSurrogate(text[i])) { yield return (char.ConvertToUtf32(text, i), i); i++; }
            else yield return (text[i], i);
        }
    }
}
