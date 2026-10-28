using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>A deliberately bounded LTR paragraph with Latin and Hebrew words.</summary>
public sealed class TextMixedParagraphRequest
{
    /// <summary>Creates a short horizontal query using one immutable static face.</summary>
    public TextMixedParagraphRequest(string text, TextFontFace face, double fontSizeMillimeters,
        double widthMillimeters, int maximumLines = 32, int maximumShapeCalls = 128)
    {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Face = face ?? throw new ArgumentNullException(nameof(face));
        if (!FinitePositive(fontSizeMillimeters) || fontSizeMillimeters > 1000 ||
            !FinitePositive(widthMillimeters) || widthMillimeters > 1_000_000 ||
            maximumLines is < 1 or > 32 || maximumShapeCalls is < 1 or > 128)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid mixed-paragraph geometry or budget.");
        if (text.Length is < 3 or > 128)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                "A mixed paragraph needs three to 128 BMP source characters.");
        if (text[0] == ' ' || text[text.Length - 1] == ' ')
            throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                "Leading or trailing spaces exceed the mixed-paragraph profile.");
        var latin = false; var hebrew = false; var previousSpace = false;
        for (var index = 0; index < text.Length; index++)
        {
            var value = text[index];
            if (value is >= 'A' and <= 'Z' or >= 'a' and <= 'z') latin = true;
            else if (value is >= '\u05d0' and <= '\u05ea') hebrew = true;
            else if (value != ' ' || previousSpace)
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Only ASCII Latin letters, Hebrew letters and single ASCII word spaces are admitted.",
                    sourceStart: index, sourceLength: 1);
            previousSpace = value == ' ';
        }
        if (!latin || !hebrew)
            throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                "The bounded mixed paragraph needs both Latin and Hebrew words.");
        FontSizeMillimeters = fontSizeMillimeters;
        WidthMillimeters = widthMillimeters;
        MaximumLines = maximumLines;
        MaximumShapeCalls = maximumShapeCalls;
    }

    /// <summary>Unmodified logical UTF-16 paragraph source.</summary>
    public string Text { get; }
    /// <summary>One selected face used by every run.</summary>
    public TextFontFace Face { get; }
    /// <summary>Requested em size in millimeters.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>Available visible line width.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Maximum accepted final line count.</summary>
    public int MaximumLines { get; }
    /// <summary>Maximum native shape calls, including final reshaping.</summary>
    public int MaximumShapeCalls { get; }

    private static bool FinitePositive(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>One final script run or one inter-word space in logical source order.</summary>
public sealed class TextMixedParagraphRun
{
    internal TextMixedParagraphRun(TextShapedRun shaped, double xMillimeters)
    { Shaped = shaped; XMillimeters = xMillimeters; }
    /// <summary>Exact final shaped result, with absolute source clusters and visual glyph order.</summary>
    public TextShapedRun Shaped { get; }
    /// <summary>Horizontal start of this run on its line.</summary>
    public double XMillimeters { get; }
    /// <summary>Absolute UTF-16 start in the original paragraph.</summary>
    public int SourceStart => Shaped.Request.SourceStart;
    /// <summary>Unchanged logical text of the run.</summary>
    public string Text => Shaped.Request.Text;
    /// <summary>Final run direction.</summary>
    public TextDirection Direction => Shaped.Request.Direction;
}

/// <summary>A final line with word-safe source boundaries and exact reshaped runs.</summary>
public sealed class TextMixedParagraphLine
{
    internal TextMixedParagraphLine(int sourceStart, int sourceLength, double topMillimeters,
        double baselineMillimeters, double visibleWidthMillimeters, double totalAdvanceMillimeters,
        IReadOnlyList<TextMixedParagraphRun> runs)
    {
        SourceStart = sourceStart; SourceLength = sourceLength;
        TopMillimeters = topMillimeters; BaselineMillimeters = baselineMillimeters;
        VisibleWidthMillimeters = visibleWidthMillimeters;
        TotalAdvanceMillimeters = totalAdvanceMillimeters;
        Runs = new ReadOnlyCollection<TextMixedParagraphRun>(runs.ToArray());
    }
    /// <summary>Absolute UTF-16 start after any preceding line separator.</summary>
    public int SourceStart { get; }
    /// <summary>Source length, including a trailing separator at a soft break.</summary>
    public int SourceLength { get; }
    /// <summary>Top of this line relative to the query origin.</summary>
    public double TopMillimeters { get; }
    /// <summary>Baseline relative to the query origin.</summary>
    public double BaselineMillimeters { get; }
    /// <summary>Advance excluding a trailing break space.</summary>
    public double VisibleWidthMillimeters { get; }
    /// <summary>Advance including every source character on the line.</summary>
    public double TotalAdvanceMillimeters { get; }
    /// <summary>Runs in logical order; glyphs within Hebrew runs are in visual order.</summary>
    public IReadOnlyList<TextMixedParagraphRun> Runs { get; }
}

/// <summary>Completed bounded paragraph query; no output is created on failure.</summary>
public sealed class TextMixedParagraphResult
{
    internal TextMixedParagraphResult(TextMixedParagraphRequest request, IReadOnlyList<TextMixedParagraphLine> lines)
    {
        Request = request;
        Lines = new ReadOnlyCollection<TextMixedParagraphLine>(lines.ToArray());
        WidthMillimeters = lines.Max(line => line.VisibleWidthMillimeters);
        HeightMillimeters = lines.Count * Math.Max(request.Face.UnitsPerEm,
            request.Face.Ascender - request.Face.Descender + request.Face.LineGap) *
            request.FontSizeMillimeters / request.Face.UnitsPerEm;
    }
    /// <summary>Original request and immutable face.</summary>
    public TextMixedParagraphRequest Request { get; }
    /// <summary>Final lines in logical paragraph order.</summary>
    public IReadOnlyList<TextMixedParagraphLine> Lines { get; }
    /// <summary>Maximum visible line advance.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Total reserved line height.</summary>
    public double HeightMillimeters { get; }
}

/// <summary>Analyzes one constrained LTR mixed paragraph, chooses word breaks and reshapes final runs.</summary>
public sealed class TextMixedParagraphEngine
{
    private readonly ITextShaper _shaper;
    /// <summary>Uses the supplied shaper for both candidate measurement and final runs.</summary>
    public TextMixedParagraphEngine(ITextShaper shaper) => _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));

    /// <summary>Returns a complete result or a source-addressed failure.</summary>
    public TextMixedParagraphResult Layout(TextMixedParagraphRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        var text = request.Text;
        const long maximumCoverageWork = 8_000_000;
        long coverageWork = 0;
        var coverage = new Dictionary<char, bool>();
        bool Covers(char value)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (coverage.TryGetValue(value, out var cached)) return cached;
            var cost = request.Face.CoverageLookupWorkUpperBound;
            if (coverageWork > maximumCoverageWork - cost)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Mixed paragraph font coverage lookup budget exceeded.");
            coverageWork += cost;
            var result = request.Face.Covers(value);
            coverage.Add(value, result);
            return result;
        }
        for (var index = 0; index < text.Length; index++)
            if (!Covers(text[index]))
                throw new TextLayoutException(TextLayoutError.MissingGlyph,
                    "The selected face does not cover the mixed paragraph at UTF-16 offset " + index + ".",
                    sourceStart: index, sourceLength: 1);
        var words = new List<(int Start, int Length, bool Hebrew)>();
        for (var start = 0; start < text.Length;)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var end = text.IndexOf(' ', start);
            if (end < 0) end = text.Length;
            var hebrew = text[start] is >= '\u05d0' and <= '\u05ea';
            for (var index = start + 1; index < end; index++)
                if ((text[index] is >= '\u05d0' and <= '\u05ea') != hebrew)
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "A script change inside a word needs a separate bidi profile.",
                        sourceStart: index, sourceLength: 1);
            words.Add((start, end - start, hebrew));
            start = end + 1;
        }
        if (words.Count > 32)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Mixed paragraph exceeds 32 words.");
        var scale = request.FontSizeMillimeters / request.Face.UnitsPerEm;
        var height = Math.Max(request.Face.UnitsPerEm,
            request.Face.Ascender - request.Face.Descender + request.Face.LineGap) * scale;
        var baselineOffset = request.Face.Ascender * scale;
        var calls = 0; long work = 0;
        TextShapedRun Shape(int start, int length, bool hebrew)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++calls > request.MaximumShapeCalls)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Mixed paragraph exceeded the shaping-call budget.");
            var cost = 2L * request.Face.FaceByteLength + 64L * length;
            if (work > 512L * 1024 * 1024 - cost)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Mixed paragraph exceeded the cumulative shaping-work budget.");
            work += cost;
            var shapeRequest = new TextShapeRequest(text.Substring(start, length), start,
                request.Face, hebrew ? TextDirection.RightToLeft : TextDirection.LeftToRight,
                hebrew ? "Hebr" : "Latn", hebrew ? "he" : "en");
            var result = _shaper.Shape(shapeRequest, cancellationToken);
            if (result is null || !ReferenceEquals(result.Request, shapeRequest) || result.XAdvance < 0 || result.YAdvance != 0)
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "Mixed paragraph shaper returned a mismatched or vertical run.");
            return result;
        }
        var spaceAdvances = new double[words.Count - 1];
        var wordAdvances = new double[words.Count];
        for (var index = 0; index < words.Count; index++)
        {
            var word = words[index];
            wordAdvances[index] = Shape(word.Start, word.Length, word.Hebrew).XAdvance * scale;
            if (index < words.Count - 1)
                spaceAdvances[index] = Shape(word.Start + word.Length, 1, false).XAdvance * scale;
        }
        var lines = new List<TextMixedParagraphLine>();
        for (var first = 0; first < words.Count;)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var visible = 0d;
            var last = first;
            for (var index = first; index < words.Count; index++)
            {
                var candidate = visible + (index == first ? 0 : spaceAdvances[index - 1]) + wordAdvances[index];
                if (index > first && candidate > request.WidthMillimeters + 1e-8) break;
                visible = candidate;
                last = index;
            }
            if (visible > request.WidthMillimeters + 1e-8)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                    "A mixed paragraph word exceeds the available line width.",
                    sourceStart: words[first].Start, sourceLength: words[first].Length);
            if (lines.Count >= request.MaximumLines)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                    "Mixed paragraph exceeded the final-line budget.");
            var runList = new List<TextMixedParagraphRun>();
            var x = 0d;
            for (var index = first; index <= last; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var word = words[index];
                var shaped = Shape(word.Start, word.Length, word.Hebrew);
                runList.Add(new TextMixedParagraphRun(shaped, x));
                x += shaped.XAdvance * scale;
                if (index < last || last < words.Count - 1 && index == last)
                {
                    var space = Shape(word.Start + word.Length, 1, false);
                    runList.Add(new TextMixedParagraphRun(space, x));
                    x += space.XAdvance * scale;
                }
            }
            var sourceStart = words[first].Start;
            var sourceEnd = last == words.Count - 1 ? text.Length : words[last].Start + words[last].Length + 1;
            var visibleAdvance = last == words.Count - 1 ? x : x - runList[runList.Count - 1].Shaped.XAdvance * scale;
            if (visibleAdvance > request.WidthMillimeters + 1e-8)
                throw new TextLayoutException(TextLayoutError.UnsupportedOutput,
                    "Final reshaping exceeded the available line width.");
            lines.Add(new TextMixedParagraphLine(sourceStart, sourceEnd - sourceStart,
                lines.Count * height, lines.Count * height + baselineOffset,
                visibleAdvance, x, runList));
            first = last + 1;
        }
        return new TextMixedParagraphResult(request, lines);
    }
}
