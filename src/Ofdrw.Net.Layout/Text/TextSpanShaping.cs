using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>A positive caller-owned interval, in absolute UTF-16 source offsets.</summary>
public sealed class TextSourceSpan
{
    /// <summary>Creates an interval without assigning a font, style, glyph or caret position.</summary>
    public TextSourceSpan(int sourceStart, int sourceLength)
    {
        if (sourceStart < 0) throw new ArgumentOutOfRangeException(nameof(sourceStart));
        if (sourceLength <= 0 || (long)sourceStart + sourceLength > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(sourceLength));
        SourceStart = sourceStart;
        SourceLength = sourceLength;
    }

    /// <summary>Absolute UTF-16 start in the caller's source.</summary>
    public int SourceStart { get; }
    /// <summary>Positive length in UTF-16 code units.</summary>
    public int SourceLength { get; }
}

/// <summary>A complete single-face run with an ordered, contiguous source partition.</summary>
/// <remarks>The spans identify source ownership only. The caller still selects the run's
/// actual face, direction, script, language and features; this is not style or bidi analysis.</remarks>
public sealed class TextSpanShapingRequest
{
    private const int MaximumSpans = 256;

    /// <summary>Snapshots at most 256 spans covering the complete run without gaps,
    /// overlaps or surrogate-pair splits. An empty run requires an empty partition.</summary>
    public TextSpanShapingRequest(TextShapeRequest run, IReadOnlyList<TextSourceSpan> spans)
    {
        Run = run ?? throw new ArgumentNullException(nameof(run));
        if (spans is null) throw new ArgumentNullException(nameof(spans));
        var spanCount = spans.Count;
        if ((uint)spanCount > MaximumSpans)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                "A span shaping query exceeds 256 source spans.");
        var copy = new List<TextSourceSpan>(spanCount);
        var next = run.SourceStart;
        var end = run.SourceStart + run.Text.Length;
        for (var index = 0; index < spanCount; index++)
        {
            var span = spans[index];
            if (span is null || span.SourceStart != next ||
                (long)span.SourceStart + span.SourceLength > end)
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "Source spans must partition the complete run in logical source order.");
            var relativeStart = span.SourceStart - run.SourceStart;
            var relativeEnd = relativeStart + span.SourceLength;
            if (relativeStart > 0 && char.IsLowSurrogate(run.Text[relativeStart]) ||
                relativeEnd < run.Text.Length && char.IsLowSurrogate(run.Text[relativeEnd]))
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "A source span boundary splits a UTF-16 surrogate pair.",
                    sourceStart: span.SourceStart, sourceLength: span.SourceLength);
            copy.Add(span);
            next = span.SourceStart + span.SourceLength;
        }
        if (next != end)
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "Source spans do not cover the end of the complete run.");
        Spans = new ReadOnlyCollection<TextSourceSpan>(copy);
    }

    /// <summary>The complete, unmodified contextual input shaped exactly once.</summary>
    public TextShapeRequest Run { get; }
    /// <summary>Immutable source spans in logical order; list indices retain caller association.</summary>
    public IReadOnlyList<TextSourceSpan> Spans { get; }
}

/// <summary>The intersection of one source span with a glyph's source cluster.</summary>
public sealed class TextSpanSourceSlice
{
    internal TextSpanSourceSlice(int spanIndex, int sourceStart, int sourceLength)
    { SpanIndex = spanIndex; SourceStart = sourceStart; SourceLength = sourceLength; }

    /// <summary>Index into the request's immutable logical source partition.</summary>
    public int SpanIndex { get; }
    /// <summary>Absolute UTF-16 start of the intersection.</summary>
    public int SourceStart { get; }
    /// <summary>Positive UTF-16 length of the intersection.</summary>
    public int SourceLength { get; }
}

/// <summary>Source ownership of one original glyph occurrence, retaining shaping order.</summary>
/// <remarks>Multiple slices do not assign a unique style or a per-span advance.
/// Cluster intervals are not caret or selection positions.</remarks>
public sealed class TextSpanGlyphCoverage
{
    internal TextSpanGlyphCoverage(int glyphIndex, TextShapedGlyph glyph, int clusterGlyphCount,
        List<TextSpanSourceSlice> slices)
    {
        GlyphIndex = glyphIndex; ClusterStart = glyph.ClusterStart; ClusterLength = glyph.ClusterLength;
        ClusterGlyphCount = clusterGlyphCount;
        Slices = new ReadOnlyCollection<TextSpanSourceSlice>(slices);
    }

    /// <summary>Index of this occurrence in the completed run's glyph array.</summary>
    public int GlyphIndex { get; }
    /// <summary>Absolute UTF-16 start of the complete source cluster.</summary>
    public int ClusterStart { get; }
    /// <summary>UTF-16 length of the complete source cluster.</summary>
    public int ClusterLength { get; }
    /// <summary>Number of original glyph occurrences sharing this exact source cluster interval.</summary>
    public int ClusterGlyphCount { get; }
    /// <summary>Immutable source intersections in logical span order.</summary>
    public IReadOnlyList<TextSpanSourceSlice> Slices { get; }
    /// <summary>Whether this glyph's cluster belongs to more than one source span.</summary>
    public bool CrossesSpanBoundary => Slices.Count > 1;
}

/// <summary>A whole-run shape and per-occurrence source ownership, with no rendering policy.</summary>
public sealed class TextSpanShapingResult
{
    internal TextSpanShapingResult(TextSpanShapingRequest request, TextShapedRun shapedRun,
        List<TextSpanGlyphCoverage> coverage)
    {
        Request = request; ShapedRun = shapedRun;
        GlyphCoverage = new ReadOnlyCollection<TextSpanGlyphCoverage>(coverage);
    }

    /// <summary>The immutable input including its complete context and source partition.</summary>
    public TextSpanShapingRequest Request { get; }
    /// <summary>Original GIDs, positions and source clusters from the single shaping call.</summary>
    public TextShapedRun ShapedRun { get; }
    /// <summary>One entry per original glyph occurrence, in the same order as the shaped run.</summary>
    public IReadOnlyList<TextSpanGlyphCoverage> GlyphCoverage { get; }
}

/// <summary>Shapes a complete caller-selected run once, then relates clusters to source spans.</summary>
/// <remarks>No legacy Flow, font fallback, bidi, style assignment, pagination or output
/// profile is changed. A cross-span cluster may require a caller's explicit style decision.</remarks>
public sealed class TextSpanShapingEngine
{
    private const int MaximumSourceSlices = 32768;
    private readonly ITextShaper _shaper;

    /// <summary>Creates a source-association query using a conforming whole-run shaper.</summary>
    public TextSpanShapingEngine(ITextShaper shaper)
        => _shaper = shaper ?? throw new ArgumentNullException(nameof(shaper));

    /// <summary>Returns at most 32768 source slices, or fails without a partial result.
    /// Cancellation is checked before/after the shaping call and during association;
    /// it does not interrupt arbitrary instructions inside a synchronous native shaper.</summary>
    public TextSpanShapingResult Shape(TextSpanShapingRequest request,
        CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        TextShapedRun shaped;
        if (request.Run.Context is not null)
        {
            // An existing replacement shaper may legitimately implement only
            // whole-run shaping. Never silently upgrade it to external context.
            if (_shaper is not IContextualTextShaper contextual)
                throw new TextLayoutException(TextLayoutError.InvalidRequest,
                    "The selected shaper has not opted into contextual item queries.");
            shaped = contextual.ShapeContextualItem(request.Run, cancellationToken);
        }
        else shaped = _shaper.Shape(request.Run, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (shaped is null || !ReferenceEquals(shaped.Request, request.Run))
            throw new TextLayoutException(TextLayoutError.InvalidRequest,
                "The shaper did not return the exact complete input request.");

        // Count every occurrence before allocating source-slice output. Repeated
        // M:N clusters remain separate occurrences and consume their own budget.
        var count = 0;
        var clusterCounts = new Dictionary<(int Start, int Length), int>();
        foreach (var glyph in shaped.Glyphs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var key = (glyph.ClusterStart, glyph.ClusterLength);
            clusterCounts.TryGetValue(key, out var occurrences);
            clusterCounts[key] = occurrences + 1;
            var end = glyph.ClusterStart + glyph.ClusterLength;
            for (var index = FindSpan(request.Spans, glyph.ClusterStart);
                 index < request.Spans.Count && request.Spans[index].SourceStart < end; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++count > MaximumSourceSlices)
                    throw new TextLayoutException(TextLayoutError.BudgetExceeded,
                        "Span shaping source associations exceed 32768 slices.");
            }
        }

        var coverage = new List<TextSpanGlyphCoverage>(shaped.Glyphs.Count);
        for (var glyphIndex = 0; glyphIndex < shaped.Glyphs.Count; glyphIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var glyph = shaped.Glyphs[glyphIndex];
            var end = glyph.ClusterStart + glyph.ClusterLength;
            var slices = new List<TextSpanSourceSlice>();
            for (var index = FindSpan(request.Spans, glyph.ClusterStart);
                 index < request.Spans.Count && request.Spans[index].SourceStart < end; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var span = request.Spans[index];
                var start = Math.Max(span.SourceStart, glyph.ClusterStart);
                var sliceEnd = Math.Min(span.SourceStart + span.SourceLength, end);
                slices.Add(new TextSpanSourceSlice(index, start, sliceEnd - start));
            }
            coverage.Add(new TextSpanGlyphCoverage(glyphIndex, glyph,
                clusterCounts[(glyph.ClusterStart, glyph.ClusterLength)], slices));
        }
        return new TextSpanShapingResult(request, shaped, coverage);
    }

    private static int FindSpan(IReadOnlyList<TextSourceSpan> spans, int sourceStart)
    {
        var low = 0; var high = spans.Count;
        while (low < high)
        {
            var middle = low + (high - low) / 2;
            if (spans[middle].SourceStart + spans[middle].SourceLength <= sourceStart) low = middle + 1;
            else high = middle;
        }
        return low;
    }
}
