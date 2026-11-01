using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>A caller-selected immutable face and opaque style equality key for one source interval.</summary>
/// <remarks>The key declares caller-owned style compatibility. No paint, size, family-name
/// resolution or synthetic font style is inferred from it.</remarks>
public sealed class TextFontSpan
{
    /// <summary>Creates a positive absolute UTF-16 source interval with an explicitly selected face.</summary>
    public TextFontSpan(int sourceStart, int sourceLength, TextFontFace face, int styleKey = 0)
    {
        Source = new TextSourceSpan(sourceStart, sourceLength);
        Face = face ?? throw new ArgumentNullException(nameof(face));
        StyleKey = styleKey;
    }
    internal TextSourceSpan Source { get; }
    /// <summary>Absolute UTF-16 source start.</summary>
    public int SourceStart => Source.SourceStart;
    /// <summary>Positive UTF-16 source length.</summary>
    public int SourceLength => Source.SourceLength;
    /// <summary>The original immutable caller-selected face, retained even when an alias coalesces.</summary>
    public TextFontFace Face { get; }
    /// <summary>Opaque immutable caller equality identity; equal keys declare compatible style ownership.</summary>
    public int StyleKey { get; }
}

/// <summary>A bounded nonempty source partition with caller-selected fonts and uniform shaping attributes.</summary>
/// <remarks>This does not analyze bidi, discover fonts, derive features from styles, or resolve actual paint/size values.</remarks>
public sealed class TextFontSpanShapingRequest
{
    /// <summary>Snapshots up to 256 spans and features over at most 4096 UTF-16 units.
    /// At most eight distinct selected face programs totaling 64 MiB are admitted.</summary>
    public TextFontSpanShapingRequest(string text, int sourceStart, IReadOnlyList<TextFontSpan> spans,
        TextDirection direction = TextDirection.LeftToRight, string script = "Latn", string language = "en",
        IReadOnlyList<TextFeature>? features = null, int maximumShapeCalls = 256)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (spans is null) throw new ArgumentNullException(nameof(spans));
        if (text.Length == 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "A font-span query requires nonempty source text.");
        var count = spans.Count;
        if (count == 0 || (uint)count > 256)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Supply one to 256 font source spans.");
        if (maximumShapeCalls < 1 || maximumShapeCalls > 256)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "The shaping-call budget must be between one and 256.");
        var copy = new List<TextFontSpan>(count);
        var sources = new List<TextSourceSpan>(count);
        var identities = new HashSet<string>(StringComparer.Ordinal);
        long bytes = 0;
        for (var index = 0; index < count; index++)
        {
            var span = spans[index];
            if (span is null)
                throw new TextLayoutException(TextLayoutError.InvalidRequest, "A font source span is null.");
            copy.Add(span); sources.Add(span.Source);
            if (identities.Add(span.Face.FaceSha256))
            {
                bytes += span.Face.FaceByteLength;
                if (identities.Count > 8 || bytes > 64L * 1024 * 1024)
                    throw new TextLayoutException(TextLayoutError.BudgetExceeded, "The selected font programs exceed the face budget.");
            }
        }
        var featureCopy = new List<TextFeature>();
        if (features is not null)
        {
            var featureCount = features.Count;
            if ((uint)featureCount > 256)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "A font-span query exceeds 256 features.");
            for (var index = 0; index < featureCount; index++) featureCopy.Add(features[index]);
        }
        // Reuse the complete-source and contiguous UTF-16 partition contracts.
        // No face coverage check is made against characters owned by another font.
        var template = new TextShapeRequest(text, sourceStart, copy[0].Face, direction, script, language, featureCopy);
        _ = new TextSpanShapingRequest(template, sources);
        Text = template.Text; SourceStart = template.SourceStart;
        Direction = template.Direction; Script = template.Script; Language = template.Language;
        Features = template.Features;
        Spans = new ReadOnlyCollection<TextFontSpan>(copy);
        MaximumShapeCalls = maximumShapeCalls;
    }

    /// <summary>The complete unmodified logical source.</summary>
    public string Text { get; }
    /// <summary>Absolute UTF-16 start of the complete source.</summary>
    public int SourceStart { get; }
    /// <summary>Immutable logical source partition retaining every caller face and style key.</summary>
    public IReadOnlyList<TextFontSpan> Spans { get; }
    /// <summary>Uniform caller-analyzed direction; no paragraph bidi analysis is performed.</summary>
    public TextDirection Direction { get; }
    /// <summary>Uniform caller-selected script tag.</summary>
    public string Script { get; }
    /// <summary>Uniform caller-selected language tag.</summary>
    public string Language { get; }
    /// <summary>Immutable feature ranges relative to the complete source.</summary>
    public IReadOnlyList<TextFeature> Features { get; }
    /// <summary>Maximum native/contextual shaping calls, checked before the first call.</summary>
    public int MaximumShapeCalls { get; }
}

/// <summary>One coalesced font item, with original glyphs and one declared style key per occurrence.</summary>
public sealed class TextFontSpanShapedRun
{
    internal TextFontSpanShapedRun(int first, int count, TextSpanShapingResult shaping, List<int> keys)
    {
        FirstSpanIndex = first; SpanCount = count; Shaping = shaping;
        GlyphStyleKeys = new ReadOnlyCollection<int>(keys);
    }
    /// <summary>First input span index covered by this item.</summary>
    public int FirstSpanIndex { get; }
    /// <summary>Number of consecutive input spans covered by this item.</summary>
    public int SpanCount { get; }
    /// <summary>Exact contextual item result. Its coverage slice indices are local to this item's spans.</summary>
    public TextSpanShapingResult Shaping { get; }
    /// <summary>One style equality key for each native glyph occurrence, in the same native order.</summary>
    public IReadOnlyList<int> GlyphStyleKeys { get; }
}

/// <summary>Logical font items without a combined advance, global visual order or output admission.</summary>
public sealed class TextFontSpanShapingResult
{
    internal TextFontSpanShapingResult(TextFontSpanShapingRequest request, List<TextFontSpanShapedRun> runs)
    { Request = request; Runs = new ReadOnlyCollection<TextFontSpanShapedRun>(runs); }
    /// <summary>The exact immutable input, retaining all original selected face associations.</summary>
    public TextFontSpanShapingRequest Request { get; }
    /// <summary>Items in logical source order. Each retains its own face design units and native glyph order.</summary>
    public IReadOnlyList<TextFontSpanShapedRun> Runs { get; }
}

/// <summary>Coalesces an explicit font partition and refuses ambiguous cross-style glyph ownership.</summary>
/// <remarks>Requires an explicitly contextual shaper. Legacy Flow and all writer/reader profiles remain unchanged.</remarks>
public sealed class TextFontSpanShapingEngine
{
    private const long MaximumShapeWork = 512L * 1024 * 1024;
    private const int MaximumGlyphs = 16384;
    private const int MaximumSlices = 32768;
    private readonly TextSpanShapingEngine _spans;

    /// <summary>Creates a query using an explicitly opted-in contextual shaper.</summary>
    public TextFontSpanShapingEngine(IContextualTextShaper shaper)
        => _spans = new TextSpanShapingEngine(shaper ?? throw new ArgumentNullException(nameof(shaper)));

    /// <summary>Preflights all font boundaries and shape-call/work budgets before any native call.
    /// Returns at most 16384 glyph occurrences and 32768 source slices, or fails without a partial result.</summary>
    public TextFontSpanShapingResult Shape(TextFontSpanShapingRequest request,
        CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        var plans = new List<Plan>();
        long work = 0;
        for (var first = 0; first < request.Spans.Count;)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var face = request.Spans[first].Face;
            var end = first + 1;
            while (end < request.Spans.Count && string.Equals(face.FaceSha256,
                request.Spans[end].Face.FaceSha256, StringComparison.Ordinal))
            { cancellationToken.ThrowIfCancellationRequested(); end++; }
            if (plans.Count >= request.MaximumShapeCalls)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Font items exceed the shaping-call budget.");
            // Charge both face copies used by the pinned adapter and weighted
            // complete-context text work before entering any replacement shaper.
            var cost = 2L * face.FaceByteLength + 64L * request.Text.Length;
            if (work > MaximumShapeWork - cost)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Font items exceed the cumulative shaping-work budget.");
            work += cost;
            var context = new TextShapeRequest(request.Text, request.SourceStart, face,
                request.Direction, request.Script, request.Language, request.Features);
            var start = request.Spans[first].SourceStart;
            var last = request.Spans[end - 1];
            var item = TextShapeRequest.CreateContextualItem(context,
                new TextSourceSpan(start, last.SourceStart + last.SourceLength - start));
            var sources = new List<TextSourceSpan>(end - first);
            for (var index = first; index < end; index++) sources.Add(request.Spans[index].Source);
            plans.Add(new Plan(first, end - first, new TextSpanShapingRequest(item, sources)));
            first = end;
        }
        // In particular, a later font switch that splits a base/mark sequence
        // cannot run an earlier otherwise valid item before being refused.
        var runs = new List<TextFontSpanShapedRun>(plans.Count);
        long glyphs = 0, slices = 0;
        foreach (var plan in plans)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var result = _spans.Shape(plan.Query, cancellationToken);
            glyphs += result.ShapedRun.Glyphs.Count;
            if (glyphs > MaximumGlyphs)
                throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Font items exceed 16384 glyph occurrences.");
            foreach (var coverage in result.GlyphCoverage)
            {
                cancellationToken.ThrowIfCancellationRequested();
                slices += coverage.Slices.Count;
                if (slices > MaximumSlices)
                    throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Font items exceed 32768 source slices.");
            }
            var keys = new List<int>(result.GlyphCoverage.Count);
            foreach (var coverage in result.GlyphCoverage)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var key = request.Spans[plan.First + coverage.Slices[0].SpanIndex].StyleKey;
                foreach (var slice in coverage.Slices)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (request.Spans[plan.First + slice.SpanIndex].StyleKey != key)
                        throw new TextLayoutException(TextLayoutError.InvalidRequest,
                            "A shaped source cluster intersects conflicting caller style keys; unique glyph ownership is undefined.",
                            sourceStart: coverage.ClusterStart, sourceLength: coverage.ClusterLength);
                }
                keys.Add(key);
            }
            runs.Add(new TextFontSpanShapedRun(plan.First, plan.Count, result, keys));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new TextFontSpanShapingResult(request, runs);
    }

    private sealed class Plan(int first, int count, TextSpanShapingRequest query)
    {
        internal int First { get; } = first;
        internal int Count { get; } = count;
        internal TextSpanShapingRequest Query { get; } = query;
    }
}
