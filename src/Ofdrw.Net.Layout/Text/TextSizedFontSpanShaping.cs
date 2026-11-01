using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading;

namespace Ofdrw.Net.Layout.Text;

/// <summary>An existing caller-selected font/source/style association with an explicit em size.</summary>
/// <remarks>The original span and face remain immutable. StyleKey remains opaque and never supplies a size or paint value.</remarks>
public sealed class TextSizedFontSpan
{
    /// <summary>Binds an existing font span to a finite size in (0,1000] millimeters with a representable positive size/UPEM scale.</summary>
    public TextSizedFontSpan(TextFontSpan span, double fontSizeMillimeters)
    {
        Span = span ?? throw new ArgumentNullException(nameof(span));
        if (double.IsNaN(fontSizeMillimeters) || double.IsInfinity(fontSizeMillimeters) ||
            fontSizeMillimeters <= 0 || fontSizeMillimeters > 1000 || fontSizeMillimeters / span.Face.UnitsPerEm <= 0)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid explicit font size.",
                sourceStart: span.SourceStart, sourceLength: span.SourceLength);
        FontSizeMillimeters = fontSizeMillimeters;
    }
    /// <summary>The exact caller span, including its source interval, selected face and opaque style key.</summary>
    public TextFontSpan Span { get; }
    /// <summary>Explicit em size in millimeters, independent of the opaque style key.</summary>
    public double FontSizeMillimeters { get; }
}

/// <summary>A nonempty source partition with explicit caller-selected faces and sizes and uniform shaping attributes.</summary>
/// <remarks>Face or size changes define shaping-item boundaries; compatible face/size spans coalesce.
/// This is an explicit query policy, not automatic family/style resolution, paint interpretation or paragraph analysis.</remarks>
public sealed class TextSizedFontSpanShapingRequest
{
    /// <summary>Snapshots one to 256 sized spans over at most 4096 UTF-16 units, using the existing eight-face/64 MiB,
    /// 256-feature and one-to-256 shaping-call input limits.</summary>
    public TextSizedFontSpanShapingRequest(string text, int sourceStart, IReadOnlyList<TextSizedFontSpan> spans,
        TextDirection direction = TextDirection.LeftToRight, string script = "Latn", string language = "en",
        IReadOnlyList<TextFeature>? features = null, int maximumShapeCalls = 256)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (spans is null) throw new ArgumentNullException(nameof(spans));
        var count = spans.Count;
        if (count == 0 || (uint)count > 256)
            throw new TextLayoutException(TextLayoutError.BudgetExceeded, "Supply one to 256 sized source spans.");
        var copy = new List<TextSizedFontSpan>(count);
        var fonts = new List<TextFontSpan>(count);
        var sizes = new List<double>(count);
        for (var index = 0; index < count; index++)
        {
            var span = spans[index];
            if (span is null) throw new TextLayoutException(TextLayoutError.InvalidRequest, "A sized source span is null.");
            copy.Add(span); fonts.Add(span.Span); sizes.Add(span.FontSizeMillimeters);
        }
        FontPartition = new TextFontSpanShapingRequest(text, sourceStart, fonts, direction, script, language,
            features, maximumShapeCalls);
        Spans = new ReadOnlyCollection<TextSizedFontSpan>(copy);
        Sizes = new ReadOnlyCollection<double>(sizes);
    }

    private TextSizedFontSpanShapingRequest(TextSizedFontSpanShapingRequest context,
        TextFontSpanShapingRequest item, List<TextSizedFontSpan> spans)
    {
        Context = context; FontPartition = item;
        Spans = new ReadOnlyCollection<TextSizedFontSpan>(spans);
        var sizes = new List<double>(spans.Count);
        foreach (var span in spans) sizes.Add(span.FontSizeMillimeters);
        Sizes = new ReadOnlyCollection<double>(sizes);
    }

    /// <summary>Selects a nonempty text-element range while retaining the exact original sized partition and native source context.</summary>
    /// <remarks>Features are intersected and rebased; clipped spans retain their face, size and key, and unchanged spans retain
    /// object identity. Nested items are refused. This does not select a legal/final line or crop already shaped glyphs.</remarks>
    public static TextSizedFontSpanShapingRequest CreateContextualItem(TextSizedFontSpanShapingRequest context, TextSourceSpan item)
    {
        if (context is null) throw new ArgumentNullException(nameof(context));
        if (item is null) throw new ArgumentNullException(nameof(item));
        if (context.Context is not null)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Select from the original complete sized partition, not another item.");
        var selected = TextFontSpanShapingRequest.CreateContextualItem(context.FontPartition, item);
        var spans = new List<TextSizedFontSpan>(selected.Spans.Count);
        var ownerIndex = 0;
        foreach (var span in selected.Spans)
        {
            while (context.Spans[ownerIndex].Span.SourceStart + context.Spans[ownerIndex].Span.SourceLength <= span.SourceStart)
                ownerIndex++;
            var owner = context.Spans[ownerIndex];
            spans.Add(ReferenceEquals(span, owner.Span) ? owner : new TextSizedFontSpan(span, owner.FontSizeMillimeters));
        }
        return new TextSizedFontSpanShapingRequest(context, selected, spans);
    }

    internal TextFontSpanShapingRequest FontPartition { get; }
    internal IReadOnlyList<double> Sizes { get; }
    /// <summary>The exact complete sized request for a selected item, or null for a whole partition.</summary>
    public TextSizedFontSpanShapingRequest? Context { get; }
    /// <summary>Unmodified logical Unicode for the whole partition or selected item.</summary>
    public string Text => FontPartition.Text;
    /// <summary>Absolute UTF-16 start of the whole partition or selected item.</summary>
    public int SourceStart => FontPartition.SourceStart;
    /// <summary>Immutable logical partition retaining original span, face, size and opaque key associations.</summary>
    public IReadOnlyList<TextSizedFontSpan> Spans { get; }
    /// <summary>Uniform caller-analyzed direction; no paragraph bidi analysis is performed.</summary>
    public TextDirection Direction => FontPartition.Direction;
    /// <summary>Uniform caller-selected script tag.</summary>
    public string Script => FontPartition.Script;
    /// <summary>Uniform caller-selected language tag.</summary>
    public string Language => FontPartition.Language;
    /// <summary>Immutable item-relative features; selected requests retain complete original features in Context.</summary>
    public IReadOnlyList<TextFeature> Features => FontPartition.Features;
    /// <summary>Maximum native/contextual calls, preflighted before any call.</summary>
    public int MaximumShapeCalls => FontPartition.MaximumShapeCalls;
}

/// <summary>One compatible face/size item, with exact native shaping and its millimeter projection.</summary>
public sealed class TextSizedFontSpanShapedRun
{
    internal TextSizedFontSpanShapedRun(TextFontSpanShapedRun shaped, double size, TextPositionedRun positioned)
    {
        FirstSpanIndex = shaped.FirstSpanIndex; SpanCount = shaped.SpanCount; Shaping = shaped.Shaping;
        GlyphStyleKeys = shaped.GlyphStyleKeys; FontSizeMillimeters = size; PositionedRun = positioned;
    }
    /// <summary>First selected input span index covered by this item.</summary>
    public int FirstSpanIndex { get; }
    /// <summary>Number of consecutive compatible face/size spans covered by this item.</summary>
    public int SpanCount { get; }
    /// <summary>Exact native contextual result; coverage slice indices are local to this item's spans.</summary>
    public TextSpanShapingResult Shaping { get; }
    /// <summary>Explicit em size used for this item and its projection.</summary>
    public double FontSizeMillimeters { get; }
    /// <summary>The exact shaped run projected at the declared size, with initial origin and baseline (0,0).</summary>
    public TextPositionedRun PositionedRun { get; }
    /// <summary>Opaque style key for each glyph occurrence, in unchanged native order.</summary>
    public IReadOnlyList<int> GlyphStyleKeys { get; }
}

/// <summary>Sized font items in logical source order, without a combined visual line or advance.</summary>
public sealed class TextSizedFontSpanShapingResult
{
    internal TextSizedFontSpanShapingResult(TextSizedFontSpanShapingRequest request, List<TextSizedFontSpanShapedRun> runs)
    { Request = request; Runs = new ReadOnlyCollection<TextSizedFontSpanShapedRun>(runs); }
    /// <summary>The exact immutable sized request, including its original context when selected.</summary>
    public TextSizedFontSpanShapingRequest Request { get; }
    /// <summary>Logical/source items; each retains its own native glyph order and face/size projection.</summary>
    public IReadOnlyList<TextSizedFontSpanShapedRun> Runs { get; }
}

/// <summary>Queries explicit face/size itemization using original source context and shared millimeter positioning.</summary>
/// <remarks>All original face/size boundaries and native-call/work budgets are preflighted before native calls.
/// Global 16384-glyph/32768-slice limits and each positioned run's coordinate budget apply; cancellation or refusal
/// publishes no partial result. Native scale/hinting, actual paint/style resolution, final-line/bidi composition,
/// legacy Flow and writer/reader admission remain unchanged.</remarks>
public sealed class TextSizedFontSpanShapingEngine
{
    private readonly TextFontSpanShapingEngine _fonts;
    /// <summary>Creates the query using an explicitly opted-in contextual shaper.</summary>
    public TextSizedFontSpanShapingEngine(IContextualTextShaper shaper) => _fonts = new TextFontSpanShapingEngine(shaper);
    /// <summary>Shapes compatible face/size items, retaining original context and projecting each exact result at its declared size.</summary>
    public TextSizedFontSpanShapingResult Shape(TextSizedFontSpanShapingRequest request, CancellationToken cancellationToken = default)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        var shaped = _fonts.ShapeSized(request.FontPartition, request.Sizes, (request.Context ?? request).Sizes, cancellationToken);
        var runs = new List<TextSizedFontSpanShapedRun>(shaped.Runs.Count);
        foreach (var run in shaped.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var size = request.Spans[run.FirstSpanIndex].FontSizeMillimeters;
            var positioned = TextPositionedRun.Create(run.Shaping.ShapedRun, size, cancellationToken: cancellationToken);
            runs.Add(new TextSizedFontSpanShapedRun(run, size, positioned));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new TextSizedFontSpanShapingResult(request, runs);
    }
}
