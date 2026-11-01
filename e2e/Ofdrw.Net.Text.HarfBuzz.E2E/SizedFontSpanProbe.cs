#nullable enable
using System.Text.Json;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

internal static class SizedFontSpanProbe
{
    internal static async Task Run(string output, TextFontFace dejavu, TextFontFace noto)
    {
        var rows = new List<object>();
        TextSizedFontSpan Span(int start, int length, double size, int key = 7, TextFontFace? face = null) =>
            new(new TextFontSpan(start, length, face ?? dejavu, key), size);
        TextSizedFontSpanShapingRequest Request(string text, int start, params TextSizedFontSpan[] spans) => new(text, start, spans);
        object Input(TextSizedFontSpanShapingRequest request) => new {
            request.Text, request.SourceStart, direction = request.Direction.ToString(), request.Script, request.Language,
            request.Features, request.MaximumShapeCalls,
            spans = request.Spans.Select(s => new { s.Span.SourceStart, s.Span.SourceLength, s.Span.StyleKey,
                s.Span.Face.FaceSha256, s.Span.Face.UnitsPerEm, s.FontSizeMillimeters }) };
        void Capture(string role, TextSizedFontSpanShapingRequest request, params uint[][] expected)
        {
            var meter = new Meter(); var result = new TextSizedFontSpanShapingEngine(meter).Shape(request);
            if (!ReferenceEquals(result.Request, request) || result.Runs.Count != expected.Length || meter.Calls != expected.Length)
                throw new InvalidDataException("Sized query item count/identity failed: " + role);
            for (var index = 0; index < expected.Length; index++)
            {
                var run = result.Runs[index]; var shape = run.Shaping.ShapedRun;
                if (!ReferenceEquals(shape, run.PositionedRun.ShapedRun) ||
                    !shape.Glyphs.Select(g => g.GlyphId).SequenceEqual(expected[index]) ||
                    run.PositionedRun.FontSizeMillimeters != request.Spans[run.FirstSpanIndex].FontSizeMillimeters ||
                    shape.Request.Context?.Text != (request.Context ?? request).Text ||
                    shape.Request.Context?.SourceStart != (request.Context ?? request).SourceStart)
                    throw new InvalidDataException("Sized query original-context/geometry identity failed: " + role);
            }
            rows.Add(new { role, outcome = "accepted", nativeCalls = meter.Calls, input = Input(request),
                original = Input(request.Context ?? request), selected = request.Context is not null,
                runs = result.Runs.Select(run => new { run.FirstSpanIndex, run.SpanCount, run.FontSizeMillimeters,
                    run.GlyphStyleKeys, run.Shaping.GlyphCoverage,
                    text = run.Shaping.ShapedRun.Request.Text, sourceStart = run.Shaping.ShapedRun.Request.SourceStart,
                    contextText = run.Shaping.ShapedRun.Request.Context?.Text,
                    contextSourceStart = run.Shaping.ShapedRun.Request.Context?.SourceStart,
                    run.Shaping.ShapedRun.Request.Features,
                    contextFeatures = run.Shaping.ShapedRun.Request.Context?.Features,
                    run.Shaping.ShapedRun.Request.Face.FaceSha256, run.Shaping.ShapedRun.Request.Face.UnitsPerEm,
                    run.Shaping.ShapedRun.EngineVersion, run.Shaping.ShapedRun.XAdvance, run.Shaping.ShapedRun.YAdvance,
                    designGlyphs = run.Shaping.ShapedRun.Glyphs, run.PositionedRun.OriginXMillimeters,
                    run.PositionedRun.BaselineYMillimeters, run.PositionedRun.XAdvanceMillimeters,
                    run.PositionedRun.YAdvanceMillimeters, run.PositionedRun.InkBoundsMillimeters,
                    glyphs = run.PositionedRun.Glyphs }) });
        }
        void Refuse(string role, Func<TextSizedFontSpanShapingRequest> create, TextLayoutError expected, int calls)
        {
            var meter = new Meter();
            try { new TextSizedFontSpanShapingEngine(meter).Shape(create()); }
            catch (TextLayoutException error) when (error.Error == expected && meter.Calls == calls)
            { rows.Add(new { role, outcome = "rejected", nativeCalls = meter.Calls, error = error.Error.ToString(), error.SourceStart, error.SourceLength }); return; }
            throw new InvalidDataException("Sized query refusal failed: " + role);
        }
        Capture("equal-face-size-ffi", Request("ffi", 20, Span(20, 1, 4), Span(21, 2, 4)), new uint[] { 5044 });
        Capture("equal-size-AV-opaque-keys", Request("AV", 20, Span(20, 1, 4), Span(21, 1, 4, 9)), new uint[] { 36, 57 });
        Capture("AV-size-4-6", Request("AV", 20, Span(20, 1, 4), Span(21, 1, 6, 9)), new uint[] { 36 }, new uint[] { 57 });
        Capture("ffi-size-4-6", Request("ffi", 20, Span(20, 1, 4), Span(21, 2, 6, 9)), new uint[] { 73 }, new uint[] { 5042 });
        Capture("shared-mark-equal-size", Request("a\u0323\u0301", 70, Span(70, 1, 4), Span(71, 2, 4)), new uint[] { 2463, 690 });
        Capture("arabic-size-4-6", new("ببب", 100, new[] { Span(100, 1, 4), Span(101, 2, 6, 9) },
            TextDirection.RightToLeft, "Arab", "ar"), new uint[] { 5259 }, new uint[] { 5258, 5260 });
        var mixed = new TextSizedFontSpanShapingRequest("بببAV", 100,
            new[] { Span(100, 2, 4), Span(102, 1, 6, 9), Span(103, 2, 5, 11, noto) }, TextDirection.RightToLeft, "Arab", "ar");
        Capture("selected-arabic-face-size-clips", TextSizedFontSpanShapingRequest.CreateContextualItem(mixed, new(101, 3)),
            new uint[] { 5260 }, new uint[] { 5258 }, new uint[] { 36 });
        foreach (var disables in new[] { false, true })
        {
            var features = disables ? new[] { new TextFeature("liga", 1, 2, 3), new TextFeature("liga", 0, 0, 7) } :
                new[] { new TextFeature("liga", 0, 0, 7), new TextFeature("liga", 1, 2, 3) };
            var original = new TextSizedFontSpanShapingRequest("XAffiVZ", 70,
                new[] { Span(70, 2, 4), Span(72, 3, 6, 9), Span(75, 2, 4, 11) }, features: features);
            Capture(disables ? "selected-features-last-off" : "selected-features-last-on",
                TextSizedFontSpanShapingRequest.CreateContextualItem(original, new(71, 5)), new uint[] { 36 },
                disables ? new uint[] { 73, 73, 76 } : new uint[] { 5044 }, new uint[] { 57 });
        }
        var start = int.MaxValue - 5;
        var nonBmp = Request("\U00010300a\u0323A", start, Span(start, 2, 4), Span(start + 2, 2, 6, 9), Span(start + 4, 1, 3, 11, noto));
        Capture("nonbmp-max-endpoint-selected-upems", TextSizedFontSpanShapingRequest.CreateContextualItem(nonBmp, new(start + 2, 3)),
            new uint[] { 2463 }, new uint[] { 36 });
        Capture("two-face-own-upem-size", Request("AV", 20, Span(20, 1, 6, 7, noto), Span(21, 1, 3, 9)), new uint[] { 36 }, new uint[] { 57 });
        var badBoundary = Request("Xa\u0323V", 0, Span(0, 2, 4), Span(2, 1, 6), Span(3, 1, 4));
        Refuse("refused-size-mark-cut", () => badBoundary, TextLayoutError.InvalidRequest, 0);
        Refuse("refused-outside-selected-size-mark-cut", () => TextSizedFontSpanShapingRequest.CreateContextualItem(badBoundary, new(0, 1)), TextLayoutError.InvalidRequest, 0);
        Refuse("refused-shared-ffi-conflicting-keys", () => Request("ffi", 20, Span(20, 1, 4), Span(21, 2, 4, 9)), TextLayoutError.InvalidRequest, 1);
        Refuse("refused-size-call-budget", () => new("AV", 20, new[] { Span(20, 1, 4), Span(21, 1, 6) }, maximumShapeCalls: 1), TextLayoutError.BudgetExceeded, 0);
        Refuse("refused-zero-size", () => Request("A", 20, Span(20, 1, 0)), TextLayoutError.InvalidRequest, 0);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel(); var idle = new Meter();
        try { new TextSizedFontSpanShapingEngine(idle).Shape(Request("A", 20, Span(20, 1, 4)), cancellation.Token); }
        catch (OperationCanceledException error) when (error.CancellationToken == cancellation.Token && idle.Calls == 0)
        { rows.Add(new { role = "refused-caller-cancellation", outcome = "canceled", nativeCalls = 0 }); }
        if (rows.Count != 17) throw new InvalidDataException("Sized query cancellation/corpus incomplete.");
        await File.WriteAllTextAsync(Path.Combine(output, "r22-sized-font-span-query.json"), JsonSerializer.Serialize(new {
            sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"), cases = rows,
            scope = "Explicit caller-selected face/size query; original context, logical groups and unchanged native glyph order; each exact run projected separately at declared size. No combined visual line, automatic style/paint resolver, Flow or new writer/reader/script output admission."
        }, new JsonSerializerOptions { WriteIndented = true }) + "\n");
        Console.WriteLine("R22 explicit face/size, original-context and positioned-run queries passed from candidate packages.");
    }
    private sealed class Meter : IContextualTextShaper
    {
        internal int Calls;
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("Explicit contextual entry required.");
        public TextShapedRun ShapeContextualItem(TextShapeRequest request, CancellationToken cancellationToken = default)
        { Calls++; return new HarfBuzzTextShaper().ShapeContextualItem(request, cancellationToken); }
    }
}
