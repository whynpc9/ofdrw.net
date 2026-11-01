using System.Text.Json;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

internal static class PositionedRunProbe
{
    internal static async Task Run(string output, TextFontFace dejavu, TextFontFace noto)
    {
        var native = new HarfBuzzTextShaper();
        var rows = new List<object>();
        void Capture(string role, TextShapedRun shape, double size, double x, double y, params uint[] expected)
        {
            var result = TextPositionedRun.Create(shape, size, x, y);
            if (!ReferenceEquals(result.ShapedRun, shape) ||
                !result.Glyphs.Select(g => g.GlyphId).SequenceEqual(expected) ||
                !result.Glyphs.Select(g => (g.ClusterStart, g.ClusterLength)).SequenceEqual(
                    shape.Glyphs.Select(g => (g.ClusterStart, g.ClusterLength))))
                throw new InvalidDataException("Positioning lost exact shaped identity: " + role);
            rows.Add(new { role, outcome = "accepted", shape.EngineVersion,
                text = shape.Request.Text, sourceStart = shape.Request.SourceStart,
                direction = shape.Request.Direction.ToString(), shape.Request.Script, shape.Request.Language,
                shape.Request.Features, contextText = shape.Request.Context?.Text,
                contextSourceStart = shape.Request.Context?.SourceStart,
                shape.Request.Face.FaceSha256, shape.Request.Face.UnitsPerEm,
                shape.XAdvance, shape.YAdvance, designGlyphs = shape.Glyphs,
                result.FontSizeMillimeters, result.OriginXMillimeters, result.BaselineYMillimeters,
                result.XAdvanceMillimeters, result.YAdvanceMillimeters,
                result.InkBoundsMillimeters, glyphs = result.Glyphs });
        }
        TextShapedRun Shape(string text, int start = 20) => native.Shape(new TextShapeRequest(text, start, dejavu));
        Capture("native-AV-kern", Shape("AV"), 4, 10, 20, 36, 57);
        Capture("native-ffi", Shape("ffi"), 5, 10, 20, 5044);
        Capture("native-shared-mark-cluster", Shape("a\u0323\u0301"), 4, 10, 20, 2463, 690);
        var whole = new TextShapeRequest("ببب", 100, dejavu, TextDirection.RightToLeft, "Arab", "ar");
        Capture("contextual-arabic-middle", native.ShapeContextualItem(
            TextShapeRequest.CreateContextualItem(whole, new TextSourceSpan(101, 1))), 4, 10, 20, 5260);
        Capture("contextual-arabic-native-order", native.ShapeContextualItem(
            TextShapeRequest.CreateContextualItem(whole, new TextSourceSpan(101, 2))), 4, 10, 20, 5258, 5260);
        var nonBmp = Shape("\U00010300", int.MaxValue - 2);
        Capture("nonbmp-maximum-endpoint", nonBmp, 4, -10, -20, 5373);
        Capture("noto-own-upem-size", native.Shape(new TextShapeRequest("A", 12, noto)), 6, 10, 20, 36);
        Capture("dejavu-own-upem-size", Shape("A", 13), 3, 10, 20, 36);
        var controlled = new TextShapedRun(new TextShapeRequest("AV", 20, dejavu), new[] {
            new TextShapedGlyph(36, 20, 1, 100, 20, 5, 10, new(-2, -5, 12, 15), false),
            new TextShapedGlyph(57, 21, 1, 80, -10, -7, -3, null, true)
        }, "controlled-2d-not-native");
        Capture("controlled-2d-pen-offset-ink", controlled, 4, 10, 30, 36, 57);
        var projected = TextPositionedRun.Create(controlled, 4, 10, 30);
        if (projected.Glyphs[1].OriginXMillimeters != 10.181640625 ||
            projected.Glyphs[1].OriginYMillimeters != 29.966796875 ||
            projected.XAdvanceMillimeters != 0.3515625 || projected.YAdvanceMillimeters != -0.01953125 ||
            projected.InkBoundsMillimeters?.Y != 29.9609375)
            throw new InvalidDataException("Controlled two-dimensional positioning oracle failed.");
        Capture("empty", new TextShapedRun(new TextShapeRequest("", 0, dejavu),
            Array.Empty<TextShapedGlyph>(), "controlled-empty"), 4, 0, 0);
        void Refuse(string role, Action action, TextLayoutError expected)
        {
            try { action(); }
            catch (TextLayoutException error) when (error.Error == expected)
            { rows.Add(new { role, outcome = "rejected", error = error.Error.ToString() }); return; }
            throw new InvalidDataException("Positioning refusal failed: " + role);
        }
        Refuse("refused-zero-size", () => TextPositionedRun.Create(controlled, 0), TextLayoutError.InvalidRequest);
        Refuse("refused-output-budget", () => TextPositionedRun.Create(controlled, 4, 1_000_000), TextLayoutError.BudgetExceeded);
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        try { TextPositionedRun.Create(controlled, 4, cancellationToken: cancellation.Token); }
        catch (OperationCanceledException error) when (error.CancellationToken == cancellation.Token)
        { rows.Add(new { role = "refused-caller-cancellation", outcome = "canceled" }); }
        if (rows.Count != 13) throw new InvalidDataException("Positioning cancellation did not preserve the caller token.");
        await File.WriteAllTextAsync(Path.Combine(output, "r22-positioned-run-query.json"), JsonSerializer.Serialize(new {
            sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"), cases = rows,
            scope = "Pure managed projection of exact shaped runs to millimeters; controlled 2D case is not native. No reshaping, style/paint resolution, final-line/bidi composition, pagination, Flow or new writer/reader admission."
        }, new JsonSerializerOptions { WriteIndented = true }) + "\n");
        Console.WriteLine("R22 exact shaped-run millimeter positioning queries passed from candidate packages.");
    }
}
