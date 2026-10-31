using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using System.Globalization;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class NativeBoldStrokePlannerTests
{
    [Fact]
    public void CompletedProofsSurviveLaterBudgetExhaustionWithoutStartingMoreAnalysis()
    {
        var face = CounterFont(Font(), nested: false);
        using var cancellation = new CancellationTokenSource();
        var planner = new NativeBoldStrokePlanner(cancellation.Token);
        var width = planner.Width(face, false, 4, "中");
        var outline = Assert.IsType<NativeProtectedGlyph>(planner.ProtectedGlyph(face, 4, '中'));
        var budget = typeof(NativeBoldStrokePlanner).GetField("budget",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(planner)!;
        var work = budget.GetType().GetField("work",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        // Reach the real shared cap on the next uncached glyph. The fixture
        // stays small; no production limit or separate analysis budget changes.
        work.SetValue(budget, 50_000_000);
        Assert.Equal(.04, planner.Width(face, false, 4, "国"));
        Assert.Equal(50_000_001, work.GetValue(budget));
        Assert.Equal(width, planner.Width(face, false, 4, "中"));
        Assert.Same(outline, planner.ProtectedGlyph(face, 4, '中'));
        Assert.Null(planner.ProtectedGlyph(face, 4, '国'));
        Assert.Null(planner.ProtectedGlyph(face, 5, '中'));
        Assert.Null(planner.ProtectedGlyph(face.ToArray(), 4, '中'));
        Assert.Equal(.04, planner.Width(face, false, 4, "中国"));
        Assert.Equal(50_000_001, work.GetValue(budget));
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => planner.Width(face, false, 4, "中"));
        Assert.Throws<OperationCanceledException>(() => planner.ProtectedGlyph(face, 4, '中'));
    }

    [Theory]
    [InlineData(false, 3.701)]
    [InlineData(true, 3.701)]
    [InlineData(false, 3.702)]
    [InlineData(true, 3.702)]
    [InlineData(false, 3.7041666667)]
    [InlineData(true, 3.7041666667)]
    public void ProtectedCurvesAndCounterBoxesUseDeclaredSizeWithForceIntegerPpem(bool force, double requested)
    {
        var source = new OpenTypeFace(CounterFont(Font(), nested: false));
        var head = source.Table("head");
        OpenTypeFace.Put16(head, 16, force ? OpenTypeFace.U16(head, 16) | 8 : OpenTypeFace.U16(head, 16) & ~8);
        var face = source.Build(); var snapshot = face.ToArray();
        var written = double.Parse(requested.ToString("0.###", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        var baseline = 2.346 + written;
        var planner = new NativeBoldStrokePlanner(default);
        var text = new OfdTextElement
        {
            EmitSyntheticBoldStroke = true,
            SyntheticBoldStrokeWidthMillimeters = planner.Width(face, false, requested, "中"),
            Text = "中", FontSizeMillimeters = requested, Weight = 700,
            XMillimeters = 1.23456, YMillimeters = 2.34567,
            Runs = { new OfdTextRun { Text = "中", XMillimeters = .12345, YMillimeters = requested } }
        };
        Assert.True(new NativeProtectedRowStrokeBuilder(planner, default).TryCreate(new[] { text }, face,
            new OfdPage { WidthMillimeters = 60, HeightMillimeters = 80 }, out var paths));
        var path = Assert.Single(paths);
        var tokens = path.AbbreviatedData.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var points = new List<(double X, double Y)>();
        for (var index = 0; index < tokens.Length;)
        {
            if (tokens[index] == "C") { index++; continue; }
            Assert.Contains(tokens[index], new[] { "M", "L" });
            points.Add((double.Parse(tokens[index + 1], CultureInfo.InvariantCulture),
                double.Parse(tokens[index + 2], CultureInfo.InvariantCulture))); index += 3;
        }
        // Original design-unit rectangles at the separately serialized
        // boundary/run origins. ForcePPEMToInt must not enlarge the vector
        // outlines from3.704mm to3.7041667mm (10.5pt at72dpi).
        Assert.InRange(Math.Abs(points.Min(p => p.X) - 1.358), 0, .0000015);
        Assert.InRange(Math.Abs(points.Max(p => p.X) - (1.358 + written)), 0, .0000015);
        Assert.InRange(Math.Abs(points.Min(p => p.Y) - 2.346), 0, .0000015);
        Assert.InRange(Math.Abs(points.Max(p => p.Y) - baseline), 0, .0000015);
        var clip = Assert.Single(path.Clips!.Items).Areas.Single().Path;
        var clipTokens = clip.AbbreviatedData.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var hole = Enumerable.Range(0, clipTokens.Length).Where(i => clipTokens[i] == "M").Skip(1).Single();
        const double pad = .5005 * 25.4 / 72;
        Assert.InRange(Math.Abs(double.Parse(clipTokens[hole + 1], CultureInfo.InvariantCulture) -
            Math.Floor((1.358 + written * .1 - pad) * 1e6) / 1e6), 0, .000002);
        Assert.InRange(Math.Abs(double.Parse(clipTokens[hole + 2], CultureInfo.InvariantCulture) -
            Math.Floor((baseline - written * .55 - pad) * 1e6) / 1e6), 0, .000002);
        Assert.Equal(snapshot, face);
    }

    [Fact]
    public void HighlyCoincidentContoursRetainLegacyWithinTheIntersectionAdmissionBudget()
    {
        // Small, finite construction:64 identical filled rectangles produce
        // many coincident/touching edge pairs despite their simple final union.
        // Refuse that analysis workload before calling the clipping engine.
        (int X, int Y)[] rectangle = [(0,0),(0,1000),(1000,1000),(1000,0)];
        var face = BuildCounterFont(Font(), Enumerable.Repeat(rectangle, 64).ToArray());
        var snapshot = face.ToArray();
        var planner = new NativeBoldStrokePlanner(default);
        Assert.Equal(.04, planner.Width(face, false, 4, "中"));
        Assert.Null(planner.ProtectedGlyph(face, 4, '中'));
        Assert.Equal(snapshot, face);
    }

    [Fact]
    public void ProtectedOriginalCurvesAndOverlapHolesUseTheSerializedTextPositions()
    {
        var face = BuildCounterFont(Font(), new (int X, int Y)[][]
        {
            [(0,0),(0,1000),(100,1000),(100,0)],
            [(900,0),(900,1000),(1000,1000),(1000,0)],
            [(0,550),(0,1000),(1000,1000),(1000,550)],
            [(0,0),(0,450),(1000,450),(1000,0)]
        });
        var snapshot = face.ToArray();
        var planner = new NativeBoldStrokePlanner(default);
        var glyph = Assert.IsType<NativeProtectedGlyph>(planner.ProtectedGlyph(face, 2.1166667, '中'));
        Assert.Equal(new OpenTypeCmap(new OpenTypeFace(face)).DirectGlyph('中'), glyph.GlyphId);
        var counter = Assert.Single(glyph.Counters);
        const double size = 2.117, points = size * 72 / 25.4;
        Assert.InRange(counter.Left, .1 * points - .00001, .1 * points + .00001);
        Assert.InRange(counter.Right, .9 * points - .00001, .9 * points + .00001);
        Assert.InRange(counter.Top + glyph.BaselineOffsetPoints, -.55 * points - .00001, -.55 * points + .00001);
        Assert.InRange(counter.Bottom + glyph.BaselineOffsetPoints, -.45 * points - .00001, -.45 * points + .00001);
        var text = new OfdTextElement
        {
            EmitSyntheticBoldStroke = true, SyntheticBoldStrokeWidthMillimeters = .021,
            Text = "中中", FontSizeMillimeters = 2.1166667, Weight = 700,
            XMillimeters = 1.23456, YMillimeters = 2.34567,
            Runs = { new OfdTextRun { Text = "中中", XMillimeters = .12345,
                YMillimeters = 2.1166667, DeltaX = "2.123456" } }
        };
        var builder = new NativeProtectedRowStrokeBuilder(planner, default);
        Assert.True(builder.TryCreate(new[] { text }, face, new OfdPage
            { WidthMillimeters = 60, HeightMillimeters = 80 }, out var paths));
        var path = Assert.Single(paths);
        Assert.Null(path.SourceXml);
        Assert.Equal(.074, path.LineWidthMillimeters);
        Assert.True(OfdPathDataSyntax.IsValidLiteral(path.AbbreviatedData, default, requireDrawableSegment: true));
        var tokens = path.AbbreviatedData.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var starts = Enumerable.Range(0, tokens.Length).Where(i => tokens[i] == "M").ToArray();
        Assert.Equal(8, starts.Length); // two original four-piece glyphs, not union contours
        // A closed TrueType contour may start at its last on-curve point.
        // Check all four independent source rectangles, not an arbitrary
        // callback start vertex. The two origins use the writer's separate
        // boundary/run grid and the unchanged six-decimal advance.
        for (var glyphIndex = 0; glyphIndex < 2; glyphIndex++)
        {
            var x = 1.358 + glyphIndex * 2.123456;
            var rectangles = new[]
            {
                (Left:x, Top:4.463-size, Right:x+size*.1, Bottom:4.463),
                (Left:x+size*.9, Top:4.463-size, Right:x+size, Bottom:4.463),
                (Left:x, Top:4.463-size, Right:x+size, Bottom:4.463-size*.55),
                (Left:x, Top:4.463-size*.45, Right:x+size, Bottom:4.463)
            };
            for (var contour = 0; contour < 4; contour++)
            {
                var startIndex = glyphIndex * 4 + contour;
                var end = startIndex + 1 < starts.Length ? starts[startIndex + 1] : tokens.Length;
                var pointsInContour = new List<(double X, double Y)>();
                for (var index = starts[startIndex]; index < end;)
                {
                    if (tokens[index] == "C") { index++; continue; }
                    Assert.Contains(tokens[index], new[] { "M", "L" });
                    pointsInContour.Add((double.Parse(tokens[index+1], CultureInfo.InvariantCulture),
                        double.Parse(tokens[index+2], CultureInfo.InvariantCulture)));
                    index += 3;
                }
                var expected = rectangles[contour];
                // Six-decimal mm serialization plus float font callbacks.
                Assert.InRange(Math.Abs(pointsInContour.Min(p => p.X)-expected.Left), 0, .0000015);
                Assert.InRange(Math.Abs(pointsInContour.Max(p => p.X)-expected.Right), 0, .0000015);
                Assert.InRange(Math.Abs(pointsInContour.Min(p => p.Y)-expected.Top), 0, .0000015);
                Assert.InRange(Math.Abs(pointsInContour.Max(p => p.Y)-expected.Bottom), 0, .0000015);
            }
        }
        Assert.Equal(2, path.Clips!.Items.Count); // every overlap-built opening, both placements
        var clip = path.Clips.Items[0].Areas.Single().Path;
        Assert.Equal(OfdFillRule.EvenOdd, clip.FillRule);
        var clipTokens = clip.AbbreviatedData.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var holeStart = Enumerable.Range(0, clipTokens.Length).Where(i => clipTokens[i] == "M").Skip(1).Single();
        const double pad = .5005 * 25.4 / 72;
        Assert.Equal(Math.Floor((1.358 + size * .1 - pad) * 1e6) / 1e6,
            double.Parse(clipTokens[holeStart + 1], CultureInfo.InvariantCulture));
        Assert.Equal(Math.Floor((4.463 - size * .55 - pad) * 1e6) / 1e6,
            double.Parse(clipTokens[holeStart + 2], CultureInfo.InvariantCulture));
        Assert.Equal(snapshot, face);
        Assert.Equal("中中", text.Text);
        Assert.Equal(.021, text.SyntheticBoldStrokeWidthMillimeters);
        var sibling = new OfdTextElement
        {
            EmitSyntheticBoldStroke = true, SyntheticBoldStrokeWidthMillimeters = .021,
            Text = "中", FontSizeMillimeters = 2.1166667, Weight = 700,
            XMillimeters = 1.73456, YMillimeters = 2.34567,
            Runs = { new OfdTextRun { Text = "中", XMillimeters = .12345, YMillimeters = 2.1166667 } }
        };
        // These original glyphs overlap. Per-TextObject-only masks allow a
        // neighbor's added stroke to reach the other object's white opening.
        Assert.True(builder.TryCreate(new[] { text, sibling }, face,
            new OfdPage { WidthMillimeters = 60, HeightMillimeters = 80 }, out var groupPaths));
        Assert.Equal(2, groupPaths.Count);
        Assert.All(groupPaths, p => Assert.Equal(3, p.Clips!.Items.Count));
        Assert.Equal(groupPaths[0].Clips!.Items.Select(c => c.Areas.Single().Path.AbbreviatedData),
            groupPaths[1].Clips!.Items.Select(c => c.Areas.Single().Path.AbbreviatedData));
    }

    [Fact]
    public void AFailedMemberDoesNotPartiallySupplementAnEqualFormatRowGroup()
    {
        var face = CounterFont(Font(), nested: false);
        OfdTextElement Text(string value) => new()
        {
            EmitSyntheticBoldStroke = true, SyntheticBoldStrokeWidthMillimeters = .04,
            Text = value, FontSizeMillimeters = 4, Weight = 700,
            Runs = { new OfdTextRun { Text = value, YMillimeters = 4,
                DeltaX = value.Length > 1 ? "4" : null } }
        };
        var first = Text("中"); var mixed = Text("中A");
        var builder = new NativeProtectedRowStrokeBuilder(new NativeBoldStrokePlanner(default), default);
        Assert.False(builder.TryCreate(new[] { first, mixed }, face,
            new OfdPage { WidthMillimeters = 60, HeightMillimeters = 80 }, out var paths));
        Assert.Empty(paths);
        Assert.Equal(.04, first.SyntheticBoldStrokeWidthMillimeters);
        Assert.Equal(.04, mixed.SyntheticBoldStrokeWidthMillimeters);
        Assert.Null(first.SourceXml); Assert.Null(mixed.SourceXml);
        // This failure occurs after the first member's path was built, unlike
        // the mixed-script admission refusal above. No partial group escapes.
        var missing = Text("\u9fff");
        Assert.Equal(0, new OpenTypeCmap(new OpenTypeFace(face)).DirectGlyph('\u9fff'));
        Assert.False(builder.TryCreate(new[] { first, missing }, face,
            new OfdPage { WidthMillimeters = 60, HeightMillimeters = 80 }, out paths));
        Assert.Empty(paths);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        var canceledBuilder = new NativeProtectedRowStrokeBuilder(planner: new NativeBoldStrokePlanner(default), canceled.Token);
        Assert.Throws<OperationCanceledException>(() => canceledBuilder.TryCreate(new[] { first }, face,
            new OfdPage { WidthMillimeters = 60, HeightMillimeters = 80 }, out _));
    }

    [Theory]
    [InlineData(6d)]
    [InlineData(12d)]
    [InlineData(18d)]
    public void OverlappingSameDirectionPiecesCannotBypassTheAbsoluteCounterFloor(double points)
    {
        var original = Font();
        // Four overlapping clockwise rectangles have exactly the same filled
        // region as the explicit rectangular-counter control below. No input
        // contour is itself a hole; their union encloses the100-unit opening.
        var face = BuildCounterFont(original, new (int X, int Y)[][]
        {
            [(0,0),(0,1000),(100,1000),(100,0)],
            [(900,0),(900,1000),(1000,1000),(1000,0)],
            [(0,550),(0,1000),(1000,1000),(1000,550)],
            [(0,0),(0,450),(1000,450),(1000,0)]
        });
        var snapshot = face.ToArray();
        var originalGid = new OpenTypeCmap(new OpenTypeFace(original)).DirectGlyph('中');
        Assert.Equal(originalGid, new OpenTypeCmap(new OpenTypeFace(face)).DirectGlyph('中'));
        var size = points * 25.4 / 72;
        var written = double.Parse(size.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture),
            System.Globalization.CultureInfo.InvariantCulture);
        var width = new NativeBoldStrokePlanner(default).Width(face, false, size, "中");
        var legacy = size * .01;
        if (points == 6) Assert.Equal(legacy, width); // even unstroked radius is below0.4pt
        else
        {
            // Equal filled regions must retain the positive12/18pt controls;
            // always falling back for every overlap is not this correction.
            var explicitCounter = CounterFont(original, nested: false);
            var control = new NativeBoldStrokePlanner(default).Width(explicitCounter, false, size, "中");
            Assert.True(width > legacy);
            Assert.Equal(control, width);
        }
        if (width > legacy + 1e-9)
        {
            var radius = written * 72 / 25.4 * .05;
            var remaining = radius - width * 72 / 25.4 / 2;
            Assert.True(remaining >= Math.Max(.4, radius * .6) + .0005,
                $"{points}pt overlap-built counter retains only{remaining}pt after an increase.");
        }
        Assert.Equal(snapshot, face);
    }

    [Fact]
    public void NestedCounterCannotBorrowTheLargeInnerWhiteRegion()
    {
        var original = Font();
        var planner = new NativeBoldStrokePlanner(default);
        Assert.True(planner.Width(original, false, 4, "中") > .04); // plain rectangle control
        var nested = CounterFont(original, nested: true);
        var snapshot = nested.ToArray();
        // The outer white ring is only10 design units wide. Every point in
        // its9x9 grid lies inside the large nested inner hole. Those points
        // belong to the inner counter and cannot justify thickening the ring.
        Assert.Equal(.04, planner.Width(nested, false, 4, "中"));
        Assert.Equal(snapshot, nested);
    }

    [Fact]
    public void OpticalMinimumOnlyAllowsAnIncreaseWhenTheSerializedSizePreservesItsOpening()
    {
        var face = CounterFont(Font(), nested: false);
        var planner = new NativeBoldStrokePlanner(default);
        const double small = 25.4 * 6 / 72, medium = 25.4 * 12 / 72, large = 25.4 * 18 / 72;
        Assert.Equal(small * .01, planner.Width(face, false, small, "中"));
        var mediumWidth = planner.Width(face, false, medium, "中");
        var largeWidth = planner.Width(face, false, large, "中");
        Assert.InRange(mediumWidth, .139, .141);
        Assert.InRange(largeWidth, .221, .223);
        // The rectangular hole is100/1000em high. These are independent
        // analytic clearances at the serialized12pt/18pt sizes, including the
        // stricter target before the planner subtracts its sampling allowance.
        var mediumRadius = 4.233 * 72 / 25.4 * .05;
        Assert.True(mediumRadius - mediumWidth * 72 / 25.4 / 2 >= .4005);
        Assert.True(.9 - largeWidth * 72 / 25.4 / 2 >= .5405);
        // The6pt hole cannot meet0.4pt even without a stroke. The legacy
        // fallback above does not claim to satisfy the increased-width guard.
    }

    [Fact]
    public void MissingGlyphAndUnsupportedFontKeepTheLegacyStroke()
    {
        var original = Font();
        var planner = new NativeBoldStrokePlanner(default);
        Assert.Equal(.04, planner.Width(original, false, 4, "量"));
        var variable = new OpenTypeFace(original);
        variable.Tables["fvar"] = new byte[16];
        Assert.Equal(.04, planner.Width(variable.Build(), false, 4, "中"));
        Assert.Equal(.04, planner.Width(original, true, 4, "中"));
    }

    [Fact]
    public void CancellationIsNotConvertedIntoAnOpticalFallback()
    {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var planner = new NativeBoldStrokePlanner(cancellation.Token);
        Assert.Throws<OperationCanceledException>(() => planner.Width(Font(), false, 4, "中"));
    }

    private static byte[] Font()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
    }

    private static byte[] CounterFont(byte[] original, bool nested)
    {
        (int X, int Y)[] points = nested ? new (int, int)[]
        {
            (0,0),(0,1000),(1000,1000),(1000,0),
            (100,100),(900,100),(900,900),(100,900),
            (110,110),(110,890),(890,890),(890,110),
            (150,150),(850,150),(850,850),(150,850)
        } : new (int, int)[]
        {
            (0,0),(0,1000),(1000,1000),(1000,0),
            (100,450),(900,450),(900,550),(100,550)
        };
        return BuildCounterFont(original, points.Chunk(4).ToArray());
    }

    private static byte[] BuildCounterFont(byte[] original, (int X, int Y)[][] shapes)
    {
        var face = new OpenTypeFace(original); var gid = new OpenTypeCmap(face).DirectGlyph('中');
        var points = shapes.SelectMany(shape => shape).ToArray();
        var contours = shapes.Length;
        var flags = 12 + contours * 2;
        var xCoordinates = flags + points.Length;
        var yCoordinates = xCoordinates + points.Length * 2;
        var glyph = new byte[yCoordinates + points.Length * 2];
        OpenTypeFace.Put16(glyph, 0, contours); OpenTypeFace.Put16(glyph, 6, 1000); OpenTypeFace.Put16(glyph, 8, 1000);
        var endPoint = -1;
        for (var i = 0; i < contours; i++)
        {
            endPoint += shapes[i].Length;
            OpenTypeFace.Put16(glyph, 10 + i * 2, endPoint);
        }
        for (var i = 0; i < points.Length; i++)
        {
            glyph[flags + i] = 1;
            OpenTypeFace.Put16(glyph, xCoordinates + i * 2, unchecked((ushort)(points[i].X - (i == 0 ? 0 : points[i - 1].X))));
            OpenTypeFace.Put16(glyph, yCoordinates + i * 2, unchecked((ushort)(points[i].Y - (i == 0 ? 0 : points[i - 1].Y))));
        }
        var loca = face.Table("loca"); var shortLoca = OpenTypeFace.U16(face.Table("head"), 50) == 0;
        var count = OpenTypeFace.U16(face.Table("maxp"), 4);
        int Offset(int index) => shortLoca ? OpenTypeFace.U16(loca, index * 2) * 2 : (int)OpenTypeFace.U32(loca, index * 4);
        var start = Offset(gid); var end = Offset(gid + 1); var old = face.Table("glyf");
        var replacement = new byte[old.Length + glyph.Length - (end - start)];
        Buffer.BlockCopy(old, 0, replacement, 0, start); Buffer.BlockCopy(glyph, 0, replacement, start, glyph.Length);
        Buffer.BlockCopy(old, end, replacement, start + glyph.Length, old.Length - end);
        face.Tables["glyf"] = replacement;
        for (var i = count; i > gid; i--)
        {
            var offset = Offset(i) + glyph.Length - (end - start);
            if (shortLoca) OpenTypeFace.Put16(loca, i * 2, offset / 2);
            else OpenTypeFace.Put32(loca, i * 4, (uint)offset);
        }
        OpenTypeFace.Put16(face.Table("maxp"), 6, Math.Max(OpenTypeFace.U16(face.Table("maxp"), 6), points.Length));
        OpenTypeFace.Put16(face.Table("maxp"), 8, Math.Max(OpenTypeFace.U16(face.Table("maxp"), 8), contours));
        return face.Build();
    }
}
