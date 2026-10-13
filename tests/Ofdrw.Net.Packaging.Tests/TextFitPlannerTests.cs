using Ofdrw.Net.Layout.Editing.Internal;

namespace Ofdrw.Net.Packaging.Tests;

public class TextFitPlannerTests
{
    internal static OfdTextFitBox Ink => new(.5, -2, 2.5, 0);
    internal static OfdTextFitMetric Metric(int scalar, int start = 0, double advance = 3, OfdTextFitBox? ink = null, bool inkless = false,
        double advanceY = 0, uint glyphId = 1, int? length = null) => new(scalar, start, length ?? (scalar > 0xffff ? 2 : 1), glyphId, advance, advanceY, inkless ? null : ink ?? Ink);
    internal static OfdTextFitMetric[] Metrics(string text, double advance = 3, OfdTextFitBox? ink = null, bool inkless = false)
    {
        var result = new List<OfdTextFitMetric>();
        for (int i = 0; i < text.Length;)
        { int scalar = char.ConvertToUtf32(text, i); result.Add(Metric(scalar, i, advance, ink, inkless)); i += scalar > 0xffff ? 2 : 1; }
        return result.ToArray();
    }
    internal static OfdTextFitSourceCluster[] Sources(string text = "ABC", double x = 0, double advance = 3, OfdTextFitBox? ink = null, bool inkless = false)
        => Metrics(text, advance, ink, inkless).Select((metric, i) => new OfdTextFitSourceCluster(metric, x + i * advance, 4, "slot", metric.Utf16Start, metric.Utf16Length)).ToArray();
    internal static OfdTextFitObject Object(string text = "ABC", IReadOnlyList<OfdTextFitSourceCluster>? sources = null,
        string occurrence = "object", string owner = "owner", string version = "version", string scope = "scope", string style = "font-style",
        OfdTextFitAffine? affine = null, OfdTextFitBox? clip = null, OfdTextFitBox? boundary = null, bool complete = true,
        OfdTextFitProfile profile = OfdTextFitProfile.SimpleScalarMetrics, int page = 0)
        => new(page, occurrence, owner, version, scope, style, text, sources ?? Sources(text), affine ?? new OfdTextFitAffine(1, 0, 0, 1),
            boundary ?? new OfdTextFitBox(0, 0, 20, 10), clip ?? new OfdTextFitBox(0, 0, 20, 10), complete, profile);
    internal static OfdTextFitRequest Request(string replacement = "--", int start = 1, int count = 1, string expected = "B", string occurrence = "object",
        IReadOnlyList<OfdTextFitMetric>? metrics = null, string style = "font-style")
        => new(occurrence, start, count, expected, replacement, style, metrics ?? Metrics(replacement, 1.25, new OfdTextFitBox(.25, -1.5, 1, -.5)));
    internal static OfdTextFitInput Input(IReadOnlyList<OfdTextFitObject>? objects = null, IReadOnlyList<OfdTextFitRequest>? requests = null,
        IReadOnlyList<OfdTextFitNeighbor>? neighbors = null)
        => new("version", "scope", objects ?? new[] { Object() }, requests ?? new[] { Request() }, neighbors ?? Array.Empty<OfdTextFitNeighbor>());
    internal static OfdTextFitPlanResult Plan(OfdTextFitInput? input = null, OfdTextFitLimits? limits = null)
        => OfdTextFitPlanner.Plan(input ?? Input(), limits ?? new OfdTextFitLimits());
    internal static void Failure(OfdTextFitFailure expected, OfdTextFitInput input)
    { var result = Plan(input); Assert.False(result.FitsSuppliedGeometry); Assert.Equal(expected, result.Failure); Assert.Empty(result.Plans); }

    [Fact]
    public void VariableScalarCountPreservesUnchangedSourceAnchorsAndMapping()
    {
        var obj = Object(); var result = Plan(Input(new[] { obj })); Assert.True(result.FitsSuppliedGeometry);
        var plan = Assert.Single(result.Plans);
        Assert.Equal("A--C", plan.Text); Assert.Equal(new[] { 0d, 3, 4.25, 6 }, plan.Placements.Select(p => p.X));
        Assert.Equal(5.5, plan.Placements[2].X + plan.Placements[2].Metric.AdvanceX);
        Assert.Same(obj.Clusters[0], plan.Placements[0].Source); Assert.Same(obj.Clusters[2], plan.Placements[3].Source);
        Assert.Equal(2, plan.Placements[3].Source!.Metric.Utf16Start); Assert.Equal(3, plan.Placements[3].OutputUtf16Start);
        Assert.Null(plan.Placements[1].Source); Assert.Null(plan.Placements[2].Source);
        Assert.Equal(OfdTextFitDisposition.ReplaceScalars, Assert.Single(plan.Requests).Disposition);
        Assert.Equal(4, result.Statistics.OutputPositions); Assert.Equal(4, result.Statistics.OutputUtf16);
    }

    [Theory]
    [InlineData("", "AC", 2)]
    [InlineData("-", "A-C", 3)]
    [InlineData("--", "A--C", 4)]
    public void WidthRatherThanScalarCountControlsFit(string replacement, string expected, int count)
    { var plan = Assert.Single(Plan(Input(requests: new[] { Request(replacement) })).Plans); Assert.Equal(expected, plan.Text); Assert.Equal(count, plan.Placements.Count); Assert.Equal(6, plan.Placements.Last().X); }

    [Theory]
    [InlineData(3.25, .25, -1.5, 1, -.5)]
    [InlineData(1.25, .25, -2.125, 1, -.5)]
    [InlineData(1.25, .25, -1.5, 1, .125)]
    [InlineData(1.25, -.125, -1.5, 1, -.5)]
    public void WidthHeightBottomAndNegativeBearingCannotExpandRegion(double advance, double left, double top, double right, double bottom)
        => Failure(OfdTextFitFailure.ReplacementOutsideRegion, Input(requests: new[] { Request("-", metrics: Metrics("-", advance, new(left, top, right, bottom))) }));

    [Theory]
    [InlineData(3, true)]
    [InlineData(3.0009765625, false)]
    public void InklessSpaceStillReservesTerminalAdvance(double advance, bool fits)
    {
        var result = Plan(Input(requests: new[] { Request(" ", metrics: Metrics(" ", advance, inkless: true)) }));
        Assert.Equal(fits, result.FitsSuppliedGeometry);
        if (!fits) { Assert.Equal(OfdTextFitFailure.ReplacementOutsideRegion, result.Failure); Assert.Empty(result.Plans); }
    }

    [Fact]
    public void AllInklessTargetCannotInventHeightButCanBeDeleted()
    {
        var obj = Object(" ", Sources(" ", inkless: true));
        Failure(OfdTextFitFailure.UnprovedHeight, Input(new[] { obj }, new[] { Request("-", 0, 1, " ") }));
        var plan = Assert.Single(Plan(Input(new[] { obj }, new[] { Request("", 0, 1, " ") })).Plans);
        Assert.True(plan.RetainRootWithCanonicalEmptyTextCode); Assert.Equal("", plan.Text); Assert.Empty(plan.Placements);
        Assert.Equal(0, plan.EmptyAnchorX); Assert.Equal(4, plan.EmptyAnchorY); Assert.Equal("owner", plan.Owner);
        Assert.Equal(OfdTextFitDisposition.RetainRootWithCanonicalEmptyTextCode, Assert.Single(plan.Requests).Disposition);
    }

    [Fact]
    public void FullDeletionRetainsRootAndOriginalFirstAnchor()
    {
        var obj = Object(sources: Sources(x: 2));
        var plan = Assert.Single(Plan(Input(new[] { obj }, new[] { Request("", 0, 3, "ABC") })).Plans);
        Assert.Empty(plan.Placements); Assert.True(plan.RetainRootWithCanonicalEmptyTextCode); Assert.Equal(2, plan.EmptyAnchorX); Assert.Equal(4, plan.EmptyAnchorY);
        Assert.Equal(obj.FontStyle, plan.FontStyle); Assert.Equal(obj.Occurrence, plan.Occurrence);
    }

    [Fact]
    public void TranslationAppliedOnceAndLocalAnchorsPreserved()
    {
        var obj = Object(affine: new(1, 0, 0, 1, 10, 20), clip: new(10, 20, 30, 30));
        var plan = Assert.Single(Plan(Input(new[] { obj })).Plans); Assert.Equal(3, plan.Placements[1].X); Assert.Equal(4, plan.Placements[1].Y);
        Failure(OfdTextFitFailure.Collision, Input(new[] { obj }, neighbors: new[] { new OfdTextFitNeighbor(0, "graphic", new(13, 22, 14, 23)) }));
        Failure(OfdTextFitFailure.OriginalOutsideRegion, Input(new[] { Object(affine: new(1, 0, 0, 1, 10, 20), clip: new(13.5, 20, 30, 30)) }));
    }

    [Theory]
    [InlineData(2, 0, 0, 1)]
    [InlineData(-1, 0, 0, 1)]
    [InlineData(0, 1, -1, 0)]
    [InlineData(1, 0, 1, 1)]
    public void UnsupportedLinearTransformsRejectEvenNoOp(double a, double b, double c, double d)
        => Failure(OfdTextFitFailure.UnsupportedTransform, Input(new[] { Object(affine: new(a, b, c, d)) }, new[] { Request("B", metrics: Metrics("B")) }));

    [Fact]
    public void NonfiniteCoordinatesAndOverflowingEndpointsReject()
    {
        Failure(OfdTextFitFailure.UnsupportedTransform, Input(new[] { Object(affine: new(1, 0, 0, 1, double.PositiveInfinity)) }));
        var sources = Sources(); sources[1] = new(Metric('B', 1, double.MaxValue), double.MaxValue, 4, "slot", 1, 1);
        Failure(OfdTextFitFailure.InvalidGeometry, Input(new[] { Object(sources: sources) }));
        Failure(OfdTextFitFailure.InvalidGeometry, Input(new[] { Object(boundary: new(0, 0, double.NaN, 10)) }));
        Failure(OfdTextFitFailure.OriginalOutsideRegion, Input(new[] { Object(boundary: new(0, 0, 5, 10)) }));
    }

    [Fact]
    public void FiniteButCollapsedProgressionAndTranslatedInkReject()
    {
        const double x = 1e16;
        var sources = new[] { new OfdTextFitSourceCluster(Metric('A', advance: 4, ink: new(0, -2, 2, 0)), x, 4, "slot", 0, 1) };
        var obj = Object("A", sources, boundary: new(0, 0, x + 32, 10), clip: new(0, 0, x + 32, 10));
        Failure(OfdTextFitFailure.UnsupportedProgression, Input(new[] { obj }, new[] { Request("-", 0, 1, "A", metrics: Metrics("-", .25, new(0, -1, 2, 0))) }));
        var collapsedSource = Object(sources: Sources(ink: new(.5, -2, 1, 0)), affine: new(1, 0, 0, 1, x, 0), clip: new(x, 0, x + 32, 10));
        Failure(OfdTextFitFailure.InvalidGeometry, Input(new[] { collapsedSource }));
        var translated = Object("A", Sources("A", advance: 4, ink: new(0, -2, 2, 0)), affine: new(1, 0, 0, 1, x, 0), clip: new(x, 0, x + 32, 10));
        Failure(OfdTextFitFailure.InvalidGeometry, Input(new[] { translated }, new[] { Request("-", 0, 1, "A", metrics: Metrics("-", 2, new(.5, -1, 1, 0))) }));
    }

    [Theory]
    [InlineData(5.5, true)]
    [InlineData(5.4990234375, false)]
    public void EdgeTouchAllowedButTinyStrictOverlapRejected(double neighborLeft, bool fits)
    {
        var result = Plan(Input(neighbors: new[] { new OfdTextFitNeighbor(0, "graphic", new(neighborLeft, 2, 6, 4)) }));
        Assert.Equal(fits, result.FitsSuppliedGeometry); if (!fits) Assert.Equal(OfdTextFitFailure.Collision, result.Failure);
    }

    [Fact]
    public void UnchangedTextOnAnotherObjectAndNoOpRemainObstacles()
    {
        var second = Object("D", Sources("D", 3), "second", "owner2");
        Failure(OfdTextFitFailure.Collision, Input(new[] { Object(), second }, new[] { Request(), Request("D", 0, 1, "D", "second", Metrics("D")) }));
        Failure(OfdTextFitFailure.Collision, Input(new[] { Object(), second }));
    }

    [Fact]
    public void AllInsertedFootprintsCheckedJointlyAcrossBatch()
    {
        var second = Object(occurrence: "second", owner: "owner2");
        Failure(OfdTextFitFailure.Collision, Input(new[] { Object(), second }, new[] { Request(), Request(occurrence: "second") }));
        Assert.True(Plan(Input(new[] { Object(), Object(occurrence: "second", owner: "owner2", page: 1) }, new[] { Request(), Request(occurrence: "second") })).FitsSuppliedGeometry);
    }

    [Fact]
    public void IncompleteCoverageAnywhereOnChangedPageRejects()
    {
        Failure(OfdTextFitFailure.UnprovedNeighbors, Input(new[] { Object(complete: false) }));
        Failure(OfdTextFitFailure.UnprovedNeighbors, Input(new[] { Object(), Object(occurrence: "second", owner: "owner2", complete: false) }));
        Assert.True(Plan(Input(new[] { Object(complete: false) }, new[] { Request("B", metrics: Metrics("B")) })).FitsSuppliedGeometry);
    }

    [Fact]
    public void SemanticOwnerAliasAndDuplicateOccurrenceReject()
    {
        Failure(OfdTextFitFailure.AliasedOwner, Input(new[] { Object(), Object(occurrence: "second", page: 1) }));
        Failure(OfdTextFitFailure.InvalidIdentity, Input(new[] { Object(), Object(owner: "owner2") }));
    }

    [Theory]
    [InlineData(1, 1, "B")]
    [InlineData(0, 2, "AB")]
    public void OverlapRejectsIncludingNoOps(int start, int count, string expected)
        => Failure(OfdTextFitFailure.OverlappingRequests, Input(requests: new[] { Request(), Request(expected, start, count, expected, metrics: Metrics(expected)) }));

    [Fact]
    public void AdjacentIntervalsCannotBorrowEachOthersWidth()
    {
        var requests = new[] { Request("-", 0, 1, "A"), Request("-", 1, 1, "B") };
        var plan = Assert.Single(Plan(Input(requests: requests)).Plans); Assert.Equal("--C", plan.Text); Assert.Equal(new[] { 0d, 3, 6 }, plan.Placements.Select(p => p.X));
        Failure(OfdTextFitFailure.ReplacementOutsideRegion, Input(requests: new[] { Request("-", 0, 1, "A", metrics: Metrics("-", 3.25)), Request("", 1, 1, "B") }));
    }

    [Theory]
    [InlineData(0, 3)]
    [InlineData(-1, 3)]
    [InlineData(4, 3)]
    public void ChangedObjectsRejectZeroNegativeOrOverlappingAdvance(double advance, double nextX)
    {
        var sources = Sources(); sources[0] = new(Metric('A', 0, advance), 0, 4, "slot", 0, 1);
        sources[1] = new(Metric('B', 1), nextX, 4, "slot", 1, 1);
        Failure(OfdTextFitFailure.UnsupportedProgression, Input(new[] { Object(sources: sources) }));
    }

    [Fact]
    public void DifferentBaselineAndAdvanceYRejectForChanges()
    {
        var sources = Sources(); sources[2] = new(Metric('C', 2), 6, 4.125, "slot", 2, 1);
        Failure(OfdTextFitFailure.UnsupportedProgression, Input(new[] { Object(sources: sources) }));
        sources[2] = new(Metric('C', 2, advanceY: .125), 6, 4, "slot", 2, 1);
        Failure(OfdTextFitFailure.UnsupportedProgression, Input(new[] { Object(sources: sources) }));
    }

    [Fact]
    public void NoOpPreservesOriginalMetricsAndSkipsChangedFit()
    {
        var sources = Sources(); sources[1] = new(Metric('B', 1, 0), 3, 4, "slot", 1, 1);
        var obj = Object(sources: sources, complete: false);
        var plan = Assert.Single(Plan(Input(new[] { obj }, new[] { Request("B", metrics: Metrics("B", 100)) })).Plans);
        Assert.Equal(OfdTextFitDisposition.Unchanged, Assert.Single(plan.Requests).Disposition);
        Assert.Same(sources[1].Metric, plan.Placements[1].Metric); Assert.Equal(0, plan.Placements[1].Metric.AdvanceX);
    }

    [Fact]
    public void RawSpanAliasesRejectButDifferentSlotsMayReuseOffsets()
    {
        var sources = Sources(); sources[1] = new(Metric('B', 1), 3, 4, "slot", 0, 1);
        Failure(OfdTextFitFailure.InvalidRawMapping, Input(new[] { Object(sources: sources) }));
        sources[1] = new(Metric('B', 1), 3, 4, "other-slot", 0, 1);
        Assert.True(Plan(Input(new[] { Object(sources: sources) })).FitsSuppliedGeometry);
        sources[1] = new(Metric('B', 1), 3, 4, "slot", int.MaxValue, 1);
        Failure(OfdTextFitFailure.InvalidRawMapping, Input(new[] { Object(sources: sources) }));
    }

    [Fact]
    public void SupplementaryScalarCoveragePreservesUtf16Offsets()
    {
        var obj = Object("A😀C");
        var plan = Assert.Single(Plan(Input(new[] { obj }, new[] { Request("--", expected: "😀") })).Plans);
        Assert.Equal("A--C", plan.Text); Assert.Equal(3, plan.Placements.Last().OutputUtf16Start);
        Assert.Equal(3, plan.Placements.Last().Source!.RawStart);
        var unicodeReplacement = Request("😀", metrics: Metrics("😀", 1.25, new(.25, -1.5, 1, -.5)));
        var result = Assert.Single(Plan(Input(requests: new[] { unicodeReplacement })).Plans);
        Assert.Equal("A😀C", result.Text); Assert.Equal(3, result.Placements.Last().OutputUtf16Start);
    }

    [Theory]
    [InlineData(0, 2, 1)]
    [InlineData(1, 1, 0)]
    [InlineData(2, 1, 1)]
    public void PartialWrongOrMissingGlyphMappingRejects(int start, int length, uint glyph)
    {
        var sources = Sources(); sources[1] = new(Metric('B', start, glyphId: glyph, length: length), 3, 4, "slot", 1, 1);
        Failure(OfdTextFitFailure.InvalidScalarMapping, Input(new[] { Object(sources: sources) }));
    }

    [Fact]
    public void LoneSurrogateControlAndIncompleteMetricCoverageReject()
    {
        var source = new[] { new OfdTextFitSourceCluster(Metric(0xd800), 0, 4, "slot", 0, 1) };
        Failure(OfdTextFitFailure.InvalidScalarMapping, Input(new[] { Object("\ud800", source) }));
        source = new[] { new OfdTextFitSourceCluster(Metric('\n'), 0, 4, "slot", 0, 1) };
        Failure(OfdTextFitFailure.InvalidScalarMapping, Input(new[] { Object("\n", source) }));
        Failure(OfdTextFitFailure.InvalidScalarMapping, Input(requests: new[] { Request("--", metrics: Metrics("-")) }));
        Failure(OfdTextFitFailure.InvalidScalarMapping, Input(requests: new[] { Request("-", metrics: Metrics("-", inkless: true)) }));
    }

    [Fact]
    public void VersionScopeStyleProfileAndMatchedTextMustAgree()
    {
        Failure(OfdTextFitFailure.InvalidIdentity, Input(new[] { Object(version: "other") }));
        Failure(OfdTextFitFailure.InvalidIdentity, Input(new[] { Object(scope: "other") }));
        Failure(OfdTextFitFailure.InvalidRequest, Input(requests: new[] { Request(style: "other") }));
        Failure(OfdTextFitFailure.UnsupportedProfile, Input(new[] { Object(profile: OfdTextFitProfile.Unsupported) }));
        Failure(OfdTextFitFailure.InvalidRequest, Input(requests: new[] { Request(expected: "C") }));
        Failure(OfdTextFitFailure.InvalidRequest, Input(requests: new[] { Request(start: int.MaxValue, count: int.MaxValue) }));
    }

    [Fact]
    public void MutableConstructorArraysAreCapturedAndResultCollectionsRejectMutation()
    {
        var sources = Sources(); var originalB = sources[1]; var obj = Object(sources: sources);
        sources[1] = Sources("X")[0]; Assert.Same(originalB, obj.Clusters[1]);
        var metrics = Metrics("--", 1.25, new(.25, -1.5, 1, -.5)); var request = Request(metrics: metrics);
        metrics[0] = Metric('W'); Assert.Equal('-', request.Metrics[0].Scalar);
        var objects = new[] { obj }; var requests = new[] { request }; var neighbors = new[] { new OfdTextFitNeighbor(1, "offpage", new(0, 0, 1, 1)) };
        var input = Input(objects, requests, neighbors); objects[0] = Object("X"); requests[0] = Request(""); neighbors[0] = new(0, "collision", new(3, 2, 4, 3));
        var plan = Assert.Single(Plan(input).Plans); Assert.Equal("A--C", plan.Text);
        Assert.False(plan.Placements is OfdTextFitPlacement[]);
        Assert.Throws<NotSupportedException>(() => ((IList<OfdTextFitPlacement>)plan.Placements)[0] = plan.Placements[1]);
        Assert.Throws<NotSupportedException>(() => ((IList<OfdTextFitSourceCluster>)obj.Clusters)[0] = originalB);
    }

    [Fact]
    public void EmptyRequestListPublishesEmptyPlanAndLastBatchFailurePublishesNothing()
    {
        Assert.Empty(Plan(Input(requests: Array.Empty<OfdTextFitRequest>())).Plans);
        var result = Plan(Input(requests: new[] { Request("-", 0, 1, "A"), Request("W", metrics: Metrics("W", 4)) }));
        Assert.Equal(OfdTextFitFailure.ReplacementOutsideRegion, result.Failure); Assert.Equal(1, result.RequestIndex); Assert.Empty(result.Plans);
    }
}
