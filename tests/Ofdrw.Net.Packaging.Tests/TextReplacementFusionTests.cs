using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementFusionTests(Xunit.Abstractions.ITestOutputHelper log)
{
    [Theory]
    [InlineData("path", 1)]
    [InlineData("rotation", 1)]
    public async Task FullQueryNonHitPageDoesNotBroadenFitNeighborsOrObjectCapacity(string kind, int count)
    {
        var source = await TwoPages(kind, count);
        var originalOther = source.Pages[1].Elements.Select(SourceXml).ToArray();
        var hit = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits);
        using var external = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, [new(hit, "--")]);
        var expected = external.CreatePackageCopy();
        var provider = new TextReplacementTests.Metrics();
        using var fused = new OfdTextReplacer(provider).Replace(source, "B", "--");
        Assert.Single(fused.Statistics.Analyses); Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Disposals);
        var actual = fused.CreatePackageCopy();
        Assert.Equal(TextReplacementTests.Text(expected).SourceXml, TextReplacementTests.Text(actual).SourceXml);
        Assert.Equal(originalOther, actual.Pages[1].Elements.Select(SourceXml));
        Assert.Equal(count, actual.Pages[1].Elements.Count);
        var readback = await TextReplacementTests.Roundtrip(actual);
        Assert.Equal("A--C", TextReplacementTests.Text(readback).Text);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(513)]
    public async Task EveryNeighborAndObjectLimitStillAppliesOnTheAffectedPage(int count)
    {
        var source = await TwoPages("path", count, affected: true);
        var error = Assert.ThrowsAny<Exception>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "-"));
        if (count == 1) Assert.Equal("UnprovedReplacementNeighbors", error.Message);
        else
        {
            Assert.IsType<InvalidDataException>(error); // Full operation reaches its independent material cap first.
            var isolated = Assert.Throws<InvalidDataException>(() => IsolateFitView(source));
            Assert.Equal("ReplacementPageObjectLimit", isolated.Message);
        }
    }

    [Fact]
    public async Task Original513NonHitObjectsRetainOuterLimitEvidenceAndIsolatedAffectedPageProof()
    {
        var source = await TwoPages("path", 513);
        var hit = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits);
        var external = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, [new(hit, "-")]));
        var fused = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "-"));
        Assert.Equal(513, source.Pages[1].Elements.Count);
        Assert.Single(IsolateFitView(source));
        log.WriteLine(JsonSerializer.Serialize(new { count = 513, fullExternalRefusal = external.Message, fullFusedRefusal = fused.Message,
            externalPhase = external.Data["Ofdrw.TextReplacement.Phase"], fusedPhase = fused.Data["Ofdrw.TextReplacement.Phase"],
            note = "Full operation does not pass. Internal live-capture test isolates affected-page filtering without whole-package callback guards; product guards and fixture count are unchanged." }));
    }

    private static List<OfdTextReplacementRoot> IsolateFitView(OfdDocumentPackage source)
    {
        using var operation = new OfdTextReplacementOperation(source, new(), default); List<OfdTextReplacementRoot>? roots = null;
        // Deliberately isolates fit scope; not a replacement/publication path or acceptance of the complete operation.
        new OfdKeywordAnalysis(new TextReplacementTests.Metrics(), new(), capture: capture =>
            roots = OfdTextReplacementFitAdapter.Plan(capture, [new(Assert.Single(capture.CurrentHits), "-")],
                operation, default, new HashSet<int> { 0 })).Find(source, [0, 1], "B");
        return roots!;
    }

    [Fact]
    public async Task UnknownSourceOnNonHitPageStillRejectsTheCompleteQuery()
    {
        var source = await TwoPages("path", 1);
        source.Pages[1].Elements.Add(new OfdRawElement { ObjectId = "9999", LayerId = "21", Xml = "<ofd:Unknown xmlns:ofd='http://www.ofdspec.org/2016' ID='9999'/>" });
        Assert.Throws<NotSupportedException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "-"));
    }

    [Theory]
    [InlineData("open", "mutate")]
    [InlineData("measure", "mutate")]
    [InlineData("replacement", "mutate")]
    [InlineData("statistics", "mutate")]
    [InlineData("dispose", "mutate")]
    [InlineData("open", "throw")]
    [InlineData("measure", "throw")]
    [InlineData("replacement", "throw")]
    [InlineData("statistics", "throw")]
    [InlineData("dispose", "throw")]
    [InlineData("open", "cancel")]
    [InlineData("measure", "cancel")]
    [InlineData("replacement", "cancel")]
    [InlineData("statistics", "cancel")]
    [InlineData("dispose", "cancel")]
    public async Task EveryActualFusedCallbackRemainsGuardedAndTerminal(string phase, string action)
    {
        var source = await TwoPages("path", 1); var before = TextReplacementTests.Text(source).SourceXml;
        using var cts = new CancellationTokenSource(); var provider = new TextReplacementTests.Metrics();
        provider.Callback = value =>
        {
            if (value != (phase == "replacement" ? "measure" : phase) || (phase == "replacement" && provider.Measures != 2)) return;
            if (action == "mutate") source.Pages[1].Elements[0].XMillimeters++;
            else if (action == "throw") throw new ApplicationException("Injected fused callback");
            else cts.Cancel();
        };
        var error = Assert.ThrowsAny<Exception>(() => new OfdTextReplacer(provider).Replace(source, "B", "-", cancellationToken: cts.Token));
        Assert.IsType(action switch { "mutate" => typeof(InvalidDataException), "throw" => typeof(ApplicationException), _ => typeof(OperationCanceledException) }, error);
        Assert.Equal(phase == "open" && action == "throw" ? 0 : 1, provider.Disposals);
        Assert.Equal(before, TextReplacementTests.Text(source).SourceXml);
        Assert.NotNull(error.Data["Ofdrw.TextReplacement.Trace"]);
    }

    [Theory]
    [InlineData("count")]
    [InlineData("indexer")]
    public async Task CallerSelectionCannotMutatePreviouslyCapturedWholeSource(string phase)
    {
        var source = await TwoPages("path", 1);
        var selection = new CallbackSelection(phase, () => source.Options.Metadata.Title = "selection mutation");
        Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, selection, "B", "-"));
    }

    [Theory]
    [InlineData("B", "--", true)]
    [InlineData("B", "B", false)]
    [InlineData("missing", "-", false)]
    [InlineData("ABC", "", true)]
    public async Task BoundedTraceSeparatesObservedPreparationTransferAndIndependentNoOpCopy(string keyword, string replacement, bool changed)
    {
        var source = await TextReplacementTests.Read(); var provider = new TextReplacementTests.Metrics();
        var result = new OfdTextReplacer(provider).Replace(source, keyword, replacement);
        var owned = result.Claim();
        try
        {
            var prep = owned.TraceSnapshot(); var prepWork = owned.Context.Material.ValidationWorkBytes;
            Assert.Equal(prepWork, prep.Sum(value => value.ValidationWorkBytes));
            Assert.Equal(owned.Context.Material.CloneCopyBytes, prep.Sum(value => value.CopyBytes));
            Assert.Equal(1, prep.Sum(value => value.SourceCaptures));
            Assert.Equal(changed ? 1 : 0, prep.Sum(value => value.OwnedCloneAttempts));
            Assert.Equal(changed ? 1 : 0, prep.Sum(value => value.CandidateCaptureAttempts));
            Assert.Equal(provider.Opens + provider.Measures + provider.StatisticsCalls + provider.Disposals, prep.Sum(value => value.CallbackBoundaryAttempts));
            Assert.Equal(0, prep.Sum(value => value.RetainedFontDigestBytes));
            var package = owned.TakePackage(default); var transferred = owned.TraceSnapshot();
            Assert.Equal(changed ? 1 : 0, transferred.Sum(value => value.CandidateGuardAttempts));
            Assert.Equal(changed ? 0 : 1, transferred.Sum(value => value.IndependentNoOpCopyOperations));
            Assert.Equal(changed ? 0 : 1, transferred.Sum(value => value.IndependentNoOpCloneAttempts));
            Assert.True(owned.Context.Material.ValidationWorkBytes > prepWork);
            Assert.Equal(owned.Context.Material.ValidationWorkBytes, transferred.Sum(value => value.ValidationWorkBytes));
            log.WriteLine(JsonSerializer.Serialize(new { keyword, replacement, prep, transferred, analyses = result.Statistics.Analyses,
                independentNoOpCopy = !changed, note = "Trace counts attempts at owned API boundaries; shared clone internal hashes are inferred, not directly observed." }));
            Assert.NotSame(source, package);
        }
        finally { owned.Dispose(); }
    }

    [Fact]
    public async Task ExternalHitAuthenticationStillRehashesRetainedFontsAndUsesFreshAnalysis()
    {
        var source = await TextReplacementTests.Read(); var hit = Assert.Single(new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits);
        var provider = new TextReplacementTests.Metrics(); var result = new OfdTextReplacer(provider).Replace(source, [new(hit, "-")]);
        var owned = result.Claim();
        try
        {
            var trace = owned.TraceSnapshot(); Assert.Equal(6, trace.Sum(value => value.RetainedFontDigestBytes));
            Assert.Equal(2, trace.Sum(value => value.RetainedFontDigestOperations)); Assert.Equal(1, provider.Opens);
            log.WriteLine(JsonSerializer.Serialize(new { external = true, trace, analyses = result.Statistics.Analyses }));
        }
        finally { owned.Dispose(); }
    }

    [Fact]
    public async Task OwnedRequestCountAndCumulativeTextBoundsRejectInsideLiveCaptureBeforeFitOrClone()
    {
        var source = await TextReplacementTests.Read([("AAA", 2d, "3 3")]);
        foreach (var limits in new[] { new OfdTextReplacementLimits(maxRequests: 2), new OfdTextReplacementLimits(maxReplacementUtf16Units: 2) })
        {
            var provider = new TextReplacementTests.Metrics();
            var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, "A", "-", new(limits)));
            Assert.Equal(1, provider.Measures); Assert.Equal(1, provider.Disposals);
            var trace = Assert.IsType<OfdTextReplacementTraceCell[]>(error.Data["Ofdrw.TextReplacement.Trace"]);
            Assert.Equal(0, trace.Sum(value => value.OwnedCloneAttempts));
            Assert.Equal("AuthenticateKeywordMatches", error.Data["Ofdrw.TextReplacement.Phase"]);
        }
    }

    [Fact]
    public async Task ChangedExternalBatchRetainsIdentityRequestPageInFitScope()
    {
        var source = await TwoPages("identity-rotation", 1);
        var hits = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits; Assert.Equal(2, hits.Count);
        var error = Assert.Throws<NotSupportedException>(() => new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source,
            [new(hits[0], "-"), new(hits[1], "B")]));
        Assert.Equal("UnsupportedReplacementTransform", error.Message);
    }

    [Fact]
    public async Task FusedOverlappingMatchesRejectWithoutClone()
    {
        var source = await TextReplacementTests.Read([("AAA", 2d, "3 3")]); var provider = new TextReplacementTests.Metrics();
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, "AA", "-"));
        Assert.Equal("OverlappingReplacementRequests", error.Message);
        Assert.Equal(0, Assert.IsType<OfdTextReplacementTraceCell[]>(error.Data["Ofdrw.TextReplacement.Trace"]).Sum(value => value.OwnedCloneAttempts));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task SupplementaryReplacementRetainsScalarAndSuffixAnchor()
    {
        var source = await TextReplacementTests.Read(); using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, "B", "𠀀");
        var readback = await TextReplacementTests.Roundtrip(result.CreatePackageCopy()); Assert.Equal("A𠀀C", TextReplacementTests.Text(readback).Text);
        var found = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(readback, "C");
        Assert.Equal(23, Assert.Single(found.Hits).Fragments[0].BaselineMillimeters.X);
    }

    [Fact]
    public async Task UntrustedExceptionDataCannotReplaceTheOriginalProviderError()
    {
        var source = await TextReplacementTests.Read(); var sentinel = new RefusingDiagnosticException();
        var provider = new TextReplacementTests.Metrics { Callback = _ => throw sentinel };
        Assert.Same(sentinel, Assert.Throws<RefusingDiagnosticException>(() => new OfdTextReplacer(provider).Replace(source, "B", "-")));
        Assert.Equal(0, provider.Disposals);
    }

    [Fact]
    public async Task FailedTraceReservationDoesNotAllocateAFailureReportOrOpenProvider()
    {
        var source = await TextReplacementTests.Read(); var provider = new TextReplacementTests.Metrics();
        // Core's limit exception already carries one private limit marker.
        // Replacement reporting must preserve it without adding any entry.
        InvalidDataException? error = null;
        try { new OfdTextReplacer(provider).Replace(source, "B", "-", new(new OfdTextReplacementLimits(maxMaterializedBytes: 1))); }
        catch (InvalidDataException refused) { error = refused; }
        Assert.NotNull(error); Assert.Equal(0, provider.Opens);
        Assert.True(OfdTraversalBudget.IsLimit(error)); Assert.Single(error.Data.Cast<System.Collections.DictionaryEntry>());
    }

    private sealed class RefusingDiagnosticException : Exception
    { public override System.Collections.IDictionary Data => throw new InvalidOperationException("Diagnostic access refused"); }

    private static async Task<OfdDocumentPackage> TwoPages(string kind, int count, bool affected = false)
    {
        var source = await TextReplacementTests.Read();
        var page = new OfdPage { Id = "2", WidthMillimeters = 200, HeightMillimeters = 200 };
        source.Pages.Add(page); var target = affected ? source.Pages[0] : page;
        for (var i = 0; i < count; i++)
        {
            OfdElement element = kind == "path" ? new OfdPathElement { AbbreviatedData = "M 0 0 L 1 1" } :
                new OfdTextElement { FontResourceId = "10", FontName = "Fixture", FontSizeMillimeters = 4, Text = kind == "identity-rotation" ? "B" : "D",
                    Runs = { new() { Text = kind == "identity-rotation" ? "B" : "D", XMillimeters = 2, YMillimeters = 8 } }, Transform = new[] { 0d, 1, -1, 0, 20, 0 } };
            element.ObjectId = (1000 + i).ToString(); element.LayerId = affected ? "20" : "21";
            element.XMillimeters = 80; element.YMillimeters = 80; element.WidthMillimeters = 20; element.HeightMillimeters = 20;
            target.Elements.Add(element);
        }
        return await TextReplacementTests.Roundtrip(source);
    }
    private static string? SourceXml(OfdElement element) => element switch
    { OfdTextElement text => text.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
    private sealed class CallbackSelection(string phase, Action callback) : IReadOnlyList<int>
    {
        public int Count { get { if (phase == "count") callback(); return 2; } }
        public int this[int index] { get { if (phase == "indexer") callback(); return index; } }
        public IEnumerator<int> GetEnumerator() => new[] { 0, 1 }.AsEnumerable().GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
