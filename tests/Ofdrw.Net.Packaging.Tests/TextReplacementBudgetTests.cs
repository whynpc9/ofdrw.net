using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TextReplacementBudgetTests(Xunit.Abstractions.ITestOutputHelper log)
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LongFullAndLateHitsHaveLinearCorrelationVisits(bool late)
    {
        const int count = 1024; var value = new string('A', count); var delta = string.Join(" ", Enumerable.Repeat("3", count - 1));
        var source = await TextReplacementTests.Read([(value, 2d, delta)], 4000, 5000);
        var found = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, late ? "A" : value);
        var hit = late ? found.Hits[count - 1] : Assert.Single(found.Hits);
        using var result = new OfdTextReplacer(new TextReplacementTests.Metrics()).Replace(source, new[] { new OfdTextReplacement(hit, "-") });
        var statistics = Assert.Single(result.Statistics.Analyses);
        Assert.InRange(statistics.CorrelationVisits, count * 2, count * 4 + 32);
        Assert.True(statistics.LogicalValidationWorkUnits > 0);
        Assert.InRange(statistics.ReservedLogicalValidationWorkUnits - statistics.LogicalValidationWorkUnits, 0, 65535);
        log.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { count, late, statistics.CorrelationVisits,
            statistics.LogicalValidationWorkUnits, statistics.ReservedLogicalValidationWorkUnits, result.Statistics.Fit }));
        Assert.Equal(late ? value.Substring(0, count - 1) + "-" : "-", TextReplacementTests.Text(result.CreatePackageCopy()).Text);
    }

    [Fact]
    public async Task ManySlotsAndReversedDisjointRequestsKeepLinearInsertionCorrelation()
    {
        const int count = 128;
        var slots = Enumerable.Range(0, count).Select(i => ("A", 2d + i * 3, "")).ToArray();
        var source = await TextReplacementTests.Read(slots, 500, 600);
        var hits = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "A").Hits;
        var requests = Enumerable.Range(0, count).Where(i => i % 2 == 0).Reverse().Select(i => new OfdTextReplacement(hits[i], "--")).ToArray();
        var provider = new TextReplacementTests.Metrics(); var measures = 0; provider.Callback = value => { if (value == "measure") measures++; };
        using var result = new OfdTextReplacer(provider).Replace(source, requests);
        Assert.Equal(2, measures); Assert.Equal(1, provider.Disposals);
        Assert.Equal(3, Assert.Single(result.Statistics.Analyses).ReturnedProviderClusters);
        Assert.Equal(253, Assert.Single(result.Statistics.Analyses).ReusedGeometryClusters);
        log.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { count, requests = requests.Length, measures,
            analysis = result.Statistics.Analyses, result.Statistics.Fit }));
        Assert.InRange(Assert.Single(result.Statistics.Analyses).CorrelationVisits, count * 3, count * 10);
        var output = result.CreatePackageCopy();
        Assert.Equal(Enumerable.Range(0, count).Select(i => i % 2 == 0 ? "--" : "A"), TextReplacementTests.Text(output).Runs.Select(run => run.Text));
        Assert.Equal(slots.Select(slot => slot.Item2), TextReplacementTests.Text(output).Runs.Select(run => run.XMillimeters));
    }

    [Fact]
    public void ChunkedUtf8CountPreservesSupplementaryScalarAcrossBufferBoundary()
    {
        var value = "<x>" + new string('a', 8191 - 3) + "𠀋</x>";
        using var operation = new OfdKeywordOperation(); var budget = new OfdKeywordLocatorBudget(new(), operation);
        var element = budget.Parse(value);
        Assert.EndsWith("𠀋", element.Value); Assert.Equal(Encoding.UTF8.GetByteCount(value), budget.XmlBytes);
    }

    [Theory]
    [InlineData("count", "metadata")]
    [InlineData("indexer", "metadata")]
    [InlineData("count", "attachment")]
    [InlineData("indexer", "attachment")]
    public async Task PublicRequestCollectionCallbacksCannotChangeUnqueriedSerializedFields(string phase, string field)
    {
        var source = await TextReplacementTests.Read(); source.Attachments.Add(new() { Id = "40", Name = "unqueried", Data = [1, 2, 3] });
        var hit = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits[0];
        var requests = new CallbackRequests(new(hit, "-"), phase, () =>
        { if (field == "metadata") source.Options.Metadata.Title = "changed by request callback"; else source.Attachments[0].Data[0]++; });
        var provider = new TextReplacementTests.Metrics();
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, requests));
        Assert.Equal("AuthenticateHits", error.Data["Ofdrw.TextReplacement.Phase"]); Assert.Equal(0, provider.Opens);
    }
    private sealed class CallbackRequests(OfdTextReplacement request, string phase, Action callback) : IReadOnlyList<OfdTextReplacement>
    {
        public int Count { get { if (phase == "count") callback(); return 1; } }
        public OfdTextReplacement this[int index] { get { if (phase == "indexer") callback(); return request; } }
        public IEnumerator<OfdTextReplacement> GetEnumerator() => new[] { request }.AsEnumerable().GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Theory]
    [InlineData(16383, false)]
    [InlineData(16384, true)]
    public void SnapshotFixedBufferReservationPrecedesTheAllocationBoundary(long capacity, bool reached)
    {
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = capacity });
        using var operation = new OfdKeywordOperation(material: context.Material); var entered = new List<string>();
        var budget = new OfdKeywordLocatorBudget(new(), operation, entered.Add);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordSourceSnapshot(new(), budget));
        Assert.Equal(reached, entered.Contains("SnapshotFixedState"));
    }

    [Fact]
    public void LargePreservedEntrySetCannotGrowPastPrepaidKeys()
    {
        var source = new OfdDocumentPackage(); for (var i = 0; i < 1000; i++) source.PreservedEntries.Add("opaque/" + i, Array.Empty<byte>());
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1_000_000 });
        context.Material.ChargeCopy(1_000_000 - (16384 + 10 * 128), default);
        using var operation = new OfdKeywordOperation(material: context.Material); var entries = 0;
        var budget = new OfdKeywordLocatorBudget(new(), operation, phase => { if (phase == "SnapshotEntryKey") entries++; });
        Assert.Throws<InvalidDataException>(() => new OfdKeywordSourceSnapshot(source, budget));
        Assert.Equal(10, entries); Assert.Equal(1000, source.PreservedEntries.Count);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task LargeControlUtf8AndBomCopiesArePrepaidBeforeDecode(bool bom, bool exact)
    {
        var source = await TextReplacementTests.Read(); var text = Encoding.UTF8.GetString(source.PreservedEntries["OFD.xml"]);
        text = text.Replace("</ofd:OFD>", "<!--" + new string('x', 100_000) + "--></ofd:OFD>");
        var bytes = Encoding.UTF8.GetBytes((bom ? "\ufeff" : "") + text); source.PreservedEntries["OFD.xml"] = bytes;
        const long cap = 2_000_000;
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = cap });
        using var operation = new OfdKeywordOperation(material: context.Material); var decodeEntered = 0;
        var budget = new OfdKeywordLocatorBudget(new(), operation, phase => { if (phase == "ControlUtf8Decode") decodeEntered++; });
        using var snapshot = new OfdKeywordSourceSnapshot(source, budget);
        var allowance = 128L + bytes.LongLength + 256 + 4L * bytes.LongLength + 256 - (exact ? 0 : 1);
        context.Material.ChargeCopy(cap - context.Material.CloneCopyBytes - allowance, default);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordSourceAdmission(source, snapshot, budget).PrepareDocument());
        Assert.Equal(exact ? 1 : 0, decodeEntered); Assert.Same(bytes, source.PreservedEntries["OFD.xml"]);
    }

    [Fact]
    public void MatcherAndPlannerCannotResetOrPostchargeTheOuterWorkLedger()
    {
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1_000_000 });
        using var operation = new OfdKeywordOperation(material: context.Material);
        var matcher = new OfdOrdinalMatcher("A", operation);
        // The matcher already holds one prepaid block. Exhaust the outer
        // ledger, then force the same operation to request further work.
        context.Material.ChargeWork(1_000_000 - context.Material.ValidationWorkBytes, default);
        Assert.Throws<InvalidDataException>(() => operation.Work(65536));
        Assert.Equal(65536, operation.ReservedLogicalValidationWorkUnits);
        Assert.Equal(1_000_000, context.Material.ValidationWorkBytes);
        var plannerEntered = false;
        Assert.Throws<InvalidDataException>(() => OfdTextFitPlanner.Plan(TextFitPlannerTests.Input(), new(), chargeWork: count =>
        { operation.Work(65536); plannerEntered = true; }));
        Assert.False(plannerEntered);
    }

    [Fact]
    public void SnapshotSeekCannotChangePositionBeforeItsLogicalWorkReservation()
    {
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 1_000_000 });
        using var operation = new OfdKeywordOperation(material: context.Material);
        var font = operation.CaptureFont(new OfdFontResource { Id = "10", Data = [1, 2, 3] }, "10");
        using var stream = font.OpenRead();
        context.Material.ChargeWork(1_000_000 - context.Material.ValidationWorkBytes, default);
        Assert.Throws<InvalidDataException>(() => stream.Seek(1, SeekOrigin.Begin));
        Assert.Equal(0, stream.Position); Assert.Equal(0, operation.FontSeekOperations);
    }

    [Fact]
    public async Task ExplicitNullSelectionDoesNotBecomeTheAllPagesOverload()
    {
        var source = await TextReplacementTests.Read(); var provider = new TextReplacementTests.Metrics();
        Assert.Throws<ArgumentNullException>(() => new OfdTextReplacer(provider).Replace(source, (IReadOnlyList<int>)null!, "B", "-"));
        Assert.Equal(0, provider.Opens);
    }

    [Fact]
    public async Task OversizedReplacementRejectsLengthBeforeScanningInvalidUnicode()
    {
        var source = await TextReplacementTests.Read(); var provider = new TextReplacementTests.Metrics();
        var options = new OfdTextReplacementOptions(new(maxReplacementUtf16Units: 4));
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(provider).Replace(source, "B", "\ud840xxxxx", options));
        Assert.Equal("ReplacementTextLimit", error.Message); Assert.Equal(0, provider.Opens);
    }

    [Fact]
    public void OddOffsetSupplementarySequenceHasBoundedCooperativeCheckpoints()
    {
        var text = "A" + string.Concat(Enumerable.Repeat("𠀀", 20_000));
        var checkpoints = OfdTextReplacer.ValidateLiteral(text, default);
        Assert.InRange(checkpoints, text.Length / 257, text.Length / 256 + 2);
        using var cts = new CancellationTokenSource(); cts.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdTextReplacer.ValidateLiteral(text, cts.Token));
    }

    [Theory]
    [InlineData("open", false)]
    [InlineData("measure", false)]
    [InlineData("statistics", false)]
    [InlineData("dispose", false)]
    [InlineData("open", true)]
    [InlineData("measure", true)]
    [InlineData("statistics", true)]
    [InlineData("dispose", true)]
    public async Task ThrowingAndCancellingCallbacksLeaveNoResultAndCloseOwnedSession(string phase, bool cancel)
    {
        var source = await TextReplacementTests.Read(); var hit = new OfdKeywordLocator(new TextReplacementTests.Metrics()).Find(source, "B").Hits[0];
        using var cts = new CancellationTokenSource(); var original = TextReplacementTests.Text(source).SourceXml;
        var provider = new TextReplacementTests.Metrics { Callback = value =>
        { if (value != phase) return; if (cancel) cts.Cancel(); else throw new ApplicationException("Injected provider failure"); } };
        if (cancel) Assert.Throws<OperationCanceledException>(() => new OfdTextReplacer(provider).Replace(source, new[] { new OfdTextReplacement(hit, "-") }, cancellationToken: cts.Token));
        else Assert.Throws<ApplicationException>(() => new OfdTextReplacer(provider).Replace(source, new[] { new OfdTextReplacement(hit, "-") }));
        Assert.Equal(phase == "open" && !cancel ? 0 : 1, provider.Disposals);
        Assert.Equal(original, TextReplacementTests.Text(source).SourceXml);
    }
}
