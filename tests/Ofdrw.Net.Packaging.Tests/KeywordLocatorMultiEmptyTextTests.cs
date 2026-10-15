using System.Globalization;
using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed partial class KeywordLocatorEmptyTextTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper output;
    public KeywordLocatorEmptyTextTests(Xunit.Abstractions.ITestOutputHelper output) { this.output = output; }
    private static string Code(string value = "", string? x = "2", string? y = "8", string attrs = "") =>
        $"<o:TextCode{(x is null ? "" : $" X=\"{x}\"")}{(y is null ? "" : $" Y=\"{y}\"")} {attrs}>{value}</o:TextCode>";
    public static IEnumerable<object[]> AuthenticatedEmptySlots()
    {
        yield return [Empty + Code(x: "4"), "", 2]; // Former authenticated two-empty negative.
        yield return [Code("A") + Code(x: "4"), "A", 2]; // Former authenticated mixed-empty negative.
        yield return [Empty + Code("A"), "A", 2];
        yield return [Code("A") + Code(x: "4") + Code("B", "4"), "AB", 3];
        yield return [Empty + Code("A") + Code(x: "4") + Code(x: "4") + Code("B", "4") + Code(x: "6"), "AB", 6];
        yield return [Empty + Code(x: "4") + Code(x: "6"), "", 3];
    }
    [Theory]
    [MemberData(nameof(AuthenticatedEmptySlots))]
    public async Task AuthenticatedLeadingInteriorTrailingMultipleAndAllEmptySlotsRoundtrip(string codes, string keyword, int slots)
    {
        var source = await Read(Object(codes)); var beforeXml = Text(source).SourceXml;
        var measured = new List<string>(); var provider = keyword.Length == 0 ? NeverCalledProvider() : new KeywordLocatorTests.Provider
        { Callback = phase => { if (phase == "measure") measured.Add("measure"); } };
        var locator = new OfdKeywordLocator(provider); var before = locator.Find(source, keyword.Length == 0 ? "absent" : keyword);
        Assert.True(before.IsComplete); Assert.Equal(slots, Text(source).Runs.Count); Assert.Equal(beforeXml, Text(source).SourceXml);
        using var destination = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, destination);
        destination.Position = 0; var afterSource = await new OfdReader().ReadAsync(destination);
        var after = locator.Find(afterSource, keyword.Length == 0 ? "absent" : keyword);
        Assert.True(after.IsComplete); Assert.Empty(after.Diagnostics); Assert.Equal(slots, Text(afterSource).Runs.Count);
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { slots, keyword, phase = "authenticated-default-writer-fresh-reader", after.Statistics, measuredCalls = measured.Count }));
        Assert.Equal("30", Text(afterSource).ObjectId); Assert.Equal("10", Text(afterSource).FontResourceId);
        Assert.Equal(new byte[] { 1, 2, 3 }, afterSource.Fonts[0].Data);
        Assert.Contains("\n  <ofd:Area>", Encoding.UTF8.GetString(afterSource.PreservedEntries[PagePath]));
        if (keyword.Length == 0)
        { AssertEmpty(before); AssertEmpty(after); Assert.Equal(0, provider.Opens); Assert.Equal(0, provider.Disposals); }
        else
        {
            var a = Assert.Single(before.Hits); var b = Assert.Single(after.Hits); Assert.True(a.CanReplace); Assert.True(b.CanReplace);
            Assert.Equal(keyword.Length, b.Fragments.Count); Assert.Equal(keyword.Length, after.Statistics.ReservedScalarGlyphs);
            Assert.Equal(keyword.Length, after.Statistics.Provider.GlyphRequests);
            Assert.Equal(2 * keyword.Length, measured.Count); Assert.Equal(2, provider.Opens); Assert.Equal(2, provider.Disposals);
            Assert.Equal(a.Fragments.Select(f => (f.TextCodePath, f.BaselineMillimeters.X, f.BaselineMillimeters.Y)), b.Fragments.Select(f => (f.TextCodePath, f.BaselineMillimeters.X, f.BaselineMillimeters.Y)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MultiAndMixedEmptyCannotManufactureOriginalEnvelope(bool fakeSourcePath)
    {
        foreach (var codes in new[] { Empty + Code(x: "4"), Code("A") + Code(x: "4") })
        {
            var source = KeywordLocatorTests.Model(""); var text = Text(source); text.ObjectId = "30";
            text.SourceXml = Object(codes).Replace("<o:TextObject", $"<o:TextObject xmlns:o=\"{Ns}\"");
            if (fakeSourcePath) source.Pages[0].SourceEntryPath = PagePath;
            Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete);
            text.SourceXml = null; text.Text = "A"; text.Runs.Clear(); text.Runs.Add(new() { Text = "", XMillimeters = 2, YMillimeters = 8 });
            text.Runs.Add(new() { Text = "A", XMillimeters = 4, YMillimeters = 8 });
            Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete);
        }
    }

    [Theory]
    [InlineData(null, null, 0, 4, 21, 34)]
    [InlineData("8", null, 8, 4, 23, 34)]
    [InlineData(null, "12", 0, 12, 21, 35)]
    public async Task NonemptyInheritanceUsesPreviousSourceStartAndValidatesReaderDefaults(string? x, string? y,
        double typedX, double typedY, double pageX, double pageY)
    {
        var source = await Read(Object(Code(x: "6", y: "11") + Code("A", x, y)));
        Assert.Equal(typedX, Text(source).Runs[1].XMillimeters); Assert.Equal(typedY, Text(source).Runs[1].YMillimeters);
        var hit = Assert.Single(new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, "A").Hits);
        Assert.True(hit.CanReplace); Assert.EndsWith("TextCode[1]", hit.Fragments[0].TextCodePath);
        Assert.Equal(pageX, hit.Fragments[0].BaselineMillimeters.X); Assert.Equal(pageY, hit.Fragments[0].BaselineMillimeters.Y);
        if (x is null) Text(source).Runs[1].XMillimeters = 6; else Text(source).Runs[1].YMillimeters = 11;
        Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete);
    }

    [Fact]
    public async Task EmptyStartDoesNotForceOrDestroyGeometricContinuity()
    {
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        var source = await Read(Object(Code("A") + Code(x: "4") + Code("B", null, null)));
        Assert.Single(locator.Find(source, "AB").Hits); Assert.Equal(0, Text(source).Runs[2].XMillimeters); Assert.Equal(4, Text(source).Runs[2].YMillimeters);
        var gap = await Read(Object(Code("A") + Code(x: "8") + Code("B", null, null)));
        Assert.Empty(locator.Find(gap, "AB").Hits); Assert.Single(locator.Find(gap, "B").Hits);
        var differentBaseline = await Read(Object(Code("A") + Code(x: "4", y: "9") + Code("B", null, null)));
        Assert.Empty(locator.Find(differentBaseline, "AB").Hits);
        // Table46 inherits the previous start (2), never AB's last origin (4).
        var previousStart = await Read(Object(Code("AB", attrs: "DeltaX=\"2\"") + Code("C", null, null) + Code(x: "8")));
        Assert.Empty(locator.Find(previousStart, "ABC").Hits);
        Assert.Equal(17, Assert.Single(locator.Find(previousStart, "C").Hits).Fragments[0].BaselineMillimeters.X);
    }

    [Fact]
    public async Task ExplicitSurvivorCoordinatesKeepFrozenAnchorAfterPrecedingSlotChanges()
    {
        var original = await Read(Object(Code(x: "6", y: "11") + Code("A", null, null)));
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider());
        var frozen = Assert.Single(locator.Find(original, "A").Hits).Fragments[0].BaselineMillimeters;
        // Independent post-edit fixtures: no product Replace or source-certificate fabrication.
        var explicitSurvivor = await Read(Object(Code(x: "40", y: "50") + Code("A", "6", "11")));
        using var destination = new MemoryStream(); await new OfdPackageWriter().WriteAsync(explicitSurvivor, destination); destination.Position = 0;
        var reread = await new OfdReader().ReadAsync(destination); var after = Assert.Single(locator.Find(reread, "A").Hits).Fragments[0].BaselineMillimeters;
        Assert.Equal(frozen.X, after.X); Assert.Equal(frozen.Y, after.Y);
        var unfixed = await Read(Object(Code(x: "40", y: "50") + Code("A", null, null)));
        Assert.NotEqual(frozen.X, Assert.Single(locator.Find(unfixed, "A").Hits).Fragments[0].BaselineMillimeters.X);
    }

    [Fact]
    public async Task RawEscapedScalarSpansAndOriginalSlotOrdinalSurviveSkippedEmpties()
    {
        var source = await Read(Object(Empty + Code("\\0020\\005C\\D840\\DC0B", "6", "8", "DeltaX=\"1 2\"") + Code(x: "12")));
        var result = new OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(source, " \\𠀋");
        var hit = Assert.Single(result.Hits); Assert.True(result.IsComplete); Assert.True(hit.CanReplace);
        Assert.Equal(new[] { 0, 1, 2 }, hit.Fragments.Select(f => f.LogicalUtf16Start));
        Assert.Equal(new[] { 1, 1, 2 }, hit.Fragments.Select(f => f.LogicalUtf16Length));
        Assert.Equal(new[] { 0, 5, 10 }, hit.Fragments.Select(f => f.SourceUtf16Start));
        Assert.Equal(new[] { 5, 5, 10 }, hit.Fragments.Select(f => f.SourceUtf16Length));
        Assert.All(hit.Fragments, f => Assert.EndsWith("TextCode[1]", f.TextCodePath));
        Assert.Equal(new[] { 21d, 22, 24 }, hit.Fragments.Select(f => f.BaselineMillimeters.X));
        Assert.Equal(3, result.Statistics.ReservedScalarGlyphs); Assert.Equal(3, result.Statistics.Provider.GlyphRequests);
        Assert.Equal("\\0020\\005C\\D840\\DC0B", Text(source).Text);
    }

    [Theory]
    [InlineData("<o:TextCode Y=\"8\"/>")]
    [InlineData("<o:TextCode X=\"4\"/>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\" DeltaX=\"0\"/>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\" DeltaY=\"\"/>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\"><![CDATA[]]></o:TextCode>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\"><!--x--></o:TextCode>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\"> </o:TextCode>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\"><?x y?></o:TextCode>")]
    [InlineData("<o:TextCode X=\"4\" Y=\"8\"><v:X xmlns:v=\"urn:v\"/></o:TextCode>")]
    public async Task MixedObjectDoesNotHideNoncanonicalLaterEmpty(string invalid)
    {
        var source = await Read(Object(Code("A") + invalid)); var provider = NeverCalledProvider();
        Assert.False(new OfdKeywordLocator(provider).Find(source, "A", Incomplete()).IsComplete); Assert.Equal(0, provider.Opens);
    }

    [Theory]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\">A<!--x--></o:TextCode>")]
    [InlineData("<o:TextCode X=\"2\" Y=\"8\"><![CDATA[A]]></o:TextCode>")]
    public async Task EmptyBearingRootDoesNotAdmitNonemptyCommentOrCdata(string invalid)
    { var source = await Read(Object(invalid + Empty)); Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete); }

    [Theory]
    [InlineData("aggregate")]
    [InlineData("raw-vs-decoded")]
    [InlineData("order")]
    [InlineData("count")]
    [InlineData("first-anchor")]
    [InlineData("last-anchor")]
    [InlineData("empty-delta")]
    [InlineData("nonempty-delta")]
    [InlineData("boundary")]
    [InlineData("size")]
    [InlineData("matrix")]
    [InlineData("matrix-presence")]
    [InlineData("fill")]
    public async Task EmptyBearingAllSlotsAndRootTypedProjectionRemainAuthoritative(string kind)
    {
        var source = await Read(Object(Empty + Code("\\0020A", "4", "8", "DeltaX=\"1\"") + Code(x: "8")));
        var text = Text(source);
        switch (kind)
        {
            case "aggregate": text.Text += "A"; break;
            case "raw-vs-decoded": text.Runs[1].Text = " A"; break;
            case "order": text.Runs.Reverse(); break;
            case "count": text.Runs.RemoveAt(2); break;
            case "first-anchor": text.Runs[0].XMillimeters++; break;
            case "last-anchor": text.Runs[2].YMillimeters++; break;
            case "empty-delta": text.Runs[2].DeltaX = ""; break;
            case "nonempty-delta": text.Runs[1].DeltaX = "2"; break;
            case "boundary": text.WidthMillimeters++; break;
            case "size": text.FontSizeMillimeters++; break;
            case "matrix": text.Transform![4]++; break;
            case "matrix-presence": text.Transform = null; break;
            case "fill": text.FillColor = OfdColor.Black; break;
        }
        Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A", Incomplete()).IsComplete);
    }

    [Theory]
    [InlineData("1 0 0 1 0 0")]
    [InlineData("0 1 -1 0 20 0")]
    [InlineData("1 0 0.5 1 0 0")]
    public async Task AllEmptyCoordinatesKeepTransformChecksWithoutInkOrClip(string matrix)
    {
        var source = await Read(Object(Empty + Code(x: "1000000", y: "-1000000"), matrix: matrix));
        AssertEmpty(new OfdKeywordLocator(NeverCalledProvider()).Find(source, "A"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EveryEmptyIncludingMiddleOrLastChecksComposedOverflowBeforeProvider(bool mixed)
    {
        foreach (var codes in new[] { Empty + Code(x: "1e308") , Empty + Code(x: "1e308") + Empty })
        {
            var source = await Read(Object((mixed ? Code("A") : "") + codes, boundary: "1e308 30 100 50"));
            var provider = NeverCalledProvider();
            Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A")); Assert.Equal(0, provider.Opens);
        }
    }

    [Fact]
    public async Task ExactFinitePrecisionAndAbsentRootProjectionSurviveDefaultWriter()
    {
        const string x = "2.1234567891234567";
        var obj = Object(Code(x: "4") + Code("A", x)).Replace(" CTM=\"1 0 0 1 0 0\"", "").Replace("<o:FillColor Value=\"0 80 160\"/>", "");
        var source = await Read(obj); Assert.Null(Text(source).Transform); Assert.Equal(OfdColor.Black.Red, Text(source).FillColor.Red);
        var locator = new OfdKeywordLocator(new KeywordLocatorTests.Provider()); var before = Assert.Single(locator.Find(source, "A").Hits).Fragments[0];
        Assert.Equal(double.Parse(x, CultureInfo.InvariantCulture) + 15, before.BaselineMillimeters.X);
        using var destination = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, destination); destination.Position = 0;
        var after = await new OfdReader().ReadAsync(destination); Assert.Null(Text(after).Transform);
        Assert.Equal(before.BaselineMillimeters.X, Assert.Single(locator.Find(after, "A").Hits).Fragments[0].BaselineMillimeters.X);
        Text(after).Transform = [1, 0, 0, 1, 0, 0];
        Assert.False(new OfdKeywordLocator(NeverCalledProvider()).Find(after, "A", Incomplete()).IsComplete);
    }

    [Theory]
    [InlineData("open", 0)]
    [InlineData("measure", 2)]
    [InlineData("statistics", 4)]
    [InlineData("dispose", 0)]
    public async Task MixedCallbacksCannotMutateRetainedEmptySlots(string phase, int slot)
    {
        var source = await Read(Object(Empty + Code("A") + Code(x: "4") + Code("B", "4") + Code(x: "6")));
        var provider = new KeywordLocatorTests.Provider { Callback = step => { if (step == phase) Text(source).Runs[slot].XMillimeters++; } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "AB")); Assert.Equal(1, provider.Disposals);
    }

    [Theory]
    [InlineData("raw")]
    [InlineData("xml")]
    [InlineData("font")]
    [InlineData("run-order")]
    public async Task MixedDisposeCannotMutateOtherSourceEvidence(string kind)
    {
        var source = await Read(Object(Empty + Code("A") + Code(x: "4")));
        var provider = new KeywordLocatorTests.Provider { Callback = phase =>
        {
            if (phase != "dispose") return;
            switch (kind)
            {
                case "raw": Text(source).Text += "A"; break;
                case "xml": Text(source).SourceXml += " "; break;
                case "font": source.Fonts[0].Data[0]++; break;
                case "run-order": Text(source).Runs.Reverse(); break;
            }
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A")); Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task ManyZeroGlyphSlotsConsumeLinearNodeFieldReferenceAndXmlBudgets()
    {
        async Task<OfdDocumentPackage> Many(int count) => await Read(Object(string.Concat(Enumerable.Repeat(Empty, count))));
        var small = await Many(16); var large = await Many(32); var provider = NeverCalledProvider();
        var options = new OfdKeywordSearchOptions(new(maxTextUtf16Units: 1, maxClusterVisits: 1), new(maxGlyphOperations: 1));
        var one = new OfdKeywordLocator(provider).Find(small, "A", options); var two = new OfdKeywordLocator(provider).Find(large, "A", options);
        var third = new OfdKeywordLocator(provider).Find(await Many(48), "A", options);
        AssertEmpty(one); AssertEmpty(two); AssertEmpty(third);
        Assert.Equal(two.Statistics.ValidationFields - one.Statistics.ValidationFields, third.Statistics.ValidationFields - two.Statistics.ValidationFields);
        Assert.Equal(two.Statistics.AdmissionNodes - one.Statistics.AdmissionNodes, third.Statistics.AdmissionNodes - two.Statistics.AdmissionNodes);
        Assert.True(two.Statistics.AdmissionNodes > one.Statistics.AdmissionNodes);
        Assert.True(two.Statistics.ValidationFields > one.Statistics.ValidationFields);
        Assert.InRange(two.Statistics.ValidationFields - one.Statistics.ValidationFields, 16, 200L * 16);
        AssertEmpty(new OfdKeywordLocator(provider, new(maxNodes: (int)two.Statistics.AdmissionNodes, maxValidationFields: two.Statistics.ValidationFields,
            maxValidationCharacters: two.Statistics.ValidationCharacters, maxXmlBytes: two.Statistics.AdmissionXmlBytes)).Find(large, "A", options));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxNodes: (int)two.Statistics.AdmissionNodes - 1)).Find(large, "A", options));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxValidationFields: two.Statistics.ValidationFields - 1)).Find(large, "A", options));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxValidationCharacters: two.Statistics.ValidationCharacters - 1)).Find(large, "A", options));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxXmlBytes: two.Statistics.AdmissionXmlBytes - 1)).Find(large, "A", options));
        var references = Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider, new(maxReferences: (int)two.Statistics.TraversalReferences + 32)).Find(large, "A", options));
        Assert.Contains("admission references", references.Message); Assert.Equal(0, provider.Opens);
        output.WriteLine(System.Text.Json.JsonSerializer.Serialize(new { phase = "zero-slot-capacity", slots16 = one.Statistics, slots32 = two.Statistics, slots48 = third.Statistics,
            configuredRejectedReferenceCap = two.Statistics.TraversalReferences + 32, referenceRejection = references.Message, providerOpens = provider.Opens }));
        using var cts = new CancellationTokenSource(); cts.Cancel();
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(provider).Find(large, "A", options, cts.Token));
    }

    [Fact]
    public async Task MixedCallbackCancellationStillDisposesExactlyOnce()
    {
        var source = await Read(Object(string.Concat(Enumerable.Repeat(Empty, 32)) + Code("A")));
        using var cts = new CancellationTokenSource();
        var provider = new KeywordLocatorTests.Provider { Callback = phase => { if (phase == "measure") cts.Cancel(); } };
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(provider).Find(source, "A", cancellationToken: cts.Token)); Assert.Equal(1, provider.Disposals);
    }
}
