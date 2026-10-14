using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class KeywordLocatorTests
{
    [Fact]
    public void OverlapsAreOrdinalAndNoDeltaKeepsCoincidentOrigins()
    {
        var source = Model("aaa"); var provider = new Provider();
        var result = new OfdKeywordLocator(provider).Find(source, "aa");
        Assert.True(result.IsComplete); Assert.Equal("SelectedDocumentBody", result.Scope);
        Assert.Equal("ContinuousBaseline", result.MatchPolicy); Assert.Equal(2, result.Hits.Count);
        Assert.Equal(new[] { 0, 1 }, result.Hits.Select(h => h.Fragments[0].SourceUtf16Start));
        Assert.All(result.Hits.SelectMany(h => h.Fragments), f => { Assert.Equal(17, f.BaselineMillimeters.X); Assert.Equal(31, f.BaselineMillimeters.Y); });
        Assert.Equal(1, result.Statistics.FontHashOperations); Assert.Equal(1, provider.Opens); Assert.Equal(1, provider.Disposals);
        Assert.True(Assert.Single(new OfdKeywordLocator(new Provider()).Find(source, "aaa").Hits).CanReplace);
        Assert.Empty(new OfdKeywordLocator(new Provider()).Find(source, "A").Hits);
        Assert.Equal("aaa", Text(source).Text); Assert.Null(Text(source).SourceXml);
    }

    [Fact]
    public void SpaceHasAdvanceWithoutAnInventedBoxAndSupplementaryCountsAsOneGlyph()
    {
        var source = Model("A 𠀋"); Text(source).Runs[0].DeltaX = "2 1";
        var result = new OfdKeywordLocator(new Provider()).Find(source, " 𠀋");
        var hit = Assert.Single(result.Hits); Assert.Equal(2, hit.Fragments.Count);
        Assert.Null(hit.Fragments[0].InkQuadrilateralMillimeters); Assert.Equal(1, hit.Fragments[0].AdvanceXMillimeters);
        Assert.Equal(2, hit.Fragments[1].SourceUtf16Length); Assert.Equal(20, hit.Fragments[1].BaselineMillimeters.X);
        Assert.Equal(4, result.Statistics.ExpandedDeltaValues);
        var spaces = new OfdKeywordLocator(new Provider()).Find(Model(" "), " ");
        Assert.Null(Assert.Single(spaces.Hits).BoundsMillimeters);
    }

    [Fact]
    public void MultipleTextCodesAndAdjacentObjectsJoinOnlyWithProvenContinuity()
    {
        var source = Model("A"); var first = Text(source); first.Runs.Add(new() { Text = "B", XMillimeters = 4, YMillimeters = 8 });
        first.Text = "AB";
        var second = Text(Model("C")); second.Runs[0].XMillimeters = 6; source.Pages[0].Elements.Add(second);
        var locator = new OfdKeywordLocator(new Provider()); var found = locator.Find(source, "ABC");
        Assert.True(found.IsComplete); var hit = Assert.Single(found.Hits);
        Assert.Equal("ABC", hit.MatchedText); Assert.Equal(3, hit.Fragments.Count); Assert.False(hit.CanReplace);
        Assert.Equal(new[] { 17d, 19d, 21d }, hit.Fragments.Select(fragment => fragment.BaselineMillimeters.X));
        Assert.NotEqual(hit.Fragments[1].OccurrencePath[3], hit.Fragments[2].OccurrencePath[3]);
        Assert.True(Assert.Single(locator.Find(source, "AB").Hits).CanReplace); // multiple TextCodes, one occurrence
        Assert.True(Assert.Single(locator.Find(source, "C").Hits).CanReplace); // match starts in the later owner of this segment
        second.Runs[0].XMillimeters = 10;
        Assert.Empty(new OfdKeywordLocator(new Provider()).Find(source, "ABC").Hits);
        second.Runs[0].XMillimeters = 6; second.Runs[0].YMillimeters = 9;
        Assert.Empty(new OfdKeywordLocator(new Provider()).Find(source, "ABC").Hits);
    }

    [Theory]
    [InlineData("raw")]
    [InlineData("path")]
    [InlineData("layer")]
    public void StructuralBarriersPreventTextConcatenation(string barrier)
    {
        var source = Model("A"); var second = Text(Model("B")); second.Runs[0].XMillimeters = 4;
        if (barrier == "raw") source.Pages[0].Elements.Add(new OfdRawElement { Xml = "<PageBlock><TextObject /></PageBlock>" });
        if (barrier == "path") source.Pages[0].Elements.Add(new OfdPathElement());
        if (barrier == "layer") second.LayerId = "different";
        source.Pages[0].Elements.Add(second);
        var result = new OfdKeywordLocator(new Provider()).Find(source, "AB", new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.Empty(result.Hits); Assert.Equal(barrier != "raw", result.IsComplete);
        if (barrier == "raw") Assert.Throws<NotSupportedException>(() => new OfdKeywordLocator(new Provider()).Find(source, "AB"));
    }

    [Theory]
    [InlineData("1", 3)]
    [InlineData("1 2 3", 3)]
    [InlineData("1 2 garbage", 3)]
    [InlineData("g 2 1 garbage", 3)]
    [InlineData("g 0 1", 3)]
    [InlineData("g -1 1", 3)]
    [InlineData("g 2147483648 1", 3)]
    [InlineData("g 2147483647 1", 3)]
    [InlineData("g 2", 3)]
    [InlineData("NaN 2", 3)]
    [InlineData("1e309 2", 3)]
    [InlineData(" ", 3)]
    [InlineData("1", 1)]
    public void EntireDeltaArrayMustBeValid(string delta, int length)
    {
        var source = Model(new string('A', length)); Text(source).Runs[0].DeltaX = delta;
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider()).Find(source, "A",
            new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete)));
    }

    [Fact]
    public void RepeatAndNegativeDeltasUseScalarCountAndMissingAxisStaysZero()
    {
        var source = Model("𠀋AA"); Text(source).Runs[0].DeltaX = "g 2 -1";
        var hit = Assert.Single(new OfdKeywordLocator(new Provider()).Find(source, "𠀋AA").Hits);
        Assert.Equal(new[] { 17d, 16d, 15d }, hit.Fragments.Select(f => f.BaselineMillimeters.X));
        Assert.All(hit.Fragments, f => Assert.Equal(31, f.BaselineMillimeters.Y));
        Text(source).Runs[0].DeltaY = "1 -1";
        Assert.Empty(new OfdKeywordLocator(new Provider()).Find(source, "𠀋AA").Hits); // No cross-line matching.
    }

    [Theory]
    [InlineData("gap")]
    [InlineData("overlap")]
    [InlineData("surrogate")]
    [InlineData("cluster")]
    [InlineData("empty")]
    [InlineData("null")]
    public void CustomProviderCoverageCannotBypassValidation(string kind)
    {
        var text = kind == "surrogate" ? "𠀋" : "AA";
        var provider = new Provider { Override = request => kind switch
        {
            "gap" => new([new(1, 1, 1, null, 2)]),
            "overlap" => new([new(0, 1, 1, null, 2), new(0, 1, 1, null, 2)]),
            "surrogate" => new([new(0, 1, 1, null, 2), new(1, 1, 1, null, 2)]),
            "cluster" => new([new(0, 2, 1, null, 2)]),
            "empty" => new([]), _ => null!
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(Model(text), text));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public void UnsupportedProviderIsExplicitlyIncompleteAndCannotReturnANoHitClaim()
    {
        var provider = new Provider { Override = _ => new([], "Unsupported mapping") };
        var locator = new OfdKeywordLocator(provider);
        Assert.Throws<NotSupportedException>(() => locator.Find(Model("A"), "z"));
        var result = locator.Find(Model("A"), "z", new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.False(result.IsComplete); Assert.Empty(result.Hits); Assert.Equal("UnsupportedFontGeometry", Assert.Single(result.Diagnostics).Code);
    }

    [Theory]
    [InlineData("open")]
    [InlineData("measure")]
    [InlineData("statistics")]
    [InlineData("dispose")]
    public void EveryProviderCallbackIsFollowedBySourceValidation(string phase)
    {
        var source = Model("A"); var provider = new Provider();
        provider.Callback = name => { if (name == phase) Text(source).Runs[0].XMillimeters++; };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A"));
        Assert.Equal(1, provider.Disposals);
    }

    [Theory]
    [InlineData("later-text")]
    [InlineData("matrix")]
    [InlineData("font-bytes")]
    [InlineData("font-owner")]
    [InlineData("page-order")]
    public void CapturePrecedesCallbacksAndRejectsMutatingLaterInputs(string kind)
    {
        var source = Model("A"); var second = Text(Model("B")); source.Pages[0].Elements.Add(second);
        var first = Text(source); first.Transform = [1, 0, 0, 1, 0, 0];
        var provider = new Provider { Callback = phase =>
        {
            if (phase != "open") return;
            switch (kind)
            {
                case "later-text": second.Runs[0].Text = "C"; break;
                case "matrix": first.Transform[4]++; break;
                case "font-bytes": source.Fonts[0].Data[0]++; break;
                case "font-owner": source.Fonts[0] = new() { Id = "10", Data = [1] }; break;
                case "page-order": source.Pages.Add(new()); break;
            }
        } };
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A"));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public void ImmutableHitBindsExactPackageAndRejectsStaleSameIdSource()
    {
        var source = Model("A"); var hit = Assert.Single(new OfdKeywordLocator(new Provider()).Find(source, "A").Hits);
        hit.Source.ValidateUnchanged(source);
        Assert.Throws<InvalidDataException>(() => hit.Source.ValidateUnchanged(Model("A")));
        source.Fonts[0].Data[0]++;
        Assert.Throws<InvalidDataException>(() => hit.Source.ValidateUnchanged(source));
        Assert.IsAssignableFrom<IList<OfdKeywordFragment>>(hit.Fragments);
        Assert.Throws<NotSupportedException>(() => ((IList<OfdKeywordFragment>)hit.Fragments).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<int>)hit.Fragments[0].OccurrencePath)[0] = 99);
    }

    [Fact]
    public void LimitsAreCumulativeInclusiveAndFailuresDoNotReturnTruncatedSuccess()
    {
        var source = Model("aaa");
        var ok = new OfdKeywordLocator(new Provider(), new(maxDeltaValues: 4, maxFragments: 4)).Find(source, "aa",
            new(limits: new(maxTextUtf16Units: 3, maxClusterVisits: 3, maxHits: 2)));
        Assert.Equal(2, ok.Hits.Count); Assert.Equal(4, ok.Statistics.ReturnedFragments);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider(), new(maxDeltaValues: 3)).Find(source, "aa"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider(), new(maxFragments: 3)).Find(source, "aa"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider()).Find(source, "aa", new(limits: new(maxHits: 1))));
        var provider = new Provider();
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "aa", new(limits: new(maxTextUtf16Units: 2))));
        Assert.Equal(0, provider.Opens);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider(), new(maxValidationFields: 1)).Find(source, "a"));
    }

    [Fact]
    public void WorkLimitsAlsoCoverReturnedClustersTokensAndOrdinalComparisons()
    {
        var source = Model("AAA"); Text(source).Runs[0].DeltaX = "g 2 1";
        var ok = new OfdKeywordLocator(new Provider(), new(maxDeltaTokens: 3)).Find(source, "A",
            new(limits: new(maxMatchComparisons: 3), geometryLimits: new(maxGlyphOperations: 3)));
        Assert.Equal(3, ok.Statistics.MatchComparisons); Assert.Equal(3, ok.Statistics.ReturnedProviderClusters);
        Assert.Equal(3, ok.Statistics.ReservedScalarGlyphs); Assert.Equal(3, ok.Statistics.DeltaTokens);
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider(), new(maxDeltaTokens: 2)).Find(source, "A"));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider()).Find(source, "A", new(limits: new(maxMatchComparisons: 2))));
        var provider = new Provider();
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(provider).Find(source, "A", new(geometryLimits: new(maxGlyphOperations: 2))));
        Assert.Equal(0, provider.Opens);
    }

    [Fact]
    public void CumulativeOverflowIsFatalEvenWhenAnEarlierGlyphWouldBeClipped()
    {
        var source = Model("AAA"); Text(source).Runs[0].DeltaX = "1e308 1e308";
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider()).Find(source, "A",
            new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete)));
    }

    [Theory]
    [InlineData("boundary")]
    [InlineData("page")]
    [InlineData("singular")]
    public void UnprovedClippingAndDegenerateTransformsHaveExplicitUnsupportedResults(string kind)
    {
        var source = Model("A"); var text = Text(source);
        if (kind == "boundary") text.WidthMillimeters = 2.5;
        if (kind == "page") source.Pages[0].WidthMillimeters = 17.5;
        if (kind == "singular") text.Transform = [1, 1, 1, 1, 0, 0];
        var result = new OfdKeywordLocator(new Provider()).Find(source, "A", new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete));
        Assert.False(result.IsComplete); Assert.Empty(result.Hits);
        Assert.Contains(Assert.Single(result.Diagnostics).Code, new[] { "BoundaryClipping", "PhysicalBoxClipping", "DegenerateTextTransform" });
    }

    [Fact]
    public void NameOnlyAndAmbiguousFontsNeverFallBackAndSharedDefinitionsAreNotEditable()
    {
        var source = Model("A"); var provider = new Provider(); source.Fonts[0].Data = [];
        Assert.False(new OfdKeywordLocator(provider).Find(source, "A", new(unsupportedContent: OfdUnsupportedContentBehavior.ReportIncomplete)).IsComplete);
        Assert.Equal(0, provider.Opens);
        source.Fonts.Add(new() { Id = "10", Data = [1] });
        Assert.Throws<NotSupportedException>(() => new OfdKeywordLocator(provider).Find(source, "A"));
        var shared = Model("A"); var page = new OfdPage { XMillimeters = 5, YMillimeters = 7, WidthMillimeters = 200, HeightMillimeters = 200 };
        page.Elements.Add(Text(shared)); shared.Pages.Add(page);
        var result = new OfdKeywordLocator(new Provider()).Find(shared, "A");
        Assert.Equal(2, result.Hits.Count); Assert.All(result.Hits, hit => Assert.False(hit.CanReplace));
        Assert.NotEqual(result.Hits[0].Fragments[0].OccurrencePath[0], result.Hits[1].Fragments[0].OccurrencePath[0]);
    }

    [Fact]
    public void DistinctTextDefinitionsSharingAMutableRunCannotAdvertiseEditableOwnership()
    {
        var source = Model("A"); var first = Text(source); var second = Text(Model("A"));
        second.Runs.Clear(); second.Runs.Add(first.Runs[0]); source.Pages[0].Elements.Add(second);
        var result = new OfdKeywordLocator(new Provider()).Find(source, "A");
        Assert.True(result.IsComplete); Assert.Equal(2, result.Hits.Count);
        Assert.All(result.Hits, hit => Assert.False(hit.CanReplace));
        Assert.NotEqual(result.Hits[0].Fragments[0].OccurrencePath[3], result.Hits[1].Fragments[0].OccurrencePath[3]);
    }

    [Fact]
    public void SnapshotStreamsAreReleasedOnProviderFailureAndCancelledSupplementaryScan()
    {
        Stream? leaked = null;
        var provider = new Provider { Override = request => { leaked = request.Font.OpenRead(); throw new InvalidOperationException("backend failure"); } };
        Assert.Throws<InvalidOperationException>(() => new OfdKeywordLocator(provider).Find(Model("A"), "A"));
        Assert.NotNull(leaked); Assert.False(leaked.CanRead); Assert.Equal(1, provider.Disposals);
        using var cts = new CancellationTokenSource();
        var longText = "A" + string.Concat(Enumerable.Repeat("𠀋", 2048));
        var cancellation = new Provider { Callback = phase => { if (phase == "open") cts.Cancel(); } };
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(cancellation).Find(Model(longText), "𠀋", cancellationToken: cts.Token));
        Assert.Equal(1, cancellation.Disposals);
    }

    [Fact]
    public void ExcludedTemplateAndAnnotationScopesAreDeclaredWithoutClaimingFullPageSearch()
    {
        var source = Model("A"); var template = new OfdTemplateContent(); template.Elements.Add(Text(Model("template")));
        source.Pages[0].Templates.Add(template); source.Pages[0].AnnotationAppearances.Add(Text(Model("annotation")));
        var result = new OfdKeywordLocator(new Provider()).Find(source, "A");
        Assert.True(result.IsComplete); Assert.Single(result.Hits);
        Assert.Equal(new[] { "ExcludedTemplateScope", "ExcludedAnnotationScope" }, result.Diagnostics.Select(d => d.Code));
        Assert.All(result.Diagnostics, diagnostic => Assert.False(diagnostic.MakesScopeIncomplete));
    }

    [Fact]
    public void SelectionAndCancellationNeverSilentlyFallBackToAllPages()
    {
        var source = Model("A"); source.Pages.Add(Model("B").Pages[0]);
        var locator = new OfdKeywordLocator(new Provider());
        var selected = locator.Find(source, new[] { 1 }, "B"); Assert.Equal(new[] { 1 }, selected.PageIndexes);
        Assert.Equal(1, Assert.Single(selected.Hits).PageIndex);
        Assert.Empty(locator.Find(source, Array.Empty<int>(), "A").Hits);
        Assert.Throws<ArgumentException>(() => locator.Find(source, new[] { -1 }, "A"));
        Assert.Throws<ArgumentException>(() => locator.Find(source, new[] { 0, 0 }, "A"));
        Assert.Throws<ArgumentException>(() => locator.Find(source, ""));
        Assert.Throws<InvalidDataException>(() => locator.Find(source, "\ud800"));
        using var cts = new CancellationTokenSource();
        var provider = new Provider { Callback = phase => { if (phase == "measure") cts.Cancel(); } };
        Assert.Throws<OperationCanceledException>(() => new OfdKeywordLocator(provider).Find(source, "A", cancellationToken: cts.Token));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public void ACallerSelectionIndexerCannotChangePageOwnershipDuringTheCopy()
    {
        var source = Model("A"); source.Pages.Add(Model("B").Pages[0]); source.Pages.Add(Model("C").Pages[0]);
        var selection = new MutatingSelection(() => source.Pages.Reverse(0, 2));
        Assert.Throws<InvalidDataException>(() => new OfdKeywordLocator(new Provider()).Find(source, selection, "B"));
    }
    private sealed class MutatingSelection(Action mutate) : IReadOnlyList<int>
    {
        public int Count => 2;
        public int this[int index] { get { if (index == 1) mutate(); return index == 0 ? 0 : 2; } }
        public IEnumerator<int> GetEnumerator() => throw new NotSupportedException();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    internal static OfdDocumentPackage Model(string text)
    {
        var package = new OfdDocumentPackage(); package.Fonts.Add(new() { Id = "10", FontName = "fake", Data = [1, 2, 3] });
        var page = new OfdPage { XMillimeters = 5, YMillimeters = 7, WidthMillimeters = 200, HeightMillimeters = 200 };
        var element = new OfdTextElement { Text = text, FontResourceId = "10", XMillimeters = 20, YMillimeters = 30,
            WidthMillimeters = 100, HeightMillimeters = 50, LayerId = "1" };
        element.Runs.Add(new() { Text = text, XMillimeters = 2, YMillimeters = 8 }); page.Elements.Add(element); package.Pages.Add(page);
        return package;
    }
    internal static OfdTextElement Text(OfdDocumentPackage source) => (OfdTextElement)source.Pages[0].Elements[0];
    internal sealed class Provider : IOfdGlyphGeometryProvider
    {
        internal int Opens, Disposals; internal Action<string>? Callback;
        internal Func<OfdGlyphGeometryRequest, OfdGlyphGeometry>? Override;
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken cancellationToken = default)
        { Opens++; Callback?.Invoke("open"); return new Session(this); }
        private sealed class Session(Provider owner) : IOfdGlyphGeometrySession
        {
            private long glyphs;
            public OfdGlyphGeometryStatistics Statistics { get { owner.Callback?.Invoke("statistics"); return new(glyphRequests: glyphs); } }
            public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken cancellationToken = default)
            {
                owner.Callback?.Invoke("measure"); if (owner.Override is not null) return owner.Override(request);
                var result = new List<OfdGlyphClusterGeometry>();
                for (var i = 0; i < request.Text.Length;)
                {
                    var length = char.IsHighSurrogate(request.Text[i]) ? 2 : 1; var space = request.Text[i] == ' ';
                    result.Add(new(i, length, 1, space ? null : new(0, -2, 1, 0), space ? 1 : 2)); i += length; glyphs++;
                }
                return new(result);
            }
            public void Dispose() { owner.Disposals++; owner.Callback?.Invoke("dispose"); }
        }
    }
}
