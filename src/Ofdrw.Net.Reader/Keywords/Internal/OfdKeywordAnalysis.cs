using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Reader.Keywords.Internal;

/// <summary>Finds ordinal keywords in supported selected-document body text using exact original-font geometry.</summary>
/// <remarks>Templates/annotations are excluded named scopes. Matching never infers reading order or whitespace.
/// Source-XML deltas follow GB/T33190-2016 table46: absent axes do not move; later X/Y inherit prior TextCode starts.
/// Source-backed empty slots preserve their ordinals and have no glyphs. An all-empty owner is a continuity barrier; font bytes remain verified.
/// All inputs must remain stable. Observable callback changes to captured source/geometry/ownership inputs are rejected; concurrent transient changes
/// restored before observation are not detectable. This service does not implement replacement or write pages.</remarks>
internal sealed class OfdKeywordAnalysis
{
    private enum Finalization { Independent, ImmediateOwnedFileFactory }
    private readonly IOfdGlyphGeometryProvider provider;
    private readonly OfdKeywordLocatorLimits limits;
    private readonly OfdModelMaterialization? material;
    private readonly Action? validateWholeSource;
    private readonly Action<OfdKeywordAnalysisCapture>? capture;
    private readonly Action<string>? beforeCallback;
    private readonly PrivateOwnedAnalysis? privateOwned;

    /// <summary>Creates a locator with an explicit provider and finite cumulative source-admission limits.</summary>
    internal OfdKeywordAnalysis(IOfdGlyphGeometryProvider geometryProvider, OfdKeywordLocatorLimits? locatorLimits = null,
        OfdModelMaterialization? material = null, Action? validateWholeSource = null,
        Action<OfdKeywordAnalysisCapture>? capture = null, Action<string>? beforeCallback = null)
    { provider = geometryProvider ?? throw new ArgumentNullException(nameof(geometryProvider)); limits = locatorLimits ?? new();
        this.material = material; this.validateWholeSource = validateWholeSource; this.capture = capture; this.beforeCallback = beforeCallback; }

    private OfdKeywordAnalysis(PrivateOwnedAnalysis binding, OfdKeywordLocatorLimits limits,
        Action<OfdKeywordAnalysisCapture> capture, Action<string> beforeCallback)
    {
        privateOwned = binding; provider = binding.Adapter; this.limits = limits;
        material = binding.Context.Material; this.capture = capture; this.beforeCallback = beforeCallback;
    }

    // Sole production caller is Layout's fresh ticket kernel. A generic provider,
    // old model session, runtime provider type or bool cannot enter this route.
    internal static OfdKeywordSearchResult FindPrivateOwned(OfdOwnedAnalysisClaim claim,
        OfdOwnedGlyphGeometryHost exactHost, OfdDocumentPackage source, OfdControlledEditContext context,
        int[] pages, string keyword, OfdKeywordLocatorLimits limits, OfdKeywordSearchOptions options,
        Action<OfdKeywordAnalysisCapture> capture, Action<string> beforeCallback,
        out OfdOwnedHostClosure actualClosure, CancellationToken token)
    {
        // Failure belongs to the claim's real owner. No foreign context/host
        // is touched or cleaned up before this identity/phase/CAS admission.
        OfdOwnedSaveAuthority.RequireActualAnalysis(claim, exactHost, source, context, token);
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            context.Material.ChargeCopy(4096, token); // binding, adapter and analysis shells
            var binding = new PrivateOwnedAnalysis(claim, exactHost, context);
            var analysis = new OfdKeywordAnalysis(binding, limits, capture, beforeCallback);
            var result = analysis.FindCore(source, pages, keyword, options, token, Finalization.ImmediateOwnedFileFactory);
            actualClosure = binding.Closure ?? throw new InvalidOperationException("MissingActualHostClosure");
            return result;
        }
        catch { context.Fail(); throw; }
    }

    /// <summary>Searches all current page positions, restricted to the result's declared body/visibility/matching scope.</summary>
    internal OfdKeywordSearchResult Find(OfdDocumentPackage source, string keyword, OfdKeywordSearchOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (source.Pages.Count > limits.MaxPages) throw new InvalidDataException("Selected page capacity exceeded.");
        return Find(source, Enumerable.Range(0, source.Pages.Count).ToArray(), keyword, options, cancellationToken);
    }

    /// <summary>Searches a copied distinct page selection in requested order. Invalid/duplicate positions reject.</summary>
    internal OfdKeywordSearchResult Find(OfdDocumentPackage source, IReadOnlyList<int> pageIndexes, string keyword,
        OfdKeywordSearchOptions? options = null, CancellationToken cancellationToken = default)
        => FindCore(source, pageIndexes, keyword, options, cancellationToken, Finalization.Independent);

    // Only the private file job uses this entry and immediately enters the
    // shared synchronous clone/patch/Writer factory. It is not a freshness
    // credential and cannot be carried across a callback, await or return.
    internal OfdKeywordSearchResult FindForOwnedFile(OfdDocumentPackage source, int[] pageIndexes, string keyword,
        OfdKeywordSearchOptions options, CancellationToken cancellationToken)
    {
        if (material is null || validateWholeSource is null || capture is null)
            throw new InvalidOperationException("Owned file analysis requires the enclosing private operation.");
        return FindCore(source, pageIndexes, keyword, options, cancellationToken, Finalization.ImmediateOwnedFileFactory);
    }

    private OfdKeywordSearchResult FindCore(OfdDocumentPackage source, IReadOnlyList<int> pageIndexes, string keyword,
        OfdKeywordSearchOptions? options, CancellationToken cancellationToken, Finalization finalization)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (pageIndexes is null) throw new ArgumentNullException(nameof(pageIndexes));
        cancellationToken.ThrowIfCancellationRequested();
        var selectionCount = pageIndexes.Count;
        if (selectionCount < 0 || selectionCount > limits.MaxPages) throw new InvalidDataException("Selected page capacity exceeded.");
        material?.ChargeCopy(512L + 64L * selectionCount, cancellationToken);
        var selection = new int[selectionCount]; var seen = new HashSet<int>(); var pages = new OfdPage[selectionCount];
        for (var i = 0; i < selectionCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested(); var index = pageIndexes[i];
            if (index < 0 || index >= source.Pages.Count || !seen.Add(index)) throw new ArgumentException("Page positions must be distinct valid members.", nameof(pageIndexes));
            selection[i] = index; pages[i] = source.Pages[index];
        }
        for (var i = 0; i < selectionCount; i++)
            if (selection[i] >= source.Pages.Count || !ReferenceEquals(pages[i], source.Pages[selection[i]]))
                throw new InvalidDataException("Source page ownership changed while copying the selection.");
        using var operation = new OfdKeywordOperation(options, cancellationToken, material);
        try
        {
        var budget = new OfdKeywordLocatorBudget(limits, operation);
        if (keyword is null) throw new ArgumentNullException(nameof(keyword));
        budget.Copy(512L + 32L * keyword.Length);
        var matcher = new OfdOrdinalMatcher(keyword, operation);
        // The existing A0 pass has its own finite structure budget. Reserve its
        // bounded temporary positions in the caller ledger before that pass.
        if (material is not null)
            budget.Copy(checked(256L * limits.MaxNodes + 32L * limits.MaxReferences + 256L * limits.MaxOccurrences +
                12L * material.Structure.SourceXmlBytes));
        // Exactly one shared graphics traversal. It provides ownership, not glyph exactness.
        var traversal = OfdModelTraversal.ReadGraphics(source, pages, new OfdTraversalOptions
        {
            IncludeTemplates = false, IncludeAnnotations = false, VisibleOnly = false, RequireSupportedSource = false,
            Limits = new OfdModelLimits { MaxNodeCount = limits.MaxNodes, MaxReferenceCount = limits.MaxReferences,
                MaxExpandedOccurrenceCount = limits.MaxOccurrences, MaxDepth = limits.MaxDepth, MaxSourceXmlBytes = limits.MaxXmlBytes }
        }, cancellationToken);
        // Fixed analysis/admission containers; snapshot and variable-size
        // arrays/strings are independently reserved at their allocation sites.
        budget.Copy(16_384);
        var diagnostics = new List<OfdKeywordDiagnostic>(); var hits = new List<OfdKeywordHit>();
        var reserved = new Dictionary<OfdGraphicOccurrence, int>();
        budget.Copy(256L + 128L * traversal.Occurrences.Count);
        budget.Work(32L * traversal.Occurrences.Count);
        foreach (var occurrence in traversal.Occurrences)
        {
            budget.Check();
            if (occurrence.Element is not OfdTextElement text) continue;
            long runLength = 0;
            foreach (var run in text.Runs) { budget.Check(); runLength += run.Text?.Length ?? 0; }
            var count = Math.Max(runLength, text.Text?.Length ?? 0);
            if (count > int.MaxValue) throw new InvalidDataException("Source text capacity exceeded.");
            budget.Text((int)count); reserved.Add(occurrence, (int)count);
        }
        using var snapshot = new OfdKeywordSourceSnapshot(source, budget);
        var admission = new OfdKeywordSourceAdmission(source, snapshot, budget);
        var rejected = new HashSet<string>(StringComparer.Ordinal);
        foreach (var diagnostic in traversal.Diagnostics)
            if (diagnostic.Kind is OfdTraversalDiagnosticKind.RawObject or OfdTraversalDiagnosticKind.UnsupportedSource)
            {
                var path = Key(diagnostic.OccurrencePath); rejected.Add(path);
                Notice(diagnostic.Kind == OfdTraversalDiagnosticKind.RawObject ? "RawBodyContent" : "UnsupportedBodySource",
                    diagnostic.OccurrencePath[0], path, true);
            }
        var certificates = new Dictionary<int, Dictionary<int, (int Layer, string Path)>>();
        var documentSupported = true;
        try { admission.PrepareDocument(); }
        catch (NotSupportedException exception) { documentSupported = false; Notice(exception.Message, null, null, true); }
        foreach (var index in selection)
        {
            var page = source.Pages[index];
            if (page.Templates.Count > 0) Notice("ExcludedTemplateScope", index, null, false);
            if (page.AnnotationAppearances.Count > 0) Notice("ExcludedAnnotationScope", index, null, false);
            if (!documentSupported) continue;
            try { certificates.Add(index, admission.Page(page, index)); }
            catch (NotSupportedException exception) { Notice(exception.Message, index, page.SourceEntryPath, true); }
        }
        var candidates = new Dictionary<OfdGraphicOccurrence, OfdKeywordTextSource>();
        var positionedSources = new Dictionary<OfdKeywordTextSource, List<OfdPositionedCluster>>();
        var needsProvider = false;
        foreach (var occurrence in traversal.Occurrences)
        {
            budget.Check();
            if (occurrence.Element is not OfdTextElement || rejected.Contains(Key(occurrence.OccurrencePath)) ||
                !certificates.TryGetValue(occurrence.OccurrencePath[0], out var certificate)) continue;
            try
            {
                if (!certificate.TryGetValue(occurrence.OccurrencePath[3], out var locator)) throw new NotSupportedException("UnprovedSourceLocator");
                var captured = OfdKeywordInputCapture.CaptureSource(occurrence, locator.Layer, locator.Path, budget, reserved[occurrence], admission);
                captured.Editable &= !snapshot.HasSharedRun((OfdTextElement)occurrence.Element, budget);
                if (!captured.Visible) continue; // Visibility exclusion is also a continuity barrier.
                foreach (var run in captured.Runs)
                {
                    var scalars = 0;
                    for (var i = 0; i < run.Text.Length; i++) { budget.Check(); scalars++; if (char.IsHighSurrogate(run.Text[i])) i++; }
                    budget.ReserveGlyphs(scalars);
                }
                var binding = admission.ResolveFont(captured.FontId);
                captured.Font = operation.CaptureFont(binding.Resource, captured.FontId);
                snapshot.Bind(binding.Resource, captured.Font, binding.Entry, budget);
                candidates.Add(occurrence, captured);
                positionedSources.Add(captured, new List<OfdPositionedCluster>());
                needsProvider |= !captured.AllRunsEmpty;
            }
            catch (NotSupportedException exception) { Notice(exception.Message, occurrence.OccurrencePath[0], Key(occurrence.OccurrencePath), true); }
        }
        snapshot.Seal(budget);
        snapshot.ValidateUnchanged(source, budget); snapshot.ValidateFontEntries(budget);
        var providerStatistics = new OfdGlyphGeometryStatistics();
        var geometryCache = capture is null ? null : new OfdKeywordGeometryCache(budget);
        OwnedProvider? owned = null; IOfdGlyphGeometrySession? session = null;
        if (needsProvider || privateOwned is not null)
        {
            owned = new OwnedProvider(provider, privateOwned?.Context);
            beforeCallback?.Invoke("OpenProvider");
            session = operation.OpenSession(owned);
            snapshot.ValidateUnchanged(source, budget); validateWholeSource?.Invoke();
        }
        var segmentText = new StringBuilder(); var segment = new List<OfdPositionedCluster>();
        var providerClusters = 0;
        foreach (var occurrence in traversal.Occurrences)
        {
            budget.Check();
            if (!candidates.TryGetValue(occurrence, out var captured)) { Flush(); continue; }
            if (captured.AllRunsEmpty) { Flush(); continue; } // Retained empty owner is a continuity barrier.
            foreach (var run in captured.Runs)
            {
                budget.Check();
                if (run.Text.Length == 0) continue; // Mixed empty slots do not discard or bridge actual glyph geometry.
                OfdGlyphGeometry geometry = null!;
                var reused = geometryCache is not null && geometryCache.TryGet(captured.Font!, run.Text, captured.Size, captured.Weight, captured.Italic, out geometry!);
                if (!reused)
                {
                    var request = operation.Request(captured.Font!, run.Text, captured.Size, captured.Weight, captured.Italic);
                    budget.Copy(512L * run.Text.Length + 256);
                    beforeCallback?.Invoke("MeasureSource");
                    geometry = session!.Measure(request, cancellationToken);
                    snapshot.ValidateUnchanged(source, budget); validateWholeSource?.Invoke();
                }
                operation.Check();
                if (geometry is null) throw new InvalidDataException("Provider returned no geometry.");
                if (!geometry.IsSupported)
                {
                    Flush(); Notice("UnsupportedFontGeometry", captured.PageIndex, run.Path, true); continue;
                }
                if (reused) budget.Reused(geometry.Clusters.Count); else budget.Returned(geometry.Clusters.Count);
                List<OfdPositionedCluster> positioned;
                try
                {
                    positioned = OfdStrictTextGeometry.Position(captured, run, geometry, budget,
                        operation.Options.Limits.MaxClusterVisits - providerClusters);
                    if (!reused) geometryCache?.Add(captured.Font!, run.Text, captured.Size, captured.Weight, captured.Italic, geometry);
                    providerClusters += geometry.Clusters.Count;
                }
                catch (NotSupportedException exception)
                {
                    // Returned clusters still consume cumulative work when clipping rejects geometry.
                    if (geometry.Clusters.Count > operation.Options.Limits.MaxClusterVisits - providerClusters)
                        throw new InvalidDataException("Provider cluster capacity exceeded.");
                    providerClusters += geometry.Clusters.Count;
                    Flush(); Notice(exception.Message, captured.PageIndex, run.Path, true); continue;
                }
                if (capture is not null) positionedSources[captured].AddRange(positioned);
                foreach (var cluster in positioned)
                {
                    budget.Check();
                    if (segment.Count > 0 && (!OfdStrictTextGeometry.Continuous(segment[segment.Count - 1], cluster,
                            ReferenceEquals(segment[segment.Count - 1].Run, cluster.Run)) || !OnBaseline(segment[0], cluster))) Flush();
                    segmentText.Append(run.Text, cluster.Geometry.Utf16Start, cluster.Geometry.Utf16Length); segment.Add(cluster);
                }
            }
        }
        Flush();
        if (capture is not null)
        {
            budget.Copy(2176L + 8L * hits.Count);
            var currentHits = hits.ToArray();
            capture(new OfdKeywordAnalysisCapture(source, traversal, candidates, positionedSources,
                operation, budget, snapshot, admission, session, validateWholeSource, beforeCallback, geometryCache, currentHits));
        }
        if (session is not null)
        {
            beforeCallback?.Invoke("ProviderStatistics");
            providerStatistics = session.Statistics ?? throw new InvalidDataException("Provider returned no statistics.");
            snapshot.ValidateUnchanged(source, budget); validateWholeSource?.Invoke();
            beforeCallback?.Invoke("DisposeProvider");
            owned!.Close(); // Actual user Dispose is a callback; call it once while font snapshots are still available.
            snapshot.ValidateUnchanged(source, budget); validateWholeSource?.Invoke();
        }
        beforeCallback?.Invoke("FinalFontValidation");
        snapshot.ValidateUnchanged(source, budget); snapshot.ValidateFontEntries(budget);
        operation.ValidateSourceFontsUnchanged();
        // The private factory's real fresh entry check discharges this final
        // whole-source obligation. Every actual provider callback guard and
        // both final font/entry comparisons above remain unchanged.
        if (finalization == Finalization.Independent) validateWholeSource?.Invoke();
        operation.Check();
        privateOwned?.Complete(token: cancellationToken);
        var statistics = new OfdKeywordLocatorStatistics(budget, operation, traversal.NodeCount, traversal.ReferenceCount, providerStatistics);
        return new OfdKeywordSearchResult(hits, diagnostics, selection, statistics);

        void Notice(string code, int? page, string? path, bool incomplete)
        {
            budget.Diagnostic();
            if (incomplete && operation.Options.UnsupportedContent == OfdUnsupportedContentBehavior.Fail)
                throw new NotSupportedException("Keyword body source is unsupported: " + code);
            diagnostics.Add(new OfdKeywordDiagnostic(code, page, path, incomplete));
        }
        void Flush()
        {
            if (segment.Count == 0) return;
            budget.Copy(512L * segment.Count + 8L * segmentText.Length + 256);
            var mappings = new OfdGlyphClusterGeometry[segment.Count]; var offset = 0;
            for (var i = 0; i < segment.Count; i++)
            {
                budget.Check(); var original = segment[i].Geometry;
                mappings[i] = new OfdGlyphClusterGeometry(offset, original.Utf16Length, original.GlyphId,
                    original.FontBoundsMillimeters, original.AdvanceXMillimeters, original.AdvanceYMillimeters);
                offset += original.Utf16Length;
            }
            foreach (var match in matcher.Find(segmentText.ToString(), mappings))
            {
                budget.Fragment(match.ClusterCount);
                budget.Copy(1024L * match.ClusterCount + 256);
                var fragments = new List<OfdKeywordFragment>(match.ClusterCount); var editable = true;
                for (var i = match.ClusterStart; i < match.ClusterStart + match.ClusterCount; i++)
                { budget.Check(); var cluster = segment[i]; fragments.Add(new OfdKeywordFragment(cluster)); editable &= cluster.Source.Editable; }
                hits.Add(new OfdKeywordHit(segment[match.ClusterStart].Source.PageIndex, keyword, fragments, editable, snapshot));
            }
            segment.Clear(); segmentText.Clear();
        }
        }
        catch { privateOwned?.Context.Fail(); throw; }
    }

    private static string Key(IReadOnlyList<int> path) => string.Join("/", path);
    private static bool OnBaseline(OfdPositionedCluster first, OfdPositionedCluster next)
    {
        var m = first.Source.Matrix; var length = Math.Sqrt(m[0] * m[0] + m[1] * m[1]);
        if (!OfdGlyphBox.Finite(length) || length == 0) return false;
        return Math.Abs((next.Baseline.X - first.Baseline.X) * -m[1] / length +
            (next.Baseline.Y - first.Baseline.Y) * m[0] / length) <= 0.01;
    }
    private sealed class PrivateOwnedAnalysis
    {
        private readonly OfdOwnedAnalysisClaim claim;
        private readonly OfdOwnedGlyphGeometryHost host;
        internal readonly OfdControlledEditContext Context;
        internal readonly PrivateHostAdapter Adapter;
        internal OfdOwnedHostClosure? Closure { get; private set; }
        internal PrivateOwnedAnalysis(OfdOwnedAnalysisClaim claim, OfdOwnedGlyphGeometryHost host, OfdControlledEditContext context)
        {
            this.claim = claim ?? throw new ArgumentNullException(nameof(claim));
            this.host = host ?? throw new ArgumentNullException(nameof(host)); Context = context;
            Adapter = new PrivateHostAdapter(host);
        }
        internal void Complete(CancellationToken token)
        {
            if (!Adapter.Closed) throw new InvalidOperationException("HostNotActuallyClosed");
            Closure = OfdOwnedSaveAuthority.CompleteClosedHost(claim, host, token);
        }
    }
    // This adapter is constructed only by the typed private binding above.
    // It never escapes into a public IProvider entry and cannot install hooks.
    private sealed class PrivateHostAdapter(OfdOwnedGlyphGeometryHost host) : IOfdGlyphGeometryProvider, IOfdGlyphGeometrySession
    {
        private bool opened;
        internal bool Closed { get; private set; }
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits limits, CancellationToken token = default)
        {
            if (opened || Closed) throw new InvalidOperationException("HostSessionAlreadyConsumed");
            opened = true; host.Open(limits, token); return this;
        }
        public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken token = default) => host.Measure(request, token);
        public OfdGlyphGeometryStatistics Statistics => host.ReadStatistics();
        public void Dispose()
        {
            if (Closed || !opened) return;
            Closed = true; host.Close();
        }
    }
    private sealed class OwnedProvider(IOfdGlyphGeometryProvider inner, OfdControlledEditContext? privateContext = null) : IOfdGlyphGeometryProvider, IOfdGlyphGeometrySession
    {
        private IOfdGlyphGeometrySession? actual;
        public IOfdGlyphGeometrySession OpenSession(OfdGlyphGeometryLimits value, CancellationToken token = default)
        {
            actual = inner.OpenSession(value, token) ?? throw new InvalidDataException("Provider returned no session."); return this;
        }
        public OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken token = default) => actual!.Measure(request, token);
        public OfdGlyphGeometryStatistics Statistics => actual!.Statistics;
        internal void Close() { var current = actual; actual = null; current?.Dispose(); }
        public void Dispose()
        {
            if (privateContext is null) Close();
            else { try { Close(); } catch { } } // the private actual Close is called explicitly above
        }
    }
}
