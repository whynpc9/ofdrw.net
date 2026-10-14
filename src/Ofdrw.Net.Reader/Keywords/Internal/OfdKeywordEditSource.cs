using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// Valid only inside the shared analysis callback, while its original font
// snapshots and single provider session are alive. Never a public credential.
internal sealed class OfdKeywordAnalysisCapture
{
    internal OfdDocumentPackage Package { get; }
    internal OfdGraphicTraversalResult Traversal { get; }
    internal Dictionary<OfdGraphicOccurrence, OfdKeywordTextSource> Sources { get; }
    internal Dictionary<OfdKeywordTextSource, List<OfdPositionedCluster>> Positions { get; }
    internal OfdKeywordLocatorBudget Budget { get; }
    internal IReadOnlyList<OfdKeywordHit> CurrentHits { get; }
    private readonly OfdKeywordOperation operation;
    private readonly OfdKeywordSourceSnapshot snapshot;
    internal OfdKeywordSourceAdmission Admission { get; }
    internal string SourceVersion => snapshot.Version;
    private readonly IOfdGlyphGeometrySession? session;
    private readonly Action? validateWholeSource;
    private readonly Action<string>? beforeCallback;
    private readonly OfdKeywordGeometryCache? geometryCache;
    private readonly Dictionary<string, KeyValuePair<OfdGraphicOccurrence, OfdKeywordTextSource>> owners = new(StringComparer.Ordinal);
    private readonly Dictionary<OfdKeywordTextSource, Dictionary<(string Path, int Offset), int>> indices = new();
    internal OfdKeywordAnalysisCapture(OfdDocumentPackage package, OfdGraphicTraversalResult traversal,
        Dictionary<OfdGraphicOccurrence, OfdKeywordTextSource> sources,
        Dictionary<OfdKeywordTextSource, List<OfdPositionedCluster>> positions, OfdKeywordOperation operation,
        OfdKeywordLocatorBudget budget, OfdKeywordSourceSnapshot snapshot, OfdKeywordSourceAdmission admission, IOfdGlyphGeometrySession? session,
        Action? validateWholeSource, Action<string>? beforeCallback, OfdKeywordGeometryCache? geometryCache,
        OfdKeywordHit[] currentHits)
    {
        Package = package; Traversal = traversal; Sources = sources; Positions = positions;
        this.operation = operation; Budget = budget; this.snapshot = snapshot; Admission = admission;
        this.session = session; this.validateWholeSource = validateWholeSource;
        this.beforeCallback = beforeCallback;
        this.geometryCache = geometryCache;
        CurrentHits = Array.AsReadOnly(currentHits);
        budget.Copy(1024L + 256L * sources.Count + 128L * positions.Sum(value => (long)value.Value.Count));
        foreach (var pair in sources)
        {
            budget.Correlate(); budget.Copy(64L + 24L * pair.Key.OccurrencePath.Count);
            owners.Add(string.Join("/", pair.Key.OccurrencePath), pair);
            var index = new Dictionary<(string Path, int Offset), int>(); indices.Add(pair.Value, index);
            var clusters = positions[pair.Value];
            for (var i = 0; i < clusters.Count; i++)
            {
                var cluster = clusters[i]; budget.Correlate(cluster.Run.Path.Length);
                index.Add((cluster.Run.Path, cluster.Geometry.Utf16Start), i);
            }
        }
    }
    internal KeyValuePair<OfdGraphicOccurrence, OfdKeywordTextSource> FindOwner(OfdKeywordFragment fragment)
    {
        Budget.Correlate(); Budget.Copy(64L + 24L * fragment.OccurrencePath.Count);
        if (!owners.TryGetValue(string.Join("/", fragment.OccurrencePath), out var owner))
            throw new NotSupportedException("UnprovedEditOwner");
        return owner;
    }
    internal int FindPosition(OfdKeywordTextSource source, OfdKeywordFragment fragment)
    {
        Budget.Correlate(fragment.TextCodePath.Length);
        return indices[source].TryGetValue((fragment.TextCodePath, fragment.LogicalUtf16Start), out var index) ? index : -1;
    }
    internal void AuthenticateCurrentHit(OfdKeywordHit hit)
    {
        if (!ReferenceEquals(hit.Source, snapshot) || hit.SourceVersion != snapshot.Version)
            throw new InvalidDataException("MixedKeywordSourceVersions");
        if (!hit.CanReplace || hit.Fragments.Count == 0 || hit.Fragments.Any(fragment => !fragment.IsAuthoritativeXml))
            throw new NotSupportedException("UneditableKeywordSource");
        OfdKeywordEditSource.Authenticate(this, hit);
    }
    internal OfdGlyphGeometry Measure(OfdKeywordTextSource source, string text)
    {
        operation.Check();
        if (text.Length == 0) { Budget.Copy(256); return new OfdGlyphGeometry(Array.Empty<OfdGlyphClusterGeometry>()); }
        var count = 0;
        for (var i = 0; i < text.Length; i++, count++) if (char.IsHighSurrogate(text[i])) i++;
        Budget.ReserveGlyphs(count); Budget.Copy(512L * count + 256);
        if (geometryCache is not null && geometryCache.TryGet(source.Font!, text, source.Size, source.Weight, source.Italic, out var reused))
        { Budget.Reused(reused.Clusters.Count); return reused; }
        beforeCallback?.Invoke("MeasureReplacement");
        var geometry = session!.Measure(operation.Request(source.Font!, text, source.Size, source.Weight, source.Italic), operation.Token);
        snapshot.ValidateUnchanged(Package, Budget); validateWholeSource?.Invoke(); operation.Check();
        if (geometry is null) throw new InvalidDataException("Provider returned no replacement geometry.");
        if (!geometry.IsSupported) throw new NotSupportedException("UnsupportedReplacementFontGeometry");
        Budget.Returned(geometry.Clusters.Count);
        OfdStrictTextGeometry.ValidateCoverage(text, geometry, Budget);
        geometryCache?.Add(source.Font!, text, source.Size, source.Weight, source.Italic, geometry);
        return geometry;
    }
}

internal static class OfdKeywordEditSource
{
    internal static (long Bytes, long Operations) ValidateHits(OfdDocumentPackage package, IReadOnlyList<OfdKeywordHit> hits,
        OfdKeywordLocatorLimits limits, OfdKeywordSearchOptions options, OfdModelMaterialization material, CancellationToken token)
    {
        if (hits.Count == 0) return (0, 0);
        using var operation = new OfdKeywordOperation(options, token, material);
        var budget = new OfdKeywordLocatorBudget(limits, operation); var snapshot = hits[0].Source;
        foreach (var hit in hits)
        {
            budget.Check();
            if (!ReferenceEquals(snapshot, hit.Source) || hit.SourceVersion != snapshot.Version)
                throw new InvalidDataException("MixedKeywordSourceVersions");
            if (!hit.CanReplace || hit.Fragments.Count == 0 || hit.Fragments.Any(fragment => !fragment.IsAuthoritativeXml))
                throw new NotSupportedException("UneditableKeywordSource");
        }
        snapshot.ValidateForEdit(package, budget);
        return (budget.RetainedFontDigestBytes, budget.RetainedFontDigestOperations);
    }

    internal static OfdKeywordSearchResult Analyze(OfdDocumentPackage package, IReadOnlyList<int> pages, string keyword,
        IOfdGlyphGeometryProvider provider, OfdKeywordLocatorLimits limits, OfdKeywordSearchOptions options,
        OfdModelMaterialization material, Action validateWholeSource, Action<OfdKeywordAnalysisCapture>? capture,
        CancellationToken token, Action<string>? beforeCallback = null)
        => new OfdKeywordAnalysis(provider, limits, material, validateWholeSource, capture, beforeCallback)
            .Find(package, pages, keyword, options, token);

    internal static OfdKeywordSearchResult AnalyzeForOwnedFile(OfdDocumentPackage package, int[] pages, string keyword,
        IOfdGlyphGeometryProvider provider, OfdKeywordLocatorLimits limits, OfdKeywordSearchOptions options,
        OfdModelMaterialization material, Action validateWholeSource, Action<OfdKeywordAnalysisCapture> capture,
        CancellationToken token, Action<string> beforeCallback)
        => new OfdKeywordAnalysis(provider, limits, material, validateWholeSource, capture, beforeCallback)
            .FindForOwnedFile(package, pages, keyword, options, token);

    internal static (OfdKeywordTextSource Source, int Start, int Count) Authenticate(OfdKeywordAnalysisCapture capture, OfdKeywordHit hit)
    {
        var fragments = hit.Fragments; var first = fragments[0]; var budget = capture.Budget;
        var pair = capture.FindOwner(first);
        if (pair.Key is null || !pair.Value.Editable || !pair.Value.HasXml)
            throw new NotSupportedException("UnprovedEditOwner");
        var source = pair.Value; var positions = capture.Positions[source];
        budget.Copy(128L + 4L * hit.MatchedText.Length);
        var logical = new System.Text.StringBuilder(hit.MatchedText.Length); var start = capture.FindPosition(source, first); var offset = 0;
        foreach (var fragment in fragments)
        {
            budget.Correlate(checked(fragment.TextCodePath.Length + fragment.FontResourceId.Length + fragment.FontSha256.Length +
                (fragment.SourceEntryPath?.Length ?? 0) + 8 * (fragment.OccurrencePath.Count + fragment.DefinitionPath.Count)));
            if (!fragment.OccurrencePath.SequenceEqual(first.OccurrencePath) ||
                !fragment.DefinitionPath.SequenceEqual(pair.Key.DefinitionPath) || fragment.SourceEntryPath != source.EntryPath ||
                fragment.FontResourceId != source.FontId || fragment.FontSha256 != source.Font!.ContentSha256 ||
                fragment.FontByteCount != source.Font.ByteCount || fragment.FontSizeMillimeters != source.Size ||
                fragment.Weight != source.Weight || fragment.Italic != source.Italic ||
                fragment.FontBold != source.Font.Bold || fragment.FontItalic != source.Font.Italic || !fragment.IsAuthoritativeXml)
                throw new InvalidDataException("KeywordFragmentIdentityMismatch");
            var index = start + offset++;
            if (start < 0 || index >= positions.Count)
                throw new InvalidDataException("KeywordFragmentRangeMismatch");
            var position = positions[index]; var geometry = position.Geometry; var run = position.Run;
            if (run.Path != fragment.TextCodePath || geometry.Utf16Start != fragment.LogicalUtf16Start ||
                geometry.Utf16Length != fragment.LogicalUtf16Length || geometry.GlyphId != fragment.GlyphId ||
                run.SourceOffsets[geometry.Utf16Start] != fragment.SourceUtf16Start ||
                run.SourceOffsets[geometry.Utf16Start + geometry.Utf16Length] - fragment.SourceUtf16Start != fragment.SourceUtf16Length ||
                position.Baseline.X != fragment.BaselineMillimeters.X || position.Baseline.Y != fragment.BaselineMillimeters.Y ||
                position.Advance.X != fragment.AdvanceXMillimeters || position.Advance.Y != fragment.AdvanceYMillimeters ||
                !SameInk(position.Quad, fragment.InkQuadrilateralMillimeters))
                throw new InvalidDataException("KeywordFragmentGeometryMismatch");
            logical.Append(run.Text, geometry.Utf16Start, geometry.Utf16Length);
        }
        budget.Work(2L * hit.MatchedText.Length);
        if (logical.ToString() != hit.MatchedText) throw new InvalidDataException("KeywordFragmentTextMismatch");
        return (source, start, fragments.Count);
    }
    private static bool SameInk(OfdKeywordQuadrilateral? left, OfdKeywordQuadrilateral? right)
    {
        if (left is null || right is null) return left is null && right is null;
        for (var i = 0; i < 4; i++)
            if (left.Corners[i].X != right.Corners[i].X || left.Corners[i].Y != right.Corners[i].Y) return false;
        return true;
    }
}
