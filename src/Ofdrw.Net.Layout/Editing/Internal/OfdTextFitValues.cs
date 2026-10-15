using System.Collections.ObjectModel;
using System.Threading;

namespace Ofdrw.Net.Layout.Editing.Internal;

// These labels and metrics are supplied values, never source authentication.
internal enum OfdTextFitProfile { SimpleScalarMetrics, Unsupported }
internal enum OfdTextFitDisposition { Unchanged, ReplaceScalars, RetainRootWithCanonicalEmptyTextCode }
internal enum OfdTextFitFailure
{
    None, InvalidIdentity, UnsupportedProfile, UnsupportedTransform, InvalidGeometry,
    InvalidScalarMapping, InvalidRawMapping, InvalidRequest, OverlappingRequests,
    AliasedOwner, UnprovedNeighbors, UnsupportedProgression, UnprovedHeight,
    OriginalOutsideRegion, ReplacementOutsideRegion, Collision
}

internal readonly struct OfdTextFitBox
{
    internal OfdTextFitBox(double left, double top, double right, double bottom)
    { Left = left; Top = top; Right = right; Bottom = bottom; }
    internal double Left { get; }
    internal double Top { get; }
    internal double Right { get; }
    internal double Bottom { get; }
}

internal readonly struct OfdTextFitAffine
{
    internal OfdTextFitAffine(double a = 1, double b = 0, double c = 0, double d = 1, double tx = 0, double ty = 0)
    { A = a; B = b; C = c; D = d; Tx = tx; Ty = ty; }
    internal double A { get; }
    internal double B { get; }
    internal double C { get; }
    internal double D { get; }
    internal double Tx { get; }
    internal double Ty { get; }
}

internal sealed class OfdTextFitMetric
{
    internal OfdTextFitMetric(int scalar, int utf16Start, int utf16Length, uint glyphId,
        double advanceX, double advanceY, OfdTextFitBox? ink)
    { Scalar = scalar; Utf16Start = utf16Start; Utf16Length = utf16Length; GlyphId = glyphId; AdvanceX = advanceX; AdvanceY = advanceY; Ink = ink; }
    internal int Scalar { get; }
    internal int Utf16Start { get; }
    internal int Utf16Length { get; }
    internal uint GlyphId { get; }
    internal double AdvanceX { get; }
    internal double AdvanceY { get; }
    internal OfdTextFitBox? Ink { get; }
}

internal sealed class OfdTextFitSourceCluster
{
    internal OfdTextFitSourceCluster(OfdTextFitMetric metric, double x, double y, string slot, int rawStart, int rawLength)
    { Metric = metric ?? throw new ArgumentNullException(nameof(metric)); X = x; Y = y; Slot = OfdTextFitCapture.Token(slot); RawStart = rawStart; RawLength = rawLength; }
    internal OfdTextFitMetric Metric { get; }
    internal double X { get; }
    internal double Y { get; }
    internal string Slot { get; }
    internal int RawStart { get; }
    internal int RawLength { get; }
}

internal sealed class OfdTextFitObject
{
    internal OfdTextFitObject(int page, string occurrence, string owner, string version, string scope,
        string fontStyle, string text, IReadOnlyList<OfdTextFitSourceCluster> clusters,
        OfdTextFitAffine affine, OfdTextFitBox boundary, OfdTextFitBox pageClip,
        bool completeNeighbors = true, OfdTextFitProfile profile = OfdTextFitProfile.SimpleScalarMetrics,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Page = page; Occurrence = OfdTextFitCapture.Token(occurrence); Owner = OfdTextFitCapture.Token(owner);
        Version = OfdTextFitCapture.Token(version); Scope = OfdTextFitCapture.Token(scope); FontStyle = OfdTextFitCapture.Token(fontStyle);
        Text = OfdTextFitCapture.Text(text); Clusters = OfdTextFitCapture.Copy(clusters, OfdTextFitLimits.HardClusters, cancellationToken);
        Affine = affine; Boundary = boundary; PageClip = pageClip; CompleteNeighbors = completeNeighbors; Profile = profile;
    }
    internal int Page { get; }
    internal string Occurrence { get; }
    internal string Owner { get; }
    internal string Version { get; }
    internal string Scope { get; }
    internal string FontStyle { get; }
    internal string Text { get; }
    internal IReadOnlyList<OfdTextFitSourceCluster> Clusters { get; }
    internal int CapturedPositions => Clusters.Count;
    internal OfdTextFitAffine Affine { get; }
    internal OfdTextFitBox Boundary { get; }
    internal OfdTextFitBox PageClip { get; }
    internal bool CompleteNeighbors { get; }
    internal OfdTextFitProfile Profile { get; }
}

internal sealed class OfdTextFitRequest
{
    internal OfdTextFitRequest(string occurrence, int scalarStart, int scalarCount, string expected,
        string replacement, string fontStyle, IReadOnlyList<OfdTextFitMetric> metrics, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Occurrence = OfdTextFitCapture.Token(occurrence); ScalarStart = scalarStart; ScalarCount = scalarCount;
        Expected = OfdTextFitCapture.Text(expected); Replacement = OfdTextFitCapture.Text(replacement);
        FontStyle = OfdTextFitCapture.Token(fontStyle); Metrics = OfdTextFitCapture.Copy(metrics, OfdTextFitLimits.HardClusters, cancellationToken);
    }
    internal string Occurrence { get; }
    internal int ScalarStart { get; }
    internal int ScalarCount { get; }
    internal string Expected { get; }
    internal string Replacement { get; }
    internal string FontStyle { get; }
    internal IReadOnlyList<OfdTextFitMetric> Metrics { get; }
    internal int CapturedPositions => Metrics.Count;
}

internal sealed class OfdTextFitNeighbor
{
    internal OfdTextFitNeighbor(int page, string key, OfdTextFitBox box)
    { Page = page; Key = OfdTextFitCapture.Token(key); Box = box; }
    internal int Page { get; }
    internal string Key { get; }
    internal OfdTextFitBox Box { get; }
}

internal sealed class OfdTextFitInput
{
    internal OfdTextFitInput(string version, string scope, IReadOnlyList<OfdTextFitObject> objects,
        IReadOnlyList<OfdTextFitRequest> requests, IReadOnlyList<OfdTextFitNeighbor> neighbors, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        Version = OfdTextFitCapture.Token(version); Scope = OfdTextFitCapture.Token(scope);
        Objects = OfdTextFitCapture.Copy(objects, OfdTextFitLimits.HardObjects, cancellationToken);
        Requests = OfdTextFitCapture.Copy(requests, OfdTextFitLimits.HardRequests, cancellationToken);
        Neighbors = OfdTextFitCapture.Copy(neighbors, OfdTextFitLimits.HardNeighbors, cancellationToken);
    }
    internal string Version { get; }
    internal string Scope { get; }
    internal IReadOnlyList<OfdTextFitObject> Objects { get; }
    internal IReadOnlyList<OfdTextFitRequest> Requests { get; }
    internal IReadOnlyList<OfdTextFitNeighbor> Neighbors { get; }
    internal int CapturedPositions => Objects.Count + Requests.Count + Neighbors.Count;
}

internal sealed class OfdTextFitPlacement
{
    internal OfdTextFitPlacement(OfdTextFitMetric metric, int outputUtf16Start, double x, double y, OfdTextFitSourceCluster? source)
    { Metric = metric; OutputUtf16Start = outputUtf16Start; X = x; Y = y; Source = source; }
    internal OfdTextFitMetric Metric { get; }
    internal int OutputUtf16Start { get; }
    internal double X { get; }
    internal double Y { get; }
    internal OfdTextFitSourceCluster? Source { get; }
}

internal sealed class OfdTextFitRequestPlan
{
    internal OfdTextFitRequestPlan(int requestIndex, int scalarStart, int scalarCount, OfdTextFitDisposition disposition)
    { RequestIndex = requestIndex; ScalarStart = scalarStart; ScalarCount = scalarCount; Disposition = disposition; }
    internal int RequestIndex { get; }
    internal int ScalarStart { get; }
    internal int ScalarCount { get; }
    internal OfdTextFitDisposition Disposition { get; }
}

internal sealed class OfdTextFitObjectPlan
{
    internal OfdTextFitObjectPlan(OfdTextFitObject original, string text, List<OfdTextFitPlacement> placements,
        List<OfdTextFitRequestPlan> requests)
    {
        Occurrence = original.Occurrence; Owner = original.Owner; FontStyle = original.FontStyle;
        Text = text; Placements = Array.AsReadOnly(placements.ToArray()); Requests = Array.AsReadOnly(requests.ToArray());
        RetainRootWithCanonicalEmptyTextCode = text.Length == 0;
        EmptyAnchorX = original.Clusters.Count == 0 ? 0 : original.Clusters[0].X;
        EmptyAnchorY = original.Clusters.Count == 0 ? 0 : original.Clusters[0].Y;
    }
    internal string Occurrence { get; }
    internal string Owner { get; }
    internal string FontStyle { get; }
    internal string Text { get; }
    internal IReadOnlyList<OfdTextFitPlacement> Placements { get; }
    internal IReadOnlyList<OfdTextFitRequestPlan> Requests { get; }
    internal bool RetainRootWithCanonicalEmptyTextCode { get; }
    internal double EmptyAnchorX { get; }
    internal double EmptyAnchorY { get; }
}

internal sealed class OfdTextFitStatistics
{
    internal OfdTextFitStatistics(long capturedPositions, long visits, long outputPositions, long outputUtf16, long comparisons)
    { CapturedPositions = capturedPositions; Visits = visits; OutputPositions = outputPositions; OutputUtf16 = outputUtf16; Comparisons = comparisons; }
    // Positions copied when the referenced snapshot factories ran, counting distinct
    // child snapshots once. Each Plan charges these existing captures again; this is
    // not a claim that Plan performed those copies or bounded upstream memory/RSS.
    internal long CapturedPositions { get; }
    internal long Visits { get; }
    internal long OutputPositions { get; }
    internal long OutputUtf16 { get; }
    internal long Comparisons { get; }
}

internal sealed class OfdTextFitPlanResult
{
    internal OfdTextFitPlanResult(OfdTextFitFailure failure, int objectIndex, int requestIndex,
        IReadOnlyList<OfdTextFitObjectPlan> plans, OfdTextFitStatistics statistics)
    { Failure = failure; ObjectIndex = objectIndex; RequestIndex = requestIndex; Plans = plans; Statistics = statistics; }
    internal bool FitsSuppliedGeometry => Failure == OfdTextFitFailure.None;
    internal OfdTextFitFailure Failure { get; }
    internal int ObjectIndex { get; }
    internal int RequestIndex { get; }
    internal IReadOnlyList<OfdTextFitObjectPlan> Plans { get; }
    internal OfdTextFitStatistics Statistics { get; }
}

internal static class OfdTextFitCapture
{
    internal static string Token(string value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length > 256) throw new OfdTextFitLimitException();
        return value;
    }
    internal static string Text(string value)
    {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length > OfdTextFitLimits.HardUtf16) throw new OfdTextFitLimitException();
        return value;
    }
    internal static ReadOnlyCollection<T> Copy<T>(IReadOnlyList<T> values, int maximum, CancellationToken token) where T : class
    {
        token.ThrowIfCancellationRequested();
        if (values == null) throw new ArgumentNullException(nameof(values));
        int count = values.Count;
        if (count < 0 || count > maximum) throw new OfdTextFitLimitException();
        var copy = new T[count];
        for (int i = 0; i < count; i++) { token.ThrowIfCancellationRequested(); copy[i] = values[i] ?? throw new ArgumentException("Null frozen value."); }
        token.ThrowIfCancellationRequested();
        return Array.AsReadOnly(copy);
    }
}
