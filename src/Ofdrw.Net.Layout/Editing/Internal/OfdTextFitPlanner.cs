using System.Text;
using System.Threading;

namespace Ofdrw.Net.Layout.Editing.Internal;

// Conservative scalar-metric foundation. Correlation labels, complete-neighbor flags and
// metrics are assertions of a future adapter; a successful result authenticates none of them.
internal static class OfdTextFitPlanner
{
    internal static OfdTextFitPlanResult Plan(OfdTextFitInput input, OfdTextFitLimits limits, CancellationToken cancellationToken = default,
        Action<long>? chargeWork = null)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (limits == null) throw new ArgumentNullException(nameof(limits));
        var budget = new OfdTextFitBudget(limits, cancellationToken, chargeWork);
        try { return Run(input, limits, budget); }
        catch (Rejection failure)
        {
            budget.Check();
            return new OfdTextFitPlanResult(failure.Reason, failure.ObjectIndex, failure.RequestIndex,
                Array.Empty<OfdTextFitObjectPlan>(), budget.Snapshot());
        }
    }

    private static OfdTextFitPlanResult Run(OfdTextFitInput input, OfdTextFitLimits limits, OfdTextFitBudget budget)
    {
        Preflight(input, limits, budget);
        Need(Identity(input.Version) && Identity(input.Scope), OfdTextFitFailure.InvalidIdentity);
        var indices = new Dictionary<string, int>(StringComparer.Ordinal);
        var owners = new Dictionary<string, int>(StringComparer.Ordinal);
        var aliasedOwners = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < input.Objects.Count; i++)
        {
            budget.Visit();
            var obj = input.Objects[i];
            Need(Identity(obj.Occurrence) && Identity(obj.Owner) && Identity(obj.FontStyle) && obj.Page >= 0 &&
                obj.Version == input.Version && obj.Scope == input.Scope, OfdTextFitFailure.InvalidIdentity, i);
            Need(!indices.ContainsKey(obj.Occurrence), OfdTextFitFailure.InvalidIdentity, i);
            indices.Add(obj.Occurrence, i);
            if (owners.ContainsKey(obj.Owner)) aliasedOwners.Add(obj.Owner); else owners.Add(obj.Owner, i);
            Need(obj.Profile == OfdTextFitProfile.SimpleScalarMetrics, OfdTextFitFailure.UnsupportedProfile, i);
            var affine = obj.Affine;
            Need(affine.A == 1 && affine.B == 0 && affine.C == 0 && affine.D == 1 && Finite(affine.Tx) && Finite(affine.Ty),
                OfdTextFitFailure.UnsupportedTransform, i);
            Need(BoxValid(obj.Boundary) && obj.Boundary.Left == 0 && obj.Boundary.Top == 0 && BoxValid(obj.PageClip),
                OfdTextFitFailure.InvalidGeometry, i);
            ValidateSource(obj, i, budget);
        }
        var neighborKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var neighbor in input.Neighbors)
        {
            budget.Visit();
            Need(neighbor.Page >= 0 && Identity(neighbor.Key) && BoxValid(neighbor.Box), OfdTextFitFailure.InvalidGeometry);
            Need(neighborKeys.Add(neighbor.Page.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + neighbor.Key),
                OfdTextFitFailure.InvalidIdentity);
        }
        var edits = new List<Edit>(input.Requests.Count);
        long projectedPositions = 0, projectedUtf16 = 0;
        var touched = new HashSet<int>();
        for (int r = 0; r < input.Requests.Count; r++)
        {
            budget.Visit();
            var request = input.Requests[r];
            Need(indices.TryGetValue(request.Occurrence, out int o), OfdTextFitFailure.InvalidRequest, -1, r);
            var obj = input.Objects[o];
            Need(request.FontStyle == obj.FontStyle && request.ScalarStart >= 0 && request.ScalarCount > 0 &&
                (long)request.ScalarStart + request.ScalarCount <= obj.Clusters.Count, OfdTextFitFailure.InvalidRequest, o, r);
            ValidateMetrics(request.Replacement, request.Metrics, o, r, budget);
            int first = obj.Clusters[request.ScalarStart].Metric.Utf16Start;
            int end = request.ScalarStart + request.ScalarCount;
            int utf16End = end == obj.Clusters.Count ? obj.Text.Length : obj.Clusters[end].Metric.Utf16Start;
            Need(request.Expected.Length == utf16End - first &&
                string.CompareOrdinal(obj.Text, first, request.Expected, 0, request.Expected.Length) == 0,
                OfdTextFitFailure.InvalidRequest, o, r);
            bool changed = request.Expected != request.Replacement;
            if (changed) Need(!aliasedOwners.Contains(obj.Owner), OfdTextFitFailure.AliasedOwner, o, r);
            if (touched.Add(o)) { projectedPositions += obj.Clusters.Count; projectedUtf16 += obj.Text.Length; }
            if (changed) { projectedPositions += request.Metrics.Count - request.ScalarCount; projectedUtf16 += request.Replacement.Length - request.Expected.Length; }
            edits.Add(new Edit(o, r, request, changed));
        }
        // Sort with a budgeted insertion comparator. Unlike framework Sort, this does
        // not wrap cancellation/limit exceptions or invoke an external comparer.
        for (int i = 1; i < edits.Count; i++)
        {
            var value = edits[i]; int j = i - 1;
            while (j >= 0)
            {
                budget.Compare();
                var prior = edits[j];
                if (prior.ObjectIndex < value.ObjectIndex || prior.ObjectIndex == value.ObjectIndex && prior.Request.ScalarStart <= value.Request.ScalarStart) break;
                edits[j + 1] = prior; j--;
            }
            edits[j + 1] = value;
        }
        for (int i = 1; i < edits.Count; i++)
        {
            budget.Compare(); var a = edits[i - 1]; var b = edits[i];
            Need(a.ObjectIndex != b.ObjectIndex || a.End <= b.Request.ScalarStart, OfdTextFitFailure.OverlappingRequests, b.ObjectIndex, b.RequestIndex);
        }
        // Validate all interval conflicts before applying projected-count deltas. This
        // prevents overlapping deletions from disguising an oversized output count.
        OfdTextFitBudget.Add(0, projectedPositions, limits.OutputClusters);
        OfdTextFitBudget.Add(0, projectedUtf16, limits.OutputUtf16);
        var checkedProgressions = new HashSet<int>();
        var checkedCoveragePages = new HashSet<int>();
        foreach (var edit in edits)
        {
            budget.Check();
            if (!edit.Changed) continue;
            var obj = input.Objects[edit.ObjectIndex];
            // Coverage is a page assertion: another supplied object cannot disagree.
            if (checkedCoveragePages.Add(obj.Page))
                foreach (var pageObject in input.Objects)
                {
                    budget.Compare();
                    if (pageObject.Page == obj.Page) Need(pageObject.CompleteNeighbors, OfdTextFitFailure.UnprovedNeighbors, edit.ObjectIndex, edit.RequestIndex);
                }
            if (checkedProgressions.Add(edit.ObjectIndex)) ValidateProgression(obj, edit, budget);
            Place(obj, edit, budget);
        }
        var obstacles = BuildUnchangedObstacles(input, edits, budget);
        var inserted = new List<Footprint>();
        foreach (var edit in edits)
        {
            foreach (var footprint in edit.Footprints)
            {
                budget.Check();
                foreach (var obstacle in obstacles)
                {
                    budget.Compare();
                    Need(footprint.Page != obstacle.Page || !Intersects(footprint.Box, obstacle.Box),
                        OfdTextFitFailure.Collision, edit.ObjectIndex, edit.RequestIndex);
                }
                foreach (var prior in inserted)
                {
                    budget.Compare();
                    Need(footprint.Page != prior.Page || !Intersects(footprint.Box, prior.Box),
                        OfdTextFitFailure.Collision, edit.ObjectIndex, edit.RequestIndex);
                }
                inserted.Add(footprint);
            }
        }
        var plans = new List<OfdTextFitObjectPlan>(touched.Count);
        int cursor = 0;
        while (cursor < edits.Count)
        {
            budget.Check(); int end = cursor + 1;
            while (end < edits.Count && edits[end].ObjectIndex == edits[cursor].ObjectIndex) { budget.Visit(); end++; }
            plans.Add(BuildPlan(input.Objects[edits[cursor].ObjectIndex], edits, cursor, end, budget));
            cursor = end;
        }
        budget.Check();
        return new OfdTextFitPlanResult(OfdTextFitFailure.None, -1, -1, plans.AsReadOnly(), budget.Snapshot());
    }

    private static void Preflight(OfdTextFitInput input, OfdTextFitLimits limits, OfdTextFitBudget budget)
    {
        OfdTextFitBudget.Add(0, input.Objects.Count, limits.Objects);
        OfdTextFitBudget.Add(0, input.Requests.Count, limits.Requests);
        OfdTextFitBudget.Add(0, input.Neighbors.Count, limits.Neighbors);
        long clusters = 0, utf16 = 0, tokens = 0;
        long copies = input.CapturedPositions;
        var capturedObjects = new HashSet<OfdTextFitObject>();
        var capturedRequests = new HashSet<OfdTextFitRequest>();
        tokens = OfdTextFitBudget.Add(tokens, input.Version.Length + input.Scope.Length, limits.TokenUtf16);
        foreach (var obj in input.Objects)
        {
            budget.Visit();
            clusters = OfdTextFitBudget.Add(clusters, obj.Clusters.Count, limits.Clusters);
            if (capturedObjects.Add(obj)) copies += obj.CapturedPositions;
            utf16 = OfdTextFitBudget.Add(utf16, obj.Text.Length, limits.InputUtf16);
            tokens = OfdTextFitBudget.Add(tokens, obj.Occurrence.Length + obj.Owner.Length + obj.Version.Length + obj.Scope.Length + obj.FontStyle.Length, limits.TokenUtf16);
            foreach (var cluster in obj.Clusters) { budget.Visit(); tokens = OfdTextFitBudget.Add(tokens, cluster.Slot.Length, limits.TokenUtf16); }
        }
        foreach (var request in input.Requests)
        {
            budget.Visit(); clusters = OfdTextFitBudget.Add(clusters, request.Metrics.Count, limits.Clusters);
            if (capturedRequests.Add(request)) copies += request.CapturedPositions;
            utf16 = OfdTextFitBudget.Add(utf16, (long)request.Expected.Length + request.Replacement.Length, limits.InputUtf16);
            tokens = OfdTextFitBudget.Add(tokens, request.Occurrence.Length + request.FontStyle.Length, limits.TokenUtf16);
        }
        foreach (var neighbor in input.Neighbors) { budget.Visit(); tokens = OfdTextFitBudget.Add(tokens, neighbor.Key.Length, limits.TokenUtf16); }
        budget.CapturedPositions = copies;
    }

    private static void ValidateSource(OfdTextFitObject obj, int o, OfdTextFitBudget budget)
    {
        int offset = 0;
        var slots = new Dictionary<string, List<OfdTextFitSourceCluster>>(StringComparer.Ordinal);
        foreach (var source in obj.Clusters)
        {
            budget.Visit();
            ValidateMetric(obj.Text, source.Metric, ref offset, o, -1);
            Need(Finite(source.X) && Finite(source.Y), OfdTextFitFailure.InvalidGeometry, o);
            // Positive disjoint raw spans are necessary numeric conditions only.
            // Slot order, raw decoding and ownership require the future source adapter.
            Need(Identity(source.Slot) && source.RawStart >= 0 && source.RawLength > 0 && (long)source.RawStart + source.RawLength <= int.MaxValue,
                OfdTextFitFailure.InvalidRawMapping, o);
            if (!slots.TryGetValue(source.Slot, out var spans)) { spans = new List<OfdTextFitSourceCluster>(); slots.Add(source.Slot, spans); }
            foreach (var previous in spans)
            {
                budget.Compare();
                Need((long)previous.RawStart + previous.RawLength <= source.RawStart || (long)source.RawStart + source.RawLength <= previous.RawStart,
                    OfdTextFitFailure.InvalidRawMapping, o);
            }
            spans.Add(source);
            // Even a no-op must not carry overflowing coordinates/advance endpoints.
            var extent = Extent(source.X, source.Y, source.Metric, o, -1);
            PageBox(extent, obj.Affine, o, -1);
            if (source.Metric.Ink.HasValue)
                Need(BoxValid(PageBox(InkAt(source.X, source.Y, source.Metric.Ink.Value, o, -1), obj.Affine, o, -1)), OfdTextFitFailure.InvalidGeometry, o);
        }
        Need(offset == obj.Text.Length, OfdTextFitFailure.InvalidScalarMapping, o);
    }

    private static void ValidateMetrics(string text, IReadOnlyList<OfdTextFitMetric> metrics, int o, int r, OfdTextFitBudget budget)
    {
        int offset = 0;
        foreach (var metric in metrics) { budget.Visit(); ValidateMetric(text, metric, ref offset, o, r); }
        Need(offset == text.Length, OfdTextFitFailure.InvalidScalarMapping, o, r);
    }

    private static void ValidateMetric(string text, OfdTextFitMetric metric, ref int offset, int o, int r)
    {
        int scalar = metric.Scalar;
        Need(scalar >= 0 && scalar <= 0x10ffff && !(scalar >= 0xd800 && scalar <= 0xdfff) &&
            scalar >= 0x20 && !(scalar >= 0x7f && scalar <= 0x9f) && scalar != 0x2028 && scalar != 0x2029,
            OfdTextFitFailure.InvalidScalarMapping, o, r);
        int length = scalar > 0xffff ? 2 : 1;
        Need(metric.Utf16Start == offset && metric.Utf16Length == length && (long)offset + length <= text.Length && metric.GlyphId != 0,
            OfdTextFitFailure.InvalidScalarMapping, o, r);
        int decoded = text[offset];
        if (length == 2)
        {
            Need(char.IsHighSurrogate(text[offset]) && char.IsLowSurrogate(text[offset + 1]), OfdTextFitFailure.InvalidScalarMapping, o, r);
            decoded = char.ConvertToUtf32(text[offset], text[offset + 1]);
        }
        Need(decoded == scalar, OfdTextFitFailure.InvalidScalarMapping, o, r);
        Need(Finite(metric.AdvanceX) && Finite(metric.AdvanceY) && (!metric.Ink.HasValue || BoxValid(metric.Ink.Value)), OfdTextFitFailure.InvalidGeometry, o, r);
        // A null box is an explicit inkless metric. Only Unicode spaces are admitted.
        Need(metric.Ink.HasValue || IsSpace(scalar), OfdTextFitFailure.InvalidScalarMapping, o, r);
        offset += length;
    }

    private static bool IsSpace(int scalar) => scalar == 0x20 || scalar == 0xa0 || scalar == 0x1680 ||
        scalar >= 0x2000 && scalar <= 0x200a || scalar == 0x202f || scalar == 0x205f || scalar == 0x3000;

    private static void ValidateProgression(OfdTextFitObject obj, Edit edit, OfdTextFitBudget budget)
    {
        double y = obj.Clusters[0].Y;
        for (int i = 0; i < obj.Clusters.Count; i++)
        {
            budget.Visit(); var source = obj.Clusters[i];
            Need(source.Y == y && source.Metric.AdvanceY == 0 && source.Metric.AdvanceX > 0,
                OfdTextFitFailure.UnsupportedProgression, edit.ObjectIndex, edit.RequestIndex);
            double terminal = Sum(source.X, source.Metric.AdvanceX, edit.ObjectIndex, edit.RequestIndex);
            Need(terminal > source.X, OfdTextFitFailure.UnsupportedProgression, edit.ObjectIndex, edit.RequestIndex);
            if (i + 1 < obj.Clusters.Count) Need(source.X < obj.Clusters[i + 1].X && terminal <= obj.Clusters[i + 1].X,
                OfdTextFitFailure.UnsupportedProgression, edit.ObjectIndex, edit.RequestIndex);
        }
    }

    private static void Place(OfdTextFitObject obj, Edit edit, OfdTextFitBudget budget)
    {
        int o = edit.ObjectIndex, r = edit.RequestIndex;
        var first = obj.Clusters[edit.Request.ScalarStart]; var last = obj.Clusters[edit.End - 1];
        double right = Sum(last.X, last.Metric.AdvanceX, o, r);
        if (edit.End < obj.Clusters.Count) right = Math.Min(right, obj.Clusters[edit.End].X);
        double top = double.PositiveInfinity, bottom = double.NegativeInfinity;
        for (int i = edit.Request.ScalarStart; i < edit.End; i++)
        {
            budget.Visit(); var source = obj.Clusters[i];
            var extent = Extent(source.X, source.Y, source.Metric, o, r);
            Need(extent.Left >= first.X && extent.Right <= right && Contains(obj.Boundary, extent) && Contains(obj.PageClip, PageBox(extent, obj.Affine, o, r)),
                OfdTextFitFailure.OriginalOutsideRegion, o, r);
            if (source.Metric.Ink.HasValue)
            {
                var ink = InkAt(source.X, source.Y, source.Metric.Ink.Value, o, r);
                top = Math.Min(top, ink.Top); bottom = Math.Max(bottom, ink.Bottom);
            }
        }
        if (edit.Request.Metrics.Count == 0) return;
        Need(Finite(top) && Finite(bottom) && top < bottom, OfdTextFitFailure.UnprovedHeight, o, r);
        Need(first.X < right, OfdTextFitFailure.ReplacementOutsideRegion, o, r);
        double x = first.X;
        foreach (var metric in edit.Request.Metrics)
        {
            budget.Visit();
            Need(metric.AdvanceX > 0 && metric.AdvanceY == 0, OfdTextFitFailure.UnsupportedProgression, o, r);
            double terminal = Sum(x, metric.AdvanceX, o, r);
            Need(terminal > x, OfdTextFitFailure.UnsupportedProgression, o, r);
            var extent = Extent(x, first.Y, metric, o, r);
            Need(extent.Left >= first.X && extent.Right <= right && extent.Top >= top && extent.Bottom <= bottom &&
                Contains(obj.Boundary, extent) && Contains(obj.PageClip, PageBox(extent, obj.Affine, o, r)), OfdTextFitFailure.ReplacementOutsideRegion, o, r);
            var corridor = new OfdTextFitBox(x, top, terminal, bottom);
            if (metric.Ink.HasValue)
            {
                var ink = InkAt(x, first.Y, metric.Ink.Value, o, r);
                Need(BoxValid(PageBox(ink, obj.Affine, o, r)), OfdTextFitFailure.InvalidGeometry, o, r);
                corridor = Union(corridor, ink);
            }
            Need(BoxValid(corridor) && BoxValid(PageBox(corridor, obj.Affine, o, r)) && Contains(obj.Boundary, corridor) && Contains(obj.PageClip, PageBox(corridor, obj.Affine, o, r)), OfdTextFitFailure.ReplacementOutsideRegion, o, r);
            edit.Placements.Add(new OfdTextFitPlacement(metric, 0, x, first.Y, null));
            edit.Footprints.Add(new Footprint(obj.Page, PageBox(corridor, obj.Affine, o, r)));
            x = terminal;
        }
    }

    private static List<Footprint> BuildUnchangedObstacles(OfdTextFitInput input, List<Edit> edits, OfdTextFitBudget budget)
    {
        var obstacles = new List<Footprint>();
        int e = 0;
        for (int o = 0; o < input.Objects.Count; o++)
        {
            budget.Check(); var obj = input.Objects[o];
            while (e < edits.Count && edits[e].ObjectIndex < o) { budget.Visit(); e++; }
            int next = e;
            for (int s = 0; s < obj.Clusters.Count; s++)
            {
                budget.Visit();
                while (next < edits.Count && edits[next].ObjectIndex == o && edits[next].End <= s) { budget.Visit(); next++; }
                bool removed = next < edits.Count && edits[next].ObjectIndex == o && edits[next].Changed && edits[next].Request.ScalarStart <= s && s < edits[next].End;
                var source = obj.Clusters[s];
                if (!removed && source.Metric.Ink.HasValue)
                    obstacles.Add(new Footprint(obj.Page, PageBox(InkAt(source.X, source.Y, source.Metric.Ink.Value, o, -1), obj.Affine, o, -1)));
            }
        }
        foreach (var neighbor in input.Neighbors) { budget.Visit(); obstacles.Add(new Footprint(neighbor.Page, neighbor.Box)); }
        return obstacles;
    }

    private static OfdTextFitObjectPlan BuildPlan(OfdTextFitObject obj, List<Edit> edits, int start, int end, OfdTextFitBudget budget)
    {
        int count = obj.Clusters.Count, length = obj.Text.Length;
        for (int e = start; e < end; e++)
        {
            budget.Visit(); var edit = edits[e];
            if (edit.Changed) { count += edit.Request.Metrics.Count - edit.Request.ScalarCount; length += edit.Request.Replacement.Length - edit.Request.Expected.Length; }
        }
        budget.Output(count, length);
        var text = new StringBuilder(length);
        var placements = new List<OfdTextFitPlacement>(count);
        var requests = new List<OfdTextFitRequestPlan>(end - start);
        int scalar = 0;
        for (int e = start; e < end; e++)
        {
            var edit = edits[e];
            while (scalar < edit.Request.ScalarStart) AppendOriginal(obj.Clusters[scalar++], obj.Text, text, placements, budget);
            if (edit.Changed)
            {
                int offset = text.Length;
                foreach (var placement in edit.Placements)
                {
                    budget.Visit(); placements.Add(new OfdTextFitPlacement(placement.Metric, offset + placement.Metric.Utf16Start, placement.X, placement.Y, null));
                }
                text.Append(edit.Request.Replacement); scalar = edit.End;
            }
            else while (scalar < edit.End) AppendOriginal(obj.Clusters[scalar++], obj.Text, text, placements, budget);
            var disposition = !edit.Changed ? OfdTextFitDisposition.Unchanged : count == 0 ? OfdTextFitDisposition.RetainRootWithCanonicalEmptyTextCode : OfdTextFitDisposition.ReplaceScalars;
            requests.Add(new OfdTextFitRequestPlan(edit.RequestIndex, edit.Request.ScalarStart, edit.Request.ScalarCount, disposition));
        }
        while (scalar < obj.Clusters.Count) AppendOriginal(obj.Clusters[scalar++], obj.Text, text, placements, budget);
        budget.Check();
        return new OfdTextFitObjectPlan(obj, text.ToString(), placements, requests);
    }

    private static void AppendOriginal(OfdTextFitSourceCluster source, string original, StringBuilder text, List<OfdTextFitPlacement> placements, OfdTextFitBudget budget)
    {
        budget.Visit(); placements.Add(new OfdTextFitPlacement(source.Metric, text.Length, source.X, source.Y, source));
        text.Append(original, source.Metric.Utf16Start, source.Metric.Utf16Length);
    }

    private static OfdTextFitBox Extent(double x, double y, OfdTextFitMetric metric, int o, int r)
    {
        double tx = Sum(x, metric.AdvanceX, o, r), ty = Sum(y, metric.AdvanceY, o, r);
        var advance = new OfdTextFitBox(Math.Min(x, tx), Math.Min(y, ty), Math.Max(x, tx), Math.Max(y, ty));
        return CheckedBox(metric.Ink.HasValue ? Union(advance, InkAt(x, y, metric.Ink.Value, o, r)) : advance, o, r);
    }
    private static OfdTextFitBox InkAt(double x, double y, OfdTextFitBox ink, int o, int r)
    {
        var box = CheckedBox(new OfdTextFitBox(Sum(x, ink.Left, o, r), Sum(y, ink.Top, o, r), Sum(x, ink.Right, o, r), Sum(y, ink.Bottom, o, r)), o, r);
        Need(BoxValid(box), OfdTextFitFailure.InvalidGeometry, o, r); return box;
    }
    private static OfdTextFitBox PageBox(OfdTextFitBox box, OfdTextFitAffine affine, int o, int r) =>
        CheckedBox(new OfdTextFitBox(Sum(box.Left, affine.Tx, o, r), Sum(box.Top, affine.Ty, o, r), Sum(box.Right, affine.Tx, o, r), Sum(box.Bottom, affine.Ty, o, r)), o, r);
    private static OfdTextFitBox CheckedBox(OfdTextFitBox box, int o, int r)
    {
        Need(Finite(box.Right - box.Left) && Finite(box.Bottom - box.Top) && box.Left <= box.Right && box.Top <= box.Bottom, OfdTextFitFailure.InvalidGeometry, o, r); return box;
    }
    private static double Sum(double a, double b, int o, int r)
    { double value = a + b; Need(Finite(value), OfdTextFitFailure.InvalidGeometry, o, r); return value; }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static bool BoxValid(OfdTextFitBox box) => Finite(box.Left) && Finite(box.Top) && Finite(box.Right) && Finite(box.Bottom) &&
        box.Left < box.Right && box.Top < box.Bottom && Finite(box.Right - box.Left) && Finite(box.Bottom - box.Top);
    private static bool Contains(OfdTextFitBox outer, OfdTextFitBox inner) => outer.Left <= inner.Left && outer.Top <= inner.Top && outer.Right >= inner.Right && outer.Bottom >= inner.Bottom;
    private static bool Intersects(OfdTextFitBox a, OfdTextFitBox b) => a.Left < b.Right && b.Left < a.Right && a.Top < b.Bottom && b.Top < a.Bottom;
    private static OfdTextFitBox Union(OfdTextFitBox a, OfdTextFitBox b) => new OfdTextFitBox(Math.Min(a.Left, b.Left), Math.Min(a.Top, b.Top), Math.Max(a.Right, b.Right), Math.Max(a.Bottom, b.Bottom));
    private static bool Identity(string text) => text.Length > 0 && !string.IsNullOrWhiteSpace(text);
    private static void Need(bool condition, OfdTextFitFailure reason, int o = -1, int r = -1) { if (!condition) throw new Rejection(reason, o, r); }
    private sealed class Rejection : Exception
    {
        internal Rejection(OfdTextFitFailure reason, int o, int r) { Reason = reason; ObjectIndex = o; RequestIndex = r; }
        internal OfdTextFitFailure Reason { get; }
        internal int ObjectIndex { get; }
        internal int RequestIndex { get; }
    }
    private sealed class Edit
    {
        internal Edit(int objectIndex, int requestIndex, OfdTextFitRequest request, bool changed)
        { ObjectIndex = objectIndex; RequestIndex = requestIndex; Request = request; Changed = changed; }
        internal int ObjectIndex { get; }
        internal int RequestIndex { get; }
        internal OfdTextFitRequest Request { get; }
        internal bool Changed { get; }
        internal int End => Request.ScalarStart + Request.ScalarCount;
        internal List<OfdTextFitPlacement> Placements { get; } = new List<OfdTextFitPlacement>();
        internal List<Footprint> Footprints { get; } = new List<Footprint>();
    }
    private readonly struct Footprint
    {
        internal Footprint(int page, OfdTextFitBox box) { Page = page; Box = box; }
        internal int Page { get; }
        internal OfdTextFitBox Box { get; }
    }
}
