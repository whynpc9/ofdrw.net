using System.IO;
using System.Globalization;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Layout.Editing.Internal;

// Frozen only from the same live analysis and real compiler. The capsule proves
// prepared inputs/placements, never a fresh measurement or successful save.
internal sealed class OfdTextGeometryExpectation : IDisposable
{
    internal sealed class Scalar
    {
        internal Scalar(OfdTextFitMetric metric, int slot, int index, int logicalStart, int rawStart, int rawLength,
            double x, double y, bool inserted, int originalRawStart, int originalRawLength)
        { Metric = metric; Slot = slot; Index = index; LogicalStart = logicalStart; RawStart = rawStart; RawLength = rawLength;
          X = x; Y = y; Inserted = inserted; OriginalRawStart = originalRawStart; OriginalRawLength = originalRawLength; }
        internal OfdTextFitMetric Metric { get; }
        internal int Slot { get; }
        internal int Index { get; }
        internal int LogicalStart { get; }
        internal int RawStart { get; }
        internal int RawLength { get; }
        internal double X { get; }
        internal double Y { get; }
        internal bool Inserted { get; }
        internal int OriginalRawStart { get; }
        internal int OriginalRawLength { get; }
    }
    internal sealed class Slot
    {
        internal Slot(string raw, string logical, int[] offsets, IReadOnlyList<OfdKeywordInputAttribute> attributes,
            double x, double y, string? dx, string? dy, Scalar[] scalars)
        { Raw = raw; Logical = logical; Offsets = Array.AsReadOnly(offsets); Attributes = attributes; X = x; Y = y;
          DeltaX = dx; DeltaY = dy; Scalars = Array.AsReadOnly(scalars); }
        internal string Raw { get; }
        internal string Logical { get; }
        internal IReadOnlyList<int> Offsets { get; }
        internal IReadOnlyList<OfdKeywordInputAttribute> Attributes { get; }
        internal double X { get; }
        internal double Y { get; }
        internal string? DeltaX { get; }
        internal string? DeltaY { get; }
        internal IReadOnlyList<Scalar> Scalars { get; }
    }
    internal sealed class Root
    {
        private readonly OfdTextElement originalOwner;
        private readonly OfdPage originalPage;
        private readonly OfdFontResource originalFont;
        internal Root(OfdKeywordTextSource source, OfdKeywordRootInput metadata, OfdKeywordFontInput font,
            OfdGlyphFontSnapshot identity, Slot[] slots, bool measuredInputs, OfdTextElement owner, OfdPage page, OfdFontResource fontOwner)
        {
            originalOwner = owner; originalPage = page; originalFont = fontOwner;
            Page = source.PageIndex; Layer = source.LayerOrdinal; Element = source.ElementOrdinal; Path = source.Path; Entry = source.EntryPath!;
            FontId = source.FontId; Size = source.Size; Weight = source.Weight; Italic = source.Italic;
            X = source.X; Y = source.Y; Width = source.Width; Height = source.Height;
            PageX = source.PageX; PageY = source.PageY; PageWidth = source.PageWidth; PageHeight = source.PageHeight;
            A = source.Matrix[0]; B = source.Matrix[1]; C = source.Matrix[2]; D = source.Matrix[3]; Tx = source.Matrix[4]; Ty = source.Matrix[5];
            Metadata = metadata; Font = font; FontSha256 = identity.ContentSha256; FontBytes = identity.ByteCount;
            Occurrence = Array.AsReadOnly(source.OccurrencePath.ToArray()); Definition = Array.AsReadOnly(source.DefinitionPath.ToArray());
            Slots = Array.AsReadOnly(slots); IdentityOnly = !measuredInputs;
        }
        internal int Page { get; } internal int Layer { get; } internal int Element { get; }
        internal string Path { get; } internal string Entry { get; } internal string FontId { get; }
        internal double Size { get; } internal int Weight { get; } internal bool Italic { get; }
        internal double X { get; } internal double Y { get; } internal double Width { get; } internal double Height { get; }
        internal double PageX { get; } internal double PageY { get; } internal double PageWidth { get; } internal double PageHeight { get; }
        internal double A { get; } internal double B { get; } internal double C { get; } internal double D { get; } internal double Tx { get; } internal double Ty { get; }
        internal IReadOnlyList<int> Occurrence { get; } internal IReadOnlyList<int> Definition { get; }
        internal OfdKeywordRootInput Metadata { get; } internal OfdKeywordFontInput Font { get; }
        internal string FontSha256 { get; } internal int FontBytes { get; } internal bool IdentityOnly { get; }
        internal IReadOnlyList<Slot> Slots { get; }
        // Retain real original handles for the later approved09 mapping seam,
        // without exposing a model/array getter or caller-issued ID authority.
        internal void AssertOriginalOwnership(OfdDocumentPackage source, OfdKeywordLocatorBudget budget)
        {
            budget.Field(3);
            if (Page >= source.Pages.Count || !ReferenceEquals(source.Pages[Page], originalPage) || Element >= originalPage.Elements.Count ||
                !ReferenceEquals(originalPage.Elements[Element], originalOwner)) throw new InvalidDataException("GeometryOriginalOwnerBinding");
            var found = false;
            foreach (var font in source.Fonts) { budget.Field(); if (ReferenceEquals(font, originalFont)) { found = true; break; } }
            if (!found) throw new InvalidDataException("GeometryOriginalFontBinding");
        }
    }
    private readonly OfdTextReplacementOperation operation;
    private readonly OfdControlledEditContext context;
    private readonly Root[] roots;
    private readonly OfdKeywordPageInput[] pages;
    private int disposed;
    private OfdTextGeometryExpectation(OfdTextReplacementOperation operation, string version, Root[] roots, OfdKeywordPageInput[] pages)
    { this.operation = operation; context = operation.Context; SourceVersion = version; this.roots = roots; this.pages = pages;
      Roots = Array.AsReadOnly(roots); Pages = Array.AsReadOnly(pages); }
    internal string SourceVersion { get; }
    // This component has no general deterministic provider claim. Production
    // qualification is the audited private CLI composition using this preset.
    internal const string ProductionProfile = "OfdTrustedEmbeddedFontGeometryProvider/SixLabors.Fonts-1.0.1/Regular-simple/v1";
    internal IReadOnlyList<Root> Roots { get; }
    internal IReadOnlyList<OfdKeywordPageInput> Pages { get; }
    internal void AssertBinding(OfdTextReplacementOperation actual, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); context.AssertActive();
        if (disposed != 0 || !ReferenceEquals(actual, operation) || !ReferenceEquals(actual.Context, context))
            throw new InvalidOperationException("ForeignOrRevokedGeometryExpectation");
    }
    public void Dispose() { if (Interlocked.Exchange(ref disposed, 1) != 0) return; Array.Clear(roots, 0, roots.Length); Array.Clear(pages, 0, pages.Length); }

    internal sealed class Builder : IDisposable
    {
        private sealed class Pending(OfdKeywordTextSource source, OfdKeywordRootInput metadata, OfdKeywordFontInput font,
            OfdGlyphFontSnapshot identity, Slot[] slots, OfdTextElement owner, OfdPage page, OfdFontResource fontOwner)
        {
            internal readonly OfdTextElement Owner = owner;
            internal readonly OfdPage Page = page;
            internal readonly OfdFontResource FontOwner = fontOwner;
            internal readonly OfdKeywordTextSource Source = source;
            internal readonly OfdKeywordRootInput Metadata = metadata;
            internal readonly OfdKeywordFontInput Font = font;
            internal readonly OfdGlyphFontSnapshot Identity = identity;
            internal Slot[] Slots = slots;
            internal readonly HashSet<OfdTextFitMetric> Metrics = new();
        }
        private readonly OfdTextReplacementOperation operation;
        private readonly string version;
        private readonly OfdKeywordLocatorBudget budget;
        private readonly Dictionary<OfdTextElement, Pending> roots = new();
        private readonly Dictionary<int, OfdKeywordPageInput> pages = new();
        private bool closed;
        private Root[]? frozenRoots;
        private OfdKeywordPageInput[]? frozenPages;
        private Builder(OfdTextReplacementOperation operation, OfdKeywordAnalysisCapture capture, int pageCount)
        { this.operation = operation; version = capture.SourceVersion; budget = capture.Budget; }
        internal static Builder Create(OfdTextReplacementOperation operation, OfdKeywordAnalysisCapture capture, int pageCount)
        {
            capture.Budget.Copy(checked(4096L + 512L * pageCount)); capture.Budget.Retain(capture.SourceVersion);
            return new Builder(operation, capture, pageCount);
        }
        internal void CaptureRoot(OfdKeywordAnalysisCapture capture, OfdGraphicOccurrence occurrence, OfdTextReplacementRoot root)
        {
            budget.Check(); if (closed || roots.ContainsKey(root.Owner)) throw new InvalidOperationException("GeometryExpectationOwnerAlias");
            var source = root.Source;
            // These exact strings were created by admission/P0, not retained
            // by the original model inventory. Admit them before further root
            // retention; shared snapshot digests deduplicate by object identity.
            budget.Retain(source.Path); budget.Retain(source.Font?.ContentSha256);
            var strict = OfdStrictTextGeometry.CapturePrivate(occurrence, source.LayerOrdinal, source.Path, budget,
                checked(source.Runs.Sum(r => r.RawText.Length)), capture.Admission.HasOriginalEnvelope);
            var font = capture.Admission.CaptureFontInput(source.FontId, root.Owner);
            if (!pages.ContainsKey(source.PageIndex)) pages.Add(source.PageIndex, capture.Admission.CapturePageInput(source.PageIndex));
            budget.Copy(checked(2048L + 512L * source.Runs.Count + 512L * root.Fit.Clusters.Count + 16L * source.OccurrencePath.Length));
            budget.Copy(checked(1024L + 256L * source.Runs.Count + 128L * capture.Positions[source].Count));
            var positionsByRun = new Dictionary<OfdKeywordSourceRun, List<OfdPositionedCluster>>(source.Runs.Count);
            foreach (var run in source.Runs) { budget.Field(); positionsByRun.Add(run, new List<OfdPositionedCluster>()); }
            foreach (var position in capture.Positions[source]) { budget.Correlate(); positionsByRun[position.Run].Add(position); }
            var slots = new Slot[source.Runs.Count]; var logicalOffset = 0; var metricCursor = 0;
            for (var r = 0; r < slots.Length; r++)
            {
                var run = source.Runs[r]; var positions = positionsByRun[run];
                budget.Copy(checked(512L + 4L * run.SourceOffsets.Length + 512L * positions.Count));
                var metrics = new Scalar[positions.Count];
                for (var i = 0; i < positions.Count; i++)
                {
                    budget.Field(); var position = positions[i]; var original = root.Fit.Clusters[metricCursor++];
                    var start = position.Geometry.Utf16Start; var length = position.Geometry.Utf16Length;
                    if (original.Metric.Utf16Start != logicalOffset + start || !SameMetric(original.Metric, position.Geometry))
                        throw new InvalidDataException("GeometryMetricAssociation");
                    metrics[i] = new Scalar(original.Metric, r, i, start, run.SourceOffsets[start],
                        run.SourceOffsets[start + length] - run.SourceOffsets[start], position.RawX, position.RawY, false,
                        original.RawStart, original.RawLength);
                }
                budget.Retain(run.RawText); budget.Retain(run.Text); budget.Retain(run.DeltaX); budget.Retain(run.DeltaY);
                slots[r] = new Slot(run.RawText, run.Text, run.SourceOffsets.ToArray(), strict.PrivateInput!.Slots[r].Attributes,
                    run.X, run.Y, run.DeltaX, run.DeltaY, metrics); logicalOffset += run.Text.Length;
            }
            if (metricCursor != root.Fit.Clusters.Count || source.Font is null) throw new InvalidDataException("IncompleteGeometryMetrics");
            var pending = new Pending(strict, strict.PrivateInput!, font, source.Font, slots, root.Owner, occurrence.Page, capture.Admission.ResolveFont(source.FontId).Resource);
            foreach (var scalar in slots.SelectMany(s => s.Scalars)) pending.Metrics.Add(scalar.Metric);
            roots.Add(root.Owner, pending);
        }
        internal void RecordReplacementMetrics(OfdTextReplacementRoot root, OfdGlyphGeometry geometry, OfdTextFitMetric[] metrics)
        {
            budget.Check(); if (closed || !roots.TryGetValue(root.Owner, out var pending) || geometry.Clusters.Count != metrics.Length)
                throw new InvalidOperationException("UnboundReplacementMetrics");
            budget.Copy(checked(512L + 128L * metrics.Length));
            for (var i = 0; i < metrics.Length; i++)
            {
                budget.Field(); if (!SameMetric(metrics[i], geometry.Clusters[i])) throw new InvalidDataException("GeometryMetricAssociation");
                pending.Metrics.Add(metrics[i]);
            }
        }
        private static bool SameMetric(OfdTextFitMetric metric, OfdGlyphClusterGeometry geometry)
        {
            var ink = geometry.FontBoundsMillimeters;
            return metric.GlyphId == geometry.GlyphId && metric.Utf16Length == geometry.Utf16Length &&
                metric.AdvanceX == geometry.AdvanceXMillimeters && metric.AdvanceY == geometry.AdvanceYMillimeters &&
                metric.Ink.HasValue == (ink is not null) && (ink is null ||
                metric.Ink!.Value.Left == ink.Left && metric.Ink.Value.Top == ink.Top && metric.Ink.Value.Right == ink.Right && metric.Ink.Value.Bottom == ink.Bottom);
        }
        internal void RecordCompiled(OfdTextReplacementRoot root, OfdTextSlotEdit[] edits, List<OfdTextFitPlacement>[] actual)
        {
            budget.Check(); if (closed || !roots.TryGetValue(root.Owner, out var pending)) throw new InvalidOperationException("UnboundCompiledGeometry");
            if (root.Plan.Requests.All(r => r.Disposition == OfdTextFitDisposition.Unchanged)) return;
            budget.Copy(checked(1024L + 512L * edits.Length + 512L * root.Plan.Placements.Count));
            var slots = new Slot[edits.Length];
            for (var r = 0; r < edits.Length; r++)
            {
                budget.Check(); var edit = edits[r]; var placements = actual[r];
                var decoded = OfdStrictTextGeometry.Decode(edit.RawOfdText, budget);
                var scalars = new Scalar[placements.Count]; var offset = 0;
                for (var i = 0; i < placements.Count; i++)
                {
                    budget.Field(); var placement = placements[i]; var length = placement.Metric.Utf16Length;
                    if (!pending.Metrics.Contains(placement.Metric)) throw new InvalidDataException("UnmeasuredCompiledMetric");
                    if (offset + length > decoded.Text.Length || char.ConvertToUtf32(decoded.Text, offset) != placement.Metric.Scalar)
                        throw new InvalidDataException("CompiledGeometryMetricAssociation");
                    scalars[i] = new Scalar(placement.Metric, r, i, offset, decoded.Offsets[offset],
                        decoded.Offsets[offset + length] - decoded.Offsets[offset], placement.X - pending.Source.Matrix[4],
                        placement.Y - pending.Source.Matrix[5], placement.Source is null,
                        placement.Source?.RawStart ?? -1, placement.Source?.RawLength ?? 0); offset += length;
                }
                if (offset != decoded.Text.Length) throw new InvalidDataException("CompiledGeometryCoverage");
                budget.Copy(1024); var x = edit.X.ToString("R", CultureInfo.InvariantCulture); var y = edit.Y.ToString("R", CultureInfo.InvariantCulture);
                budget.Retain(x); budget.Retain(y); budget.Retain(decoded.Text); budget.Retain(edit.RawOfdText); budget.Retain(edit.DeltaX); budget.Retain(edit.DeltaY);
                var attrs = new List<OfdKeywordInputAttribute>(4);
                // The approved fixed-slot patch preserves existing attribute
                // positions and appends newly explicit axes/deltas. Preserve
                // that original order using actual compiler values only.
                foreach (var original in pending.Metadata.Slots[r].Attributes)
                {
                    var value = Value(original.Name);
                    if (value is not null) attrs.Add(new("", original.Name, value));
                }
                foreach (var name in new[] { "X", "Y", "DeltaX", "DeltaY" })
                    if (!attrs.Any(a => a.Name == name) && Value(name) is { } value) attrs.Add(new("", name, value));
                string? Value(string name) => name switch
                { "X" => x, "Y" => y, "DeltaX" => edit.DeltaX, "DeltaY" => edit.DeltaY, _ => throw new InvalidDataException("UnexpectedCompiledSlotAttribute") };
                slots[r] = new Slot(edit.RawOfdText, decoded.Text, decoded.Offsets, Array.AsReadOnly(attrs.ToArray()), edit.X, edit.Y,
                    edit.DeltaX, edit.DeltaY, scalars);
            }
            pending.Slots = slots;
        }
        internal void FreezeBeforeProviderClosure()
        {
            operation.Context.AssertActive(); budget.Check();
            if (closed || frozenRoots is not null) throw new InvalidOperationException("GeometryExpectationAlreadyFrozen");
            // Immutable records and final wrapper shells are prepaid while the
            // live analysis budget still exists. Seal later performs no scan or
            // provider callback and cannot use that disposed KeywordOperation.
            budget.Copy(checked(4096L + 2048L * roots.Count + 32L * pages.Count));
            frozenRoots = roots.Values.OrderBy(p => p.Source.PageIndex).ThenBy(p => p.Source.LayerOrdinal).ThenBy(p => p.Source.ElementOrdinal)
                .Select(p => new Root(p.Source, p.Metadata, p.Font, p.Identity, p.Slots, p.Metrics.Count != 0, p.Owner, p.Page, p.FontOwner)).ToArray();
            frozenPages = pages.Values.OrderBy(p => p.Position).ToArray(); roots.Clear(); pages.Clear();
        }
        internal OfdTextGeometryExpectation SealAfterProviderClosure()
        {
            operation.Context.AssertActive();
            if (closed || frozenRoots is null || frozenPages is null) throw new InvalidOperationException("GeometryExpectationNotFrozen");
            closed = true;
            var result = new OfdTextGeometryExpectation(operation, version, frozenRoots, frozenPages);
            frozenRoots = null; frozenPages = null; return result;
        }
        public void Dispose()
        {
            closed = true; roots.Clear(); pages.Clear();
            if (frozenRoots is not null) Array.Clear(frozenRoots, 0, frozenRoots.Length);
            if (frozenPages is not null) Array.Clear(frozenPages, 0, frozenPages.Length);
            frozenRoots = null; frozenPages = null;
        }
    }
}
