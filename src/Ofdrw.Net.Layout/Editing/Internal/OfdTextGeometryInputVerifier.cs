using System.IO;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal readonly struct OfdTextGeometryInputStatistics(int pages, int roots, int slots, int scalars, long hashBytes, int hashes)
{
    internal int Pages { get; } = pages;
    internal int Roots { get; } = roots;
    internal int Slots { get; } = slots;
    internal int Scalars { get; } = scalars;
    internal long FreshFontHashBytes { get; } = hashBytes;
    internal int FreshFontHashOperations { get; } = hashes;
    // No provider is accepted or called by this input comparison.
    internal int ProviderOpens => 0;
    internal int ProviderMeasures => 0;
}

// Identity-bound component comparison only. A relocated output can be admitted
// solely by the future real09 owner/resource receipt; this class never guesses
// Writer paths/IDs or grants package/atomic publication authority.
internal static class OfdTextGeometryInputVerifier
{
    internal static OfdTextGeometryInputStatistics CompareIdentityBoundComponent(OfdTextGeometryExpectation expectation,
        OfdTextReplacementOperation operation, OfdDocumentPackage actual, CancellationToken token)
    {
        var context = operation.Context;
        try
        {
            expectation.AssertBinding(operation, token);
            context.Material.ChargeCopy(checked(2048L + 128L * expectation.Roots.Count + 8L * expectation.Pages.Count), token);
            var selected = expectation.Pages.Select(p => p.Position).ToArray();
            using var capture = OfdKeywordInputCapture.CaptureActual(actual, selected, context, token);
            var budget = capture.Budget;
            Need(capture.Sources.Count == expectation.Roots.Count, "GeometryRootCount");
            foreach (var page in expectation.Pages) ComparePage(page, capture.Admission.CapturePageInput(page.Position), budget);
            var hashes = new Dictionary<byte[], byte[]>(); long hashBytes = 0; var slots = 0; var scalars = 0;
            for (var r = 0; r < expectation.Roots.Count; r++)
            {
                budget.Check(); var expected = expectation.Roots[r]; var source = capture.Sources[r];
                expected.AssertOriginalOwnership(operation.Source ?? throw new InvalidOperationException("RevokedGeometrySource"), budget);
                budget.Field(24);
                Need(expected.Page == source.PageIndex && expected.Layer == source.LayerOrdinal && expected.Element == source.ElementOrdinal &&
                    expected.Path == source.Path && expected.Entry == source.EntryPath && expected.FontId == source.FontId &&
                    expected.Size == source.Size && expected.Weight == source.Weight && expected.Italic == source.Italic &&
                    expected.X == source.X && expected.Y == source.Y && expected.Width == source.Width && expected.Height == source.Height &&
                    expected.PageX == source.PageX && expected.PageY == source.PageY && expected.PageWidth == source.PageWidth && expected.PageHeight == source.PageHeight &&
                    expected.A == source.Matrix[0] && expected.B == source.Matrix[1] && expected.C == source.Matrix[2] && expected.D == source.Matrix[3] &&
                    expected.Tx == source.Matrix[4] && expected.Ty == source.Matrix[5], "GeometryRootOrUnboundRelocation");
                SameInts(expected.Occurrence, source.OccurrencePath, budget); SameInts(expected.Definition, source.DefinitionPath, budget);
                var input = source.PrivateInput ?? throw new InvalidDataException("IncompletePrivateRootInput");
                Same(expected.Metadata.Namespace, input.Namespace, budget);
                Attributes(expected.Metadata.Attributes, input.Attributes, budget);
                SameStrings(expected.Metadata.Children, input.Children, budget);
                Need((expected.Metadata.Fill is null) == (input.Fill is null), "GeometryFillPresence");
                if (expected.Metadata.Fill is not null) Attributes(expected.Metadata.Fill, input.Fill!, budget);
                Need(expected.Slots.Count == source.Runs.Count && input.Slots.Count == source.Runs.Count, "GeometrySlotCount");
                var occurrence = capture.Traversal.Occurrences[r];
                var owner = (OfdTextElement)occurrence.Element;
                var font = capture.Admission.CaptureFontInput(source.FontId, owner);
                CompareFont(expected.Font, font, budget);
                var resource = capture.Admission.ResolveFont(source.FontId).Resource;
                if (resource.Data is null || !actual.PreservedEntries.TryGetValue(font.Entry, out var entry) || entry is null)
                    throw new InvalidDataException("MissingActualGeometryFontBytes");
                Hash(resource.Data, expected.FontBytes, expected.FontSha256); Hash(entry, expected.FontBytes, expected.FontSha256);
                budget.Copy(1024);
                var prepared = new OfdKeywordTextSource
                {
                    X = expected.X, Y = expected.Y, Width = expected.Width, Height = expected.Height,
                    PageX = expected.PageX, PageY = expected.PageY, PageWidth = expected.PageWidth, PageHeight = expected.PageHeight,
                    Matrix = new[] { expected.A, expected.B, expected.C, expected.D, expected.Tx, expected.Ty }
                };
                var rootScalars = 0;
                for (var s = 0; s < expected.Slots.Count; s++)
                {
                    budget.Check(); slots++; var slot = expected.Slots[s]; var run = source.Runs[s];
                    Attributes(slot.Attributes, input.Slots[s].Attributes, budget);
                    Same(slot.Raw, run.RawText, budget); Same(slot.Logical, run.Text, budget);
                    Same(slot.DeltaX, run.DeltaX, budget); Same(slot.DeltaY, run.DeltaY, budget);
                    SameInts(slot.Offsets, run.SourceOffsets, budget);
                    Need(slot.X == run.X && slot.Y == run.Y, "GeometrySlotStart");
                    var dx = OfdStrictTextGeometry.Deltas(run.DeltaX, Math.Max(0, slot.Scalars.Count - 1), budget);
                    var dy = OfdStrictTextGeometry.Deltas(run.DeltaY, Math.Max(0, slot.Scalars.Count - 1), budget);
                    var x = run.X; var y = run.Y; var offset = 0;
                    if (slot.Scalars.Count == 0)
                    {
                        budget.Copy(128); _ = OfdStrictTextGeometry.Transform(source, x, y, false);
                        Need(run.Text.Length == 0, "GeometryEmptyMetricCoverage"); continue;
                    }
                    for (var i = 0; i < slot.Scalars.Count; i++)
                    {
                        budget.Check(); budget.Field(12); budget.Copy(512);
                        var scalar = slot.Scalars[i]; var metric = scalar.Metric;
                        if (i != 0) { x += dx[i - 1]; y += dy[i - 1]; }
                        Need(scalar.Slot == s && scalar.Index == i && scalar.LogicalStart == offset && offset < run.Text.Length &&
                            char.ConvertToUtf32(run.Text, offset) == metric.Scalar && metric.GlyphId != 0 &&
                            metric.Utf16Length == (char.IsHighSurrogate(run.Text[offset]) ? 2 : 1) &&
                            scalar.RawStart == run.SourceOffsets[offset] && scalar.RawLength == run.SourceOffsets[offset + metric.Utf16Length] - run.SourceOffsets[offset] &&
                            x == scalar.X && y == scalar.Y, "GeometryScalarMetricPlacement");
                        // Reuse the exact measured Position transform/clipping
                        // kernel with retained metrics. This is replay, not Measure.
                        Point(OfdStrictTextGeometry.Transform(source, x, y, false), OfdStrictTextGeometry.Transform(prepared, scalar.X, scalar.Y, false));
                        var ax = source.Matrix[0] * metric.AdvanceX + source.Matrix[2] * metric.AdvanceY;
                        var ay = source.Matrix[1] * metric.AdvanceX + source.Matrix[3] * metric.AdvanceY;
                        Need(OfdGlyphBox.Finite(ax) && OfdGlyphBox.Finite(ay) &&
                            ax == expected.A * metric.AdvanceX + expected.C * metric.AdvanceY &&
                            ay == expected.B * metric.AdvanceX + expected.D * metric.AdvanceY, "GeometryAdvance");
                        if (metric.Ink is { } ink)
                        {
                            Corner(ink.Left, ink.Top); Corner(ink.Right, ink.Top); Corner(ink.Right, ink.Bottom); Corner(ink.Left, ink.Bottom);
                        }
                        offset += metric.Utf16Length; scalars++; rootScalars++;
                        void Corner(double cx, double cy) => Point(OfdStrictTextGeometry.Transform(source, x + cx, y + cy, true),
                            OfdStrictTextGeometry.Transform(prepared, scalar.X + cx, scalar.Y + cy, true));
                    }
                    Need(offset == run.Text.Length, "GeometryScalarCoverage");
                }
                Need((rootScalars == 0 || !expected.IdentityOnly) && source.AllRunsEmpty == (rootScalars == 0), "GeometryIdentityOnlyQualification");
            }
            expectation.AssertBinding(operation, token);
            return new OfdTextGeometryInputStatistics(selected.Length, expectation.Roots.Count, slots, scalars, hashBytes, hashes.Count);

            void Hash(byte[] current, int length, string baseline)
            {
                budget.Check(); Need(current.Length == length, "GeometryFontLength");
                if (!hashes.TryGetValue(current, out var digest))
                {
                    context.Material.ChargeCopy(128, token);
                    digest = context.Material.HashBytes(current, token); hashes.Add(current, digest); hashBytes = checked(hashBytes + current.LongLength);
                }
                budget.Characters(baseline.Length); Need(baseline.Length == 64, "GeometryFontIdentity");
                for (var i = 0; i < digest.Length; i++)
                    Need(Hex(baseline[2 * i]) == digest[i] / 16 && Hex(baseline[2 * i + 1]) == digest[i] % 16, "GeometryFontIdentity");
            }
        }
        catch { context.Fail(); throw; }
    }
    private static int Hex(char c) => c is >= '0' and <= '9' ? c - '0' : c is >= 'a' and <= 'f' ? c - 'a' + 10 : -1;
    private static void ComparePage(OfdKeywordPageInput a, OfdKeywordPageInput b, OfdKeywordLocatorBudget budget)
    {
        budget.Field(4); Need(a.Position == b.Position && a.OwnArea == b.OwnArea && a.Layers.Count == b.Layers.Count, "GeometryPageAuthority");
        Same(a.Id, b.Id, budget); Same(a.Entry, b.Entry, budget); Same(a.Namespace, b.Namespace, budget); Same(a.PhysicalBox, b.PhysicalBox, budget);
        Attributes(a.Reference, b.Reference, budget);
        for (var i = 0; i < a.Layers.Count; i++)
        { budget.Field(); Need(a.Layers[i].Objects == b.Layers[i].Objects, "GeometryLayerOrder");
          Same(a.Layers[i].Namespace, b.Layers[i].Namespace, budget); Attributes(a.Layers[i].Attributes, b.Layers[i].Attributes, budget); }
    }
    private static void CompareFont(OfdKeywordFontInput a, OfdKeywordFontInput b, OfdKeywordLocatorBudget budget)
    {
        budget.Field(3); Need(a.Bold == b.Bold && a.Italic == b.Italic && a.Ordinal == b.Ordinal, "GeometryFontStyleOrDeclarationOrder");
        Same(a.OriginEntry, b.OriginEntry, budget); Same(a.DeclarationKind, b.DeclarationKind, budget); Same(a.DeclarationValue, b.DeclarationValue, budget);
        Same(a.Namespace, b.Namespace, budget); Same(a.ResourceEntry, b.ResourceEntry, budget); Same(a.Entry, b.Entry, budget);
        Same(a.FontFile, b.FontFile, budget); Same(a.FontName, b.FontName, budget); Same(a.FamilyName, b.FamilyName, budget); Same(a.Charset, b.Charset, budget);
        Attributes(a.ResourceAttributes, b.ResourceAttributes, budget); Attributes(a.Attributes, b.Attributes, budget);
    }
    private static void Attributes(IReadOnlyList<OfdKeywordInputAttribute> a, IReadOnlyList<OfdKeywordInputAttribute> b, OfdKeywordLocatorBudget budget)
    {
        budget.Field(); Need(a.Count == b.Count, "GeometryAttributePresence");
        for (var i = 0; i < a.Count; i++)
        { budget.Field(); Same(a[i].Namespace, b[i].Namespace, budget); Same(a[i].Name, b[i].Name, budget); Same(a[i].Value, b[i].Value, budget); }
    }
    private static void SameStrings(IReadOnlyList<string> a, IReadOnlyList<string> b, OfdKeywordLocatorBudget budget)
    { budget.Field(); Need(a.Count == b.Count, "GeometryChildOrder"); for (var i = 0; i < a.Count; i++) Same(a[i], b[i], budget); }
    private static void SameInts(IReadOnlyList<int> a, IReadOnlyList<int> b, OfdKeywordLocatorBudget budget)
    { budget.Field(); Need(a.Count == b.Count, "GeometryOffsetCount"); for (var i = 0; i < a.Count; i++) { budget.Field(); Need(a[i] == b[i], "GeometryOffset"); } }
    private static void Same(string? a, string? b, OfdKeywordLocatorBudget budget)
    { budget.Characters(Math.Max(a?.Length ?? 0, b?.Length ?? 0)); Need(string.Equals(a, b, StringComparison.Ordinal), "GeometryRawInputDifference"); }
    private static void Point(OfdKeywordPoint a, OfdKeywordPoint b) => Need(a.X == b.X && a.Y == b.Y, "GeometryProjectedPlacement");
    private static void Need(bool match, string reason) { if (!match) throw new InvalidDataException(reason); }
}
