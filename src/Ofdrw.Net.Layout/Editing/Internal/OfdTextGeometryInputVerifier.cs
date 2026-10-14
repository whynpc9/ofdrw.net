using System.IO;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
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
        => Compare(expectation, operation, actual, token, null);

    internal static OfdTextGeometryInputStatistics CompareReceiptBoundOwned(OfdTextGeometryExpectation expectation,
        OfdTextReplacementOperation operation, OfdOwnedGeometryExpectation issuer, OfdOwnedCompareScope scope,
        OfdDocumentPackage actual, OfdOwnedTextRootKey[] keys, OfdReadbackRootAddress[] addresses, CancellationToken token)
    {
        operation.Context.Material.ChargeCopy(checked(4096L + 256L * keys.Length), token);
        var owned = new OwnedMapping(expectation, issuer, scope, actual, keys, addresses, operation.Context, token);
        return Compare(expectation, operation, actual, token, owned);
    }
    private static OfdTextGeometryInputStatistics Compare(OfdTextGeometryExpectation expectation,
        OfdTextReplacementOperation operation, OfdDocumentPackage actual, CancellationToken token, OwnedMapping? owned)
    {
        var context = operation.Context;
        try
        {
            expectation.AssertBinding(operation, token);
            owned?.BindOriginal(operation.Source ?? throw new InvalidOperationException("RevokedGeometrySource"));
            context.Material.ChargeCopy(checked(2048L + 128L * expectation.Roots.Count + 8L * expectation.Pages.Count), token);
            var selected = owned is null ? expectation.Pages.Select(p => p.Position).ToArray() : owned.SelectedPages();
            using var capture = OfdKeywordInputCapture.CaptureActual(actual, selected, context, token);
            var budget = capture.Budget;
            Need(capture.Sources.Count == expectation.Roots.Count, "GeometryRootCount");
            foreach (var page in expectation.Pages)
            {
                if (owned is null) ComparePage(page, capture.Admission.CapturePageInput(page.Position), budget);
                else owned.Page(page, capture.Admission, budget);
            }
            var hashes = new Dictionary<byte[], byte[]>(); long hashBytes = 0; var slots = 0; var scalars = 0;
            for (var r = 0; r < expectation.Roots.Count; r++)
            {
                budget.Check(); var expected = expectation.Roots[r];
                var sourceIndex = owned is null ? r : owned.SourceIndex(r, capture, budget);
                var source = capture.Sources[sourceIndex];
                expected.AssertOriginalOwnership(operation.Source ?? throw new InvalidOperationException("RevokedGeometrySource"), budget);
                budget.Field(24);
                if (owned is null)
                    Need(expected.Page == source.PageIndex && expected.Layer == source.LayerOrdinal && expected.Element == source.ElementOrdinal &&
                        expected.Path == source.Path && expected.Entry == source.EntryPath && expected.FontId == source.FontId, "GeometryRootOrUnboundRelocation");
                else owned.Root(r, source, budget);
                Need(expected.Size == source.Size && expected.Weight == source.Weight && expected.Italic == source.Italic &&
                    expected.X == source.X && expected.Y == source.Y && expected.Width == source.Width && expected.Height == source.Height &&
                    expected.PageX == source.PageX && expected.PageY == source.PageY && expected.PageWidth == source.PageWidth && expected.PageHeight == source.PageHeight &&
                    expected.A == source.Matrix[0] && expected.B == source.Matrix[1] && expected.C == source.Matrix[2] && expected.D == source.Matrix[3] &&
                    expected.Tx == source.Matrix[4] && expected.Ty == source.Matrix[5], "GeometryRootOrUnboundRelocation");
                if (owned is null) { SameInts(expected.Occurrence, source.OccurrencePath, budget); SameInts(expected.Definition, source.DefinitionPath, budget); }
                else Need(source.OccurrencePath.Length == 4 && source.DefinitionPath.Length == 2, "OwnedRootTopology");
                var input = source.PrivateInput ?? throw new InvalidDataException("IncompletePrivateRootInput");
                Same(expected.Metadata.Namespace, input.Namespace, budget);
                if (owned is null) Attributes(expected.Metadata.Attributes, input.Attributes, budget);
                else owned.RootAttributes(r, expected.Metadata.Attributes, input.Attributes, budget);
                SameStrings(expected.Metadata.Children, input.Children, budget);
                Need((expected.Metadata.Fill is null) == (input.Fill is null), "GeometryFillPresence");
                if (expected.Metadata.Fill is not null) Attributes(expected.Metadata.Fill, input.Fill!, budget);
                Need(expected.Slots.Count == source.Runs.Count && input.Slots.Count == source.Runs.Count, "GeometrySlotCount");
                var occurrence = capture.Traversal.Occurrences[sourceIndex];
                var owner = (OfdTextElement)occurrence.Element;
                var font = capture.Admission.CaptureFontInput(source.FontId, owner);
                if (owned is null) CompareFont(expected.Font, font, budget);
                else owned.Font(r, expected, font, budget);
                var resource = capture.Admission.ResolveFont(source.FontId).Resource;
                if (resource.Data is null || !actual.PreservedEntries.TryGetValue(font.Entry, out var entry) || entry is null)
                    throw new InvalidDataException("MissingActualGeometryFontBytes");
                if (owned is null) { Hash(resource.Data, expected.FontBytes, expected.FontSha256); Hash(entry, expected.FontBytes, expected.FontSha256); }
                else owned.CurrentFont(r, expected, resource, budget);
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
    // Concrete private mapping over actual authenticated scalar receipt records.
    // No delegate, provider, cached hash flag or candidate-derived expectation.
    private sealed class OwnedMapping
    {
        private readonly OfdTextGeometryExpectation expected;
        private readonly OfdOwnedGeometryExpectation issuer;
        private readonly OfdOwnedCompareScope scope;
        private readonly OfdDocumentPackage actual;
        private readonly OfdReadbackRootAddress[] rows;
        private readonly HashSet<int> usedSources = new();
        private readonly Dictionary<int, OfdKeywordPageInput> currentPages = new();
        internal OwnedMapping(OfdTextGeometryExpectation expected, OfdOwnedGeometryExpectation issuer, OfdOwnedCompareScope scope,
            OfdDocumentPackage actual, OfdOwnedTextRootKey[] keys, OfdReadbackRootAddress[] addresses, OfdControlledEditContext context, CancellationToken token)
        {
            this.expected = expected; this.issuer = issuer; this.scope = scope; this.actual = actual;
            context.Material.ChargeWork(checked(16L * expected.Roots.Count * expected.Pages.Count), token);
            Need(keys.Length == addresses.Length && keys.Length == expected.Roots.Count, "OwnedAddressCoverage");
            rows = new OfdReadbackRootAddress[keys.Length]; var used = new HashSet<int>();
            for (var i = 0; i < keys.Length; i++)
            {
                for (var j = 0; j < addresses.Length; j++)
                {
                    context.Material.ChargeWork(16, token);
                    if (addresses[j] is not null && ReferenceEquals(keys[i], addresses[j].Key))
                    { Need(rows[i] is null && used.Add(j), "OwnedAddressAlias"); rows[i] = addresses[j]; }
                }
                Need(rows[i] is not null, "MissingOwnedAddress");
            }
        }
        private OfdReadbackRootAddress PageRow(int originalPage)
        {
            for (var i = 0; i < expected.Roots.Count; i++) if (expected.Roots[i].Page == originalPage) return rows[i];
            throw new InvalidDataException("MissingOwnedPage");
        }
        internal int[] SelectedPages() => expected.Pages.Select(p => PageRow(p.Position).PageIndex).ToArray();
        internal void Page(OfdKeywordPageInput original, OfdKeywordSourceAdmission admission, OfdKeywordLocatorBudget budget)
        {
            var row = PageRow(original.Position); var current = admission.CapturePageInput(row.PageIndex);
            Need(!currentPages.ContainsKey(row.PageIndex), "OwnedPageAlias"); currentPages.Add(row.PageIndex, current);
            budget.Field(4); Need(original.OwnArea == current.OwnArea && original.Layers.Count == current.Layers.Count, "GeometryPageAuthority");
            Same(row.OutputPageId, current.Id, budget); Same(row.EntryPath, current.Entry, budget);
            Same(original.Namespace, current.Namespace, budget); Same(original.PhysicalBox, current.PhysicalBox, budget);
            ExpandedName(row.SourcePageReferenceExpandedName, original.Namespace, "Page", budget);
            ExpandedName(row.OutputPageReferenceExpandedName, current.Namespace, "Page", budget);
            Same(row.SourcePageReferenceOwnerEntry, expectedDocumentEntry(), budget);
            Same(row.OutputPageReferenceOwnerEntry, actual.DocumentEntryPath, budget);
            Same(row.SourcePageBaseLocRaw, Attribute(original.Reference, "BaseLoc"), budget);
            Same(row.OutputPageBaseLocRaw, Attribute(current.Reference, "BaseLoc"), budget);
            MappedAttributes(original.Reference, current.Reference, budget, "ID", row.OutputPageId, "BaseLoc", row.OutputPageBaseLocRaw);
            for (var l = 0; l < original.Layers.Count; l++)
            {
                budget.Field(); Need(original.Layers[l].Objects == current.Layers[l].Objects, "GeometryLayerOrder");
                Same(original.Layers[l].Namespace, current.Layers[l].Namespace, budget);
                string? layerId = null; var hasRow = false;
                for (var r = 0; r < expected.Roots.Count; r++)
                    if (expected.Roots[r].Page == original.Position && expected.Roots[r].Layer == l)
                    {
                        var other = rows[r];
                        Need(other.PageIndex == row.PageIndex && other.LayerOrdinal == l, "OwnedLayerRelocationUnsupported");
                        Same(other.OutputPageId, row.OutputPageId, budget); Same(other.EntryPath, row.EntryPath, budget);
                        if (hasRow) Same(layerId, other.OutputLayerId, budget); else { hasRow = true; layerId = other.OutputLayerId; }
                    }
                if (hasRow) MappedAttributes(original.Layers[l].Attributes, current.Layers[l].Attributes, budget, "ID", layerId);
                else Attributes(original.Layers[l].Attributes, current.Layers[l].Attributes, budget);
            }
            string? expectedDocumentEntry() => expectedOperationSource?.DocumentEntryPath;
        }
        // Original owner document path is supplied by the existing private operation,
        // not guessed by stripping an output entry prefix.
        private OfdDocumentPackage? expectedOperationSource;
        internal void BindOriginal(OfdDocumentPackage original) => expectedOperationSource = original;
        internal int SourceIndex(int r, OfdKeywordInputCapture capture, OfdKeywordLocatorBudget budget)
        {
            var row = rows[r]; var found = -1;
            for (var i = 0; i < capture.Sources.Count; i++)
            {
                budget.Field(3); var source = capture.Sources[i];
                if (source.PageIndex != row.PageIndex || source.LayerOrdinal != row.LayerOrdinal) continue;
                Need(currentPages.TryGetValue(source.PageIndex, out var page), "MissingOwnedPageCapture"); var local = source.ElementOrdinal;
                for (var l = 0; l < source.LayerOrdinal; l++) { budget.Field(); local -= page!.Layers[l].Objects; }
                if (local == row.RootOrdinal) { Need(found == -1, "OwnedRootAddressAlias"); found = i; }
            }
            Need(found != -1 && usedSources.Add(found), "OwnedRootCoverage"); return found;
        }
        internal void Root(int r, OfdKeywordTextSource source, OfdKeywordLocatorBudget budget)
        {
            var row = rows[r]; Same(row.EntryPath, source.EntryPath, budget); Same(row.OutputFontResourceId, source.FontId, budget);
        }
        internal void RootAttributes(int r, IReadOnlyList<OfdKeywordInputAttribute> original,
            IReadOnlyList<OfdKeywordInputAttribute> current, OfdKeywordLocatorBudget budget)
            => MappedAttributes(original, current, budget, "ID", rows[r].OutputObjectId, "Font", rows[r].OutputFontResourceId);
        internal void Font(int r, OfdTextGeometryExpectation.Root original, OfdKeywordFontInput current, OfdKeywordLocatorBudget budget)
        {
            var a = original.Font; var d = rows[r].FontDeclaration;
            Same(d.SourceFontResourceId, original.FontId, budget); Same(d.OutputFontResourceId, rows[r].OutputFontResourceId, budget);
            Same(d.SourceDeclarationOriginEntry, a.OriginEntry, budget); Same(d.SourceDeclarationKind, a.DeclarationKind, budget);
            Same(d.SourceDeclarationValue, a.DeclarationValue, budget); Origin(a, d.SourceDeclarationLinkPresent, d.SourceRawLinkValue, d.SourceDeclarationExpandedName, budget);
            Same(d.SourceResourceEntry, a.ResourceEntry, budget); Same(d.SourceFontEntry, a.Entry, budget);
            Same(d.SourceFontFileRaw, a.FontFile, budget); Same(d.SourceResBaseLocRaw, Attribute(a.ResourceAttributes, "BaseLoc"), budget);
            ExpandedName(d.SourceFontExpandedName!, a.Namespace, "Font", budget);
            Need(d.SourceFontOrdinalWithinFonts == a.Ordinal && d.SourceFontElementPreorderOrdinal == a.ElementPreorder, "OwnedOriginalFontLocation");
            Same(d.OutputDeclarationOriginEntry, current.OriginEntry, budget); Same(d.OutputDeclarationKind, current.DeclarationKind, budget);
            Same(d.OutputDeclarationValue, current.DeclarationValue, budget); Origin(current, d.OutputDeclarationLinkPresent, d.OutputRawLinkValue, d.OutputDeclarationExpandedName, budget);
            Same(d.OutputResourceEntry, current.ResourceEntry, budget); Same(d.OutputFontEntry, current.Entry, budget);
            Same(d.OutputFontFileRaw, current.FontFile, budget); Same(d.OutputResBaseLocRaw, Attribute(current.ResourceAttributes, "BaseLoc"), budget);
            ExpandedName(d.OutputFontExpandedName, current.Namespace, "Font", budget);
            Need(d.OutputFontOrdinalWithinFonts == current.Ordinal && d.OutputFontElementPreorderOrdinal == current.ElementPreorder, "OwnedActualFontLocation");
            Same(a.Namespace, current.Namespace, budget); Same(a.FontName, current.FontName, budget);
            Same(a.FamilyName, current.FamilyName, budget); Same(a.Charset, current.Charset, budget);
            budget.Field(2); Need(a.Bold == current.Bold && a.Italic == current.Italic, "GeometryFontStyle");
            MappedAttributes(a.Attributes, current.Attributes, budget, "ID", d.OutputFontResourceId);
            MappedAttributes(a.ResourceAttributes, current.ResourceAttributes, budget, "BaseLoc", d.OutputResBaseLocRaw);
        }
        internal void CurrentFont(int r, OfdTextGeometryExpectation.Root original, OfdFontResource current, OfdKeywordLocatorBudget budget)
        {
            var summary = issuer.CurrentFont(scope, actual, current);
            Same(summary.ContentProfile, "raw-embedded-font-bytes/sha256/v1", budget);
            Same(summary.Sha256Hex, original.FontSha256, budget); Same(summary.ResolvedPayloadSha256Hex, original.FontSha256, budget);
            budget.Field(4); Need(summary.ByteLength == original.FontBytes && summary.ResolvedPayloadByteLength == original.FontBytes &&
                summary.Bold == original.Font.Bold && summary.Italic == original.Font.Italic, "OwnedFreshFontIdentity");
            Same(summary.FontName, original.Font.FontName ?? original.Font.FamilyName, budget);
            Same(summary.FamilyName, original.Font.FamilyName, budget); Same(summary.Charset, original.Font.Charset, budget);
            var d = rows[r].FontDeclaration; var currentDeclaration = summary.Declaration;
            Same(summary.ModelFileName, d.OutputFontFileRaw, budget);
            Same(currentDeclaration.SourceFontResourceId, original.FontId, budget);
            Same(currentDeclaration.OutputFontResourceId, d.OutputFontResourceId, budget);
            Same(currentDeclaration.SourceResourceEntry, d.SourceResourceEntry, budget); Same(currentDeclaration.SourceFontEntry, d.SourceFontEntry, budget);
            Same(currentDeclaration.OutputResourceEntry, d.OutputResourceEntry, budget); Same(currentDeclaration.OutputFontEntry, d.OutputFontEntry, budget);
            Same(currentDeclaration.SourceDeclarationOriginEntry, d.SourceDeclarationOriginEntry, budget);
            Same(currentDeclaration.OutputDeclarationOriginEntry, d.OutputDeclarationOriginEntry, budget);
            // RequireVerifiedFontSummary binds both current arrays and the actual
            // resource/receipt/scope. This path performs no second F hash.
        }
        private static void Origin(OfdKeywordFontInput font, bool present, string? raw, string? qname, OfdKeywordLocatorBudget budget)
        {
            if (font.DeclarationKind == "FixedPageRes") { Need(!present && raw is null && qname is null, "FixedPageResHasNoLink"); Same(font.DeclarationValue, "PageRes.xml", budget); }
            else { Need(present, "MissingFontOriginLink"); Same(raw, font.DeclarationValue, budget); ExpandedName(qname!, font.Namespace, font.DeclarationKind, budget); }
        }
        private static string? Attribute(IReadOnlyList<OfdKeywordInputAttribute> values, string name)
        { foreach (var value in values) if (value.Namespace.Length == 0 && value.Name == name) return value.Value; return null; }
        private static void ExpandedName(string value, string ns, string local, OfdKeywordLocatorBudget budget)
        {
            budget.Characters(checked(value.Length + ns.Length + local.Length));
            Need(value.Length == ns.Length + local.Length + 2 && value[0] == '{' && value[ns.Length + 1] == '}' &&
                string.CompareOrdinal(value, 1, ns, 0, ns.Length) == 0 && string.CompareOrdinal(value, ns.Length + 2, local, 0, local.Length) == 0, "OwnedExpandedName");
        }
        private static void MappedAttributes(IReadOnlyList<OfdKeywordInputAttribute> a, IReadOnlyList<OfdKeywordInputAttribute> b,
            OfdKeywordLocatorBudget budget, string field, string? output, string? field2 = null, string? output2 = null)
        {
            budget.Field(); Need(a.Count == b.Count, "GeometryAttributePresence");
            for (var i = 0; i < a.Count; i++)
            {
                Same(a[i].Namespace, b[i].Namespace, budget); Same(a[i].Name, b[i].Name, budget);
                if (a[i].Namespace.Length == 0 && a[i].Name == field) Same(output, b[i].Value, budget);
                else if (field2 is not null && a[i].Namespace.Length == 0 && a[i].Name == field2) Same(output2, b[i].Value, budget);
                else Same(a[i].Value, b[i].Value, budget);
            }
        }
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
