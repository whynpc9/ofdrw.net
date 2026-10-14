using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging;

public sealed partial class OfdPackageWriter
{
    internal static Task<OfdOwnedWrittenSpool> WritePrivateOwnedAsync(OfdOwnedWriteScope scope,
        OfdRetainedTextEditBatch batch, CancellationToken token) => OwnedWriteReceipt.CreateWrittenAsync(scope, batch, token);
    private static Dictionary<byte[], byte[]> FreezeOwnedEntries(Dictionary<string, byte[]> entries,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(checked(1024L + 512L * entries.Count), token);
        var keys = entries.Keys.ToArray(); var copies = new Dictionary<byte[], byte[]>();
        foreach (var key in keys)
        {
            material.ChargeWork(16, token); var source = entries[key];
            if (!copies.TryGetValue(source, out var frozen))
            {
                material.ChargeCopy(checked(128L + source.LongLength), token);
                // This fixed implementation owns both operands and records the
                // exact array copy, uninterrupted before the fresh C walk.
                frozen = new byte[source.Length]; Array.Copy(source, frozen, source.Length); copies.Add(source, frozen);
            }
            entries[key] = frozen;
        }
        return copies;
    }

    internal sealed class OwnedWriteReceipt
    {
        internal static async Task<OfdOwnedWrittenSpool> CreateWrittenAsync(OfdOwnedWriteScope scope,
            OfdRetainedTextEditBatch batch, CancellationToken token)
        {
            OfdOwnedEditCandidate? candidate = null; OfdAccountedTemporaryFile? file = null;
            try
            {
                candidate = OfdOwnedSaveAuthority.ClonePatchAndClaim(scope, batch, out var guard, token);
                var context = candidate.Context; var material = context.Material;
                material.ChargeCopy(4096, token);
                var plan = PreparePlan(candidate.Package, material, token);
                scope.Owner.RecordActualLogicalWriterReservation(scope, candidate, plan.OutputReservation, token);
                material.ChargeCopy(checked(1024L + 512L * plan.Fonts.Resources.Count), token);
                var sourceFonts = new Dictionary<string, OfdOwnedFontDeclaration.Facts>(StringComparer.Ordinal);
                foreach (var font in plan.Fonts.Resources)
                {
                    if (font.Id != font.Resource.Id || font.Resource.Data.Length == 0)
                        throw new InvalidDataException("Owned Writer requires imported embedded font owners with identity IDs.");
                    sourceFonts.Add(font.Id, OfdOwnedFontPlanEvidence.CaptureSource(candidate.Package, font.Resource, material, token));
                }
                var entries = new OfdPackageWriter().BuildEntries(candidate.Package, material, token, plan);
                var pruning = OfdOwnedPruneDispositions.Begin(context, token);
                _ = OfdPackagePruner.Prune(candidate.Package, entries, token, material, pruning);
                material.ValidateActualEntries(entries, token, plan.OutputReservation);
                var copies = FreezeOwnedEntries(entries, material, token);
                // This is a real fresh pre-stream C; original capture is not reused.
                var current = guard.AssertUnchangedAndObservePayloads(candidate.Package, context, token);
                var receipt = OwnedWriteReceipt.Create(scope, candidate, plan, entries, copies, pruning, sourceFonts, current, token);
                current.CompletePrivateIssuance();
                file = new OfdAccountedTemporaryFile(context, 512L * 1024 * 1024, token, scope.Owner.BindActualOutputObservation(scope, candidate, token));
                await WriteZipTerminalAsync(entries, plan, file, context, token).ConfigureAwait(false);
                await file.FlushAsync(token).ConfigureAwait(false); // exact owned route only; old Stream path is unchanged
                candidate.AssertClaimedActive(); token.ThrowIfCancellationRequested(); context.AssertActive();
                var raw = await file.ObserveRawZipAsync(token).ConfigureAwait(false);
                var written = OfdOwnedWrittenSpool.FromActualWrite(scope, candidate, file, receipt, raw.Length, raw.Sha256, token);
                scope.Owner.CompleteWrite(scope, written, token);
                return written;
            }
            catch { scope.Owner.Fail(); candidate?.Dispose(); file?.Dispose(); batch?.Dispose(); throw; }
        }

        private readonly OfdControlledEditContext context;
        private readonly Dictionary<string, OfdObservedPayloadIdentity> entries;
        private readonly Dictionary<string, FontRecord> fonts;
        private readonly OfdReadbackRootAddress[] addresses;
        private readonly string documentEntry;
        private readonly string[] pageEntries;
        private int used;
        private sealed class FontRecord(OfdOwnedFontDeclaration declaration, OfdObservedPayloadIdentity expected,
            string name, string? family, string? charset, bool bold, bool italic, string? fileName)
        {
            internal readonly OfdOwnedFontDeclaration Declaration = declaration;
            internal readonly OfdObservedPayloadIdentity Expected = expected;
            internal readonly string Name = name;
            internal readonly string? Family = family, Charset = charset, FileName = fileName;
            internal readonly bool Bold = bold, Italic = italic;
        }
        private OwnedWriteReceipt(OfdControlledEditContext context, Dictionary<string, OfdObservedPayloadIdentity> entries,
            Dictionary<string, FontRecord> fonts, OfdReadbackRootAddress[] addresses, string documentEntry, string[] pageEntries)
        { this.context = context; this.entries = entries; this.fonts = fonts; this.addresses = addresses; this.documentEntry = documentEntry; this.pageEntries = pageEntries; }
        private static OwnedWriteReceipt Create(OfdOwnedWriteScope scope, OfdOwnedEditCandidate candidate, WritePlan plan,
            Dictionary<string, byte[]> frozenEntries, Dictionary<byte[], byte[]> copies, OfdOwnedPruneDispositions pruning,
            Dictionary<string, OfdOwnedFontDeclaration.Facts> sourceFonts, OfdCurrentPayloadEvidence current, CancellationToken token)
        {
            var context = candidate.Context; var material = context.Material;
            material.ChargeCopy(checked(4096L + 1024L * frozenEntries.Count + 2048L * sourceFonts.Count), token);
            var targets = new Dictionary<string, OfdObservedPayloadIdentity>(StringComparer.OrdinalIgnoreCase);
            var reverseCopies = copies.ToDictionary(pair => pair.Value, pair => pair.Key);
            foreach (var entry in frozenEntries)
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(16, token);
                if (!reverseCopies.TryGetValue(entry.Value, out var origin) || origin.LongLength != entry.Value.LongLength)
                    throw new InvalidDataException("Frozen entry is not the exact privately issued copy.");
                var identity = current.FindObservedCopySource(candidate.Package, context, origin, token)
                    ?? OfdObservedPayloadIdentity.FromActual(origin, material.HashBytes(origin, token), material, token);
                targets.Add(entry.Key, identity);
            }
            foreach (var original in candidate.Package.PreservedEntries)
            {
                material.ChargeWork(16, token);
                if (!frozenEntries.TryGetValue(original.Key, out var frozen))
                { pruning.RequireActualRemoval(original.Key, original.Value, token); continue; }
                var origin = reverseCopies[frozen];
                if (ReferenceEquals(origin, original.Value)) continue;
                var expectedXml = BuildExpectedClosedRewrite(original.Key, original.Value, candidate.Package, plan, sourceFonts, pruning, material, token);
                OfdOwnedXmlPreservation.RequireExactClosedRewrite(original.Key, expectedXml, origin, material, token);
            }
            var actualPageEntries = plan.Pages.Select(page => plan.PagePaths[page]).ToArray();
            var fontRecords = new Dictionary<string, FontRecord>(StringComparer.Ordinal);
            foreach (var binding in plan.Fonts.Resources)
            {
                var source = sourceFonts[binding.Id];
                var output = OfdOwnedFontPlanEvidence.CaptureOutput(plan.DocumentPath, actualPageEntries, frozenEntries, source, binding.Id, material, token);
                OfdOwnedFontPlanEvidence.RequireIdentityMapping(source, output);
                var declaration = OfdOwnedFontDeclaration.FromActualPlan(binding.Resource.Id, binding.Id, source, output);
                var expected = current.RequireObserved(candidate.Package, context, binding.Resource.Data, token);
                if (expected.Sha256Hex != targets[output.FontEntry].Sha256Hex || expected.ByteLength != targets[output.FontEntry].ByteLength)
                    throw new InvalidDataException("Actual planned font payload does not bind the current candidate model bytes.");
                fontRecords.Add(binding.Id, new FontRecord(declaration, expected, binding.Resource.FontName,
                    binding.Resource.FamilyName ?? binding.Resource.FontName, binding.Resource.Charset,
                    binding.Resource.Bold, binding.Resource.Italic, output.FontFileRaw));
            }
            var rootKeys = scope.Owner.GetWriterRootKeys(scope, candidate, token);
            material.ChargeCopy(checked(1024L + 4096L * rootKeys.Length), token);
            var addresses = new OfdReadbackRootAddress[rootKeys.Length];
            var sourceDoc = material.ReadXml(candidate.Package.DocumentEntryPath!, candidate.Package.PreservedEntries[candidate.Package.DocumentEntryPath!], token);
            var outputDoc = material.ReadXml(plan.DocumentPath, frozenEntries[plan.DocumentPath], token);
            var sourceReferences = IndexPageReferences(sourceDoc, material, token);
            var outputReferences = IndexPageReferences(outputDoc, material, token);
            var actualLocations = IndexActualPlanRoots(plan, frozenEntries, material, token);
            for (var i = 0; i < rootKeys.Length; i++)
            {
                material.ChargeWork(64, token);
                var root = rootKeys[i].ResolveCandidate(scope.Owner, candidate);
                if (!actualLocations.TryGetValue(root, out var location))
                    throw new InvalidDataException("Expected root has no exact actual planned output locator.");
                var page = location.Page; var path = plan.PagePaths[page];
                var pageId = plan.PageIds[page]; var objectId = plan.ElementIds[root]; var fontId = plan.Fonts.TextBindings[root].Id;
                if (!sourceReferences.TryGetValue(page.Id!, out var srcPage) || !outputReferences.TryGetValue(pageId, out var outPage))
                    throw new InvalidDataException("Actual Document lacks the planned page reference.");
                if (srcPage.Name != outPage.Name || srcPage.Attribute("BaseLoc")?.Value != outPage.Attribute("BaseLoc")?.Value || candidate.Package.DocumentEntryPath != plan.DocumentPath)
                    throw new InvalidDataException("Actual page reference requires unsupported raw-literal/QName/origin relocation.");
                addresses[i] = OfdReadbackRootAddress.FromActualPlan(rootKeys[i], page.Index, path, location.LayerOrdinal, location.RootOrdinal,
                    objectId, pageId, location.LayerId, fontId, fontRecords[fontId].Declaration,
                    srcPage.Attribute("BaseLoc")?.Value, outPage.Attribute("BaseLoc")?.Value,
                    srcPage.Name.ToString(), outPage.Name.ToString(), candidate.Package.DocumentEntryPath!, plan.DocumentPath);
            }
            return new(context, targets, fontRecords, addresses, plan.DocumentPath, actualPageEntries);
        }
        private static XDocument BuildExpectedClosedRewrite(string entry, byte[] original,
            OfdDocumentPackage candidate, WritePlan plan, Dictionary<string, OfdOwnedFontDeclaration.Facts> sourceFonts,
            OfdOwnedPruneDispositions pruning, OfdModelMaterialization material, CancellationToken token)
        {
            var expected = material.ReadXml(entry, original, token);
            var document = expected.Root ?? throw new InvalidDataException("Source control XML lacks a root.");
            var ns = document.Name.Namespace;
            material.ChargeCopy(checked(1024L + 64L * document.DescendantsAndSelf().Count()), token);
            var sourceNodes = document.DescendantsAndSelf().ToArray();
            if (entry == plan.DocumentPath)
            {
                var max = document.Element(ns + "CommonData")?.Element(ns + "MaxUnitID");
                if (max is null) throw new InvalidDataException("Private document profile requires its actual MaxUnitID.");
                max.Value = plan.Ids.MaxId.ToString(CultureInfo.InvariantCulture);
            }
            foreach (var page in plan.Pages)
            {
                material.ChargeWork(16, token);
                if (plan.PagePaths[page] != entry) continue;
                if (document.Name != ns + "Page") throw new InvalidDataException("Planned page lacks its actual source Page role.");
                foreach (var element in page.Elements)
                {
                    material.ChargeWork(32, token);
                    if (candidate.SourceIdentity is null || !candidate.SourceIdentity.TryOwner(element, out var owner) ||
                        owner.Entry != entry || owner.Slot < 0 || owner.Slot >= sourceNodes.Length ||
                        sourceNodes[owner.Slot].Name != owner.Kind || sourceNodes[owner.Slot].Attribute("ID")?.Value != owner.Id ||
                        plan.ElementIds[element] != owner.Id)
                        throw new InvalidDataException("Planned root lacks its exact imported source position and identity ID.");
                    // Core applied the authenticated fixed-slot edits to this
                    // literal. It is independent of BuildEntries' emitted bytes.
                    var literal = element switch
                    {
                        OfdTextElement textRoot => textRoot.SourceXml,
                        OfdImageElement imageRoot => imageRoot.SourceXml,
                        OfdPathElement pathRoot => pathRoot.SourceXml,
                        OfdRawElement rawRoot => rawRoot.Xml,
                        _ => null
                    };
                    if (string.IsNullOrWhiteSpace(literal))
                        throw new InvalidDataException("Private root requires its imported candidate SourceXml.");
                    material.PrepareReadXmlParser(checked(2L * literal!.Length), token);
                    var approved = XElement.Parse(literal!, LoadOptions.PreserveWhitespace);
                    if (approved.Name != owner.Kind || approved.Attribute("ID")?.Value != owner.Id)
                        throw new InvalidDataException("Candidate SourceXml changed its source root authority.");
                    if (element is OfdTextElement text)
                    {
                        var binding = plan.Fonts.TextBindings[text];
                        if (binding.Id != text.FontResourceId || binding.Resource.Data.Length == 0)
                            throw new InvalidDataException("Private text rewrite requires the actual identity embedded-font plan.");
                        approved.SetAttributeValue("Font", binding.Id);
                        approved.SetAttributeValue("Weight", text.Weight != OfdTextElement.DefaultWeight
                            ? text.Weight.ToString(CultureInfo.InvariantCulture) : null);
                        approved.SetAttributeValue("Italic", text.Italic ? "true" : null);
                    }
                    // Other imported roots are preserved literally in this minimum
                    // profile; an unsupported serializer normalization fails closed.
                    sourceNodes[owner.Slot].ReplaceWith(approved);
                    if (!candidate.SourceIdentity.TryLayer(element, out var layerOwner) ||
                        layerOwner.Entry != entry || layerOwner.Slot < 0 || layerOwner.Slot >= sourceNodes.Length ||
                        sourceNodes[layerOwner.Slot].Name != layerOwner.Kind || sourceNodes[layerOwner.Slot].Attribute("ID")?.Value != layerOwner.Id ||
                        plan.LayerIds[page][GetLayerKey(element)] != layerOwner.Id)
                        throw new InvalidDataException("Planned layer lacks its exact imported source owner.");
                    var layer = sourceNodes[layerOwner.Slot];
                    layer.SetAttributeValue("Type", string.IsNullOrWhiteSpace(element.LayerType) ? "Body" : element.LayerType);
                }
            }
            foreach (var binding in plan.Fonts.Resources)
            {
                material.ChargeWork(16, token);
                var facts = sourceFonts[binding.Id];
                if (facts.ResourceEntry != entry) continue;
                var ordinal = facts.FontElementPreorderOrdinal;
                if (ordinal < 0 || ordinal >= sourceNodes.Length || sourceNodes[ordinal].Name.ToString() != facts.FontExpandedName ||
                    sourceNodes[ordinal].Attribute("ID")?.Value != binding.Id)
                    throw new InvalidDataException("Planned font lacks its actual source resource position.");
                var font = sourceNodes[ordinal]; var model = binding.Resource;
                // The actual catalog plan's explicit modeled attributes. All
                // other attributes, text, child nodes and positions stay exact.
                font.SetAttributeValue("FontName", model.FontName);
                font.SetAttributeValue("FamilyName", model.FamilyName ?? model.FontName);
                font.SetAttributeValue("Charset", model.Charset);
                font.SetAttributeValue("Bold", model.Bold ? true : (object?)null);
                font.SetAttributeValue("Italic", model.Italic ? true : (object?)null);
            }
            pruning.ApplyActualSignatureRemovals(entry, expected, token);
            return expected;
        }
        private sealed class RootLocation(OfdPage page, int layerOrdinal, int rootOrdinal, string? layerId)
        {
            internal readonly OfdPage Page = page;
            internal readonly int LayerOrdinal = layerOrdinal, RootOrdinal = rootOrdinal;
            internal readonly string? LayerId = layerId;
        }
        private static Dictionary<string, XElement> IndexPageReferences(XDocument document,
            OfdModelMaterialization material, CancellationToken token)
        {
            var ns = document.Root!.Name.Namespace;
            var parent = document.Root.Element(ns + "Pages") ?? throw new InvalidDataException("Actual Document lacks Pages.");
            var count = 0;
            foreach (var node in parent.Elements(ns + "Page")) { material.ChargeWork(16, token); count++; }
            material.ChargeCopy(checked(512L + 512L * count), token);
            var result = new Dictionary<string, XElement>(StringComparer.Ordinal);
            foreach (var node in parent.Elements(ns + "Page"))
            {
                var id = node.Attribute("ID")?.Value;
                material.ChargeWork(checked(32L + 2L * (id?.Length ?? 0)), token);
                if (string.IsNullOrEmpty(id) || result.ContainsKey(id!)) throw new InvalidDataException("Actual page references have missing or duplicate owner IDs.");
                result.Add(id!, node);
            }
            return result;
        }
        private static Dictionary<OfdTextElement, RootLocation> IndexActualPlanRoots(WritePlan plan,
            Dictionary<string, byte[]> entries, OfdModelMaterialization material, CancellationToken token)
        {
            long maximumRoots = 0;
            foreach (var page in plan.Pages) { material.ChargeWork(16, token); maximumRoots = checked(maximumRoots + page.Elements.Count); }
            material.ChargeCopy(checked(1024L + 1024L * maximumRoots), token);
            var result = new Dictionary<OfdTextElement, RootLocation>();
            var expectedCount = 0;
            foreach (var page in plan.Pages)
            {
                var path = plan.PagePaths[page]; var pageId = plan.PageIds[page];
                material.ChargeWork(checked(32L + 2L * (path.Length + pageId.Length)), token);
                if (page.Id != pageId || page.SourceEntryPath != path)
                    throw new InvalidDataException("Private text plan requires identity page IDs and paths.");
                var owners = new Dictionary<string, OfdTextElement>(StringComparer.Ordinal);
                foreach (var element in page.Elements)
                {
                    material.ChargeWork(32, token);
                    if (element is not OfdTextElement text) continue;
                    var id = plan.ElementIds[text]; var fontId = plan.Fonts.TextBindings[text].Id;
                    material.ChargeWork(checked(32L + 2L * (id.Length + fontId.Length)), token);
                    if (text.ObjectId != id || text.FontResourceId != fontId || owners.ContainsKey(id))
                        throw new InvalidDataException("Private text plan requires identity object/font bindings.");
                    owners.Add(id, text); expectedCount++;
                }
                var pageXml = material.ReadXml(path, entries[path], token); var ns = pageXml.Root!.Name.Namespace;
                var content = pageXml.Root.Element(ns + "Content") ?? throw new InvalidDataException("Actual page lacks Content.");
                var layerOrdinal = 0;
                foreach (var layer in content.Elements(ns + "Layer"))
                {
                    var layerId = layer.Attribute("ID")?.Value; var rootOrdinal = 0;
                    foreach (var node in layer.Elements())
                    {
                        material.ChargeWork(32, token);
                        if (node.Name == ns + "TextObject")
                        {
                            var id = node.Attribute("ID")?.Value;
                            material.ChargeWork(checked(32L + 2L * (id?.Length ?? 0)), token);
                            if (id is null || !owners.TryGetValue(id, out var owner) || result.ContainsKey(owner) ||
                                owner.LayerId != layerId || node.Attribute("Font")?.Value != plan.Fonts.TextBindings[owner].Id)
                                throw new InvalidDataException("Actual planned root has foreign, ambiguous or relocated layer/font ownership.");
                            result.Add(owner, new RootLocation(page, layerOrdinal, rootOrdinal, layerId));
                        }
                        rootOrdinal++;
                    }
                    layerOrdinal++;
                }
            }
            if (result.Count != expectedCount) throw new InvalidDataException("Actual planned text-root coverage is incomplete.");
            return result;
        }
        internal Dictionary<OfdFontResource, OfdOwnedVerifiedFontSummary> VerifyActual(OfdDocumentPackage actualR,
            OfdOwnedZipCensus census, OfdCurrentPayloadEvidence current, CancellationToken token)
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (Interlocked.CompareExchange(ref used, 1, 0) != 0) throw new InvalidOperationException("Writer receipt was already consumed.");
            var physical = census.TakeForPrivateVerifier(context, token);
            if (physical.Length != entries.Count || actualR.PreservedEntries.Count != entries.Count)
                throw new InvalidDataException("Actual physical ZIP/model entry census differs from the receipt.");
            foreach (var entry in physical)
            {
                context.Material.ChargeWork(16, token);
                if (entry.Directory || entry.NormalizedName is null || entry.PhysicalName != entry.NormalizedName || entry.Bytes is null ||
                    !entries.TryGetValue(entry.NormalizedName, out var expected) ||
                    !actualR.PreservedEntries.TryGetValue(entry.NormalizedName, out var projected) || !ReferenceEquals(projected, entry.Bytes))
                    throw new InvalidDataException("Actual ZIP has an unreceipted physical or projected entry.");
                var observed = current.RequireObserved(actualR, context, entry.Bytes, token);
                if (expected.ByteLength != entry.ExpandedLength || expected.ByteLength != observed.ByteLength || expected.Sha256Hex != observed.Sha256Hex)
                    throw new InvalidDataException("Actual current ZIP payload differs from the frozen receipt target.");
            }
            if (actualR.Fonts.Count != fonts.Count) throw new InvalidDataException("Actual R font coverage differs from the complete Writer plan.");
            context.Material.ChargeCopy(checked(1024L + 512L * fonts.Count), token);
            var result = new Dictionary<OfdFontResource, OfdOwnedVerifiedFontSummary>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var actualFont in actualR.Fonts)
            {
                if (!seen.Add(actualFont.Id) || !fonts.TryGetValue(actualFont.Id, out var expected))
                    throw new InvalidDataException("Actual R has an unplanned or duplicate font model.");
                var sourceFacts = new OfdOwnedFontDeclaration.Facts(expected.Declaration.OutputDeclarationOriginEntry,
                    expected.Declaration.OutputDeclarationKind, expected.Declaration.OutputDeclarationValue,
                    expected.Declaration.OutputRawLinkValue, expected.Declaration.OutputDeclarationExpandedName,
                    expected.Declaration.OutputDeclarationLinkPresent, expected.Declaration.OutputResourceEntry,
                    expected.Declaration.OutputFontEntry, expected.Declaration.OutputFontExpandedName,
                    expected.Declaration.OutputFontOrdinalWithinFonts, expected.Declaration.OutputFontElementPreorderOrdinal,
                    expected.Declaration.OutputFontFileRaw, expected.Declaration.OutputResBaseLocRaw);
                var actualFacts = OfdOwnedFontPlanEvidence.CaptureOutput(documentEntry, pageEntries, actualR.PreservedEntries,
                    sourceFacts, actualFont.Id, context.Material, token);
                OfdOwnedFontPlanEvidence.RequireIdentityMapping(sourceFacts, actualFacts);
                var actualPayload = actualR.PreservedEntries[actualFacts.FontEntry];
                var summary = OfdOwnedVerifiedFontSummary.FromActualCompoundPass(actualFont, actualPayload, actualR, context, current, expected.Declaration, token);
                if (summary.Sha256Hex != expected.Expected.Sha256Hex || summary.ByteLength != expected.Expected.ByteLength ||
                    summary.ResolvedPayloadSha256Hex != expected.Expected.Sha256Hex || summary.ResolvedPayloadByteLength != expected.Expected.ByteLength ||
                    summary.FontName != expected.Name || summary.FamilyName != expected.Family || summary.Charset != expected.Charset ||
                    summary.Bold != expected.Bold || summary.Italic != expected.Italic || summary.ModelFileName != expected.FileName)
                    throw new InvalidDataException("Actual current model/resolved-font evidence differs from its planned declaration or candidate content.");
                result.Add(actualFont, summary);
            }
            return result;
        }
        internal OfdReadbackRootAddress[] CopyAddresses(CancellationToken token)
        {
            context.AssertActive(); context.Material.ChargeCopy(checked(128L + IntPtr.Size * (long)addresses.Length), token);
            return addresses.ToArray();
        }
    }
}
