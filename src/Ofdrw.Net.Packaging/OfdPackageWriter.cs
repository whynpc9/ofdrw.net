using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Packaging;

public sealed partial class OfdPackageWriter
{
    public async Task WriteAsync(OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        _ = await WriteWithResultAsync(package, destination, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes the package and reports removed entries, including requested attachment cleanup, and invalidated signatures.</summary>
    public async Task<OfdPackageWriteResult> WriteWithResultAsync(
        OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        cancellationToken.ThrowIfCancellationRequested();
        var material = new OfdModelMaterialization(new OfdModelLimits());
        OfdPageXmlContract.ValidateDocumentArea(package, cancellationToken);
        foreach (var page in package.Pages) OfdPageXmlContract.ValidateForRewrite(package, page, cancellationToken);
        if (package.Pages.Count == 0) OfdPageXmlContract.ValidateWritableDimensions(
            package.Options.DefaultPageWidthMillimeters, package.Options.DefaultPageHeightMillimeters);
        foreach (var text in OfdCompositeGraph.Definitions(package).OfType<OfdTextElement>())
            _ = Ofdrw.Net.Core.Fonts.EmbeddedFontCoverage.HasExplicitGlyphReferences(text);
        // Normal generation honors the explicit font embedding policy. Controlled
        // text-edit continuations retain their admitted original font payloads.
        var fontPlan = new OfdFontEmbeddingPlan(package, cancellationToken, material);
        var plan = PreparePlan(package, material, cancellationToken, fontPlan);
        var entries = BuildEntries(package, material, cancellationToken, plan);
        var result = OfdPackagePruner.Prune(package, entries, cancellationToken, material);
        result.Fonts.AddRange(fontPlan.Results);
        result.Warnings.AddRange(fontPlan.Diagnostics);
        material.ValidateActualEntries(entries, cancellationToken, plan.OutputReservation);
        FreezeEntryBuffers(entries, material, cancellationToken);
        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
        {
            var zipEntry = zip.CreateEntry(entry.Key, plan.Compression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
            using var stream = zipEntry.Open();
            await stream.WriteAsync(entry.Value, 0, entry.Value.Length, cancellationToken).ConfigureAwait(false);
        }
        return result;
    }

    // A real factory-issued candidate is mandatory. No public ID/map, optional
    // materialization or fresh-context fallback can authenticate this path.
    internal static PreparedWrite PrepareOwnedWrite(OfdOwnedEditCandidate candidate, CancellationToken token = default)
        => PreparedWrite.Create(candidate, token);

    // Closed synchronous source→clone→fixed slots→claim→plan continuation.
    // No stream or caller callback is touched here. The returned PreparedWrite
    // retains all existing caller-gap and actual I/O checks.
    internal static PreparedWrite PrepareContinuousTextEdit(OfdDocumentPackage source,
        OfdSerializedGuard sourceGuard, OfdControlledEditContext context,
        OfdRetainedTextEditBatch edits, CancellationToken token = default)
        => PreparedWrite.CreateContinuous(source, sourceGuard, context, edits, token);

    // No prepared owner/status escapes between clone/patch and the first I/O.
    // Destination is borrowed, including arbitrary Stream implementations.
    internal static Task<WrittenTextEdit> WriteContinuousTextEditAsync(OfdDocumentPackage source,
        OfdSerializedGuard sourceGuard, OfdControlledEditContext context,
        OfdRetainedTextEditBatch edits, Stream destination, CancellationToken token = default)
        => WrittenTextEdit.CreateAsync(source, sourceGuard, context, edits, destination, token);

    internal sealed class WrittenTextEdit : IDisposable
    {
        private readonly OfdOwnedEditCandidate candidate;
        private readonly OfdSerializedGuard guard;
        private int state; // written, validating, disposed, failed
        private WrittenTextEdit(OfdOwnedEditCandidate candidate, OfdSerializedGuard guard)
        { this.candidate = candidate; this.guard = guard; }

        internal static async Task<WrittenTextEdit> CreateAsync(OfdDocumentPackage source,
            OfdSerializedGuard sourceGuard, OfdControlledEditContext context,
            OfdRetainedTextEditBatch edits, Stream destination, CancellationToken token)
        {
            if (context is null) throw new ArgumentNullException(nameof(context));
            OfdOwnedEditCandidate? candidate = null;
            try
            {
                context.AssertActive(); token.ThrowIfCancellationRequested();
                // Async state, written owner, and one primary EDI reservation.
                // Charge before the eligible operation can allocate/suspend.
                context.Material.ChargeCopy(1024, token);
                if (source is null || sourceGuard is null || edits is null || destination is null)
                    throw new ArgumentNullException(source is null ? nameof(source) : sourceGuard is null ? nameof(sourceGuard) : edits is null ? nameof(edits) : nameof(destination));
                candidate = OfdOwnedEditCandidate.ClonePatchAndClaim(source, sourceGuard, context, edits, token, out var guard);
                var material = context.Material;
                var plan = PreparePlan(candidate.Package, material, token);
                var entries = new OfdPackageWriter().BuildEntries(candidate.Package, material, token, plan);
                _ = OfdPackagePruner.Prune(candidate.Package, entries, token, material);
                material.ValidateActualEntries(entries, token, plan.OutputReservation);
                FreezeEntryBuffers(entries, material, token);
                // This fresh check covers the entire closed prefix. Not even
                // CanWrite/CanSeek/Position is queried before it succeeds.
                guard.AssertUnchanged(candidate.Package, context, token);
                await WriteZipTerminalAsync(entries, plan, destination, context, token).ConfigureAwait(false);
                candidate.AssertClaimedActive();
                guard.AssertUnchanged(candidate.Package, context, token);
                candidate.AssertSourceUnchanged(token);
                token.ThrowIfCancellationRequested(); context.AssertActive();
                // Only successfully disposed ZIP + fresh post-I/O C/S can
                // issue this owner. It says nothing about actual readback.
                return new WrittenTextEdit(candidate, guard);
            }
            catch { context.Fail(); candidate?.Dispose(); edits?.Dispose(); throw; }
        }

        internal void AssertSourceAndCandidateUnchanged(CancellationToken token = default)
        {
            if (Interlocked.CompareExchange(ref state, 1, 0) != 0)
                throw new InvalidOperationException("Written text edit is disposed, failed or in use.");
            var context = candidate.Context;
            try
            {
                candidate.AssertClaimedActive(); candidate.AssertSourceUnchanged(token);
                guard.AssertUnchanged(candidate.Package, context, token); candidate.AssertClaimedActive();
            }
            catch { context.Fail(); Volatile.Write(ref state, 3); throw; }
            finally { Interlocked.CompareExchange(ref state, 0, 1); }
        }

        public void Dispose()
        {
            while (true)
            {
                var before = Volatile.Read(ref state);
                if (before == 2) return;
                if (before == 1) throw new InvalidOperationException("Cannot dispose a written text edit during validation.");
                if (Interlocked.CompareExchange(ref state, 2, before) == before) break;
            }
            candidate.Dispose();
        }
    }

    private static async Task WriteZipTerminalAsync(Dictionary<string, byte[]> entries, WritePlan plan,
        Stream destination, OfdControlledEditContext context, CancellationToken token)
    {
        ZipArchive? zip = null; Stream? entryStream = null; ExceptionDispatchInfo? primary = null;
        try
        {
            zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
            {
                token.ThrowIfCancellationRequested(); context.AssertActive();
                var zipEntry = zip.CreateEntry(entry.Key, plan.Compression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
                entryStream = zipEntry.Open();
                await entryStream.WriteAsync(entry.Value, 0, entry.Value.Length, token).ConfigureAwait(false);
                // Clear the handle before trying Dispose: a failed Dispose is
                // already an attempt and must not be retried in cleanup.
                var completed = entryStream; entryStream = null; completed.Dispose();
            }
        }
        catch (Exception error) { context.Fail(); primary = ExceptionDispatchInfo.Capture(error); }

        // Fail was observed before any accessible external error cleanup.
        // Each remaining handle is disposed once, and the first error wins.
        if (entryStream is not null)
        {
            var pending = entryStream; entryStream = null;
            try { pending.Dispose(); }
            catch (Exception error) { context.Fail(); if (primary is null) primary = ExceptionDispatchInfo.Capture(error); }
        }
        if (zip is not null)
        {
            var pending = zip; zip = null;
            try { pending.Dispose(); }
            catch (Exception error) { context.Fail(); if (primary is null) primary = ExceptionDispatchInfo.Capture(error); }
        }
        primary?.Throw();
    }

    internal sealed class PreparedWrite : IDisposable
    {
        private readonly OfdOwnedEditCandidate candidate;
        private readonly OfdSerializedGuard guard;
        private readonly WritePlan plan;
        private OfdOwnedEditCandidate? ownedLifetime;
        // Created, writing, written, validating-created, validating-written,
        // disposed, failed. No successful validation is a reusable stamp.
        private int state;
        private PreparedWrite(OfdOwnedEditCandidate candidate, OfdSerializedGuard guard, WritePlan plan)
        { this.candidate = candidate; this.guard = guard; this.plan = plan; }

        internal static PreparedWrite Create(OfdOwnedEditCandidate candidate, CancellationToken token)
        {
            if (candidate is null) throw new ArgumentNullException(nameof(candidate));
            var context = candidate.Context;
            try
            {
                var guard = candidate.ClaimWrite(token);
                var plan = PreparePlan(candidate.Package, context.Material, token);
                guard.AssertUnchanged(candidate.Package, context, token);
                return new PreparedWrite(candidate, guard, plan);
            }
            catch { context.Fail(); throw; }
        }

        internal static PreparedWrite CreateContinuous(OfdDocumentPackage source,
            OfdSerializedGuard sourceGuard, OfdControlledEditContext context,
            OfdRetainedTextEditBatch edits, CancellationToken token)
        {
            if (context is null) throw new ArgumentNullException(nameof(context));
            OfdOwnedEditCandidate? candidate = null;
            try
            {
                if (source is null || sourceGuard is null || edits is null)
                    throw new ArgumentNullException(source is null ? nameof(source) : sourceGuard is null ? nameof(sourceGuard) : nameof(edits));
                candidate = OfdOwnedEditCandidate.ClonePatchAndClaim(source, sourceGuard, context, edits, token, out var guard);
                var plan = PreparePlan(candidate.Package, context.Material, token);
                guard.AssertUnchanged(candidate.Package, context, token);
                return new PreparedWrite(candidate, guard, plan) { ownedLifetime = candidate };
            }
            catch { context.Fail(); candidate?.Dispose(); edits?.Dispose(); throw; }
        }

        // This is a non-atomic Stream write. 06 owns the private spool, readback
        // and AtomicOutput publication; callbacks may fail after bytes arrive.
        internal async Task<OfdPackageWriteResult> WriteAsync(Stream destination, CancellationToken token = default)
        {
            if (Interlocked.CompareExchange(ref state, 1, 0) != 0) throw new InvalidOperationException("Prepared write has already been consumed or is in use.");
            var context = candidate.Context; var material = context.Material;
            try
            {
            if (destination is null) throw new ArgumentNullException(nameof(destination));
            candidate.AssertClaimedActive();
            candidate.AssertSourceUnchanged(token);
            guard.AssertUnchanged(candidate.Package, context, token);
            var entries = new OfdPackageWriter().BuildEntries(candidate.Package, material, token, plan);
            var result = OfdPackagePruner.Prune(candidate.Package, entries, token, material);
            material.ValidateActualEntries(entries, token, plan.OutputReservation);
            FreezeEntryBuffers(entries, material, token);
            guard.AssertUnchanged(candidate.Package, context, token);
            using (var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
                {
                    token.ThrowIfCancellationRequested();
                    var zipEntry = zip.CreateEntry(entry.Key, plan.Compression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
                    using var stream = zipEntry.Open();
                    await stream.WriteAsync(entry.Value, 0, entry.Value.Length, token).ConfigureAwait(false);
                }
            }
            // Including destination/Dispose callbacks. Failure here cannot
            // roll back an arbitrary caller stream; no verified save is returned.
            candidate.AssertClaimedActive(); guard.AssertUnchanged(candidate.Package, context, token); candidate.AssertSourceUnchanged(token);
            Volatile.Write(ref state, 2);
            return result;
            }
            catch { context.Fail(); Volatile.Write(ref state, 6); throw; }
        }

        // Fresh checks through this exact retained owner, both before writing
        // and after one successful write. Nothing mutable or stamp-like escapes.
        internal void AssertSourceAndCandidateUnchanged(CancellationToken token = default)
        {
            var before = Volatile.Read(ref state);
            var busy = before == 0 ? 3 : 4;
            if (before is not (0 or 2) || Interlocked.CompareExchange(ref state, busy, before) != before)
                throw new InvalidOperationException("Prepared write cannot validate while consumed, failed, disposed or in use.");
            var context = candidate.Context;
            try
            {
                candidate.AssertClaimedActive();
                candidate.AssertSourceUnchanged(token);
                guard.AssertUnchanged(candidate.Package, context, token);
                candidate.AssertClaimedActive();
            }
            catch { context.Fail(); Volatile.Write(ref state, 6); throw; }
            finally { Interlocked.CompareExchange(ref state, before, busy); }
        }

        public void Dispose()
        {
            while (true)
            {
                var before = Volatile.Read(ref state);
                if (before == 5) return;
                if (before is 1 or 3 or 4) throw new InvalidOperationException("Cannot dispose a prepared write during I/O or validation.");
                if (Interlocked.CompareExchange(ref state, 5, before) == before) break;
            }
            Interlocked.Exchange(ref ownedLifetime, null)?.Dispose();
        }
    }

    private static void FreezeEntryBuffers(IDictionary<string, byte[]> entries, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(128L + (long)entries.Count * 160, token);
        var keys = new List<string>();
        foreach (var key in entries.Keys) { token.ThrowIfCancellationRequested(); material.Structure.Reference(); keys.Add(key); }
        var copies = new Dictionary<byte[], byte[]>();
        foreach (var key in keys)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference(); var value = entries[key];
            if (!copies.TryGetValue(value, out var snapshot))
            {
                material.ChargeCopy(value.LongLength, token); snapshot = value.ToArray(); copies.Add(value, snapshot);
            }
            entries[key] = snapshot;
        }
        // Destination stream callbacks cannot mutate caller-owned buffers after
        // validation. Concurrent mutation during preflight remains unsupported.
    }

    private sealed class WritePlan
    {
        internal Dictionary<string, byte[]> Entries = null!;
        internal List<OfdPage> Pages = null!;
        internal Dictionary<OfdPage, string> PageIds = null!, PagePaths = null!;
        internal Dictionary<OfdElement, string> ElementIds = null!;
        internal Dictionary<OfdAnnotation, string> AnnotationIds = new();
        internal Dictionary<OfdPage, string> NewAnnotationPaths = new();
        internal string? NewAnnotationIndex;
        internal Dictionary<OfdPage, Dictionary<string, string>> LayerIds = null!;
        internal Dictionary<OfdImageElement, ImageResource> Images = null!;
        internal FontBindings Fonts = null!;
        internal OfdIdPlan Ids = null!;
        internal OfdResourceCatalog Resources = null!;
        internal CustomTagWritePlan CustomTags = null!;
        internal OfdAttachmentWritePlan Attachments = null!;
        internal string DocumentPath = string.Empty, Directory = string.Empty, PublicLocation = string.Empty;
        internal string? DocumentLocation;
        internal XNamespace Namespace = XNamespace.None;
        internal bool Compression;
        internal long OutputReservation;
    }

    private static WritePlan PreparePlan(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token, OfdFontEmbeddingPlan? fontPlan = null)
    {
        OfdCompositeGraph.Validate(package, material, token, rejectRepeatedOwners: true);
        material.InspectStorage(package, token); material.GuardPreservedXml(package, token);
        var before = material.OutputUpperBoundBytes; material.EstimateWriter(package, token);
        material.ChargeCopy(1024L + (long)package.Pages.Count * IntPtr.Size * 8, token);
        OfdDocumentEnvelope.ValidateOwnedFragments(package, material, token);
        package.SourceEnvelope?.ValidatePageOrder(package, package.Pages.OrderBy(page => page.Index).ToArray(), material, token);
        OfdNavigationXml.Validate(package, package.Pages, token, material);
        OfdStandardXml.Validate(package, material, token);
        OfdAnnotationContract.ValidatePageSelection(package, package.Pages, material, token);
        OfdPageXmlContract.ValidateDocumentArea(package, token, material);
        foreach (var page in package.Pages) OfdPageXmlContract.ValidateForRewrite(package, page, token, material);
        if (package.Pages.Count == 0) OfdPageXmlContract.ValidateWritableDimensions(package.Options.DefaultPageWidthMillimeters, package.Options.DefaultPageHeightMillimeters);
        material.ChargeCopy(512L + (long)package.PreservedEntries.Count * 128 + (long)package.Pages.Count * 256, token);
        var plan = new WritePlan
        {
            Entries = new Dictionary<string, byte[]>(package.PreservedEntries, StringComparer.OrdinalIgnoreCase),
            Namespace = XNamespace.Get(package.Options.Namespace),
            DocumentPath = package.DocumentEntryPath ?? $"{package.Options.DocumentId}/Document.xml",
            Pages = package.Pages.OrderBy(page => page.Index).ToList(), Compression = package.Options.EnableDeflateCompression,
            OutputReservation = material.OutputUpperBoundBytes - before
        };
        plan.Directory = OfdPackagePath.GetDirectory(plan.DocumentPath);
        if (plan.Pages.Count == 0) plan.Pages.Add(new OfdPage { Index = 0,
            WidthMillimeters = package.Options.DefaultPageWidthMillimeters, HeightMillimeters = package.Options.DefaultPageHeightMillimeters });
        plan.Ids = new OfdIdPlan(package, material, token);
        plan.PageIds = plan.Pages.ToDictionary(page => page, page => plan.Ids.Claim(page, page.Id));
        plan.PagePaths = BuildPagePaths(plan.Pages, plan.Directory, material, token);
        plan.ElementIds = new Dictionary<OfdElement, string>();
        plan.LayerIds = BuildPageObjectIds(package, plan.Pages, plan.ElementIds, plan.Ids, material, token);
        foreach (var page in plan.Pages)
        foreach (var annotation in (page.Annotations ?? page.AnnotationOwnership.Original)?.Items ?? Enumerable.Empty<OfdAnnotation>())
        {
            plan.AnnotationIds.Add(annotation, plan.Ids.Claim(annotation, annotation.Id));
            foreach (var element in annotation.Appearance.Elements) plan.ElementIds.Add(element, plan.Ids.Claim(element, element.ObjectId));
            if (annotation.Source is null && !plan.NewAnnotationPaths.ContainsKey(page))
                plan.NewAnnotationPaths.Add(page, FreshAnnotationPath("Page_" + plan.PageIds[page]));
        }
        if (plan.NewAnnotationPaths.Count > 0) plan.NewAnnotationIndex = FreshAnnotationPath("Annotations");
        foreach (var resource in package.CompositeResources?.Items ?? Enumerable.Empty<OfdCompositeResource>())
        {
            plan.Ids.Claim(resource, resource.Id);
            foreach (var element in resource.Elements) plan.ElementIds.Add(element, plan.Ids.Claim(element, element.ObjectId));
        }
        plan.Images = BuildImageResources(package, plan.Ids, material, token); plan.Fonts = BuildFonts(package, plan.Ids);
        plan.PublicLocation = package.PublicResourceLocation ?? "PublicRes.xml";
        plan.DocumentLocation = package.DocumentResourceLocation ?? (plan.Images.Count > 0 || package.CompositeResources is not null ? "DocumentRes.xml" : null);
        var publicPath = material.ResolvePath(plan.DocumentPath, plan.PublicLocation, token);
        var knownPaths = new List<string> { publicPath };
        foreach (var path in plan.PagePaths.Values) knownPaths.Add(EntryInDirectory(OfdPackagePath.GetDirectory(path), "PageRes.xml"));
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation)) knownPaths.Add(material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token));
        material.ChargeCopy(32L * plan.Directory.Length + 1024, token);
        var outputRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Role(OfdConstants.OfdRootFile, "OFD"); Role(plan.DocumentPath, "Document");
        foreach (var path in plan.PagePaths.Values) Role(path, "Page");
        foreach (var path in plan.NewAnnotationPaths.Values) Role(path, "PageAnnot");
        if (plan.NewAnnotationIndex is not null) Role(plan.NewAnnotationIndex, "Annotations");
        foreach (var path in package.Pages.SelectMany(page => page.AnnotationOwnership.Entries.Keys).Distinct(StringComparer.OrdinalIgnoreCase)) Role(path, "PageAnnot");
        Role(publicPath, "Res");
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation)) Role(material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token), "Res");
        plan.Attachments = OfdAttachmentWriter.Prepare(package, plan.DocumentPath, plan.Namespace, outputRoles.Keys, material, token);
        foreach (var path in plan.Attachments.Entries.Keys)
            Role(path, path == plan.Attachments.DescriptorPath ? "Attachments" : "AttachmentPayload");
        plan.CustomTags = BuildCustomTags(package, plan.Entries, plan.Namespace, plan.DocumentPath, material, token);
        foreach (var path in plan.CustomTags.Entries.Keys) Role(path, "CustomTags");
        plan.Resources = new OfdResourceCatalog(plan.Entries, plan.DocumentPath, plan.Namespace, knownPaths, material, token, package.SourceIdentity, package, outputRoles.Keys);
        plan.Resources.PlanDocument(publicPath, plan.Namespace);
        foreach (var font in plan.Fonts.Resources)
            plan.Resources.PlanFont(font.Id,
                fontPlan is not null && fontPlan.Resources.TryGetValue(font.Resource, out var embedded) ? embedded : font.Resource,
                publicPath, plan.Namespace, fontPlan?.CanonicalPayloads == true);
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation))
        {
            var resourcePath = material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token);
            plan.Resources.PlanDocument(resourcePath, plan.Namespace);
            foreach (var image in plan.Images.Values.GroupBy(image => image.Id).Select(group => group.First()))
                plan.Resources.PlanImage(image.Id, image.Image, image.Format, resourcePath, plan.Namespace);
        }
        token.ThrowIfCancellationRequested(); return plan;
        string FreshAnnotationPath(string stem)
        {
            var ordinal = 0; string path;
            do
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(128L + 8L * (plan.Directory.Length + stem.Length), token);
                path = EntryInDirectory(plan.Directory, "Annots/" + stem + "_" + (ordinal++).ToString(CultureInfo.InvariantCulture) + ".xml");
            } while (plan.Entries.ContainsKey(path) || plan.NewAnnotationPaths.Values.Contains(path, StringComparer.OrdinalIgnoreCase));
            return path;
        }
        void Role(string path, string role)
        {
            material.ChargeCopy(128, token);
            if (outputRoles.TryGetValue(path, out var previous) && (previous != "Res" || role != "Res"))
                throw new InvalidDataException("Output control/page/resource roles collide.");
            outputRoles[path] = role;
        }
    }

    private Dictionary<string, byte[]> BuildEntries(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken cancellationToken, WritePlan plan)
    {
        var entries = plan.Entries; var ns = plan.Namespace; var documentPath = plan.DocumentPath; var docId = plan.Directory;
        var orderedPages = plan.Pages; var idAllocator = plan.Ids; var pageIds = plan.PageIds; var pagePaths = plan.PagePaths;
        var elementIds = plan.ElementIds; var layerIds = plan.LayerIds; var imageResources = plan.Images; var fonts = plan.Fonts;
        var publicResourceLocation = plan.PublicLocation; var documentResourceLocation = plan.DocumentLocation;
        var ofdXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "OFD",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XAttribute("Version", "1.0"),
                new XAttribute("DocType", package.Options.DocType),
                new XElement(ns + "DocBody",
                    new XElement(ns + "DocInfo",
                        BuildDocInfo(ns, package.Options.Metadata)
                    ),
                    new XElement(ns + "DocRoot", documentPath),
                    package.PreservedDocBodyElements.Select(xml =>
                        XElement.Parse(xml, LoadOptions.PreserveWhitespace))
                )));

        plan.Resources.Apply(); plan.Resources.Flush();
        BuildComposites(package, plan, material, cancellationToken);
        BuildPages(orderedPages, entries, ns, pagePaths, pageIds, layerIds, elementIds, fonts.TextBindings, imageResources, idAllocator, material, cancellationToken);
        foreach (var entry in plan.Attachments.Entries) entries[entry.Key] = entry.Value;
        foreach (var entry in plan.CustomTags.Entries) entries[entry.Key] = entry.Value;

        var docXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Document",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CommonData",
                    new XElement(ns + "MaxUnitID", idAllocator.MaxId),
                    new XElement(ns + "PageArea",
                        new XElement(ns + "PhysicalBox", BuildBox(
                            orderedPages[0].XMillimeters,
                            orderedPages[0].YMillimeters,
                            orderedPages[0].WidthMillimeters,
                            orderedPages[0].HeightMillimeters))
                    ),
                    new XElement(ns + "PublicRes", publicResourceLocation),
                    !string.IsNullOrWhiteSpace(documentResourceLocation)
                        ? new XElement(ns + "DocumentRes", documentResourceLocation)
                        : null,
                    package.PreservedCommonDataElements.Select(xml =>
                        XElement.Parse(xml, LoadOptions.PreserveWhitespace))
                ),
                new XElement(ns + "Pages",
                    orderedPages.Select((page, i) =>
                        new XElement(ns + "Page",
                            new XAttribute("ID", pageIds[page]),
                            new XAttribute("BaseLoc", material.RelativePath(documentPath, pagePaths[page], cancellationToken)))
                    )
                ),
                plan.CustomTags.Declarations,
                plan.Attachments.Declaration,
                package.PreservedDocumentElements.Select(xml =>
                    XElement.Parse(xml, LoadOptions.PreserveWhitespace))));
        foreach (var navigation in OfdNavigationXml.Write(package, pageIds, ns, cancellationToken, material))
            OfdXmlEnvelope.Singleton(docXml.Root!, navigation.Name.LocalName, navigation);

        OfdStandardXml.WriteMetadata(package, ofdXml.Root!.Element(ns + "DocBody")!.Element(ns + "DocInfo")!);
        OfdStandardXml.WriteVersions(package, ofdXml.Root.Element(ns + "DocBody")!);
        OfdStandardXml.WriteDocument(package, docXml.Root!);

        ofdXml = OfdDocumentEnvelopeWriter.Ofd(package, ofdXml, material, cancellationToken);
        docXml = OfdDocumentEnvelopeWriter.Document(package, docXml, orderedPages, pageIds, material, cancellationToken);
        BuildAnnotations(package, plan, docXml, material, cancellationToken);
        entries[OfdConstants.OfdRootFile] = ToUtf8Bytes(ofdXml, preserveWhitespace: true);
        entries[documentPath] = ToUtf8Bytes(docXml);
        return entries;
    }

    private static void BuildComposites(OfdDocumentPackage package, WritePlan plan, OfdModelMaterialization material, CancellationToken token)
    {
        if (package.CompositeResources is not { } collection) return;
        foreach (var group in collection.Items.GroupBy(resource => resource.SourceEntryPath ?? material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token)))
        {
            token.ThrowIfCancellationRequested();
            var xml = material.ReadXml(group.Key, plan.Entries[group.Key], token); var ns = xml.Root!.Name.Namespace;
            var container = xml.Root.Element(ns + "CompositeGraphicUnits");
            if (container is null) { container = new XElement(ns + "CompositeGraphicUnits"); xml.Root.Add(container); }
            foreach (var resource in group)
            {
                var node = resource.SourceXml is null ? new XElement(ns + "CompositeGraphicUnit") : OfdXmlEnvelope.Parse(resource.SourceXml);
                node.SetAttributeValue("ID", plan.Ids.Assigned(resource)); node.SetAttributeValue("Width", ToInvariant(resource.WidthMillimeters)); node.SetAttributeValue("Height", ToInvariant(resource.HeightMillimeters));
                var children = node.Elements().ToArray();
                if (resource.SourceXml is not null && children.Length != resource.Elements.Count) throw new InvalidDataException("Composite source slots changed.");
                for (var i = 0; i < resource.Elements.Count; i++)
                {
                    var emitted = BuildPrimitive(resource.Elements[i], ns, plan.ElementIds, plan.Fonts.TextBindings, plan.Images, plan.Ids)
                        ?? throw new NotSupportedException("Unsupported composite child.");
                    if (resource.SourceXml is null) node.Add(emitted); else children[i].ReplaceWith(emitted);
                }
                if (resource.SourceXml is null) container.Add(node);
                else
                {
                    var source = xml.Root.Elements(ns + "CompositeGraphicUnits").Elements().SingleOrDefault(n => n.Name == ns + "CompositeGraphicUnit" && n.Attribute("ID")?.Value == resource.Id)
                        ?? throw new InvalidDataException("Authenticated composite resource is missing.");
                    source.ReplaceWith(node);
                }
            }
            plan.Entries[group.Key] = ToUtf8Bytes(xml);
        }
    }

    private static void BuildAnnotations(OfdDocumentPackage package, WritePlan plan, XDocument document,
        OfdModelMaterialization material, CancellationToken token)
    {
        foreach (var page in plan.Pages)
        {
            var collection = page.Annotations;
            if (collection is null) continue;
            var owner = page.AnnotationOwnership;
            // A pristine imported collection preserves source bytes, including omitted flag spellings.
            if (owner.Original is not null && owner.Baseline!.Matches(collection, material, token)) continue;
            var items = new HashSet<OfdAnnotation>(collection.Items);
            foreach (var entry in owner.Entries)
            {
                token.ThrowIfCancellationRequested();
                var source = material.ReadXml(entry.Key, plan.Entries[entry.Key], token);
                var slots = source.Root!.Elements().ToArray(); var changed = false;
                foreach (var annotation in owner.Original?.Items ?? Enumerable.Empty<OfdAnnotation>())
                {
                    if (annotation.Source!.Entry != entry.Key) continue;
                    var slot = slots[annotation.Source.Slot];
                    if (!items.Contains(annotation)) { slot.Remove(); changed = true; continue; }
                    slot.ReplaceWith(Emit(annotation)); changed = true;
                }
                if (changed) plan.Entries[entry.Key] = ToUtf8Bytes(source);
            }
            if (plan.NewAnnotationPaths.TryGetValue(page, out var newPath))
            {
                var root = new XElement(plan.Namespace + "PageAnnot");
                foreach (var annotation in collection.Items.Where(a => a.Source is null)) root.Add(Emit(annotation));
                plan.Entries[newPath] = ToUtf8Bytes(new XDocument(root));
            }
        }
        if (plan.NewAnnotationIndex is not null)
        {
            var root = new XElement(plan.Namespace + "Annotations");
            foreach (var pair in plan.NewAnnotationPaths)
                root.Add(new XElement(plan.Namespace + "Page", new XAttribute("PageID", plan.PageIds[pair.Key]),
                    new XElement(plan.Namespace + "FileLoc", material.RelativePath(plan.NewAnnotationIndex, pair.Value, token))));
            plan.Entries[plan.NewAnnotationIndex] = ToUtf8Bytes(new XDocument(root));
            document.Root!.Add(new XElement(document.Root.Name.Namespace + "Annotations", material.RelativePath(plan.DocumentPath, plan.NewAnnotationIndex, token)));
        }
        XElement Emit(OfdAnnotation annotation)
        {
            token.ThrowIfCancellationRequested();
            var node = annotation.SourceXml is null ? new XElement(plan.Namespace + "Annot") : OfdXmlEnvelope.Parse(annotation.SourceXml);
            node.SetAttributeValue("ID", plan.AnnotationIds[annotation]); node.SetAttributeValue("Type", annotation.Type.ToString());
            node.SetAttributeValue("Creator", annotation.Creator); node.SetAttributeValue("LastModDate", annotation.LastModifiedDate!.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            node.SetAttributeValue("Subtype", annotation.Subtype);
            Flag("Visible", annotation.Visible); Flag("Print", annotation.Print); Flag("NoZoom", annotation.NoZoom); Flag("NoRotate", annotation.NoRotate); Flag("ReadOnly", annotation.ReadOnly);
            var ns = node.Name.Namespace; var remark = node.Element(ns + "Remark");
            if (annotation.Remark is null) remark?.Remove();
            else if (remark is null) node.AddFirst(new XElement(ns + "Remark", annotation.Remark));
            else if (remark.Value != annotation.Remark) { remark.Nodes().OfType<XText>().Remove(); remark.AddFirst(new XText(annotation.Remark)); }
            var appearance = annotation.Appearance;
            var local = appearance.SourceXml is null ? new XElement(ns + "Appearance") : OfdXmlEnvelope.Parse(appearance.SourceXml);
            var boundary = string.Join(" ", new[] { appearance.XMillimeters, appearance.YMillimeters, appearance.WidthMillimeters, appearance.HeightMillimeters }.Select(Number));
            var old = local.Attribute("Boundary")?.Value;
            if (!OfdBoxParser.TryParse(old, out var oldBox) || oldBox.x != appearance.XMillimeters || oldBox.y != appearance.YMillimeters || oldBox.w != appearance.WidthMillimeters || oldBox.h != appearance.HeightMillimeters)
                local.SetAttributeValue("Boundary", boundary);
            local.SetAttributeValue("CTM", appearance.Transform is null ? null : string.Join(" ", appearance.Transform.Select(Number)));
            if (annotation.Source is null) foreach (var element in appearance.Elements)
                local.Add(BuildPrimitive(element, ns, plan.ElementIds, plan.Fonts.TextBindings, plan.Images, plan.Ids) ?? throw new NotSupportedException("Unsupported local primitive."));
            if (annotation.AttachmentId is string attachmentId && annotation.Source is null)
            {
                var primitive = local.Elements().Single();
                primitive.AddFirst(new XElement(ns + "Actions", new XElement(ns + "Action", new XAttribute("Event", "CLICK"),
                    new XElement(ns + "GotoA", new XAttribute("AttachID", attachmentId)))));
            }
            var original = node.Element(ns + "Appearance"); if (original is null) node.Add(local); else original.ReplaceWith(local);
            return node;
            void Flag(string name, bool? value)
            {
                var literal = node.Attribute(name)?.Value;
                if (literal is not null && value == (literal is "true" or "1")) return;
                node.SetAttributeValue(name, value?.ToString().ToLowerInvariant());
            }
        }
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    }

    private static Dictionary<OfdPage, string> BuildPagePaths(IReadOnlyList<OfdPage> pages, string documentDirectory, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        var result = new Dictionary<OfdPage, string>();
        var used = new HashSet<string>(pages.Where(page => !string.IsNullOrWhiteSpace(page.SourceEntryPath))
            .Select(page => material.ResolvePath("OFD.xml", "/" + page.SourceEntryPath, cancellationToken)), StringComparer.OrdinalIgnoreCase);
        foreach (var page in pages)
        {
            var path = page.SourceEntryPath;
            if (string.IsNullOrWhiteSpace(path))
            {
                var index = page.Index;
                do
                {
                    path = material.ResolvePath("OFD.xml", "/" + EntryInDirectory(documentDirectory, $"Pages/Page_{index++}/Content.xml"), cancellationToken);
                } while (!used.Add(path));
            }
            else
            {
                path = material.ResolvePath("OFD.xml", "/" + path, cancellationToken);
            }

            if (result.Values.Contains(path, StringComparer.OrdinalIgnoreCase))
                throw new InvalidDataException("Multiple pages share the same source entry; clone repeated pages with a new source path.");
            result.Add(page, path!);
        }

        return result;
    }

    private sealed class FontBinding
    {
        internal OfdFontResource Resource { get; set; } = new();
        internal string Id { get; set; } = string.Empty;
    }

    private sealed class FontBindings
    {
        internal List<FontBinding> Resources { get; } = new();
        internal Dictionary<OfdTextElement, FontBinding> TextBindings { get; } = new();
    }

    private static FontBindings BuildFonts(OfdDocumentPackage package, OfdIdPlan idAllocator)
    {
        var result = new FontBindings();
        var byId = new Dictionary<string, FontBinding>(StringComparer.Ordinal);
        foreach (var font in package.Fonts)
        {
            if (!string.IsNullOrEmpty(font.Id) && byId.ContainsKey(font.Id))
                throw new InvalidDataException($"Duplicate font resource ID '{font.Id}'.");
            var binding = new FontBinding { Resource = font, Id = idAllocator.Claim(font, font.Id) };
            result.Resources.Add(binding);
            if (!string.IsNullOrEmpty(font.Id)) byId.Add(font.Id, binding);
        }

        foreach (var text in OfdCompositeGraph.Definitions(package).OfType<OfdTextElement>())
        {
            FontBinding? binding = null;
            if (!string.IsNullOrEmpty(text.FontResourceId)) byId.TryGetValue(text.FontResourceId!, out binding);
            var selectedFont = OfdFontSelection.Resolve(package.Fonts, text);
            binding ??= result.Resources.FirstOrDefault(font => ReferenceEquals(font.Resource, selectedFont));
            if (binding is null)
            {
                binding = new FontBinding
                {
                    Resource = new OfdFontResource { FontName = text.FontName },
                    Id = idAllocator.Claim(new object(), text.FontResourceId)
                };
                result.Resources.Add(binding);
                if (!string.IsNullOrEmpty(text.FontResourceId)) byId[text.FontResourceId!] = binding;
            }
            result.TextBindings[text] = binding;
        }
        return result;
    }

    private sealed class ImageResource
    {
        public OfdImageElement Image { get; set; } = new();

        public string Id { get; set; } = string.Empty;

        public string FileName { get; set; } = string.Empty;

        public string Format { get; set; } = "PNG";
    }

    private static Dictionary<OfdPage, Dictionary<string, string>> BuildPageObjectIds(
        OfdDocumentPackage package, IReadOnlyList<OfdPage> pages,
        IDictionary<OfdElement, string> elementIds,
        OfdIdPlan idAllocator, OfdModelMaterialization material, CancellationToken token)
    {
        var result = new Dictionary<OfdPage, Dictionary<string, string>>();
        foreach (var page in pages)
        {
            var pageLayerIds = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var group in page.Elements.GroupBy(GetLayerKey))
            {
                var preferredLayerId = group.First().LayerId;
                pageLayerIds[group.Key] = idAllocator.ClaimLayer(page, group, preferredLayerId);
                foreach (var element in group)
                {
                    elementIds[element] = idAllocator.Claim(element, element.ObjectId);
                }
            }

            var emptyCount = 0;
            if (page.SourceEntryPath is string path && package.PreservedEntries.TryGetValue(path, out var bytes))
            {
                var source = material.ReadXml(path, bytes, token); var sourceNs = source.Root!.Name.Namespace;
                var ordinal = 0;
                foreach (var layer in source.Root.Element(sourceNs + "Content")?.Elements(sourceNs + "Layer") ?? Enumerable.Empty<XElement>())
                {
                    token.ThrowIfCancellationRequested();
                    if (!layer.HasElements && !page.Elements.Any(element => element.LayerId == layer.Attribute("ID")?.Value))
                    { pageLayerIds["__empty_" + ordinal] = idAllocator.ClaimEmptyLayer(page, layer.Attribute("ID")?.Value); emptyCount++; }
                    ordinal++;
                }
            }
            if (page.Elements.Count == 0 && emptyCount == 0) pageLayerIds["__default"] = idAllocator.Allocate();
            result[page] = pageLayerIds;
        }

        return result;
    }

    private static Dictionary<OfdImageElement, ImageResource> BuildImageResources(
        OfdDocumentPackage package,
        OfdIdPlan idAllocator, OfdModelMaterialization material, CancellationToken token)
    {
        var resources = new Dictionary<OfdImageElement, ImageResource>();

        var byContent = new Dictionary<string, ImageResource>(StringComparer.Ordinal);
        var imported = new Dictionary<OfdSourceIdentity.Owner, (string Key, ImageResource Resource)>();
        var hashes = new Dictionary<byte[], string>();
        // Establish imported owners before fresh dedup references can reuse them.
        foreach (var image in OfdCompositeGraph.Definitions(package).OfType<OfdImageElement>()
            .OrderByDescending(image => idAllocator.TryImageOwner(image, out _)))
        {
            var format = ToOfdImageFormat(image.MediaType, image.FileName);
            if (!hashes.TryGetValue(image.Data, out var hash)) hashes[image.Data] = hash = BitConverter.ToString(material.HashBytes(image.Data, token)).Replace("-", string.Empty).ToLowerInvariant();
            var key = format + ":" + hash;
            ImageResource? resource;
            if (idAllocator.TryImageOwner(image, out var owner))
            {
                if (OfdSourceIdentity.ParseId(image.ResourceId) != owner.Number)
                    throw new InvalidDataException("Imported media resource ID was changed.");
                if (imported.TryGetValue(owner, out var existing))
                {
                    if (existing.Key != key) throw new InvalidDataException("Placements of one imported resource have conflicting payloads.");
                    resource = existing.Resource;
                }
                else
                {
                    resource = new ImageResource { Image = image, Id = idAllocator.ClaimImage(image), Format = format };
                    imported.Add(owner, (key, resource));
                    if (!byContent.ContainsKey(key)) byContent.Add(key, resource);
                }
            }
            else if (!byContent.TryGetValue(key, out resource))
            {
                resource = new ImageResource { Image = image, Id = idAllocator.ClaimImage(image), Format = format };
                byContent.Add(key, resource);
            }
            resources.Add(image, resource);
        }

        return resources;
    }

    private static void BuildPages(
        List<OfdPage> pages,
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        IReadOnlyDictionary<OfdPage, string> pagePaths,
        IReadOnlyDictionary<OfdPage, string> pageIds,
        IReadOnlyDictionary<OfdPage, Dictionary<string, string>> layerIds,
        IReadOnlyDictionary<OfdElement, string> elementIds,
        IReadOnlyDictionary<OfdTextElement, FontBinding> fontBindings,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources,
        OfdIdPlan idAllocator, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
            var content = new XElement(ns + "Content");

            var layerGroups = page.Elements.GroupBy(GetLayerKey);

            foreach (var layerGroup in layerGroups)
            {
                var firstElement = layerGroup.First();
                var layerId = layerIds[page][layerGroup.Key];
                var layerType = string.IsNullOrWhiteSpace(firstElement.LayerType) ? "Body" : firstElement.LayerType;
                var layer = new XElement(
                    ns + "Layer",
                    new XAttribute("ID", layerId),
                    new XAttribute("Type", layerType));

                foreach (var element in layerGroup)
                {
                    var emitted = BuildPrimitive(element, ns, elementIds, fontBindings, imageResources, idAllocator);
                    if (emitted is not null) layer.Add(emitted);
                }

                content.Add(layer);
            }

            if (page.SourceEntryPath is not null && entries.TryGetValue(page.SourceEntryPath, out var sourceBytes))
            {
                var source = material.ReadXml(page.SourceEntryPath, sourceBytes, cancellationToken);
                var originals = source.Root?.Element(source.Root.Name.Namespace + "Content")?.Elements(source.Root.Name.Namespace + "Layer").ToList() ?? [];
                var rebuilt = content.Elements().ToDictionary(layer => layer.Attribute("ID")!.Value, StringComparer.Ordinal);
                XElement? next = null;
                for (var position = originals.Count - 1; position >= 0; position--)
                {
                    var original = originals[position];
                    var originalId = original.Attribute("ID")?.Value;
                    XElement? existing = null;
                    if (originalId is not null) rebuilt.TryGetValue(originalId, out existing);
                    if (original.HasElements) { if (existing is not null) next = existing; continue; }
                    if (existing is not null)
                    {
                        // A caller explicitly filled this original layer through its ID.
                        existing.AddFirst(original.Nodes().OfType<XComment>().Select(comment => new XComment(comment.Value)));
                        next = existing;
                        continue;
                    }
                    var emptyLayer = new XElement(original); emptyLayer.Name = ns + "Layer";
                    emptyLayer.SetAttributeValue("ID", layerIds[page]["__empty_" + position]);
                    if (next is not null) next.AddBeforeSelf(emptyLayer);
                    else content.Add(emptyLayer);
                    next = emptyLayer;
                }
            }
            if (!content.HasElements) content.Add(new XElement(ns + "Layer", new XAttribute("ID", layerIds[page]["__default"]), new XAttribute("Type", "Body")));

            var pageRoot = new XElement(
                ns + "Page",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "Area", new XElement(ns + "PhysicalBox", BuildBox(
                    page.XMillimeters,
                    page.YMillimeters,
                    page.WidthMillimeters,
                    page.HeightMillimeters))));

            foreach (var preservedXml in page.PreservedPageElements)
            {
                if (!string.IsNullOrWhiteSpace(preservedXml))
                {
                    pageRoot.Add(XElement.Parse(preservedXml, LoadOptions.PreserveWhitespace));
                }
            }

            pageRoot.Add(content);
            var pageDocument = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                pageRoot);
            entries[pagePaths[page]] = ToUtf8Bytes(pageDocument);
        }
    }

    private static XElement? BuildPrimitive(OfdElement element, XNamespace ns,
        IReadOnlyDictionary<OfdElement, string> elementIds, IReadOnlyDictionary<OfdTextElement, FontBinding> fontBindings,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources, OfdIdPlan idAllocator)
    {
        var objectId = elementIds[element];
        if (element is OfdTextElement text)
        {
            fontBindings.TryGetValue(text, out var binding);
            var fontId = binding?.Id ?? throw new InvalidDataException("Missing planned font binding.");
            var (weight, italic) = ResolveTextStyle(text, binding?.Resource);
            XElement textObject;
            if (!string.IsNullOrWhiteSpace(text.SourceXml))
            {
                if (text.Runs.Any(run => run.PositionedGlyphs is not null))
                    throw new NotSupportedException("Imported SourceXml owns its glyph mapping; clear it before creating typed glyph output.");
                textObject = XElement.Parse(text.SourceXml!, LoadOptions.PreserveWhitespace);
                textObject.SetAttributeValue("ID", objectId);
                textObject.SetAttributeValue("Font", fontId);
                // The reader parsed Weight/Italic into the model, so the model
                // is authoritative for them here (like ID and Font). Resource
                // flags are not promoted into preserved XML.
                textObject.SetAttributeValue("Weight", text.Weight != OfdTextElement.DefaultWeight
                    ? text.Weight.ToString(CultureInfo.InvariantCulture) : null);
                textObject.SetAttributeValue("Italic", text.Italic ? "true" : null);
                ApplyEmphasis(textObject, text, binding?.Resource, weight, italic, ns);
            }
            else
            {
                var width = text.WidthMillimeters <= 0 ? Math.Max(text.Text.Length * text.FontSizeMillimeters * 0.5, text.FontSizeMillimeters) : text.WidthMillimeters;
                var height = text.HeightMillimeters <= 0 ? text.FontSizeMillimeters * 1.5 : text.HeightMillimeters;
                textObject = new XElement(ns + "TextObject",
                    new XAttribute("ID", objectId),
                    new XAttribute("Boundary", BuildBox(text.XMillimeters, text.YMillimeters, width, height)),
                    new XAttribute("Font", fontId),
                    new XAttribute("Size", ToInvariant(text.FontSizeMillimeters)),
                    weight != OfdTextElement.DefaultWeight
                        ? new XAttribute("Weight", weight.ToString(CultureInfo.InvariantCulture))
                        : null,
                    italic ? new XAttribute("Italic", "true") : null,
                    text.Transform is { Length: 6 }
                        ? new XAttribute("CTM", string.Join(" ", text.Transform.Select(OfdNumericFormat.Plain)))
                        : null,
                    text.FillColor.Red != 0 ||
                    text.FillColor.Green != 0 ||
                    text.FillColor.Blue != 0 ||
                    text.FillColor.Alpha != 255
                        ? new XElement(ns + "FillColor",
                            new XAttribute(
                                "Value",
                                $"{text.FillColor.Red} {text.FillColor.Green} {text.FillColor.Blue}"),
                            text.FillColor.Alpha != 255
                                ? new XAttribute("Alpha", text.FillColor.Alpha)
                                : null)
                        : null,
                    // New typed text is literal; encode its slash once for the source TextCode grammar.
                    // Retained source envelopes above already own their encoding and remain untouched.
                    text.Runs.Count > 0
                        ? text.Runs.Select(run => new XElement(ns + "TextCode",
                            new XAttribute("X", ToInvariant(run.XMillimeters)),
                            new XAttribute("Y", ToInvariant(run.YMillimeters)),
                            !string.IsNullOrWhiteSpace(run.DeltaX)
                                ? new XAttribute("DeltaX", run.DeltaX)
                                : null,
                            !string.IsNullOrWhiteSpace(run.DeltaY)
                                ? new XAttribute("DeltaY", run.DeltaY)
                                : null,
                            OfdTextRenderProjection.EncodeLiteral(run.Text, idAllocator.Material, idAllocator.Token)))
                        : new[]
                        {
                            new XElement(ns + "TextCode",
                                new XAttribute("X", "0"),
                                new XAttribute("Y", ToInvariant(Math.Max(text.FontSizeMillimeters, 1d))),
                                OfdTextRenderProjection.EncodeLiteral(text.Text, idAllocator.Material, idAllocator.Token))
                        });
                ApplyEmphasis(textObject, text, binding?.Resource, weight, italic, ns);
            }

            if (text.Runs.Any(run => run.PositionedGlyphs is not null))
                AddPositionedGlyphs(textObject, text, ns, idAllocator);

            ApplyClipping(textObject, text, idAllocator);
            AssertNestedIds(text, textObject, idAllocator);
            return textObject;
        }
        else if (element is OfdImageElement image)
        {
            var resource = imageResources[image];
            XElement imageObject;
            if (!string.IsNullOrWhiteSpace(image.SourceXml))
            {
                imageObject = XElement.Parse(image.SourceXml!, LoadOptions.PreserveWhitespace);
                imageObject.SetAttributeValue("ID", objectId);
                // Retain the source reference spelling when it still names the authenticated owner.
                imageObject.SetAttributeValue("ResourceID", idAllocator.TryImageOwner(image, out var owner) &&
                    OfdSourceIdentity.ParseId(image.ResourceId) == owner.Number ? image.ResourceId : resource.Id);
            }
            else
            {
                imageObject = new XElement(ns + "ImageObject",
                    new XAttribute("ID", objectId),
                    new XAttribute("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters)),
                    new XAttribute("CTM", BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0)),
                    new XAttribute("ResourceID", resource.Id));
            }

            imageObject.SetAttributeValue("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters));
            imageObject.SetAttributeValue("CTM", image.Transform is { Length: 6 }
                ? string.Join(" ", image.Transform.Select(OfdNumericFormat.Plain))
                : BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0));
            imageObject.SetAttributeValue("Alpha", image.Alpha == 255 ? null : (object)Math.Max(0, Math.Min(255, image.Alpha)));
            imageObject.Elements(imageObject.Name.Namespace + "Clips").Remove();
            ApplyClipping(imageObject, image, idAllocator);

            AssertNestedIds(image, imageObject, idAllocator);
            return imageObject;
        }
        else if (element is OfdCompositeElement composite)
        {
            var node = composite.SourceXml is null ? new XElement(ns + "CompositeObject") : OfdXmlEnvelope.Parse(composite.SourceXml);
            node.SetAttributeValue("ID", objectId); node.SetAttributeValue("ResourceID", idAllocator.Assigned(composite.Resource));
            node.SetAttributeValue("Boundary", BuildBox(composite.XMillimeters, composite.YMillimeters, composite.WidthMillimeters, composite.HeightMillimeters));
            node.SetAttributeValue("CTM", OfdClipXml.Format(composite.Transform)); ApplyClipping(node, composite, idAllocator); AssertNestedIds(composite, node, idAllocator); return node;
        }
        else if (element is OfdPathElement path)
        {
            var pathObject = string.IsNullOrWhiteSpace(path.SourceXml)
                ? new XElement(ns + "PathObject")
                : XElement.Parse(path.SourceXml!, LoadOptions.PreserveWhitespace);
            pathObject.SetAttributeValue("ID", objectId);
            pathObject.SetAttributeValue(
                "Boundary",
                BuildBox(path.XMillimeters, path.YMillimeters, path.WidthMillimeters, path.HeightMillimeters));
            pathObject.SetAttributeValue("LineWidth", ToInvariant(path.LineWidthMillimeters));
            pathObject.SetAttributeValue("Stroke", path.Stroke);
            pathObject.SetAttributeValue("Fill", path.Fill);
            if (path.FillRule is { } rule) pathObject.SetAttributeValue("Rule", rule == OfdFillRule.EvenOdd ? "Even-Odd" : "NonZero");
            pathObject.SetAttributeValue(
                "CTM",
                path.Transform is { Length: 6 }
                    ? string.Join(" ", path.Transform.Select(OfdNumericFormat.Plain))
                    : null);
            SetPathColor(pathObject, ns, "StrokeColor", path.Stroke ? path.StrokeColor : null);
            SetPathColor(pathObject, ns, "FillColor", path.Fill ? path.FillColor : null);

            var abbreviatedData = pathObject.Element(pathObject.Name.Namespace + "AbbreviatedData");
            if (abbreviatedData is null)
            {
                pathObject.Add(new XElement(pathObject.Name.Namespace + "AbbreviatedData", path.AbbreviatedData));
            }
            else
            {
                var literal = string.Concat(abbreviatedData.Nodes().OfType<XText>().Select(text => text.Value));
                if (literal != path.AbbreviatedData)
                {
                    abbreviatedData.Nodes().OfType<XText>().Remove();
                    abbreviatedData.AddFirst(new XText(path.AbbreviatedData));
                }
            }

            ApplyClipping(pathObject, path, idAllocator);
            AssertNestedIds(path, pathObject, idAllocator);
            return pathObject;
        }
        else if (element is OfdRawElement raw && !string.IsNullOrWhiteSpace(raw.Xml))
        {
            var rawElement = XElement.Parse(raw.Xml, LoadOptions.PreserveWhitespace);
            idAllocator.ApplyRawRootId(raw, rawElement, objectId);
            AssertNestedIds(raw, rawElement, idAllocator);
            return rawElement;
        }        return null;
    }

    private static void AddPositionedGlyphs(XElement textObject, OfdTextElement text, XNamespace ns, OfdIdPlan ids)
    {
        if (text.SourceXml is not null || text.Runs.Count != 1 || text.Runs[0].PositionedGlyphs is not { } positioned)
            throw new NotSupportedException("Typed positioned glyph output requires one newly owned TextCode.");
        var run = text.Runs[0];
        if (text.Text != run.Text || run.Text != positioned.LogicalText ||
            run.Text.Length is < 1 or > 8 || positioned.Glyphs.Count is < 1 or > 8 ||
            text.Transform is not null || text.Weight != OfdTextElement.DefaultWeight || text.Italic)
            throw new NotSupportedException("Typed glyph output exceeds the bounded horizontal source profile.");
        var latinCombiningPair = OfdPositionedGlyphProfile.IsLatinCombiningPair(run.Text);
        if (!latinCombiningPair && run.Text.Any(code => code < 0x21 || code > 0x7e))
            throw new NotSupportedException("Typed glyph output requires printable ASCII or one bounded Latin combining pair.");
        var glyphs = positioned.Glyphs;
        if (latinCombiningPair && glyphs.Count != 1)
            throw new NotSupportedException("The Latin combining profile requires one complete original-GID cluster.");
        if (glyphs.Count == run.Text.Length)
        {
            var sourceByGid = new Dictionary<uint, char>();
            var gidBySource = new Dictionary<char, uint>();
            for (var index = 0; index < glyphs.Count; index++)
            {
                if (glyphs[index].ClusterStart != positioned.SourceStart + index || glyphs[index].ClusterLength != 1)
                    throw new NotSupportedException("One-to-one glyph output needs one source character per glyph.");
                if (sourceByGid.TryGetValue(glyphs[index].GlyphId, out var prior) && prior != run.Text[index])
                    throw new NotSupportedException("One original GID has ambiguous PDF Unicode semantics.");
                if (gidBySource.TryGetValue(run.Text[index], out var priorGid) && priorGid != glyphs[index].GlyphId)
                    throw new NotSupportedException("One source character has conflicting original GIDs in a PDF font alias.");
                sourceByGid[glyphs[index].GlyphId] = run.Text[index];
                gidBySource[run.Text[index]] = glyphs[index].GlyphId;
            }
        }
        else if (glyphs.Count != 1 || run.Text.Length is not (2 or 3) ||
                 glyphs[0].ClusterStart != positioned.SourceStart || glyphs[0].ClusterLength != run.Text.Length)
            throw new NotSupportedException("Only a complete two/three-character ligature is in this output profile.");
        foreach (var glyph in glyphs)
        {
            if (Math.Abs(glyph.OriginXMillimeters) > 1_000_000 ||
                Math.Abs(glyph.OriginYMillimeters) > 1_000_000 ||
                Math.Abs(glyph.XAdvanceMillimeters) > 1_000_000 || glyph.YAdvanceMillimeters != 0)
                throw new NotSupportedException("Typed glyph IDs or positions exceed the verified outline profile.");
        }
        var first = glyphs[0];
        var deltasX = new List<string>(); var deltasY = new List<string>();
        for (var index = 1; index < glyphs.Count; index++)
        {
            var dx = glyphs[index].OriginXMillimeters - glyphs[index - 1].OriginXMillimeters;
            var dy = glyphs[index].OriginYMillimeters - glyphs[index - 1].OriginYMillimeters;
            if (Math.Abs(dx) > 1_000_000 || Math.Abs(dy) > 1_000_000)
                throw new NotSupportedException("Typed glyph origin delta exceeds the verified profile.");
            deltasX.Add(ToInvariant(dx)); deltasY.Add(ToInvariant(dy));
        }
        var sourceCode = textObject.Elements(ns + "TextCode").Single();
        sourceCode.SetAttributeValue("X", ToInvariant(first.OriginXMillimeters));
        sourceCode.SetAttributeValue("Y", ToInvariant(first.OriginYMillimeters));
        sourceCode.SetAttributeValue("DeltaX", deltasX.Count == 0 ? null : string.Join(" ", deltasX));
        sourceCode.SetAttributeValue("DeltaY", deltasY.Any(value => value != "0") ? string.Join(" ", deltasY) : null);
        sourceCode.Value = OfdTextRenderProjection.EncodeLiteral(run.Text, ids.Material, ids.Token);
        sourceCode.AddBeforeSelf(new XElement(ns + "CGTransform",
            new XAttribute("CodePosition", "0"),
            // This first producer profile is BMP and one TextCode: source-code
            // positions are one per character, independently of HB clusters.
            new XAttribute("CodeCount", run.Text.Length),
            new XAttribute("GlyphCount", glyphs.Count),
            new XElement(ns + "Glyphs", string.Join(" ", glyphs.Select(glyph =>
                glyph.GlyphId.ToString(CultureInfo.InvariantCulture))))));
    }

    /// <summary>
    /// CT_Text Weight/Italic are what viewers actually apply; the Bold/Italic
    /// flags on a name-only font resource are ignored by most of them. A text
    /// object bound to a bold or italic resource therefore also declares the
    /// style itself, unless the model already requests one explicitly.
    /// </summary>
    private static (int Weight, bool Italic) ResolveTextStyle(OfdTextElement text, OfdFontResource? resource)
    {
        var weight = text.Weight != OfdTextElement.DefaultWeight
            ? text.Weight
            : resource?.Bold == true ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight;
        weight = Math.Max(100, Math.Min(900, weight));
        return (weight, text.Italic || resource?.Italic == true);
    }

    /// <summary>
    /// Name-only CJK faces (SimSun/SimHei) have no bold/italic file. Spec
    /// <c>Weight</c>/<c>Italic</c> are ignored by several OFD-H viewers, so also
    /// stroke the glyphs (faux bold) and shear the CTM (faux italic). Fresh native
    /// DOCX may explicitly request a bold text stroke for an embedded regular
    /// face. Imported embedded text and invisible DualLayer text stay untouched.
    /// </summary>
    private static void ApplyEmphasis(
        XElement textObject, OfdTextElement text, OfdFontResource? resource, int weight, bool italic, XNamespace ns)
    {
        if (text.FillColor.Alpha == 0) return;
        var embedded = resource is { Data.Length: > 0 };
        if (embedded && (!text.EmitSyntheticBoldStroke || !string.IsNullOrWhiteSpace(text.SourceXml))) return;

        var size = text.FontSizeMillimeters > 0 ? text.FontSizeMillimeters : 4d;
        if (weight >= 600)
        {
            textObject.SetAttributeValue("Stroke", "true");
            textObject.SetAttributeValue("Fill", "true");
            textObject.SetAttributeValue("LineWidth", ToInvariant(embedded ? FontEmphasisMetrics.BoldStrokeWidth(size, text.Text) : Math.Max(0.08, size * 0.045)));
            if (!textObject.Elements().Any(node => node.Name.LocalName == "StrokeColor"))
            {
                var stroke = new XElement(ns + "StrokeColor",
                    new XAttribute("Value", $"{text.FillColor.Red} {text.FillColor.Green} {text.FillColor.Blue}"));
                if (text.FillColor.Alpha != 255)
                    stroke.SetAttributeValue("Alpha", text.FillColor.Alpha.ToString(CultureInfo.InvariantCulture));
                var fill = textObject.Elements().FirstOrDefault(node => node.Name.LocalName == "FillColor");
                if (fill is null) textObject.AddFirst(stroke);
                else fill.AddAfterSelf(stroke);
            }
        }

        // Embedded italic rendering is independent of this bold compatibility fix.
        if (embedded || !italic) return;
        const double shear = 0.2;
        if (textObject.Attribute("CTM")?.Value is string existingMatrix)
        {
            // Only fresh baseline-normalized runs opt into composition. Preserved
            // XML and arbitrary legacy run baselines retain their current CTM.
            if (!string.IsNullOrWhiteSpace(text.SourceXml) || text.Runs.Count == 0 ||
                text.Runs.Any(run => !string.IsNullOrWhiteSpace(run.DeltaY) || ToInvariant(run.YMillimeters) != ToInvariant(size))) return;
            var matrix = existingMatrix.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
            var writtenSize = double.Parse(ToInvariant(size), CultureInfo.InvariantCulture);
            var composition = OfdTextEmphasis.ComposeNameOnlyItalic(matrix, writtenSize);
            textObject.SetAttributeValue("CTM", string.Join(" ", composition.Matrix.Select(OfdNumericFormat.Plain)));
            textObject.SetAttributeValue(OfdTextEmphasis.FauxItalicFactor,
                string.Join(" ", composition.Factor.Select(OfdNumericFormat.Plain)));
            // Graphics uses the physical page as its viewport. No width expansion.
            return;
        }
        var factor = BuildMatrix(1, 0, -shear, 1, shear * size, 0);
        textObject.SetAttributeValue("CTM", factor);
        textObject.SetAttributeValue(OfdTextEmphasis.FauxItalicFactor, factor);
        if (textObject.Attribute("Boundary")?.Value is not string box) return;
        var parts = box.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 4) return;
        if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var x) ||
            !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y) ||
            !double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var w) ||
            !double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var h))
            return;
        textObject.SetAttributeValue("Boundary", BuildBox(x, y, w + shear * h, h));
    }

    private static void ApplyClipping(XElement xml, OfdElement element, OfdIdPlan idAllocator)
    {
        if (element.Clips is null && element.ClippingXml is null) return;
        xml.Elements(xml.Name.Namespace + "Clips").Remove();
        var clip = OfdClipXml.Write(element, xml.Name.Namespace, idAllocator.Material, idAllocator.Token);
        if (clip is not null)
        {
            var actions = xml.Elements(xml.Name.Namespace + "Actions").LastOrDefault();
            if (actions is null) xml.AddFirst(clip); else actions.AddAfterSelf(clip);
        }
    }

    private static string GetLayerKey(OfdElement element)
    {
        var id = element.LayerId ?? string.Empty;
        var type = string.IsNullOrWhiteSpace(element.LayerType) ? "Body" : element.LayerType;
        return $"{id}\u001f{type}";
    }

    private static void AssertNestedIds(OfdElement model, XElement root, OfdIdPlan allocator)
    {
        allocator.AssertNestedIds(model, root);
    }

    private static void SetPathColor(
        XElement pathObject,
        XNamespace ns,
        string localName,
        OfdColor? color)
    {
        var colorElement = pathObject.Elements()
            .FirstOrDefault(x => x.Name.LocalName == localName);
        if (color is null)
        {
            colorElement?.Remove();
            return;
        }

        if (colorElement is null)
        {
            colorElement = new XElement(ns + localName);
            pathObject.AddFirst(colorElement);
        }

        colorElement.SetAttributeValue("Value", $"{color.Red} {color.Green} {color.Blue}");
        // The model color is resolved RGB. A preserved source ColorSpace or
        // palette index would reinterpret these three components on reread.
        colorElement.SetAttributeValue("ColorSpace", null);
        colorElement.SetAttributeValue("Index", null);
        colorElement.SetAttributeValue("Alpha", color.Alpha == 255 ? null : color.Alpha.ToString(CultureInfo.InvariantCulture));
    }

    // ZIP entry names are relative even when their owning document is at the package root.
    private static string EntryInDirectory(string directory, string relativePath) =>
        directory.Length == 0 ? relativePath : directory + "/" + relativePath;

    private sealed class CustomTagWritePlan
    {
        internal readonly List<XElement> Declarations = new();
        internal readonly Dictionary<string, byte[]> Entries = new(StringComparer.OrdinalIgnoreCase);
    }

    private static CustomTagWritePlan BuildCustomTags(OfdDocumentPackage package, IDictionary<string, byte[]> entries,
        XNamespace ns, string documentPath, OfdModelMaterialization material, CancellationToken token)
    {
        var result = new CustomTagWritePlan(); var source = package.CustomTagSource;
        source?.Validate(package, material, token);
        if (source is not null)
            foreach (var declaration in source.Declarations) result.Declarations.Add(OfdXmlEnvelope.Parse(declaration));
        if (source?.IsUnchanged(package) == true) return result;
        var modeledPaths = new HashSet<string>(source?.DetailPaths ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        foreach (var indexPath in source?.IndexPaths ?? Enumerable.Empty<string>())
        {
            if (!entries.TryGetValue(indexPath, out var indexBytes)) continue;
            var index = material.ReadXml(indexPath, indexBytes, token);
            foreach (var declaration in index.Root?.Elements().Where(node => node.Name.LocalName == "CustomTag") ?? Enumerable.Empty<XElement>())
            {
                var location = declaration.Elements().FirstOrDefault(node => node.Name.LocalName == "FileLoc")?.Value ?? declaration.Attribute("FileLoc")?.Value;
                if (string.IsNullOrWhiteSpace(location)) continue;
                if (modeledPaths.Contains(material.ResolvePath(indexPath, location!, token)) &&
                    (declaration.Name != ns + "CustomTag" || declaration.Attribute("TypeID")?.Value != "EMR" ||
                     declaration.Attribute("NameSpace")?.Value != OfdCustomTagSource.EmrNamespace))
                    throw new NotSupportedException("Cannot rewrite an EMR payload also referenced by an opaque custom tag schema.");
            }
        }
        var emitted = new HashSet<string>(StringComparer.Ordinal);
        var details = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in source?.DetailPaths ?? Enumerable.Empty<string>())
        {
            token.ThrowIfCancellationRequested();
            var xml = material.ReadXml(path, entries[path], token);
            foreach (var tag in xml.Root!.Elements(ns + "Tag").ToArray())
            {
                var key = tag.Attribute("Key")?.Value;
                if (string.IsNullOrWhiteSpace(key) || tag.Attribute("Value") is null) continue;
                if (package.CustomTags.TryGetValue(key!, out var value))
                {
                    if (source is null || !source.TryOriginalValue(key!, out var original) || value != original) tag.SetAttributeValue("Value", value);
                    emitted.Add(key!);
                }
                else
                {
                    tag.SetAttributeValue("Key", null); tag.SetAttributeValue("Value", null);
                    // A removed modeled value must not discard extensions carried by its slot.
                    if (!tag.HasElements && !tag.Nodes().Any() && !tag.Attributes().Any(a => !a.IsNamespaceDeclaration)) tag.Remove();
                }
            }
            details.Add(path, xml);
        }
        var additions = package.CustomTags.Where(pair => !emitted.Contains(pair.Key)).ToArray();
        if (additions.Length > 0)
        {
            var detailPath = details.Keys.FirstOrDefault();
            if (detailPath is null)
            {
                // Never overwrite an opaque payload, even at the SDK's conventional path.
                detailPath = Fresh("Tags/CustomTag_EMR", ".xml");
                details.Add(detailPath, new XDocument(new XElement(ns + "EMRTags", new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName))));
                var listPath = source?.IndexPaths.FirstOrDefault(path => entries.ContainsKey(path) &&
                    material.ReadXml(path, entries[path], token).Root?.Name == ns + "CustomTags");
                XDocument list;
                if (listPath is null)
                {
                    listPath = Fresh("Tags/CustomTags", ".xml");
                    list = new XDocument(new XElement(ns + "CustomTags", new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName)));
                    result.Declarations.Add(new XElement(ns + "CustomTags", material.RelativePath(documentPath, listPath, token)));
                }
                else list = material.ReadXml(listPath, entries[listPath], token);
                list.Root!.Add(new XElement(ns + "CustomTag", new XAttribute("TypeID", "EMR"),
                    new XAttribute("NameSpace", OfdCustomTagSource.EmrNamespace),
                    new XElement(ns + "FileLoc", material.RelativePath(listPath, detailPath, token))));
                result.Entries.Add(listPath, ToUtf8Bytes(list));
            }
            foreach (var pair in additions) details[detailPath].Root!.Add(new XElement(ns + "Tag", new XAttribute("Key", pair.Key), new XAttribute("Value", pair.Value)));
        }
        // A legacy payload without an active declaration gains a compatible index on an actual edit.
        if (details.Count > 0 && result.Declarations.Count == 0 && package.CustomTags.Count > 0)
        {
            var listPath = Fresh("Tags/CustomTags", ".xml");
            var list = new XDocument(new XElement(ns + "CustomTags", new XElement(ns + "CustomTag",
                new XAttribute("TypeID", "EMR"), new XAttribute("NameSpace", OfdCustomTagSource.EmrNamespace),
                new XElement(ns + "FileLoc", material.RelativePath(listPath, details.Keys.First(), token)))));
            result.Entries.Add(listPath, ToUtf8Bytes(list));
            result.Declarations.Add(new XElement(ns + "CustomTags", material.RelativePath(documentPath, listPath, token)));
        }
        foreach (var pair in details) result.Entries.Add(pair.Key, ToUtf8Bytes(pair.Value));
        return result;

        string Fresh(string stem, string suffix)
        {
            var directory = OfdPackagePath.GetDirectory(documentPath); var ordinal = 0; string path;
            do
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(256L + 8L * (directory.Length + stem.Length), token);
                path = EntryInDirectory(directory, stem + (ordinal == 0 ? "" : "_" + ordinal.ToString(CultureInfo.InvariantCulture)) + suffix); ordinal++;
            } while (entries.ContainsKey(path) || result.Entries.ContainsKey(path) || details.ContainsKey(path));
            return path;
        }
    }

    private static IEnumerable<XElement> BuildDocInfo(XNamespace ns, OfdMetadata metadata)
    {
        if (!string.IsNullOrWhiteSpace(metadata.Title))
        {
            yield return new XElement(ns + "Title", metadata.Title);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Author))
        {
            yield return new XElement(ns + "Author", metadata.Author);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Subject))
        {
            yield return new XElement(ns + "Subject", metadata.Subject);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Keywords))
        {
            yield return new XElement(ns + "Keywords", metadata.Keywords);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Creator))
        {
            yield return new XElement(ns + "Creator", metadata.Creator);
        }

        if (metadata.CreationDate.HasValue)
        {
            yield return new XElement(ns + "CreationDate", metadata.CreationDate.Value.ToString("O", CultureInfo.InvariantCulture));
        }

        if (metadata.ModificationDate.HasValue)
        {
            yield return new XElement(ns + "ModDate", metadata.ModificationDate.Value.ToString("O", CultureInfo.InvariantCulture));
        }
    }

    private static string BuildBox(double x, double y, double w, double h)
    {
        return string.Format(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} {3:0.###}", x, y, w, h);
    }

    private static string BuildMatrix(double a, double b, double c, double d, double e, double f)
    {
        return string.Format(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} {3:0.###} {4:0.###} {5:0.###}", a, b, c, d, e, f);
    }

    private static string ToInvariant(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string GetImageExtension(string format)
    {
        return format switch
        {
            "JPEG" => ".jpg",
            "JPG" => ".jpg",
            "BMP" => ".bmp",
            "TIFF" => ".tiff",
            _ => ".png"
        };
    }

    private static string ToOfdImageFormat(string mediaType, string fileName)
    {
        var normalizedMediaType = mediaType?.Trim().ToLowerInvariant();
        if (normalizedMediaType == "image/jpeg" || normalizedMediaType == "image/jpg")
        {
            return "JPEG";
        }

        if (normalizedMediaType == "image/bmp")
        {
            return "BMP";
        }

        if (normalizedMediaType == "image/tiff")
        {
            return "TIFF";
        }

        var extension = Path.GetExtension(fileName)?.TrimStart('.').ToUpperInvariant();
        return extension switch
        {
            "JPG" => "JPEG",
            "JPEG" => "JPEG",
            "BMP" => "BMP",
            "TIFF" => "TIFF",
            _ => "PNG"
        };
    }

    private static string ResolveEntryPath(string basePath, string relativePath)
    {
        var normalizedBase = basePath.Replace('\\', '/').TrimStart('/');
        var separator = normalizedBase.LastIndexOf('/');
        var baseDirectory = separator < 0 ? string.Empty : normalizedBase.Substring(0, separator);
        var combined = relativePath.StartsWith("/", StringComparison.Ordinal)
            ? relativePath.TrimStart('/')
            : $"{baseDirectory}/{relativePath}";
        var segments = new Stack<string>();
        foreach (var segment in combined.Split('/'))
        {
            if (string.IsNullOrWhiteSpace(segment) || segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.Pop();
                }

                continue;
            }

            segments.Push(segment);
        }

        return string.Join("/", segments.Reverse());
    }

    private static byte[] ToUtf8Bytes(XDocument xml, bool preserveWhitespace = false)
    {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(false), 1024, leaveOpen: true))
        {
            xml.Save(writer, SaveOptions.OmitDuplicateNamespaces | (preserveWhitespace ? SaveOptions.DisableFormatting : SaveOptions.None));
        }

        return ms.ToArray();
    }

}
