using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Packaging;

public sealed class OfdPackageWriter
{
    public async Task WriteAsync(OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        _ = await WriteWithResultAsync(package, destination, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes the package and reports removed page resources and invalidated signatures.</summary>
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
        var plan = PreparePlan(package, material, cancellationToken);
        var entries = BuildEntries(package, material, cancellationToken, plan);
        var result = OfdPackagePruner.Prune(package, entries, cancellationToken, material);
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
        internal Dictionary<OfdPage, Dictionary<string, string>> LayerIds = null!;
        internal Dictionary<OfdImageElement, ImageResource> Images = null!;
        internal FontBindings Fonts = null!;
        internal OfdIdPlan Ids = null!;
        internal OfdResourceCatalog Resources = null!;
        internal string DocumentPath = string.Empty, Directory = string.Empty, PublicLocation = string.Empty;
        internal string? DocumentLocation;
        internal XNamespace Namespace = XNamespace.None;
        internal bool Compression;
        internal long OutputReservation;
    }

    private static WritePlan PreparePlan(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        material.InspectStorage(package, token); material.GuardPreservedXml(package, token);
        var before = material.OutputUpperBoundBytes; material.EstimateWriter(package, token);
        material.ChargeCopy(1024L + (long)package.Pages.Count * IntPtr.Size * 8, token);
        OfdDocumentEnvelope.ValidateOwnedFragments(package, material, token);
        package.SourceEnvelope?.ValidatePageOrder(package, package.Pages.OrderBy(page => page.Index).ToArray(), material, token);
        OfdNavigationXml.Validate(package, package.Pages, token, material);
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
        plan.Images = BuildImageResources(plan.Pages, plan.Ids, material, token); plan.Fonts = BuildFonts(package, plan.Ids);
        plan.PublicLocation = package.PublicResourceLocation ?? "PublicRes.xml";
        plan.DocumentLocation = package.DocumentResourceLocation ?? (plan.Images.Count > 0 ? "DocumentRes.xml" : null);
        var publicPath = material.ResolvePath(plan.DocumentPath, plan.PublicLocation, token);
        var knownPaths = new List<string> { publicPath };
        foreach (var path in plan.PagePaths.Values) knownPaths.Add(OfdPackagePath.GetDirectory(path) + "/PageRes.xml");
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation)) knownPaths.Add(material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token));
        material.ChargeCopy(32L * plan.Directory.Length + 1024, token);
        var outputRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        Role(OfdConstants.OfdRootFile, "OFD"); Role(plan.DocumentPath, "Document");
        foreach (var path in plan.PagePaths.Values) Role(path, "Page");
        Role(publicPath, "Res");
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation)) Role(material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token), "Res");
        if (package.Attachments.Count > 0)
        {
            Role($"{plan.Directory}/Attachs/Attachments.xml", "Attachments");
            for (var index = 0; index < package.Attachments.Count; index++)
            {
                var attachment = package.Attachments[index]; if (attachment.IsExternal) continue;
                material.ChargeCopy(8L * attachment.Name.Length + 4L * plan.Directory.Length + 256, token);
                Role($"{plan.Directory}/Attachs/Attach_{index + 1}_{SanitizeFileName(attachment.Name)}", "AttachmentPayload");
            }
        }
        if (package.CustomTags.Count > 0) { Role($"{plan.Directory}/Tags/CustomTags.xml", "Tags"); Role($"{plan.Directory}/Tags/CustomTag_EMR.xml", "TagsDetail"); }
        plan.Resources = new OfdResourceCatalog(plan.Entries, plan.DocumentPath, plan.Namespace, knownPaths, material, token, package.SourceIdentity, package, outputRoles.Keys);
        plan.Resources.PlanDocument(publicPath, plan.Namespace);
        foreach (var font in plan.Fonts.Resources) plan.Resources.PlanFont(font.Id, font.Resource, publicPath, plan.Namespace);
        if (!string.IsNullOrWhiteSpace(plan.DocumentLocation))
        {
            var resourcePath = material.ResolvePath(plan.DocumentPath, plan.DocumentLocation!, token);
            plan.Resources.PlanDocument(resourcePath, plan.Namespace);
            foreach (var image in plan.Images.Values.GroupBy(image => image.Id).Select(group => group.First()))
                plan.Resources.PlanImage(image.Id, image.Image, image.Format, resourcePath, plan.Namespace);
        }
        token.ThrowIfCancellationRequested(); return plan;
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
        BuildPages(orderedPages, entries, ns, pagePaths, pageIds, layerIds, elementIds, fonts.TextBindings, imageResources, idAllocator, material, cancellationToken);
        BuildAttachments(package, entries, ns, docId);
        BuildCustomTags(package, entries, ns, docId);

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
                package.CustomTags.Count > 0
                    ? new XElement(ns + "CustomTags", "Tags/CustomTags.xml")
                    : null,
                package.Attachments.Count > 0
                    ? new XElement(ns + "Attachments", "Attachs/Attachments.xml")
                    : null,
                package.PreservedDocumentElements.Select(xml =>
                    XElement.Parse(xml, LoadOptions.PreserveWhitespace))));
        foreach (var navigation in OfdNavigationXml.Write(package, pageIds, ns, cancellationToken, material))
            OfdXmlEnvelope.Singleton(docXml.Root!, navigation.Name.LocalName, navigation);

        ofdXml = OfdDocumentEnvelopeWriter.Ofd(package, ofdXml, material, cancellationToken);
        docXml = OfdDocumentEnvelopeWriter.Document(package, docXml, orderedPages, pageIds, material, cancellationToken);
        entries[OfdConstants.OfdRootFile] = ToUtf8Bytes(ofdXml);
        entries[documentPath] = ToUtf8Bytes(docXml);
        return entries;
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
                    path = material.ResolvePath("OFD.xml", $"/{documentDirectory}/Pages/Page_{index++}/Content.xml", cancellationToken);
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

        foreach (var text in package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>())
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
        IReadOnlyList<OfdPage> pages,
        OfdIdPlan idAllocator, OfdModelMaterialization material, CancellationToken token)
    {
        var resources = new Dictionary<OfdImageElement, ImageResource>();

        var byContent = new Dictionary<string, ImageResource>(StringComparer.Ordinal);
        var imported = new Dictionary<OfdSourceIdentity.Owner, (string Key, ImageResource Resource)>();
        var hashes = new Dictionary<byte[], string>();
        // Establish imported owners before fresh dedup references can reuse them.
        foreach (var image in pages.SelectMany(x => x.Elements).OfType<OfdImageElement>()
            .OrderByDescending(image => idAllocator.TryImageOwner(image, out _)))
        {
            var format = ToOfdImageFormat(image.MediaType, image.FileName);
            if (!hashes.TryGetValue(image.Data, out var hash)) hashes[image.Data] = hash = BitConverter.ToString(material.HashBytes(image.Data, token)).Replace("-", string.Empty).ToLowerInvariant();
            var key = format + ":" + hash;
            ImageResource? resource;
            if (idAllocator.TryImageOwner(image, out var owner))
            {
                if (image.ResourceId != owner.Id) throw new InvalidDataException("Imported media resource ID was changed.");
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
                    var objectId = elementIds[element];
                    if (element is OfdTextElement text)
                    {
                        fontBindings.TryGetValue(text, out var binding);
                        var fontId = binding?.Id ?? throw new InvalidDataException("Missing planned font binding.");
                        var (weight, italic) = ResolveTextStyle(text, binding?.Resource);
                        XElement textObject;
                        if (!string.IsNullOrWhiteSpace(text.SourceXml))
                        {
                            textObject = XElement.Parse(text.SourceXml!, LoadOptions.PreserveWhitespace);
                            textObject.SetAttributeValue("ID", objectId);
                            textObject.SetAttributeValue("Font", fontId);
                            // The reader parsed Weight/Italic into the model, so the model
                            // is authoritative for them here (like ID and Font). Resource
                            // flags are not promoted into preserved XML.
                            textObject.SetAttributeValue("Weight", text.Weight != OfdTextElement.DefaultWeight
                                ? text.Weight.ToString(CultureInfo.InvariantCulture) : null);
                            textObject.SetAttributeValue("Italic", text.Italic ? "true" : null);
                            ApplyNameOnlyEmphasis(textObject, text, binding?.Resource, weight, italic, ns);
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
                                    ? new XAttribute("CTM", string.Join(" ", text.Transform.Select(ToInvariant)))
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
                                        run.Text))
                                    : new[]
                                    {
                                        new XElement(ns + "TextCode",
                                            new XAttribute("X", "0"),
                                            new XAttribute("Y", ToInvariant(Math.Max(text.FontSizeMillimeters, 1d))),
                                            text.Text)
                                    });
                            ApplyNameOnlyEmphasis(textObject, text, binding?.Resource, weight, italic, ns);
                        }

                        ApplyClipping(textObject, text);
                        AssertNestedIds(text, textObject, idAllocator);
                        layer.Add(textObject);
                    }
                    else if (element is OfdImageElement image)
                    {
                        var resource = imageResources[image];
                        XElement imageObject;
                        if (!string.IsNullOrWhiteSpace(image.SourceXml))
                        {
                            imageObject = XElement.Parse(image.SourceXml!, LoadOptions.PreserveWhitespace);
                            imageObject.SetAttributeValue("ID", objectId);
                            imageObject.SetAttributeValue("ResourceID", resource.Id);
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
                            ? string.Join(" ", image.Transform.Select(ToInvariant))
                            : BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0));
                        imageObject.SetAttributeValue("Alpha", image.Alpha == 255 ? null : (object)Math.Max(0, Math.Min(255, image.Alpha)));
                        imageObject.Elements(imageObject.Name.Namespace + "Clips").Remove();
                        ApplyClipping(imageObject, image);

                        AssertNestedIds(image, imageObject, idAllocator);
                        layer.Add(imageObject);
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
                        pathObject.SetAttributeValue(
                            "CTM",
                            path.Transform is { Length: 6 }
                                ? string.Join(" ", path.Transform.Select(ToInvariant))
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

                        ApplyClipping(pathObject, path);
                        AssertNestedIds(path, pathObject, idAllocator);
                        layer.Add(pathObject);
                    }
                    else if (element is OfdRawElement raw && !string.IsNullOrWhiteSpace(raw.Xml))
                    {
                        var rawElement = XElement.Parse(raw.Xml, LoadOptions.PreserveWhitespace);
                        idAllocator.ApplyRawRootId(raw, rawElement, objectId);
                        AssertNestedIds(raw, rawElement, idAllocator);
                        layer.Add(rawElement);
                    }
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
    /// stroke the glyphs (faux bold) and shear the CTM (faux italic). Invisible
    /// DualLayer text and embedded faces are left untouched.
    /// </summary>
    private static void ApplyNameOnlyEmphasis(
        XElement textObject, OfdTextElement text, OfdFontResource? resource, int weight, bool italic, XNamespace ns)
    {
        if (text.FillColor.Alpha == 0) return;
        if (resource is { Data.Length: > 0 }) return;

        var size = text.FontSizeMillimeters > 0 ? text.FontSizeMillimeters : 4d;
        if (weight >= 600)
        {
            textObject.SetAttributeValue("Stroke", "true");
            textObject.SetAttributeValue("Fill", "true");
            textObject.SetAttributeValue("LineWidth", ToInvariant(Math.Max(0.08, size * 0.045)));
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

        if (!italic || textObject.Attribute("CTM") is not null) return;
        const double shear = 0.2;
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

    private static void ApplyClipping(XElement xml, OfdElement element)
    {
        if (element.ClippingXml is null) return;
        xml.Elements(xml.Name.Namespace + "Clips").Remove();
        if (!string.IsNullOrWhiteSpace(element.ClippingXml))
        {
            var clip = XElement.Parse(element.ClippingXml!, LoadOptions.PreserveWhitespace);
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
        colorElement.SetAttributeValue("Alpha", color.Alpha == 255 ? null : color.Alpha.ToString(CultureInfo.InvariantCulture));
    }

    private static void BuildAttachments(OfdDocumentPackage package, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        if (package.Attachments.Count == 0)
        {
            return;
        }

        var attachments = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Attachments",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                package.Attachments.Select((attachment, index) =>
                {
                    var fileName = $"Attach_{index + 1}_{SanitizeFileName(attachment.Name)}";
                    if (!attachment.IsExternal)
                    {
                        entries[$"{docId}/Attachs/{fileName}"] = attachment.Data;
                    }

                    var fileLoc = attachment.IsExternal
                        ? attachment.ExternalPath ?? string.Empty
                        : fileName;
                    return new XElement(ns + "Attachment",
                        new XAttribute("ID", attachment.Id ?? $"Attachment_{index + 1}"),
                        new XAttribute("Name", attachment.Name),
                        new XAttribute("Format", attachment.Format ?? attachment.MediaType),
                        attachment.CreationDate.HasValue
                            ? new XAttribute("CreationDate", attachment.CreationDate.Value.ToString("O", CultureInfo.InvariantCulture))
                            : null,
                        attachment.ModificationDate.HasValue
                            ? new XAttribute("ModDate", attachment.ModificationDate.Value.ToString("O", CultureInfo.InvariantCulture))
                            : null,
                        attachment.SizeKilobytes.HasValue
                            ? new XAttribute("Size", ToInvariant(attachment.SizeKilobytes.Value))
                            : null,
                        attachment.Visible.HasValue
                            ? new XAttribute("Visible", attachment.Visible.Value)
                            : null,
                        !string.IsNullOrWhiteSpace(attachment.Usage)
                            ? new XAttribute("Usage", attachment.Usage)
                            : null,
                        new XElement(ns + "FileLoc", fileLoc));
                })));

        entries[$"{docId}/Attachs/Attachments.xml"] = ToUtf8Bytes(attachments);
    }

    private static void BuildCustomTags(OfdDocumentPackage package, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        if (package.CustomTags.Count == 0)
        {
            return;
        }

        var list = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "CustomTags",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CustomTag",
                    new XAttribute("TypeID", "EMR"),
                    new XAttribute("NameSpace", "urn:ofdrw-net:custom-tags:emr"),
                    new XElement(ns + "FileLoc", "CustomTag_EMR.xml"))));

        entries[$"{docId}/Tags/CustomTags.xml"] = ToUtf8Bytes(list);

        var detail = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "EMRTags",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                package.CustomTags.Select(x => new XElement(ns + "Tag", new XAttribute("Key", x.Key), new XAttribute("Value", x.Value)))));

        entries[$"{docId}/Tags/CustomTag_EMR.xml"] = ToUtf8Bytes(detail);
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

    private static string SanitizeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Attachment.bin";
        }

        var invalid = Path.GetInvalidFileNameChars();
        var chars = value.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        return new string(chars);
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

    private static byte[] ToUtf8Bytes(XDocument xml)
    {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(false), 1024, leaveOpen: true))
        {
            xml.Save(writer, SaveOptions.OmitDuplicateNamespaces);
        }

        return ms.ToArray();
    }

}
