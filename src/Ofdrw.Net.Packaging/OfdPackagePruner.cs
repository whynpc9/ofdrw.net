using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal sealed class OfdPackagePruner
{
    private readonly OfdModelMaterialization material;
    private readonly CancellationToken operationToken;
    private readonly OfdOwnedPruneDispositions? ownedDispositions;
    private OfdPackagePruner(OfdModelMaterialization material, CancellationToken token, OfdOwnedPruneDispositions? ownedDispositions = null) { this.material = material; operationToken = token; this.ownedDispositions = ownedDispositions; }
    internal sealed class XmlScanCache(Func<byte[], XDocument> parse)
    {
        private readonly Dictionary<string, (byte[] Bytes, XDocument? Xml, bool Opaque)> snapshots = new(StringComparer.OrdinalIgnoreCase);
        internal bool Read(string path, byte[] bytes, out XDocument xml, out bool opaque)
        {
            if (!snapshots.TryGetValue(path, out var snapshot) || !ReferenceEquals(snapshot.Bytes, bytes))
            {
                XDocument? document = null; var parsed = false;
                try { document = parse(bytes); parsed = true; } catch (XmlException) { }
                snapshot = (bytes, parsed ? document : null, !parsed && OfdXmlContentProbe.LooksLikeXml(bytes));
                snapshots[path] = snapshot;
            }
            xml = snapshot.Xml!; opaque = snapshot.Opaque; return snapshot.Xml is not null;
        }
    }
    internal static OfdPackageWriteResult Prune(
        OfdDocumentPackage package,
        IDictionary<string, byte[]> entries, CancellationToken cancellationToken = default, OfdModelMaterialization? material = null, OfdOwnedPruneDispositions? ownedDispositions = null)
    {
        if (material is null)
        {
            material = new OfdModelMaterialization(new OfdModelLimits());
            material.InspectStorage(package, cancellationToken); material.GuardPreservedXml(package, cancellationToken);
        }
        return new OfdPackagePruner(material, cancellationToken, ownedDispositions).PruneCore(package, entries, cancellationToken);
    }
    private OfdPackageWriteResult PruneCore(OfdDocumentPackage package, IDictionary<string, byte[]> entries, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new OfdPackageWriteResult();
        var scan = new XmlScanCache(Parse);
        var original = package.PreservedEntries;
        if (!original.TryGetValue("OFD.xml", out var originalRootBytes)) return result;
        var originalRoot = Parse(originalRootBytes);
        var documentPath = originalRoot.Descendants().FirstOrDefault(node => node.Name.LocalName == "DocRoot")?.Value;
        if (string.IsNullOrWhiteSpace(documentPath)) return result;
        documentPath = material.ResolvePath("OFD.xml", documentPath!, operationToken);
        if (!original.TryGetValue(documentPath, out var originalDocument)) return result;

        // Invalidated signature references must not keep deleted page payloads
        // alive during the subsequent resource reachability check.
        RemoveInvalidatedSignatures(originalRoot, original, entries, result, cancellationToken, scan);

        var retainedPaths = new HashSet<string>(package.Pages
            .Where(page => !string.IsNullOrEmpty(page.SourceEntryPath))
            .Select(page => page.SourceEntryPath!), StringComparer.OrdinalIgnoreCase);
        var retainedIdsWithoutPaths = new HashSet<string>(package.Pages
            .Where(page => string.IsNullOrEmpty(page.SourceEntryPath) && !string.IsNullOrEmpty(page.Id))
            .Select(page => page.Id!), StringComparer.Ordinal);
        var originalDocumentXml = Parse(originalDocument);
        var originalNamespace = originalDocumentXml.Root!.Name.Namespace;
        var deleted = (originalDocumentXml.Root.Element(originalNamespace + "Pages")?.Elements(originalNamespace + "Page") ?? Enumerable.Empty<XElement>())
            .Select(node => new
            {
                Id = node.Attribute("ID")?.Value ?? string.Empty,
                Path = material.ResolvePath(documentPath!, node.Attribute("BaseLoc")?.Value ?? string.Empty, operationToken)
            })
            .Where(page => !retainedPaths.Contains(page.Path) && !retainedIdsWithoutPaths.Contains(page.Id))
            .ToList();
        if (deleted.Count > 0)
        {
            var currentDocumentPath = package.DocumentEntryPath ?? $"{package.Options.DocumentId}/Document.xml";
            var currentDocumentXml = Parse(entries[currentDocumentPath]); var currentNamespace = currentDocumentXml.Root!.Name.Namespace;
            var writtenPagePaths = new HashSet<string>((currentDocumentXml.Root.Element(currentNamespace + "Pages")?.Elements(currentNamespace + "Page") ?? Enumerable.Empty<XElement>())
                .Select(node => material.ResolvePath(currentDocumentPath, node.Attribute("BaseLoc")?.Value ?? string.Empty, operationToken)),
                StringComparer.OrdinalIgnoreCase);
            var candidateIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var page in deleted)
            {
                if (original.TryGetValue(page.Path, out var bytes)) AddReferences(Parse(bytes), candidateIds);
                if (!writtenPagePaths.Contains(page.Path))
                {
                    if (ReferencesFile(entries, page.Path, scan))
                        result.Warnings.Add($"Page entry '{page.Path}' remains referenced by a template or extension and was retained as shared content.");
                    else Remove(entries, page.Path, result, OfdOwnedPruneReason.DeletedPage);
                }
            }

            RemovePageAnnotations(entries, currentDocumentPath, new HashSet<string>(deleted.Select(page => page.Id)), candidateIds, result, scan);
            var privateResourcePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            PruneUnusedTemplates(entries, currentDocumentPath, writtenPagePaths, candidateIds, privateResourcePaths, result, scan);
            foreach (var page in deleted)
            {
                var pageResources = OfdPackagePath.GetDirectory(page.Path) + "/PageRes.xml";
                if (entries.ContainsKey(page.Path) || writtenPagePaths.Any(path => OfdPackagePath.GetDirectory(path) == OfdPackagePath.GetDirectory(page.Path)) ||
                    ReferencesFile(entries, pageResources, scan)) continue;
                privateResourcePaths.Add(pageResources);
                if (entries.TryGetValue(pageResources, out var data) && TryParse(data, out var resources))
                {
                    foreach (var id in resources.Descendants().Attributes("ID")) candidateIds.Add(id.Value);
                    // Keep Res available until its exclusive font/image payloads have been pruned.
                }
            }
            PruneResources(entries, currentDocumentPath, privateResourcePaths, candidateIds, result, scan, package);
            foreach (var path in privateResourcePaths)
                if (!HasPageContentBeside(entries, path, scan) && !ReferencesFile(entries, path, scan))
                {
                    if (entries.TryGetValue(path, out var resourceBytes) && TryParse(resourceBytes, out var resourceXml) &&
                        IsTypedRoot(resourceXml, "Res", Parse(entries[currentDocumentPath]).Root!.Name.Namespace)) Remove(entries, path, result, OfdOwnedPruneReason.DeletedPageResource);
                    else if (entries.ContainsKey(path)) result.Warnings.Add($"Unmodeled page resource was retained: '{path}'.");
                }
        }

        return result;
    }

    private bool HasPageContentBeside(IDictionary<string, byte[]> entries, string resourcePath, XmlScanCache scan) =>
        entries.Where(pair => OfdPackagePath.GetDirectory(pair.Key) == OfdPackagePath.GetDirectory(resourcePath)).Any(pair =>
        {
            var parsed = scan.Read(pair.Key, pair.Value, out var xml, out var opaque);
            return opaque || parsed && xml.Root?.Name.LocalName == "Page";
        });

    private void PruneUnusedTemplates(IDictionary<string, byte[]> entries, string documentPath,
        ISet<string> pagePaths, ISet<string> candidateIds, ISet<string> privateResourcePaths, OfdPackageWriteResult result, XmlScanCache scan)
    {
        var document = Parse(entries[documentPath]);
        var templates = (document.Root!.Element(document.Root.Name.Namespace + "CommonData")?.Elements(document.Root.Name.Namespace + "TemplatePage") ?? Enumerable.Empty<XElement>()).ToList();
        var ids = new HashSet<string>(templates.Select(template => template.Attribute("ID")?.Value).OfType<string>(), StringComparer.Ordinal);
        var templatePaths = new HashSet<string>(templates.Select(template => template.Attribute("BaseLoc")?.Value)
            .Where(path => !string.IsNullOrWhiteSpace(path)).Select(path => material.ResolvePath(documentPath, path!, operationToken)), StringComparer.OrdinalIgnoreCase);
        var referencesByEntry = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var pathsByEntry = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var opaqueXml = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in entries)
        {
            // Typed XML and opaque extension XML can have any filename suffix.
            // Parse once, retaining only references relevant to these templates.
            if (scan.Read(pair.Key, pair.Value, out var xml, out var opaque))
            {
                var references = new HashSet<string>(StringComparer.Ordinal); AddReferences(xml, references);
                var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var pathValues = ReferenceValues(xml).Where(reference => reference.Attribute is not null || reference.Element.Name.LocalName is not "MaxUnitID" and not "TextCode" and not "AbbreviatedData")
                    .Where(reference => !(pair.Key.Equals(documentPath, StringComparison.OrdinalIgnoreCase) && reference.Attribute?.Name == XName.Get("BaseLoc") &&
                        reference.Element.Name == xml.Root!.Name.Namespace + "TemplatePage" && reference.Element.Parent?.Name == xml.Root.Name.Namespace + "CommonData" && reference.Element.Parent?.Parent == xml.Root));
                foreach (var reference in pathValues.Where(reference => !string.IsNullOrWhiteSpace(reference.Value)))
                {
                    try
                    {
                        foreach (var path in ResolveReferencedFiles(pair.Key, reference.Element, reference.Value)) if (templatePaths.Contains(path)) paths.Add(path);
                    }
                    catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception)) { }
                }
                pathsByEntry[pair.Key] = paths;
                references.IntersectWith(ids);
                referencesByEntry[pair.Key] = references;
            }
            else if (opaque) opaqueXml.Add(pair.Key);
        }
        foreach (var template in templates)
        {
            var id = template.Attribute("ID")?.Value;
            if (id is null) continue;
            // Scan every retained XML, excluding this declaration's own ID. Unknown references keep templates alive.
            var location = template.Attribute("BaseLoc")?.Value;
            var path = string.IsNullOrWhiteSpace(location) ? null : material.ResolvePath(documentPath, location!, operationToken);
            var live = opaqueXml.Any(entries.ContainsKey) || referencesByEntry.Any(pair => entries.ContainsKey(pair.Key) && pair.Value.Contains(id)) ||
                path is not null && pathsByEntry.Any(pair => entries.ContainsKey(pair.Key) && pair.Value.Contains(path));
            if (live) continue;
            if (path is null) continue;
            if (entries.TryGetValue(path, out var templateBytes) && (!TryParse(templateBytes, out var typedTemplate) || !IsTypedRoot(typedTemplate, "Page", document.Root.Name.Namespace)))
            {
                result.Warnings.Add($"Unmodeled template payload was retained: '{path}'.");
                continue;
            }
            template.Remove();
            entries[documentPath] = Serialize(document);
            if (pagePaths.Contains(path) || ReferencesFile(entries, path, scan)) continue;
            if (entries.TryGetValue(path, out var bytes) && TryParse(bytes, out var xmlPage)) AddReferences(xmlPage, candidateIds);
            Remove(entries, path, result, OfdOwnedPruneReason.UnusedTemplate);
            var resourcePath = OfdPackagePath.GetDirectory(path) + "/PageRes.xml";
            if (pagePaths.Any(pagePath => OfdPackagePath.GetDirectory(pagePath) == OfdPackagePath.GetDirectory(path)) ||
                ReferencesFile(entries, resourcePath, scan)) continue;
            if (entries.TryGetValue(resourcePath, out var resourceBytes) && TryParse(resourceBytes, out var resources))
                foreach (var resourceId in resources.Descendants().Attributes("ID")) candidateIds.Add(resourceId.Value);
            privateResourcePaths.Add(resourcePath);
        }
    }

    private void RemovePageAnnotations(IDictionary<string, byte[]> entries, string documentPath,
        ISet<string> deletedIds, ISet<string> candidateIds, OfdPackageWriteResult result, XmlScanCache scan)
    {
        var document = Parse(entries[documentPath]); var ns = document.Root!.Name.Namespace;
        foreach (var declaration in document.Root.Elements(ns + "Annotations"))
        {
            var path = material.ResolvePath(documentPath, declaration.Value, operationToken);
            if (!entries.TryGetValue(path, out var bytes) || !TryParse(bytes, out var xml) || !IsTypedRoot(xml, "Annotations", ns)) continue;
            var removedFiles = new List<string>();
            foreach (var page in xml.Root!.Elements(xml.Root.Name.Namespace + "Page").Where(node => deletedIds.Contains(node.Attribute("PageID")?.Value ?? string.Empty)).ToList())
            {
                var file = page.Element(xml.Root.Name.Namespace + "FileLoc")?.Value;
                if (!string.IsNullOrWhiteSpace(file)) removedFiles.Add(material.ResolvePath(path, file!, operationToken));
                page.Remove();
            }
            if (removedFiles.Count == 0) continue;
            entries[path] = Serialize(xml);
            foreach (var file in removedFiles)
            {
                if (ReferencesFile(entries, file, scan)) continue;
                if (!entries.TryGetValue(file, out var annotationBytes) || !TryParse(annotationBytes, out var annotation) || !IsTypedRoot(annotation, "PageAnnot", ns))
                {
                    result.Warnings.Add($"Unmodeled annotation payload was retained: '{file}'.");
                    continue;
                }
                AddReferences(annotation, candidateIds); Remove(entries, file, result, OfdOwnedPruneReason.DeletedAnnotation);
            }
        }
    }

    private bool IsTypedRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private void PruneResources(
        IDictionary<string, byte[]> entries, string documentPath, ISet<string> privateResourcePaths,
        ISet<string> candidateIds, OfdPackageWriteResult result, XmlScanCache scan, OfdDocumentPackage package)
    {
        var document = Parse(entries[documentPath]); var ns = document.Root!.Name.Namespace;
        var common = document.Root.Element(ns + "CommonData");
        var owned = new HashSet<string>(privateResourcePaths, StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in common?.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
            owned.Add(material.ResolvePath(documentPath, declaration.Value, operationToken));
        var pageLocations = (document.Root.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
            .Concat(common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>());
        foreach (var page in pageLocations)
        {
            var location = page.Attribute("BaseLoc")?.Value;
            if (!string.IsNullOrWhiteSpace(location)) owned.Add(OfdPackagePath.GetDirectory(material.ResolvePath(documentPath, location!, operationToken)) + "/PageRes.xml");
        }
        var documents = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        var references = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in entries)
        {
            if (!scan.Read(pair.Key, pair.Value, out var xml, out var opaque))
            {
                if (!opaque) continue;
                result.Warnings.Add($"Unused resources retained because extension XML '{pair.Key}' cannot be inspected safely.");
                return;
            }
            documents[pair.Key] = xml;
            AddReferences(xml, references);
        }

        // A deleted placement may release a whole DAG. Only reader-owned,
        // fully modeled definitions participate; ambiguous extension closure stays.
        var modeled = new HashSet<string>(StringComparer.Ordinal);
        foreach (var resource in package.CompositeResources?.Items ?? Enumerable.Empty<OfdCompositeResource>())
        {
            operationToken.ThrowIfCancellationRequested(); material.ChargeCopy(128, operationToken);
            if (resource.Id is null || resource.Elements.Any(e => e is OfdRawElement)) continue;
            var source = resource.SourceXml is null ? null : OfdXmlEnvelope.Parse(resource.SourceXml);
            if (source is not null && source.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Width" or "Height")))) continue;
            modeled.Add(resource.Id);
        }
        // Unknown reference syntax can encode an ID list or external dependency.
        // Do not infer absence of incoming edges from a scalar lookup in that case.
        if (documents.Values.Any(xml => !CanInspectCompositeClosure(xml)))
        {
            if (modeled.Count > 0) result.Warnings.Add("Composite definitions retained because opaque reference syntax prevents proving closure.");
            modeled.Clear();
        }
        var changedComposites = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        bool removed;
        do
        {
            removed = false; references.Clear();
            foreach (var documentXml in documents.Values) { operationToken.ThrowIfCancellationRequested(); AddReferences(documentXml, references); }
            foreach (var pair in documents)
            {
                if (!owned.Contains(pair.Key) || !IsTypedRoot(pair.Value, "Res", ns)) continue;
                material.ChargeCopy(128L + pair.Value.Root!.Descendants().Count() * 16L, operationToken);
                foreach (var resource in pair.Value.Root.Elements(pair.Value.Root.Name.Namespace + "CompositeGraphicUnits").Elements(pair.Value.Root.Name.Namespace + "CompositeGraphicUnit").ToArray())
                {
                    operationToken.ThrowIfCancellationRequested(); material.ChargeWork(128, operationToken); material.Structure.Reference();
                    var id = resource.Attribute("ID")?.Value;
                    if (id is null || !modeled.Contains(id) || !candidateIds.Contains(id) || references.Contains(id)) continue;
                    material.ChargeCopy(256, operationToken);
                    foreach (var reference in resource.Descendants().Attributes().Where(a => a.Name.LocalName is "Font" or "ResourceID")) candidateIds.Add(reference.Value);
                    resource.Remove(); changedComposites.Add(pair.Key); removed = true;
                }
            }
        } while (removed);
        foreach (var path in changedComposites) entries[path] = Serialize(documents[path]);
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in documents)
        {
            if (!owned.Contains(pair.Key) || !IsTypedRoot(pair.Value, "Res", ns)) continue;
            var changed = false;
            foreach (var resource in pair.Value.Root!.Elements(pair.Value.Root.Name.Namespace + "Fonts").Elements(pair.Value.Root.Name.Namespace + "Font")
                .Concat(pair.Value.Root.Elements(pair.Value.Root.Name.Namespace + "MultiMedias").Elements(pair.Value.Root.Name.Namespace + "MultiMedia")).ToList())
            {
                var id = resource.Attribute("ID")?.Value;
                if (id is null || !candidateIds.Contains(id) || references.Contains(id)) continue;
                var file = resource.Element(pair.Value.Root.Name.Namespace + "FontFile")?.Value
                    ?? resource.Element(pair.Value.Root.Name.Namespace + "MediaFile")?.Value ?? resource.Attribute("MediaFile")?.Value;
                if (!string.IsNullOrWhiteSpace(file))
                {
                    files.Add(ResolveResourceFile(pair.Key, pair.Value.Root!, file!));
                }
                resource.Remove();
                changed = true;
            }
            if (changed) entries[pair.Key] = Serialize(pair.Value);
        }

        // A second resource, attachment, template or opaque XML extension can
        // still reference the same payload. Preserve it whenever in doubt.
        foreach (var file in files)
        {
            if (!ReferencesFile(entries, file, scan)) Remove(entries, file, result, OfdOwnedPruneReason.UnreachableResourcePayload);
        }
    }

    // Deliberately partial destructive-admission grammar for C pruning.
    // An unrecognized control table stays protected; no vendor ID-list parser
    // or name-only recursive allowlist grants deletion permission.
    private bool CanInspectCompositeClosure(XDocument xml)
    {
        var root = xml.Root; if (root is null) return false;
        var ns = root.Name.Namespace;
        if (ns.NamespaceName != OfdConstants.Namespace && ns.NamespaceName != OfdConstants.StandardNamespace) return false;
        foreach (var node in xml.DescendantNodes())
        {
            operationToken.ThrowIfCancellationRequested(); material.ChargeWork(64, operationToken);
            if (node is XProcessingInstruction or XComment || node is XElement element && element.Name.Namespace != ns) return false;
        }
        bool Plain(XElement node, params string[] names) => node.Attributes().All(a => a.IsNamespaceDeclaration || a.Name.Namespace == XNamespace.None && names.Contains(a.Name.LocalName));
        bool Children(XElement node, params string[] names) => node.Elements().All(n => n.Name.Namespace == ns && names.Contains(n.Name.LocalName)) &&
            node.Nodes().All(n => n is XElement || n is XText t && string.IsNullOrWhiteSpace(t.Value));
        bool Leaf(XElement node) => !node.HasElements && Plain(node) && node.Nodes().All(n => n is XText);
        bool Graphic(XElement node)
        {
            if (node.Name == ns + "CompositeObject") return Plain(node, "ID", "Boundary", "CTM", "ResourceID", "Visible") && Children(node, "Clips") &&
                node.Elements().All(n => OfdGraphicXmlContract.HasKnownChildren(n) && n.DescendantsAndSelf().Attributes().All(OfdGraphicXmlContract.IsKnownAttribute));
            return node.Name.LocalName is "TextObject" or "ImageObject" or "PathObject" && OfdGraphicXmlContract.HasKnownChildren(node) &&
                node.DescendantsAndSelf().Attributes().All(OfdGraphicXmlContract.IsKnownAttribute) && !OfdGraphicXmlContract.HasUnsupportedReferences(node);
        }
        if (root.Name == ns + "Res")
        {
            if (!Plain(root, "BaseLoc") || !Children(root, "Fonts", "MultiMedias", "CompositeGraphicUnits")) return false;
            foreach (var container in root.Elements())
            {
                if (!Plain(container)) return false;
                var kind = container.Name.LocalName == "Fonts" ? "Font" : container.Name.LocalName == "MultiMedias" ? "MultiMedia" : "CompositeGraphicUnit";
                if (!Children(container, kind)) return false;
                foreach (var resource in container.Elements())
                {
                    operationToken.ThrowIfCancellationRequested(); material.ChargeWork(128, operationToken);
                    if (kind == "CompositeGraphicUnit")
                    { if (!Plain(resource, "ID", "Width", "Height") || resource.Elements().Any(n => !Graphic(n)) || resource.Nodes().Any(n => n is XText t && !string.IsNullOrWhiteSpace(t.Value))) return false; }
                    else if (kind == "Font")
                    { if (!Plain(resource, "ID", "FontName", "FamilyName", "Charset", "Bold", "Italic") || !Children(resource, "FontFile") || resource.Elements().Any(n => !Leaf(n))) return false; }
                    else if (!Plain(resource, "ID", "Type", "Format") || !Children(resource, "MediaFile") || resource.Elements().Any(n => !Leaf(n))) return false;
                }
            }
            return true;
        }
        if (root.Name == ns + "OFD")
        {
            if (!Plain(root, "Version", "DocType") || !Children(root, "DocBody")) return false;
            foreach (var body in root.Elements())
            {
                if (!Plain(body) || !Children(body, "DocInfo", "DocRoot")) return false;
                foreach (var child in body.Elements())
                    if (child.Name.LocalName == "DocRoot" ? !Leaf(child) : !Plain(child) || !Children(child, "DocID", "Title", "Author", "Subject", "Keywords", "Creator", "CreationDate", "ModDate") || child.Elements().Any(n => !Leaf(n))) return false;
            }
            return true;
        }
        if (root.Name == ns + "Document")
        {
            if (!Plain(root) || !Children(root, "CommonData", "Pages")) return false;
            var common = root.Element(ns + "CommonData"); var pages = root.Element(ns + "Pages");
            if (common is null || pages is null || !Plain(common) || !Children(common, "MaxUnitID", "PageArea", "PublicRes", "DocumentRes") || !Plain(pages) || !Children(pages, "Page")) return false;
            foreach (var child in common.Elements())
                if (child.Name.LocalName == "PageArea" ? !Plain(child) || !Children(child, "PhysicalBox") || child.Elements().Any(n => !Leaf(n)) : !Leaf(child)) return false;
            return pages.Elements().All(n => Plain(n, "ID", "BaseLoc") && !n.Nodes().Any());
        }
        if (root.Name == ns + "Page")
        {
            if (!Plain(root) || !Children(root, "Area", "Content")) return false;
            foreach (var child in root.Elements())
            {
                if (child.Name.LocalName == "Area") { if (!Plain(child) || !Children(child, "PhysicalBox") || child.Elements().Any(n => !Leaf(n))) return false; }
                else if (!Plain(child) || !Children(child, "Layer") || child.Elements().Any(layer => !Plain(layer, "ID", "Type") || layer.Elements().Any(n => !Graphic(n)))) return false;
            }
            return true;
        }
        return false;
    }

    private void RemoveInvalidatedSignatures(
        XDocument originalRoot,
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken, XmlScanCache scan)
    {
        var declarations = SignatureDeclarations(originalRoot).ToList();
        if (declarations.Count == 0 ||
            (original.Count == entries.Count && original.All(pair => entries.TryGetValue(pair.Key, out var bytes) && bytes.SequenceEqual(pair.Value)))) return;

        CleanSignaturesCore(original, entries, result, cancellationToken, scan);
        result.SignaturesInvalidated = true;
        result.Warnings.Add("Signature declarations were removed because the package was rewritten; sign the completed output again if required.");
    }

    internal static void CleanSignatures(
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken = default, XmlScanCache? scan = null, OfdModelMaterialization? material = null)
    {
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        var package = new OfdDocumentPackage(); foreach (var pair in original) package.PreservedEntries.Add(pair.Key, pair.Value);
        material.InspectStorage(package, cancellationToken); material.GuardPreservedXml(package, cancellationToken);
        new OfdPackagePruner(material, cancellationToken).CleanSignaturesCore(original, entries, result, cancellationToken, scan);
    }
    private void CleanSignaturesCore(
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken = default, XmlScanCache? scan = null)
    {
        scan ??= new XmlScanCache(Parse);
        cancellationToken.ThrowIfCancellationRequested();
        if (!original.TryGetValue("OFD.xml", out var bytes)) return;
        var originalRoot = Parse(bytes);
        var root = Parse(entries["OFD.xml"]);
        var removedDeclarations = SignatureDeclarations(root).ToList();
        if (removedDeclarations.Count > 0) result.SignaturesInvalidated = true;
        ownedDispositions?.RecordActualSignatureDeclarations(removedDeclarations, cancellationToken);
        removedDeclarations.Remove();
        entries["OFD.xml"] = Serialize(root);
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var knownXml = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        var inspectedLists = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var inspectedDescriptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in SignatureDeclarations(originalRoot))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var listPath = material.ResolvePath("OFD.xml", declaration.Value, operationToken);
            if (!original.TryGetValue(listPath, out var listBytes) || !inspectedLists.Add(listPath)) continue;
            if (!TryParse(listBytes, out var list) || !IsSignatureRoot(list, "Signatures", originalRoot.Root!.Name.Namespace))
            {
                result.Warnings.Add($"Unmodeled signature list was retained: '{listPath}'.");
                continue;
            }
            knownXml[listPath] = list;
            if (!IsOwnedSignaturePath(listPath))
            {
                result.Warnings.Add($"Signature payloads outside a Signs directory were retained: '{listPath}'.");
                continue;
            }
            candidates.Add(listPath);
            foreach (var record in list.Root.Elements(list.Root.Name.Namespace + "Signature"))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var location = record.Attribute("BaseLoc")?.Value;
                if (string.IsNullOrWhiteSpace(location)) continue;
                var signaturePath = material.ResolvePath(listPath, location!, operationToken);
                if (!IsOwnedSignaturePath(signaturePath) || !original.TryGetValue(signaturePath, out var signatureBytes) ||
                    !inspectedDescriptions.Add(signaturePath)) continue;
                if (!TryParse(signatureBytes, out var signature) || !IsSignatureRoot(signature, "Signature", originalRoot.Root!.Name.Namespace))
                {
                    result.Warnings.Add($"Unmodeled signature description was retained: '{signaturePath}'.");
                    continue;
                }
                knownXml[signaturePath] = signature;
                candidates.Add(signaturePath);
                var signatureNamespace = signature.Root.Name.Namespace;
                var seals = signature.Root.Elements(signatureNamespace + "SignedInfo").Elements(signatureNamespace + "Seal");
                var values = signature.Root.Elements(signatureNamespace + "SignedValue").Select(node => node.Value)
                    .Concat(seals.Elements(signatureNamespace + "BaseLoc").Select(node => node.Value))
                    .Concat(seals.Attributes("BaseLoc").Select(attribute => attribute.Value));
                foreach (var value in values)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var payload = material.ResolvePath(signaturePath, value, operationToken);
                    // A declaration is never authority to delete arbitrary document content.
                    // Values/appearances must be inside this signature's own directory.
                    var directory = OfdPackagePath.GetDirectory(signaturePath) + "/";
                    if (IsOwnedSignaturePath(payload) && payload.StartsWith(directory, StringComparison.OrdinalIgnoreCase))
                    {
                        if (entries.ContainsKey(payload)) candidates.Add(payload);
                    }
                    else result.Warnings.Add($"Unowned signature payload was retained: '{payload}'.");
                }
            }
        }
        candidates.IntersectWith(entries.Keys);
        // Scan each retained/restored XML once. Each discovered candidate becomes
        // live and joins the queue, preserving its entire transitive path closure.
        var pending = new Queue<string>(entries.Keys.Where(path => !candidates.Contains(path)));
        var scanned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = pending.Dequeue();
            if (!scanned.Add(path)) continue;
            if (!knownXml.TryGetValue(path, out var xml) && !scan.Read(path, entries[path], out xml, out var opaque))
            {
                if (!opaque) continue;
                result.Warnings.Add($"Signature payloads retained because extension XML '{path}' cannot be inspected safely.");
                return;
            }
            foreach (var reference in ReferenceValues(xml).Where(reference => !string.IsNullOrWhiteSpace(reference.Value)))
            {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var target in ResolveReferencedFiles(path, reference.Element, reference.Value))
                    if (candidates.Remove(target)) pending.Enqueue(target);
            }
        }
        foreach (var path in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Remove(entries, path, result, OfdOwnedPruneReason.InvalidatedSignature);
        }
    }

    private bool IsSignatureRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private IEnumerable<XElement> SignatureDeclarations(XDocument document) =>
        document.Root?.Elements().Where(node => node.Name == document.Root.Name.Namespace + "DocBody")
            .Elements(document.Root.Name.Namespace + "Signatures") ?? Enumerable.Empty<XElement>();

    private bool IsOwnedSignaturePath(string path) =>
        path.Split('/').Any(segment => string.Equals(segment, "Signs", StringComparison.OrdinalIgnoreCase));

    private bool ReferencesFile(IDictionary<string, byte[]> entries, string file, XmlScanCache scan)
    {
        foreach (var pair in entries)
        {
            if (!scan.Read(pair.Key, pair.Value, out var xml, out var opaque))
            {
                if (opaque) return true;
                continue;
            }
            foreach (var reference in ReferenceValues(xml).Where(reference => !string.IsNullOrWhiteSpace(reference.Value)))
                if (ResolveReferencedFiles(pair.Key, reference.Element, reference.Value).Contains(file, StringComparer.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private IEnumerable<(XElement Element, string Value, XAttribute? Attribute)> ReferenceValues(XDocument xml)
    {
        foreach (var element in xml.Descendants().Where(element => !element.HasElements)) yield return (element, element.Value, null);
        foreach (var attribute in xml.Descendants().Attributes().Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName != "ID"))
            yield return (attribute.Parent!, attribute.Value, attribute);
    }

    private IEnumerable<string> ResolveReferencedFiles(string path, XElement element, string value)
    {
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try { files.Add(material.ResolvePath(path, value, operationToken)); } catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception)) { }
        var basis = path;
        foreach (var owner in element.AncestorsAndSelf().Reverse())
        {
            var location = owner.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(location)) continue;
            try
            {
                basis = material.ResolvePath(basis, location!.TrimEnd('/') + "/_reference", operationToken);
                files.Add(material.ResolvePath(basis, value, operationToken));
            }
            catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception)) { }
        }
        return files;
    }

    private string ResolveResourceFile(string path, XElement root, string file)
    {
        var baseLocation = root.Name.LocalName == "Res" ? root.Attribute("BaseLoc")?.Value : null;
        return material.ResolvePath(path, string.IsNullOrEmpty(baseLocation) || file.StartsWith("/", StringComparison.Ordinal)
            ? file : baseLocation!.TrimEnd('/') + "/" + file, operationToken);
    }

    private void AddReferences(XDocument xml, ISet<string> references)
    {
        foreach (var attribute in xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName != "ID"))
            references.Add(attribute.Value);
        foreach (var node in xml.Descendants().Where(node => !node.HasElements &&
            node.Name.LocalName is not "MaxUnitID" and not "TextCode" and not "AbbreviatedData")) references.Add(node.Value);
    }

    private void Remove(IDictionary<string, byte[]> entries, string path, OfdPackageWriteResult result, OfdOwnedPruneReason reason)
    {
        if (!entries.TryGetValue(path, out var actual)) return;
        ownedDispositions?.RecordRemoval(path, actual, reason, operationToken);
        if (entries.Remove(path)) result.Removed.Add(path);
    }

    private XDocument Parse(byte[] bytes)
    {
        return material.ReadXml("<prune-entry>", bytes, operationToken);
    }
    private bool TryParse(byte[] bytes, out XDocument document)
    {
        try { document = Parse(bytes); return true; }
        catch (XmlException) { document = null!; return false; }
    }
    private byte[] Serialize(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
}
