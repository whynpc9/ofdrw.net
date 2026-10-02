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

internal static class OfdPackagePruner
{
    internal sealed class XmlScanCache
    {
        private readonly Dictionary<string, (byte[] Bytes, XDocument? Xml, bool Opaque)> snapshots = new(StringComparer.OrdinalIgnoreCase);
        internal bool Read(string path, byte[] bytes, out XDocument xml, out bool opaque)
        {
            if (!snapshots.TryGetValue(path, out var snapshot) || !ReferenceEquals(snapshot.Bytes, bytes))
            {
                var parsed = TryParse(bytes, out var document);
                snapshot = (bytes, parsed ? document : null, !parsed && OfdXmlContentProbe.LooksLikeXml(bytes));
                snapshots[path] = snapshot;
            }
            xml = snapshot.Xml!; opaque = snapshot.Opaque; return snapshot.Xml is not null;
        }
    }
    internal static OfdPackageWriteResult Prune(
        OfdDocumentPackage package,
        IDictionary<string, byte[]> entries, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new OfdPackageWriteResult();
        var scan = new XmlScanCache();
        var original = package.PreservedEntries;
        if (!original.TryGetValue("OFD.xml", out var originalRootBytes)) return result;
        var originalRoot = Parse(originalRootBytes);
        var documentPath = originalRoot.Descendants().FirstOrDefault(node => node.Name.LocalName == "DocRoot")?.Value;
        if (string.IsNullOrWhiteSpace(documentPath)) return result;
        documentPath = OfdPackagePath.Resolve("OFD.xml", documentPath!);
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
                Path = OfdPackagePath.Resolve(documentPath!, node.Attribute("BaseLoc")?.Value ?? string.Empty)
            })
            .Where(page => !retainedPaths.Contains(page.Path) && !retainedIdsWithoutPaths.Contains(page.Id))
            .ToList();
        if (deleted.Count > 0)
        {
            var currentDocumentPath = package.DocumentEntryPath ?? $"{package.Options.DocumentId}/Document.xml";
            var currentDocumentXml = Parse(entries[currentDocumentPath]); var currentNamespace = currentDocumentXml.Root!.Name.Namespace;
            var writtenPagePaths = new HashSet<string>((currentDocumentXml.Root.Element(currentNamespace + "Pages")?.Elements(currentNamespace + "Page") ?? Enumerable.Empty<XElement>())
                .Select(node => OfdPackagePath.Resolve(currentDocumentPath, node.Attribute("BaseLoc")?.Value ?? string.Empty)),
                StringComparer.OrdinalIgnoreCase);
            var candidateIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var page in deleted)
            {
                if (original.TryGetValue(page.Path, out var bytes)) AddReferences(Parse(bytes), candidateIds);
                if (!writtenPagePaths.Contains(page.Path))
                {
                    if (ReferencesFile(entries, page.Path, scan))
                        result.Warnings.Add($"Page entry '{page.Path}' remains referenced by a template or extension and was retained as shared content.");
                    else Remove(entries, page.Path, result);
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
            PruneResources(entries, currentDocumentPath, privateResourcePaths, candidateIds, result, scan);
            foreach (var path in privateResourcePaths)
                if (!HasPageContentBeside(entries, path, scan) && !ReferencesFile(entries, path, scan))
                {
                    if (entries.TryGetValue(path, out var resourceBytes) && TryParse(resourceBytes, out var resourceXml) &&
                        IsTypedRoot(resourceXml, "Res", Parse(entries[currentDocumentPath]).Root!.Name.Namespace)) Remove(entries, path, result);
                    else if (entries.ContainsKey(path)) result.Warnings.Add($"Unmodeled page resource was retained: '{path}'.");
                }
        }

        return result;
    }

    private static bool HasPageContentBeside(IDictionary<string, byte[]> entries, string resourcePath, XmlScanCache scan) =>
        entries.Where(pair => OfdPackagePath.GetDirectory(pair.Key) == OfdPackagePath.GetDirectory(resourcePath)).Any(pair =>
        {
            var parsed = scan.Read(pair.Key, pair.Value, out var xml, out var opaque);
            return opaque || parsed && xml.Root?.Name.LocalName == "Page";
        });

    private static void PruneUnusedTemplates(IDictionary<string, byte[]> entries, string documentPath,
        ISet<string> pagePaths, ISet<string> candidateIds, ISet<string> privateResourcePaths, OfdPackageWriteResult result, XmlScanCache scan)
    {
        var document = Parse(entries[documentPath]);
        var templates = (document.Root!.Element(document.Root.Name.Namespace + "CommonData")?.Elements(document.Root.Name.Namespace + "TemplatePage") ?? Enumerable.Empty<XElement>()).ToList();
        var ids = new HashSet<string>(templates.Select(template => template.Attribute("ID")?.Value).OfType<string>(), StringComparer.Ordinal);
        var templatePaths = new HashSet<string>(templates.Select(template => template.Attribute("BaseLoc")?.Value)
            .Where(path => !string.IsNullOrWhiteSpace(path)).Select(path => OfdPackagePath.Resolve(documentPath, path!)), StringComparer.OrdinalIgnoreCase);
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
                    catch (InvalidDataException) { }
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
            var path = string.IsNullOrWhiteSpace(location) ? null : OfdPackagePath.Resolve(documentPath, location!);
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
            Remove(entries, path, result);
            var resourcePath = OfdPackagePath.GetDirectory(path) + "/PageRes.xml";
            if (pagePaths.Any(pagePath => OfdPackagePath.GetDirectory(pagePath) == OfdPackagePath.GetDirectory(path)) ||
                ReferencesFile(entries, resourcePath, scan)) continue;
            if (entries.TryGetValue(resourcePath, out var resourceBytes) && TryParse(resourceBytes, out var resources))
                foreach (var resourceId in resources.Descendants().Attributes("ID")) candidateIds.Add(resourceId.Value);
            privateResourcePaths.Add(resourcePath);
        }
    }

    private static void RemovePageAnnotations(IDictionary<string, byte[]> entries, string documentPath,
        ISet<string> deletedIds, ISet<string> candidateIds, OfdPackageWriteResult result, XmlScanCache scan)
    {
        var document = Parse(entries[documentPath]); var ns = document.Root!.Name.Namespace;
        foreach (var declaration in document.Root.Elements(ns + "Annotations"))
        {
            var path = OfdPackagePath.Resolve(documentPath, declaration.Value);
            if (!entries.TryGetValue(path, out var bytes) || !TryParse(bytes, out var xml) || !IsTypedRoot(xml, "Annotations", ns)) continue;
            var removedFiles = new List<string>();
            foreach (var page in xml.Root!.Elements(xml.Root.Name.Namespace + "Page").Where(node => deletedIds.Contains(node.Attribute("PageID")?.Value ?? string.Empty)).ToList())
            {
                var file = page.Element(xml.Root.Name.Namespace + "FileLoc")?.Value;
                if (!string.IsNullOrWhiteSpace(file)) removedFiles.Add(OfdPackagePath.Resolve(path, file!));
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
                AddReferences(annotation, candidateIds); Remove(entries, file, result);
            }
        }
    }

    private static bool IsTypedRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private static void PruneResources(
        IDictionary<string, byte[]> entries, string documentPath, ISet<string> privateResourcePaths,
        ISet<string> candidateIds, OfdPackageWriteResult result, XmlScanCache scan)
    {
        var document = Parse(entries[documentPath]); var ns = document.Root!.Name.Namespace;
        var common = document.Root.Element(ns + "CommonData");
        var owned = new HashSet<string>(privateResourcePaths, StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in common?.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
            owned.Add(OfdPackagePath.Resolve(documentPath, declaration.Value));
        var pageLocations = (document.Root.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
            .Concat(common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>());
        foreach (var page in pageLocations)
        {
            var location = page.Attribute("BaseLoc")?.Value;
            if (!string.IsNullOrWhiteSpace(location)) owned.Add(OfdPackagePath.GetDirectory(OfdPackagePath.Resolve(documentPath, location!)) + "/PageRes.xml");
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
            if (!ReferencesFile(entries, file, scan)) Remove(entries, file, result);
        }
    }

    private static void RemoveInvalidatedSignatures(
        XDocument originalRoot,
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken, XmlScanCache scan)
    {
        var declarations = SignatureDeclarations(originalRoot).ToList();
        if (declarations.Count == 0 ||
            (original.Count == entries.Count && original.All(pair => entries.TryGetValue(pair.Key, out var bytes) && bytes.SequenceEqual(pair.Value)))) return;

        CleanSignatures(original, entries, result, cancellationToken, scan);
        result.SignaturesInvalidated = true;
        result.Warnings.Add("Signature declarations were removed because the package was rewritten; sign the completed output again if required.");
    }

    internal static void CleanSignatures(
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken = default, XmlScanCache? scan = null)
    {
        scan ??= new XmlScanCache();
        cancellationToken.ThrowIfCancellationRequested();
        if (!original.TryGetValue("OFD.xml", out var bytes)) return;
        var originalRoot = Parse(bytes);
        var root = Parse(entries["OFD.xml"]);
        SignatureDeclarations(root).Remove();
        entries["OFD.xml"] = Serialize(root);
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var knownXml = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        var inspectedLists = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var inspectedDescriptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in SignatureDeclarations(originalRoot))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var listPath = OfdPackagePath.Resolve("OFD.xml", declaration.Value);
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
                var signaturePath = OfdPackagePath.Resolve(listPath, location!);
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
                    var payload = OfdPackagePath.Resolve(signaturePath, value);
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
            Remove(entries, path, result);
        }
    }

    private static bool IsSignatureRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private static IEnumerable<XElement> SignatureDeclarations(XDocument document) =>
        document.Root?.Elements().Where(node => node.Name == document.Root.Name.Namespace + "DocBody")
            .Elements(document.Root.Name.Namespace + "Signatures") ?? Enumerable.Empty<XElement>();

    private static bool IsOwnedSignaturePath(string path) =>
        path.Split('/').Any(segment => string.Equals(segment, "Signs", StringComparison.OrdinalIgnoreCase));

    private static bool ReferencesFile(IDictionary<string, byte[]> entries, string file, XmlScanCache scan)
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

    private static IEnumerable<(XElement Element, string Value, XAttribute? Attribute)> ReferenceValues(XDocument xml)
    {
        foreach (var element in xml.Descendants().Where(element => !element.HasElements)) yield return (element, element.Value, null);
        foreach (var attribute in xml.Descendants().Attributes().Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName != "ID"))
            yield return (attribute.Parent!, attribute.Value, attribute);
    }

    private static IEnumerable<string> ResolveReferencedFiles(string path, XElement element, string value)
    {
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try { files.Add(OfdPackagePath.Resolve(path, value)); } catch (InvalidDataException) { }
        var basis = path;
        foreach (var owner in element.AncestorsAndSelf().Reverse())
        {
            var location = owner.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(location)) continue;
            try
            {
                basis = OfdPackagePath.Resolve(basis, location!.TrimEnd('/') + "/_reference");
                files.Add(OfdPackagePath.Resolve(basis, value));
            }
            catch (InvalidDataException) { }
        }
        return files;
    }

    private static string ResolveResourceFile(string path, XElement root, string file)
    {
        var baseLocation = root.Name.LocalName == "Res" ? root.Attribute("BaseLoc")?.Value : null;
        return OfdPackagePath.Resolve(path, string.IsNullOrEmpty(baseLocation) || file.StartsWith("/", StringComparison.Ordinal)
            ? file : baseLocation!.TrimEnd('/') + "/" + file);
    }

    private static void AddReferences(XDocument xml, ISet<string> references)
    {
        foreach (var attribute in xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName != "ID"))
            references.Add(attribute.Value);
        foreach (var node in xml.Descendants().Where(node => !node.HasElements &&
            node.Name.LocalName is not "MaxUnitID" and not "TextCode" and not "AbbreviatedData")) references.Add(node.Value);
    }

    private static void Remove(IDictionary<string, byte[]> entries, string path, OfdPackageWriteResult result)
    {
        if (entries.Remove(path)) result.Removed.Add(path);
    }

    private static XDocument Parse(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }
    private static bool TryParse(byte[] bytes, out XDocument document)
    {
        try { document = Parse(bytes); return true; }
        catch (XmlException) { document = null!; return false; }
    }
    private static byte[] Serialize(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
}
