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
    internal static OfdPackageWriteResult Prune(
        OfdDocumentPackage package,
        IDictionary<string, byte[]> entries, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new OfdPackageWriteResult();
        var original = package.PreservedEntries;
        if (!original.TryGetValue("OFD.xml", out var originalRootBytes)) return result;
        var originalRoot = Parse(originalRootBytes);
        var documentPath = originalRoot.Descendants().FirstOrDefault(node => node.Name.LocalName == "DocRoot")?.Value;
        if (string.IsNullOrWhiteSpace(documentPath)) return result;
        documentPath = OfdPackagePath.Resolve("OFD.xml", documentPath!);
        if (!original.TryGetValue(documentPath, out var originalDocument)) return result;

        // Invalidated signature references must not keep deleted page payloads
        // alive during the subsequent resource reachability check.
        RemoveInvalidatedSignatures(originalRoot, original, entries, result, cancellationToken);

        var retainedPaths = new HashSet<string>(package.Pages
            .Where(page => !string.IsNullOrEmpty(page.SourceEntryPath))
            .Select(page => page.SourceEntryPath!), StringComparer.OrdinalIgnoreCase);
        var retainedIdsWithoutPaths = new HashSet<string>(package.Pages
            .Where(page => string.IsNullOrEmpty(page.SourceEntryPath) && !string.IsNullOrEmpty(page.Id))
            .Select(page => page.Id!), StringComparer.Ordinal);
        var deleted = Parse(originalDocument).Descendants()
            .Where(node => node.Name.LocalName == "Page" && node.Parent?.Name.LocalName == "Pages")
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
            var writtenPagePaths = new HashSet<string>(Parse(entries[currentDocumentPath]).Descendants()
                .Where(node => node.Name.LocalName == "Page" && node.Parent?.Name.LocalName == "Pages")
                .Select(node => OfdPackagePath.Resolve(currentDocumentPath, node.Attribute("BaseLoc")?.Value ?? string.Empty)),
                StringComparer.OrdinalIgnoreCase);
            var candidateIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var page in deleted)
            {
                if (original.TryGetValue(page.Path, out var bytes)) AddReferences(Parse(bytes), candidateIds);
                if (!writtenPagePaths.Contains(page.Path))
                {
                    if (ReferencesFile(entries, page.Path, preserveOnUnknownXml: false))
                        result.Warnings.Add($"Page entry '{page.Path}' remains referenced by a template or extension and was retained as shared content.");
                    else Remove(entries, page.Path, result);
                }
            }

            RemovePageAnnotations(entries, new HashSet<string>(deleted.Select(page => page.Id)), candidateIds, result);
            var privateResourcePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            PruneUnusedTemplates(entries, currentDocumentPath, writtenPagePaths, candidateIds, privateResourcePaths, result);
            foreach (var page in deleted)
            {
                var pageResources = OfdPackagePath.GetDirectory(page.Path) + "/PageRes.xml";
                if (entries.ContainsKey(page.Path) || writtenPagePaths.Any(path => OfdPackagePath.GetDirectory(path) == OfdPackagePath.GetDirectory(page.Path)) ||
                    ReferencesFile(entries, pageResources)) continue;
                privateResourcePaths.Add(pageResources);
                if (entries.TryGetValue(pageResources, out var data) && TryParse(data, out var resources))
                {
                    foreach (var id in resources.Descendants().Attributes("ID")) candidateIds.Add(id.Value);
                    // Keep Res available until its exclusive font/image payloads have been pruned.
                }
            }
            PruneResources(entries, candidateIds, result);
            foreach (var path in privateResourcePaths)
                if (!HasPageContentBeside(entries, path) && !ReferencesFile(entries, path)) Remove(entries, path, result);
        }

        return result;
    }

    private static bool HasPageContentBeside(IDictionary<string, byte[]> entries, string resourcePath) =>
        entries.Any(pair => IsXml(pair.Key) && OfdPackagePath.GetDirectory(pair.Key) == OfdPackagePath.GetDirectory(resourcePath) &&
            TryParse(pair.Value, out var xml) && xml.Root?.Name.LocalName == "Page");

    private static void PruneUnusedTemplates(IDictionary<string, byte[]> entries, string documentPath,
        ISet<string> pagePaths, ISet<string> candidateIds, ISet<string> privateResourcePaths, OfdPackageWriteResult result)
    {
        var document = Parse(entries[documentPath]);
        foreach (var template in document.Descendants().Where(node => node.Name.LocalName == "TemplatePage").ToList())
        {
            var id = template.Attribute("ID")?.Value;
            if (id is null) continue;
            // Scan every retained XML, excluding this declaration's own ID. Unknown references keep templates alive.
            var live = entries.Where(pair => IsXml(pair.Key)).Any(pair =>
            {
                if (!TryParse(pair.Value, out var xml)) return true;
                var references = new HashSet<string>(StringComparer.Ordinal);
                AddReferences(xml, references);
                return references.Contains(id);
            });
            if (live) continue;
            var location = template.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(location)) continue;
            var path = OfdPackagePath.Resolve(documentPath, location!);
            template.Remove();
            entries[documentPath] = Serialize(document);
            if (pagePaths.Contains(path) || ReferencesFile(entries, path)) continue;
            if (entries.TryGetValue(path, out var bytes) && TryParse(bytes, out var xmlPage)) AddReferences(xmlPage, candidateIds);
            Remove(entries, path, result);
            var resourcePath = OfdPackagePath.GetDirectory(path) + "/PageRes.xml";
            if (pagePaths.Any(pagePath => OfdPackagePath.GetDirectory(pagePath) == OfdPackagePath.GetDirectory(path)) ||
                ReferencesFile(entries, resourcePath)) continue;
            if (entries.TryGetValue(resourcePath, out var resourceBytes) && TryParse(resourceBytes, out var resources))
                foreach (var resourceId in resources.Descendants().Attributes("ID")) candidateIds.Add(resourceId.Value);
            privateResourcePaths.Add(resourcePath);
        }
    }

    private static void RemovePageAnnotations(
        IDictionary<string, byte[]> entries,
        ISet<string> deletedIds,
        ISet<string> candidateIds,
        OfdPackageWriteResult result)
    {
        foreach (var pair in entries.Where(pair => IsXml(pair.Key)).ToList())
        {
            if (!TryParse(pair.Value, out var xml) || xml.Root?.Name.LocalName != "Annotations") continue;
            var removedFiles = new List<string>();
            foreach (var page in xml.Root.Elements().Where(node =>
                node.Name.LocalName == "Page" && deletedIds.Contains(node.Attribute("PageID")?.Value ?? string.Empty)).ToList())
            {
                var file = page.Elements().FirstOrDefault(node => node.Name.LocalName == "FileLoc")?.Value;
                if (!string.IsNullOrWhiteSpace(file)) removedFiles.Add(OfdPackagePath.Resolve(pair.Key, file!));
                page.Remove();
            }

            if (removedFiles.Count == 0) continue;
            entries[pair.Key] = Serialize(xml);
            foreach (var file in removedFiles)
            {
                if (ReferencesFile(entries, file)) continue;
                if (entries.TryGetValue(file, out var bytes) && TryParse(bytes, out var annotation)) AddReferences(annotation, candidateIds);
                Remove(entries, file, result);
            }
        }
    }

    private static void PruneResources(
        IDictionary<string, byte[]> entries,
        ISet<string> candidateIds,
        OfdPackageWriteResult result)
    {
        var documents = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        var references = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in entries.Where(pair => IsXml(pair.Key)))
        {
            if (!TryParse(pair.Value, out var xml))
            {
                result.Warnings.Add($"Unused resources retained because extension XML '{pair.Key}' cannot be inspected safely.");
                return;
            }
            documents[pair.Key] = xml;
            AddReferences(xml, references);
        }

        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in documents)
        {
            if (pair.Value.Root?.Name.LocalName != "Res") continue;
            var changed = false;
            foreach (var resource in pair.Value.Descendants().Where(node => node.Name.LocalName is "Font" or "MultiMedia").ToList())
            {
                var id = resource.Attribute("ID")?.Value;
                if (id is null || !candidateIds.Contains(id) || references.Contains(id)) continue;
                var file = resource.Elements().FirstOrDefault(node => node.Name.LocalName is "FontFile" or "MediaFile")?.Value
                    ?? resource.Attribute("MediaFile")?.Value;
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
            if (!ReferencesFile(entries, file)) Remove(entries, file, result);
        }
    }

    private static void RemoveInvalidatedSignatures(
        XDocument originalRoot,
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken)
    {
        var declarations = SignatureDeclarations(originalRoot).ToList();
        if (declarations.Count == 0 ||
            (original.Count == entries.Count && original.All(pair => entries.TryGetValue(pair.Key, out var bytes) && bytes.SequenceEqual(pair.Value)))) return;

        CleanSignatures(original, entries, result, cancellationToken);
        result.SignaturesInvalidated = true;
        result.Warnings.Add("Signature declarations were removed because the package was rewritten; sign the completed output again if required.");
    }

    internal static void CleanSignatures(
        IReadOnlyDictionary<string, byte[]> original,
        IDictionary<string, byte[]> entries,
        OfdPackageWriteResult result, CancellationToken cancellationToken = default)
    {
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
                    if (IsOwnedSignaturePath(payload) && payload.StartsWith(directory, StringComparison.OrdinalIgnoreCase) &&
                        !payload.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
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
        var pending = new Queue<string>(entries.Keys.Where(path => !candidates.Contains(path) && (IsXml(path) || knownXml.ContainsKey(path))));
        var scanned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = pending.Dequeue();
            if (!scanned.Add(path) || (!IsXml(path) && !knownXml.ContainsKey(path))) continue;
            if (!knownXml.TryGetValue(path, out var xml) && !TryParse(entries[path], out xml))
            {
                result.Warnings.Add($"Signature payloads retained because extension XML '{path}' cannot be inspected safely.");
                return;
            }
            var values = xml.Descendants().Where(node => !node.HasElements).Select(node => node.Value)
                .Concat(xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName != "ID").Select(attribute => attribute.Value));
            foreach (var value in values.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                cancellationToken.ThrowIfCancellationRequested();
                RestoreReference(path, value, xml.Root!, candidates, pending);
            }
        }
        foreach (var path in candidates)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Remove(entries, path, result);
        }
    }

    private static void RestoreReference(string path, string value, XElement root,
        ISet<string> candidates, Queue<string> pending)
    {
        try
        {
            var reference = OfdPackagePath.Resolve(path, value);
            if (candidates.Remove(reference)) pending.Enqueue(reference);
            reference = ResolveResourceFile(path, root, value);
            if (candidates.Remove(reference)) pending.Enqueue(reference);
        }
        catch (InvalidDataException) { }
    }

    private static bool IsSignatureRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private static IEnumerable<XElement> SignatureDeclarations(XDocument document) =>
        document.Root?.Elements().Where(node => node.Name == document.Root.Name.Namespace + "DocBody")
            .Elements(document.Root.Name.Namespace + "Signatures") ?? Enumerable.Empty<XElement>();

    private static bool IsOwnedSignaturePath(string path) =>
        path.Split('/').Any(segment => string.Equals(segment, "Signs", StringComparison.OrdinalIgnoreCase));

    private static bool ReferencesFile(IDictionary<string, byte[]> entries, string file, bool preserveOnUnknownXml = true)
    {
        foreach (var pair in entries.Where(pair => IsXml(pair.Key)))
        {
            if (!TryParse(pair.Value, out var xml))
            {
                if (preserveOnUnknownXml) return true;
                continue;
            }
            var values = xml.Descendants().Where(node => !node.HasElements).Select(node => node.Value)
                .Concat(xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName != "ID").Select(attribute => attribute.Value));
            foreach (var value in values.Where(value => !string.IsNullOrWhiteSpace(value)))
            {
                try
                {
                    if (string.Equals(OfdPackagePath.Resolve(pair.Key, value), file, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(ResolveResourceFile(pair.Key, xml.Root!, value), file, StringComparison.OrdinalIgnoreCase)) return true;
                }
                catch (InvalidDataException) { }
            }
        }
        return false;
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

    private static bool IsXml(string path) => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);
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
