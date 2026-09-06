using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static class OfdPackagePruner
{
    internal static OfdPackageWriteResult Prune(
        OfdDocumentPackage package,
        IDictionary<string, byte[]> entries)
    {
        var result = new OfdPackageWriteResult();
        var original = package.PreservedEntries;
        if (!original.TryGetValue("OFD.xml", out var originalRootBytes)) return result;
        var originalRoot = Parse(originalRootBytes);
        var documentPath = originalRoot.Descendants().FirstOrDefault(node => node.Name.LocalName == "DocRoot")?.Value;
        if (string.IsNullOrWhiteSpace(documentPath) || !original.TryGetValue(documentPath!, out var originalDocument)) return result;

        // Invalidated signature references must not keep deleted page payloads
        // alive during the subsequent resource reachability check.
        RemoveInvalidatedSignatures(originalRoot, original, entries, result);

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
            PruneResources(entries, candidateIds, result);
        }

        return result;
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
        OfdPackageWriteResult result)
    {
        var declarations = originalRoot.Descendants().Where(node => node.Name.LocalName == "Signatures").ToList();
        if (declarations.Count == 0 ||
            (original.Count == entries.Count && original.All(pair => entries.TryGetValue(pair.Key, out var bytes) && bytes.SequenceEqual(pair.Value)))) return;

        var root = Parse(entries["OFD.xml"]);
        root.Descendants().Where(node => node.Name.LocalName == "Signatures").Remove();
        entries["OFD.xml"] = Serialize(root);
        var descriptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var payloads = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in declarations)
        {
            var listPath = OfdPackagePath.Resolve("OFD.xml", declaration.Value);
            if (original.TryGetValue(listPath, out var listBytes) && TryParse(listBytes, out var list) && list.Root?.Name.LocalName == "Signatures")
            {
                descriptions.Add(listPath);
                foreach (var record in list.Descendants().Where(node => node.Name.LocalName == "Signature"))
                {
                    var location = record.Attribute("BaseLoc")?.Value;
                    if (string.IsNullOrWhiteSpace(location)) continue;
                    var signaturePath = OfdPackagePath.Resolve(listPath, location!);
                    if (original.TryGetValue(signaturePath, out var signatureBytes) && TryParse(signatureBytes, out var signature) && signature.Root?.Name.LocalName == "Signature")
                    {
                        descriptions.Add(signaturePath);
                        foreach (var reference in signature.Descendants().Where(node =>
                            node.Name.LocalName == "SignedValue" ||
                            (node.Name.LocalName == "BaseLoc" && node.Parent?.Name.LocalName == "Seal")))
                        {
                            payloads.Add(OfdPackagePath.Resolve(signaturePath, reference.Value));
                        }
                    }
                }
            }
        }
        // Only remove typed signature descriptions and payloads that no retained
        // document, resource or extension references. Never trust an arbitrary
        // SignedValue path as authority to delete a live document entry.
        var remaining = entries.Where(pair => !descriptions.Contains(pair.Key))
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        foreach (var path in descriptions)
            if (!ReferencesFile(remaining, path)) Remove(entries, path, result);
        foreach (var path in payloads)
            if (!ReferencesFile(entries, path)) Remove(entries, path, result);
        result.SignaturesInvalidated = true;
        result.Warnings.Add("Signature declarations were removed because the package was rewritten; known unreferenced signature payloads were cleaned up. Sign the completed output again if required.");
    }

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
