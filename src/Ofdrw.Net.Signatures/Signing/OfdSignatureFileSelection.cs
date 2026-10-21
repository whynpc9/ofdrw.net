using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Signatures.Signing;

internal static class OfdSignatureFileSelection
{
    internal static (string[] Protected, string[] Unprotected) Select(
        IReadOnlyDictionary<string, byte[]> entries,
        XDocument ofd,
        string signaturesPath,
        string newSignaturePath,
        OfdSignatureOptions options,
        CancellationToken cancellationToken)
    {
        var required = options.ProtectFile is null
            ? null
            : FindRequiredFiles(entries, ofd, signaturesPath, newSignaturePath, cancellationToken);
        var selected = new List<string>();
        var omitted = new List<string>();
        foreach (var path in entries.Keys.OrderBy(x => x, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var isList = string.Equals(path, signaturesPath, StringComparison.OrdinalIgnoreCase);
            var include = !isList || options.ProtectSignatureList;
            if (include && options.ProtectFile is not null)
            {
                include = options.ProtectFile("/" + path);
                if (!include && required!.Contains(path))
                    throw new InvalidDataException("Signature selection omits required document entry: /" + path);
                if (!include && isList)
                    throw new InvalidDataException("ProtectSignatureList requires /" + path);
            }
            (include ? selected : omitted).Add(path);
        }
        if (selected.Count == 0)
            throw new InvalidDataException("Signature selection contains no package entries.");
        return (selected.ToArray(), omitted.ToArray());
    }

    private static HashSet<string> FindRequiredFiles(
        IReadOnlyDictionary<string, byte[]> entries, XDocument ofd,
        string signaturesPath, string newSignaturePath,
        CancellationToken cancellationToken)
    {
        var required = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "OFD.xml" };
        var documentPaths = ofd.Descendants().Where(e => e.Name.LocalName == "DocRoot")
            .Select(e => Normalize(e.Value)).ToArray();
        if (documentPaths.Length == 0)
            throw new InvalidDataException("OFD.xml has no DocRoot for scope validation.");
        foreach (var path in documentPaths)
        {
            if (path.Length == 0 || !entries.ContainsKey(path))
                throw new InvalidDataException("Cannot prove scope for missing DocRoot: /" + path);
            required.Add(path);
        }
        var directories = documentPaths
            .Select(path => path.Substring(0, path.LastIndexOf('/') + 1)).ToArray();
        var mutableAnnotations = FindMutableAnnotations(entries, documentPaths);
        var mutableSignatureFiles = FindDeclaredSignatureFiles(entries, signaturesPath, newSignaturePath);
        var mutable = new HashSet<string>(mutableAnnotations, StringComparer.OrdinalIgnoreCase);
        mutable.UnionWith(mutableSignatureFiles);
        foreach (var entry in entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = entry.Key;
            if (IsXmlCandidate(path, entry.Value) && !mutable.Contains(path))
                required.Add(path);
            if (!mutable.Contains(path) && directories.Any(dir =>
                    path.StartsWith(dir, StringComparison.OrdinalIgnoreCase)))
                required.Add(path);
        }

        // External shared resources may sit outside a DocBody directory. A path
        // explicitly named by any non-signature XML also remains protected.
        var xmlPaths = new HashSet<string>(entries.Where(pair =>
                IsXmlCandidate(pair.Key, pair.Value) && !mutableSignatureFiles.Contains(pair.Key))
            .Select(pair => pair.Key), StringComparer.OrdinalIgnoreCase);
        xmlPaths.UnionWith(documentPaths);
        xmlPaths.UnionWith(mutableAnnotations);
        var pending = new Queue<string>(xmlPaths.OrderBy(path => path, StringComparer.Ordinal));
        while (pending.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = pending.Dequeue();
            var bytes = entries[path];
            XDocument document;
            try
            {
                using var stream = new MemoryStream(bytes, writable: false);
                using var reader = XmlReader.Create(stream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = 16 * 1024 * 1024
                });
                document = XDocument.Load(reader);
            }
            catch (Exception exception) when (exception is XmlException or InvalidOperationException)
            {
                throw new InvalidDataException(
                    "Cannot prove document dependency closure for /" + path, exception);
            }
            foreach (var table in document.Root!.DescendantsAndSelf().Where(node =>
                         node.Name.LocalName == "PublicRes" || node.Name.LocalName == "DocumentRes"))
            {
                var location = Resolve(path, table.Value.Trim());
                if (!entries.ContainsKey(location))
                    throw new InvalidDataException("Cannot prove resource table dependency: /" + path);
                required.Add(location);
                if (xmlPaths.Add(location)) pending.Enqueue(location);
            }
            foreach (var candidate in document.Root!.DescendantsAndSelf()
                         .SelectMany(node => node.Attributes().Where(a => !a.IsNamespaceDeclaration)
                             .Select(a => a.Value).Concat(
                                 node.HasElements ? Array.Empty<string>() : new[] { node.Value })))
            {
                var value = candidate.Trim();
                if (value.Length == 0) continue;
                var resolved = Resolve(path, value);
                if (entries.ContainsKey(resolved) && !mutable.Contains(resolved))
                    required.Add(resolved);
                var fromRoot = Normalize(value);
                if (entries.ContainsKey(fromRoot) && !mutable.Contains(fromRoot))
                    required.Add(fromRoot);
            }
            foreach (var page in document.Root.DescendantsAndSelf().Where(node =>
                         node.Name.LocalName == "Page" || node.Name.LocalName == "TemplatePage"))
            {
                var baseLoc = page.Attribute("BaseLoc")?.Value;
                if (string.IsNullOrWhiteSpace(baseLoc)) continue;
                var pagePath = Resolve(path, baseLoc!);
                if (!entries.ContainsKey(pagePath))
                    throw new InvalidDataException("Cannot prove page dependency: /" + path + " -> " + baseLoc);
                required.Add(pagePath);
                if (xmlPaths.Add(pagePath)) pending.Enqueue(pagePath);
            }
            AddComposedResourceFiles(path, document, entries, required);
        }
        return required;
    }

    private static HashSet<string> FindDeclaredSignatureFiles(
        IReadOnlyDictionary<string, byte[]> entries, string signaturesPath, string newSignaturePath)
    {
        var optional = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!entries.TryGetValue(signaturesPath, out var listBytes)) return optional;
        var list = ParseDependencyXml(listBytes, signaturesPath);
        foreach (var record in list.Root!.Elements().Where(node => node.Name.LocalName == "Signature"))
        {
            var baseLoc = record.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(baseLoc))
                throw new InvalidDataException("Cannot prove declared signature location in /" + signaturesPath);
            var signaturePath = Resolve(signaturesPath, baseLoc!);
            if (string.Equals(signaturePath, newSignaturePath, StringComparison.OrdinalIgnoreCase)) continue;
            if (!entries.TryGetValue(signaturePath, out var signatureBytes))
                throw new InvalidDataException("Cannot prove declared signature: /" + signaturePath);
            var signature = ParseDependencyXml(signatureBytes, signaturePath);
            var valueLoc = signature.Descendants()
                .FirstOrDefault(node => node.Name.LocalName == "SignedValue")?.Value;
            if (string.IsNullOrWhiteSpace(valueLoc))
                throw new InvalidDataException("Cannot prove SignedValue in /" + signaturePath);
            var valuePath = Resolve(signaturePath, valueLoc!.Trim());
            if (!entries.ContainsKey(valuePath))
                throw new InvalidDataException("Cannot prove SignedValue: /" + valuePath);
            optional.Add(signaturePath);
            optional.Add(valuePath);
        }
        return optional;
    }

    private static void AddComposedResourceFiles(
        string resourceXmlPath, XDocument document,
        IReadOnlyDictionary<string, byte[]> entries, HashSet<string> required)
    {
        var root = document.Root!;
        var baseLocation = root.Attribute("BaseLoc")?.Value;
        foreach (var node in root.Descendants())
        {
            string? file = null;
            if (node.Name.LocalName == "MultiMedia")
                file = node.Attribute("MediaFile")?.Value ??
                    node.Elements().FirstOrDefault(child => child.Name.LocalName == "MediaFile")?.Value;
            else if (node.Name.LocalName == "Font")
                file = node.Elements().FirstOrDefault(child => child.Name.LocalName == "FontFile")?.Value;
            if (string.IsNullOrWhiteSpace(file)) continue;
            var location = string.IsNullOrWhiteSpace(baseLocation) || file!.StartsWith("/", StringComparison.Ordinal)
                ? file!
                : baseLocation!.TrimEnd('/') + "/" + file.TrimStart('/');
            var resolved = Resolve(resourceXmlPath, location);
            if (resolved.Length == 0 || !entries.ContainsKey(resolved))
                throw new InvalidDataException("Cannot prove resource dependency: /" + resourceXmlPath + " -> " + location);
            required.Add(resolved);
        }
    }

    private static HashSet<string> FindMutableAnnotations(
        IReadOnlyDictionary<string, byte[]> entries, IEnumerable<string> documentPaths)
    {
        var optional = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var documentPath in documentPaths)
        {
            var document = ParseDependencyXml(entries[documentPath], documentPath);
            var ns = document.Root!.Name.Namespace;
            var pageIds = new HashSet<string>(document.Root.Descendants(ns + "Pages")
                .Elements(ns + "Page").Select(page => page.Attribute("ID")?.Value)
                .Where(id => !string.IsNullOrWhiteSpace(id)).Select(id => id!), StringComparer.Ordinal);
            foreach (var declaration in document.Root.Elements(ns + "Annotations"))
            {
                if (declaration.HasElements || declaration.HasAttributes || string.IsNullOrWhiteSpace(declaration.Value))
                    throw new InvalidDataException("Cannot prove annotation index location in /" + documentPath);
                var indexPath = Resolve(documentPath, declaration.Value.Trim());
                if (!entries.TryGetValue(indexPath, out var indexBytes))
                    throw new InvalidDataException("Cannot prove annotation index: /" + indexPath);
                var index = ParseDependencyXml(indexBytes, indexPath);
                if (index.Root?.Name != ns + "Annotations")
                    throw new InvalidDataException("Cannot prove annotation index structure: /" + indexPath);
                if (index.Root.Elements().Any(node => node.Name != ns + "Page"))
                    throw new InvalidDataException("Cannot prove annotation index records: /" + indexPath);
                optional.Add(indexPath);
                foreach (var page in index.Root.Elements(ns + "Page"))
                {
                    var locations = page.Elements(ns + "FileLoc").ToArray();
                    if (!pageIds.Contains(page.Attribute("PageID")?.Value ?? string.Empty) ||
                        page.Attributes().Any(attribute => attribute.Name != "PageID" && !attribute.IsNamespaceDeclaration) ||
                        page.Elements().Count() != 1 || locations.Length != 1 || locations[0].HasElements ||
                        locations[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                        string.IsNullOrWhiteSpace(locations[0].Value))
                        throw new InvalidDataException("Cannot prove annotation record in /" + indexPath);
                    var annotationPath = Resolve(indexPath, locations[0].Value.Trim());
                    if (!entries.TryGetValue(annotationPath, out var annotationBytes) ||
                        ParseDependencyXml(annotationBytes, annotationPath).Root?.Name != ns + "PageAnnot")
                        throw new InvalidDataException("Cannot prove annotation record: /" + annotationPath);
                    optional.Add(annotationPath);
                }
            }
        }
        return optional;
    }

    private static XDocument ParseDependencyXml(byte[] bytes, string path)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 16 * 1024 * 1024
            });
            return XDocument.Load(reader);
        }
        catch (Exception exception) when (exception is XmlException or InvalidOperationException)
        {
            throw new InvalidDataException("Cannot prove document dependency closure for /" + path, exception);
        }
    }

    private static bool IsXmlCandidate(string path, byte[] bytes)
    {
        if (path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return true;
        var offset = 0;
        if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf)
            offset = 3;
        else if (bytes.Length >= 2 &&
                 ((bytes[0] == 0xff && bytes[1] == 0xfe) ||
                  (bytes[0] == 0xfe && bytes[1] == 0xff)))
            return true;
        while (offset < bytes.Length &&
               (bytes[offset] == ' ' || bytes[offset] == '\t' ||
                bytes[offset] == '\r' || bytes[offset] == '\n')) offset++;
        return offset < bytes.Length && bytes[offset] == '<';
    }

    private static string Resolve(string containingPath, string value)
    {
        if (value.StartsWith("/", StringComparison.Ordinal)) return Normalize(value);
        var slash = containingPath.LastIndexOf('/');
        return Normalize((slash < 0 ? string.Empty : containingPath.Substring(0, slash + 1)) + value);
    }

    private static string Normalize(string value)
    {
        var parts = new List<string>();
        foreach (var part in value.Replace('\\', '/').TrimStart('/').Split('/'))
        {
            if (part == "..")
            {
                if (parts.Count == 0) return string.Empty;
                parts.RemoveAt(parts.Count - 1);
            }
            else if (part.Length > 0 && part != ".") parts.Add(part);
        }
        return string.Join("/", parts);
    }
}
