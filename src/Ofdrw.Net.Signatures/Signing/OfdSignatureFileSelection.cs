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
        OfdSignatureOptions options,
        CancellationToken cancellationToken)
    {
        var required = options.ProtectFile is null
            ? null
            : FindRequiredFiles(entries, ofd, cancellationToken);
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
        foreach (var entry in entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var path = entry.Key;
            if (path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) && !IsSignatureMaterial(path))
                required.Add(path);
            if (!IsSignatureMaterial(path) && directories.Any(dir =>
                    path.StartsWith(dir, StringComparison.OrdinalIgnoreCase)))
                required.Add(path);
        }

        // External shared resources may sit outside a DocBody directory. A path
        // explicitly named by any non-signature XML also remains protected.
        foreach (var entry in entries.Where(pair =>
                     pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                     !IsSignatureMaterial(pair.Key)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            XDocument document;
            try
            {
                using var stream = new MemoryStream(entry.Value, writable: false);
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
                    "Cannot prove document dependency closure for /" + entry.Key, exception);
            }
            foreach (var candidate in document.Descendants()
                         .SelectMany(node => node.Attributes().Where(a => !a.IsNamespaceDeclaration)
                             .Select(a => a.Value).Concat(
                                 node.HasElements ? Array.Empty<string>() : new[] { node.Value })))
            {
                var value = candidate.Trim();
                if (value.Length == 0 || value.Length > 1024) continue;
                var resolved = Resolve(entry.Key, value);
                if (entries.ContainsKey(resolved)) required.Add(resolved);
                var fromRoot = Normalize(value);
                if (entries.ContainsKey(fromRoot)) required.Add(fromRoot);
            }
        }
        return required;
    }

    private static bool IsSignatureMaterial(string path) =>
        path.StartsWith("Signs/", StringComparison.OrdinalIgnoreCase) ||
        path.IndexOf("/Signs/", StringComparison.OrdinalIgnoreCase) >= 0;

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
