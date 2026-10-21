using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

/// <summary>Conservative admission for editing one document while retaining the other DocBodies.</summary>
internal static class OfdMultiDocumentWriteGuard
{
    internal static void ValidateSource(OfdDocumentPackage package, CancellationToken token = default)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var bytes)) return;
        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(package, token); material.GuardPreservedXml(package, token);
        var root = material.ReadXml("OFD.xml", bytes, token);
        ValidateBefore(package, root, package.PreservedEntries, package.PreservedEntries, material, token);
    }

    internal static void ValidateBefore(OfdDocumentPackage package, XDocument root,
        IReadOnlyDictionary<string, byte[]> original, IDictionary<string, byte[]> entries,
        OfdModelMaterialization material, CancellationToken token)
    {
        var bodies = Bodies(root);
        if (bodies.Length < 2) return;
        if (package.DocumentEntryPath is null) throw Unsupported("selected DocRoot is missing");
        var roots = Roots(bodies, material, token);
        if (!roots.Contains(package.DocumentEntryPath, StringComparer.Ordinal)) throw Unsupported("selected DocRoot is not declared");
        if (roots.Any(path => !original.ContainsKey(path))) throw Unsupported("a declared document entry is missing");
        var directories = roots.Select(path => path.Split('/')[0]).ToArray();
        if (roots.Any(path => !path.Contains('/')) || directories.Distinct(StringComparer.OrdinalIgnoreCase).Count() != roots.Length)
            throw Unsupported("document directories are not distinct");
        if (bodies.Any(body => body.Elements(body.Name.Namespace + "Signatures").Any()) ||
            original.Keys.Any(path => path.Split('/').Any(part => part.Equals("Signs", StringComparison.OrdinalIgnoreCase))))
            throw Unsupported("signed or signature-bearing packages require explicit whole-package signature cleanup");
        foreach (var entry in original)
        {
            token.ThrowIfCancellationRequested();
            if (entry.Key == "OFD.xml") continue;
            var owner = directories.FirstOrDefault(dir => entry.Key.StartsWith(dir + "/", StringComparison.OrdinalIgnoreCase));
            if (owner is null) throw Unsupported("package has entries outside isolated document directories");
            if (!OfdXmlContentProbe.LooksLikeXml(entry.Value)) continue;
            XDocument xml;
            try { xml = material.ReadXml(entry.Key, entry.Value, token); }
            catch (System.Xml.XmlException) { throw Unsupported("an XML-looking entry cannot be inspected"); }
            // BaseLoc can name a directory rather than an existing archive entry.
            // A Res/FontFile (or MediaFile) is resolved relative to that base,
            // so checking its leaf value alone misses compound cross-doc paths.
            foreach (var baseLoc in xml.Descendants().Attributes().Where(attribute => attribute.Name.LocalName == "BaseLoc"))
            {
                token.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(baseLoc.Value)) continue;
                string resolved;
                try { resolved = OfdPackagePath.Resolve(entry.Key, baseLoc.Value.Trim()); }
                catch (InvalidDataException) { throw Unsupported("a BaseLoc cannot be inspected"); }
                if (!resolved.Equals(owner, StringComparison.OrdinalIgnoreCase) &&
                    !resolved.StartsWith(owner + "/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("cross-document BaseLoc");
            }
            var resourceBase = xml.Root?.Name.LocalName == "Res" ? xml.Root.Attribute("BaseLoc")?.Value : null;
            if (!string.IsNullOrWhiteSpace(resourceBase))
            {
                foreach (var file in ResourceFiles(xml.Root!))
                {
                    token.ThrowIfCancellationRequested();
                    if (string.IsNullOrWhiteSpace(file) || file.StartsWith("/", StringComparison.Ordinal)) continue;
                    string resolved;
                    try { resolved = OfdPackagePath.Resolve(entry.Key, resourceBase!.TrimEnd('/') + "/" + file); }
                    catch (InvalidDataException) { throw Unsupported("a resource path cannot be inspected"); }
                    if (!resolved.StartsWith(owner + "/", StringComparison.OrdinalIgnoreCase))
                        throw Unsupported("cross-document compound resource path");
                }
            }
            // Any path resolving to an existing entry of another document is a
            // cross-document dependency. Unknown paths are retained by the writer.
            foreach (var value in Values(xml))
            {
                token.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(value)) continue;
                string resolved;
                try { resolved = OfdPackagePath.Resolve(entry.Key, value.Trim()); }
                catch (InvalidDataException) { continue; }
                if (original.ContainsKey(resolved) && !resolved.StartsWith(owner + "/", StringComparison.OrdinalIgnoreCase))
                    throw Unsupported("cross-document resource or reference");
            }
        }
        ValidateAfter(package, root, original, entries, material, token);
    }

    internal static void ValidateAfter(OfdDocumentPackage package, XDocument root,
        IReadOnlyDictionary<string, byte[]> original, IDictionary<string, byte[]> entries,
        OfdModelMaterialization material, CancellationToken token)
    {
        var bodies = Bodies(root);
        if (bodies.Length < 2) return;
        var roots = Roots(bodies, material, token);
        var selected = package.DocumentEntryPath ?? throw Unsupported("selected DocRoot is missing");
        var others = roots.Where(path => path != selected).Select(path => path.Split('/')[0] + "/").ToArray();
        foreach (var prefix in others)
        {
            var before = original.Where(pair => pair.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToArray();
            var after = entries.Where(pair => pair.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (before.Length != after.Length || before.Any(pair => !entries.TryGetValue(pair.Key, out var bytes) || !bytes.SequenceEqual(pair.Value)))
                throw Unsupported("an unselected document entry changed");
        }
        if (!entries.TryGetValue("OFD.xml", out var written)) throw Unsupported("OFD.xml is missing from output");
        var newRoot = material.ReadXml("OFD.xml", written, token);
        var newBodies = Bodies(newRoot);
        if (newBodies.Length != bodies.Length) throw Unsupported("DocBody count changed");
        for (var i = 0; i < bodies.Length; i++)
            if (roots[i] != selected && !XNode.DeepEquals(bodies[i], newBodies[i]))
                throw Unsupported("an unselected DocBody declaration changed");
    }

    private static XElement[] Bodies(XDocument root) => root.Root?.Elements(root.Root.Name.Namespace + "DocBody").ToArray() ?? Array.Empty<XElement>();

    private static string[] Roots(XElement[] bodies, OfdModelMaterialization material, CancellationToken token) =>
        bodies.Select(body =>
        {
            var nodes = body.Elements(body.Name.Namespace + "DocRoot").ToArray();
            if (nodes.Length != 1 || string.IsNullOrWhiteSpace(nodes[0].Value)) throw Unsupported("DocBody has no unique DocRoot");
            return material.ResolvePath("OFD.xml", nodes[0].Value, token);
        }).ToArray();

    private static IEnumerable<string> Values(XDocument xml)
    {
        foreach (var node in xml.Descendants().Where(node => !node.HasElements)) yield return node.Value;
        foreach (var attribute in xml.Descendants().Attributes().Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName != "ID"))
            yield return attribute.Value;
    }

    private static IEnumerable<string> ResourceFiles(XElement root)
    {
        foreach (var fontFile in root.Descendants().Where(node => node.Name.LocalName == "FontFile"))
            yield return fontFile.Value.Trim();
        foreach (var media in root.Descendants().Where(node => node.Name.LocalName == "MultiMedia"))
        {
            var path = media.Attribute("MediaFile")?.Value ?? media.Elements().FirstOrDefault(node => node.Name.LocalName == "MediaFile")?.Value;
            if (path is not null) yield return path.Trim();
        }
    }

    private static NotSupportedException Unsupported(string reason) => new("Multi-DocBody edit requires isolated unsigned documents: " + reason + ".");
}
