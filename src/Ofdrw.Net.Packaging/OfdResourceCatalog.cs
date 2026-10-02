using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

/// <summary>Updates resources at their existing locations and gives new payloads content-based names.</summary>
internal sealed class OfdResourceCatalog
{
    private readonly IDictionary<string, byte[]> _entries;
    private readonly Dictionary<string, XDocument> _documents = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (string Path, XElement Element)> _resources = new(StringComparer.Ordinal);
    private readonly HashSet<string> _changed = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<byte[], string> _hashes = new();

    internal OfdResourceCatalog(IDictionary<string, byte[]> entries, string documentPath, XNamespace documentNamespace,
        IEnumerable<string> newResourcePaths)
    {
        _entries = entries;
        var owned = new HashSet<string>(newResourcePaths, StringComparer.OrdinalIgnoreCase);
        if (entries.TryGetValue(documentPath, out var documentBytes))
        {
            XDocument document;
            using (var input = new MemoryStream(documentBytes, false)) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
            var ns = document.Root!.Name.Namespace; var common = document.Root.Element(ns + "CommonData");
            foreach (var declaration in common?.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
                owned.Add(OfdPackagePath.Resolve(documentPath, declaration.Value));
            foreach (var page in (document.Root.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
                .Concat(common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>()))
            {
                var location = page.Attribute("BaseLoc")?.Value;
                if (!string.IsNullOrWhiteSpace(location)) owned.Add(OfdPackagePath.GetDirectory(OfdPackagePath.Resolve(documentPath, location!)) + "/PageRes.xml");
            }
        }
        foreach (var pair in entries.Where(pair => owned.Contains(pair.Key)))
        {
            XDocument xml;
            try
            {
                using var stream = new MemoryStream(pair.Value, writable: false);
                xml = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }
            catch (XmlException) { continue; }
            if (!owned.Contains(pair.Key) || xml.Root?.Name.LocalName != "Res" ||
                (xml.Root.Name.Namespace != documentNamespace && xml.Root.Name.NamespaceName != OfdConstants.Namespace && xml.Root.Name.NamespaceName != OfdConstants.StandardNamespace)) continue;
            _documents[pair.Key] = xml;
            foreach (var resource in xml.Root.Elements(xml.Root.Name.Namespace + "Fonts").Elements(xml.Root.Name.Namespace + "Font")
                .Concat(xml.Root.Elements(xml.Root.Name.Namespace + "MultiMedias").Elements(xml.Root.Name.Namespace + "MultiMedia")))
            {
                var id = resource.Attribute("ID")?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                var key = resource.Name.LocalName + "\u001f" + id;
                if (_resources.ContainsKey(key)) throw new InvalidDataException($"Duplicate resource ID '{id}' in OFD resources.");
                _resources.Add(key, (pair.Key, resource));
            }
        }
    }

    internal void EnsureDocument(string path, XNamespace ns)
    {
        if (_documents.ContainsKey(path)) return;
        if (_entries.ContainsKey(path)) throw new NotSupportedException($"Cannot overwrite an unmodeled resource entry '{path}'.");
        // Declare the "ofd" prefix like Document.xml and Content.xml do. Several
        // readers match "ofd:Res"/"ofd:MultiMedia" textually and never find the
        // image manifest when the root uses an unprefixed default namespace.
        _documents[path] = new XDocument(new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Res",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XAttribute("BaseLoc", "Res")));
        _changed.Add(path);
    }

    internal void WriteFont(string id, OfdFontResource font, string defaultPath, XNamespace ns)
    {
        var (path, element) = GetOrCreate("Font", "Fonts", id, defaultPath, ns);
        element.SetAttributeValue("FontName", font.FontName);
        element.SetAttributeValue("FamilyName", font.FamilyName ?? font.FontName);
        element.SetAttributeValue("Charset", font.Charset);
        element.SetAttributeValue("Bold", font.Bold ? true : (object?)null);
        element.SetAttributeValue("Italic", font.Italic ? true : (object?)null);
        if (font.Data.Length > 0)
        {
            var extension = string.Equals(Path.GetExtension(font.FileName), ".otf", StringComparison.OrdinalIgnoreCase)
                ? ".otf" : ".ttf";
            WritePayload(path, element, "FontFile", $"Font_{Hash(font.Data)}{extension}", font.Data);
        }
        _changed.Add(path);
    }

    internal void WriteImage(string id, OfdImageElement image, string format, string defaultPath, XNamespace ns)
    {
        var (path, element) = GetOrCreate("MultiMedia", "MultiMedias", id, defaultPath, ns);
        element.SetAttributeValue("Type", "Image");
        element.SetAttributeValue("Format", format);
        WritePayload(path, element, "MediaFile", $"Image_{Hash(image.Data)}.{format.ToLowerInvariant()}", image.Data);
        _changed.Add(path);
    }

    private (string Path, XElement Element) GetOrCreate(string kind, string container, string id, string defaultPath, XNamespace ns)
    {
        var key = kind + "\u001f" + id;
        if (_resources.TryGetValue(key, out var existing) &&
            string.Equals(existing.Path, defaultPath, StringComparison.OrdinalIgnoreCase)) return existing;
        EnsureDocument(defaultPath, ns);
        var root = _documents[defaultPath].Root!;
        var parent = root.Element(root.Name.Namespace + container);
        if (parent is null)
        {
            parent = new XElement(root.Name.Namespace + container);
            root.Add(parent);
        }
        XElement element;
        if (existing.Element is not null)
        {
            // Promote page-local resources to the shared resource file so a
            // reordered/copied page need not keep its old private PageRes path.
            element = existing.Element;
            var oldRoot = _documents[existing.Path].Root!;
            foreach (var payload in element.Elements().Where(node => node.Name == element.Name.Namespace + "FontFile" || node.Name == element.Name.Namespace + "MediaFile"))
            {
                var oldBase = oldRoot.Attribute("BaseLoc")?.Value;
                var reference = string.IsNullOrEmpty(oldBase) || payload.Value.StartsWith("/", StringComparison.Ordinal)
                    ? payload.Value : oldBase!.TrimEnd('/') + "/" + payload.Value;
                payload.Value = "/" + OfdPackagePath.Resolve(existing.Path, reference);
            }
            var legacyFile = element.Attribute("MediaFile");
            if (legacyFile is not null)
            {
                var oldBase = oldRoot.Attribute("BaseLoc")?.Value;
                var reference = string.IsNullOrEmpty(oldBase) || legacyFile.Value.StartsWith("/", StringComparison.Ordinal)
                    ? legacyFile.Value : oldBase!.TrimEnd('/') + "/" + legacyFile.Value;
                legacyFile.Value = "/" + OfdPackagePath.Resolve(existing.Path, reference);
            }
            element.Remove();
            _changed.Add(existing.Path);
        }
        else element = new XElement(root.Name.Namespace + kind, new XAttribute("ID", id));
        parent.Add(element);
        var result = (defaultPath, element);
        _resources[key] = result;
        return result;
    }

    private void WritePayload(string resourcePath, XElement resource, string name, string defaultFileName, byte[] bytes)
    {
        var element = resource.Element(resource.Name.Namespace + name);
        var file = element?.Value ?? resource.Attribute(name)?.Value ?? defaultFileName;
        if (element is null) resource.Add(new XElement(resource.Name.Namespace + name, file));
        resource.SetAttributeValue(name, null);
        var baseLocation = _documents[resourcePath].Root!.Attribute("BaseLoc")?.Value;
        var reference = string.IsNullOrEmpty(baseLocation) || file.StartsWith("/", StringComparison.Ordinal)
            ? file : baseLocation!.TrimEnd('/') + "/" + file;
        _entries[OfdPackagePath.Resolve(resourcePath, reference)] = bytes;
    }

    internal void Flush()
    {
        foreach (var path in _changed)
        {
            using var stream = new MemoryStream();
            _documents[path].Save(stream, SaveOptions.DisableFormatting);
            _entries[path] = stream.ToArray();
        }
    }

    private string Hash(byte[] data)
    {
        if (!_hashes.TryGetValue(data, out var hash)) _hashes[data] = hash = BinaryIdentity.Hash(data);
        return hash;
    }
}
