using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
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

    internal OfdResourceCatalog(IDictionary<string, byte[]> entries)
    {
        _entries = entries;
        foreach (var pair in entries.Where(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            XDocument xml;
            try
            {
                using var stream = new MemoryStream(pair.Value, writable: false);
                xml = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }
            catch (XmlException) { continue; }
            if (xml.Root?.Name.LocalName != "Res") continue;
            _documents[pair.Key] = xml;
            foreach (var resource in xml.Descendants().Where(node => node.Name.LocalName is "Font" or "MultiMedia"))
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
        _documents[path] = new XDocument(new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Res", new XAttribute("BaseLoc", "Res")));
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
        var parent = root.Elements().FirstOrDefault(node => node.Name.LocalName == container);
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
            foreach (var payload in element.Elements().Where(node => node.Name.LocalName is "FontFile" or "MediaFile"))
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
        var element = resource.Elements().FirstOrDefault(node => node.Name.LocalName == name);
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
