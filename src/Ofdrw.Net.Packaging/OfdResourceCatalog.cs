using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

/// <summary>Updates resources at their existing locations and gives new payloads content-based names.</summary>
internal sealed class OfdResourceCatalog
{
    private readonly OfdModelMaterialization material;
    private readonly CancellationToken operationToken;
    private readonly IDictionary<string, byte[]> _entries;
    private readonly Dictionary<string, XDocument> _documents = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (string Path, XElement Element)> _resources = new(StringComparer.Ordinal);
    private readonly HashSet<string> _changed = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<byte[], string> _hashes = new();
    private readonly HashSet<string> protectedPaths = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> payloads = new(StringComparer.OrdinalIgnoreCase);

    internal OfdResourceCatalog(IDictionary<string, byte[]> entries, string documentPath, XNamespace documentNamespace,
        IEnumerable<string> newResourcePaths, OfdModelMaterialization material, CancellationToken token, OfdSourceIdentity? identity = null, OfdDocumentPackage? package = null,
        IEnumerable<string>? protectedOutputPaths = null)
    {
        _entries = entries; this.material = material; operationToken = token;
        foreach (var path in (protectedOutputPaths ?? Enumerable.Empty<string>()).Concat(identity?.ControlEntries ?? Enumerable.Empty<string>()))
        { material.ChargeCopy(128, token); protectedPaths.Add(path); }
        var owned = new HashSet<string>(newResourcePaths, StringComparer.OrdinalIgnoreCase);
        if (entries.TryGetValue(documentPath, out var documentBytes))
        {
            XDocument document;
            document = material.ReadXml(documentPath, documentBytes, token);
            var ns = document.Root!.Name.Namespace; var common = document.Root.Element(ns + "CommonData");
            foreach (var declaration in common?.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
                owned.Add(material.ResolvePath(documentPath, declaration.Value, operationToken));
            foreach (var page in (document.Root.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
                .Concat(common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>()))
            {
                var location = page.Attribute("BaseLoc")?.Value;
                if (!string.IsNullOrWhiteSpace(location)) owned.Add(OfdPackagePath.GetDirectory(material.ResolvePath(documentPath, location!, operationToken)) + "/PageRes.xml");
            }
        }
        foreach (var pair in entries.Where(pair => owned.Contains(pair.Key)))
        {
            XDocument xml;
            try
            {
                xml = material.ReadXml(pair.Key, pair.Value, token);
            }
            catch (XmlException) { continue; }
            if (!owned.Contains(pair.Key) || xml.Root?.Name.LocalName != "Res" ||
                (xml.Root.Name.Namespace != documentNamespace && xml.Root.Name.NamespaceName != OfdConstants.Namespace && xml.Root.Name.NamespaceName != OfdConstants.StandardNamespace)) continue;
            if (identity is null) throw new NotSupportedException("Preserved active resource tables must be read before editing.");
            if (package is not null) identity.ValidateResourceSource(pair.Key, package);
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

    private sealed class ResourcePlan
    {
        internal string Path = string.Empty, Kind = string.Empty, Container = string.Empty, Id = string.Empty;
        internal XNamespace Namespace = XNamespace.None;
        internal XElement? Existing;
        internal string? OldPath;
        internal readonly List<(XObject Node, string Value)> Relocations = new();
        internal readonly List<(string Name, object? Value)> Attributes = new();
        internal string? PayloadName, PayloadValue, PayloadPath;
        internal byte[]? Data;
    }
    private readonly List<ResourcePlan> plans = new();
    private readonly HashSet<string> plannedKeys = new(StringComparer.Ordinal);
    private readonly Dictionary<string, XNamespace> newDocuments = new(StringComparer.OrdinalIgnoreCase);

    internal void PlanDocument(string path, XNamespace ns)
    {
        operationToken.ThrowIfCancellationRequested();
        if (_documents.ContainsKey(path)) return;
        if (_entries.ContainsKey(path)) throw new NotSupportedException($"Cannot overwrite an unmodeled resource entry '{path}'.");
        newDocuments[path] = ns;
    }

    // Input DOMs may be inspected here. No output document/node exists until
    // Apply; all authority, relocation, hash and payload-path decisions finish now.
    internal void PlanFont(string id, OfdFontResource font, string defaultPath, XNamespace ns)
    {
        var plan = PlanResource("Font", "Fonts", id, defaultPath, ns);
        plan.Attributes.Add(("FontName", font.FontName)); plan.Attributes.Add(("FamilyName", font.FamilyName ?? font.FontName));
        plan.Attributes.Add(("Charset", font.Charset)); plan.Attributes.Add(("Bold", font.Bold ? true : (object?)null));
        plan.Attributes.Add(("Italic", font.Italic ? true : (object?)null));
        if (font.Data.Length > 0)
        {
            var extension = string.Equals(Path.GetExtension(font.FileName), ".otf", StringComparison.OrdinalIgnoreCase) ? ".otf" : ".ttf";
            // PlanResource and catalog source-identity admission already bound
            // Existing to the real imported owner. A present literal, including
            // an empty or conflicting one, must still follow PlanPayload's
            // existing validation order; its unused fallback needs no hash.
            var hasLiteral = plan.Existing?.Element(plan.Existing.Name.Namespace + "FontFile") is not null ||
                plan.Existing?.Attribute("FontFile") is not null;
            var fallback = hasLiteral ? null : $"Font_{Hash(font.Data)}{extension}";
            PlanPayload(plan, "FontFile", fallback, font.Data);
        }
    }

    internal void PlanImage(string id, OfdImageElement image, string format, string defaultPath, XNamespace ns)
    {
        var plan = PlanResource("MultiMedia", "MultiMedias", id, defaultPath, ns);
        plan.Attributes.Add(("Type", "Image")); plan.Attributes.Add(("Format", format));
        PlanPayload(plan, "MediaFile", $"Image_{Hash(image.Data)}.{format.ToLowerInvariant()}", image.Data);
    }

    private ResourcePlan PlanResource(string kind, string container, string id, string path, XNamespace ns)
    {
        material.ChargeCopy(512, operationToken);
        var key = kind + "\u001f" + id;
        if (!plannedKeys.Add(key)) throw new InvalidDataException("Duplicate planned resource owner.");
        if (!_documents.ContainsKey(path))
        {
            if (_entries.ContainsKey(path)) throw new NotSupportedException($"Cannot overwrite an unmodeled resource entry '{path}'.");
            newDocuments[path] = ns;
        }
        var plan = new ResourcePlan { Path = path, Kind = kind, Container = container, Id = id, Namespace = ns };
        if (_resources.TryGetValue(key, out var existing))
        {
            plan.Existing = existing.Element;
            if (!string.Equals(existing.Path, path, StringComparison.OrdinalIgnoreCase))
            {
                if (HasOpaqueResource(existing.Element)) throw new NotSupportedException("Cannot relocate an opaque resource with unproven relative references.");
                plan.OldPath = existing.Path;
                foreach (var payload in existing.Element.Elements().Where(node => node.Name == node.Name.Namespace + "FontFile" || node.Name == node.Name.Namespace + "MediaFile"))
                    plan.Relocations.Add((payload, AbsolutePayload(existing.Path, payload.Value)));
                if (existing.Element.Attribute("MediaFile") is XAttribute legacy)
                    plan.Relocations.Add((legacy, AbsolutePayload(existing.Path, legacy.Value)));
            }
        }
        plans.Add(plan); return plan;
    }

    private string ResolvePayload(string path, string file)
    {
        var basis = _documents.TryGetValue(path, out var document) ? document.Root!.Attribute("BaseLoc")?.Value : "Res";
        material.ReservePathComposition(basis ?? string.Empty, file, operationToken);
        var reference = string.IsNullOrEmpty(basis) || file.StartsWith("/", StringComparison.Ordinal) ? file : basis!.TrimEnd('/') + "/" + file;
        return material.ResolvePath(path, reference, operationToken);
    }

    private string AbsolutePayload(string path, string file)
    {
        var resolved = ResolvePayload(path, file); material.ChargeCopy(2L * (resolved.Length + 1L), operationToken);
        return "/" + resolved;
    }

    private void PlanPayload(ResourcePlan plan, string name, string? defaultFileName, byte[] bytes)
    {
        var child = plan.Existing?.Element(plan.Existing.Name.Namespace + name);
        var attribute = plan.Existing?.Attribute(name);
        if (child is not null && attribute is not null) throw new NotSupportedException("Conflicting payload attribute and child authorities.");
        var file = child?.Value ?? attribute?.Value ?? defaultFileName
            ?? throw new InvalidOperationException("A new resource payload requires a fallback name.");
        if (plan.OldPath is not null && (child is not null || attribute is not null)) file = AbsolutePayload(plan.OldPath, file);
        var path = ResolvePayload(plan.Path, file);
        if (protectedPaths.Contains(path)) throw new InvalidDataException("A resource payload collides with a control/page/attachment output role.");
        if (payloads.TryGetValue(path, out var previous))
        {
            if (!ReferenceEquals(previous, bytes))
            {
                if (previous.Length != bytes.Length) throw new InvalidDataException("Different resource payloads share one output path.");
                for (var offset = 0; offset < bytes.Length;)
                {
                    var count = Math.Min(64 * 1024, bytes.Length - offset); material.ChargeWork(2L * count, operationToken);
                    for (var index = offset; index < offset + count; index++)
                        if (previous[index] != bytes[index]) throw new InvalidDataException("Different resource payloads share one output path.");
                    offset += count;
                }
            }
        }
        else { material.ChargeCopy(128, operationToken); payloads.Add(path, bytes); }
        plan.PayloadName = name; plan.PayloadValue = file; plan.PayloadPath = path; plan.Data = bytes;
    }

    internal void Apply()
    {
        foreach (var item in newDocuments)
        {
            operationToken.ThrowIfCancellationRequested();
            _documents.Add(item.Key, new XDocument(new XDeclaration("1.0", "UTF-8", null), new XElement(item.Value + "Res",
                new XAttribute(XNamespace.Xmlns + "ofd", item.Value.NamespaceName), new XAttribute("BaseLoc", "Res"))));
            _changed.Add(item.Key);
        }
        foreach (var plan in plans)
        {
            operationToken.ThrowIfCancellationRequested(); var root = _documents[plan.Path].Root!;
            var element = plan.Existing ?? new XElement(root.Name.Namespace + plan.Kind, new XAttribute("ID", plan.Id));
            if (plan.Existing is null || plan.OldPath is not null)
            {
                foreach (var relocation in plan.Relocations)
                { if (relocation.Node is XElement child) child.Value = relocation.Value; else ((XAttribute)relocation.Node).Value = relocation.Value; }
                if (element.Parent is not null) element.Remove(); var parent = root.Element(root.Name.Namespace + plan.Container);
                if (parent is null) { parent = new XElement(root.Name.Namespace + plan.Container); root.Add(parent); }
                parent.Add(element);
                if (plan.OldPath is not null) _changed.Add(plan.OldPath);
            }
            foreach (var attribute in plan.Attributes) element.SetAttributeValue(attribute.Name, attribute.Value);
            if (plan.PayloadName is not null)
            {
                var child = element.Element(element.Name.Namespace + plan.PayloadName);
                if (child is null) element.Add(new XElement(element.Name.Namespace + plan.PayloadName, plan.PayloadValue));
                else child.Value = plan.PayloadValue!;
                element.SetAttributeValue(plan.PayloadName, null); _entries[plan.PayloadPath!] = plan.Data!;
            }
            _changed.Add(plan.Path);
        }
    }

    private static bool HasOpaqueResource(XElement element)
    {
        var ns = element.Name.Namespace;
        var attributes = element.Name.LocalName == "Font" ? new[] { "ID", "FontName", "FamilyName", "Charset", "Bold", "Italic" } : new[] { "ID", "Type", "Format", "MediaFile" };
        if (element.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && (attribute.Name.Namespace != XNamespace.None || !attributes.Contains(attribute.Name.LocalName)))) return true;
        foreach (var node in element.Nodes())
        {
            if (node is XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) continue;
            if (node is not XElement child || child.Name.Namespace != ns || child.Name.LocalName is not ("FontFile" or "MediaFile") ||
                child.HasElements || child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) || child.Nodes().Any(n => n is not XText and not XComment)) return true;
        }
        return false;
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
        if (!_hashes.TryGetValue(data, out var hash))
        { material.ChargeCopy(512, operationToken); _hashes[data] = hash = BitConverter.ToString(material.HashBytes(data, operationToken)).Replace("-", string.Empty).ToLowerInvariant(); }
        return hash;
    }
}
