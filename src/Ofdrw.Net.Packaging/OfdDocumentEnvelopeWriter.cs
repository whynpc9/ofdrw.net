using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static class OfdDocumentEnvelopeWriter
{
    internal static XDocument Ofd(OfdDocumentPackage package, XDocument fresh, OfdModelMaterialization material, CancellationToken token)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var bytes)) return fresh;
        var original = material.ReadXml("OFD.xml", bytes, token); var root = original.Root!; var ns = fresh.Root!.Name.Namespace;
        if (root.Name != fresh.Root.Name) throw new NotSupportedException("Cannot migrate OFD source namespace.");
        var documentPath = package.DocumentEntryPath ?? package.Options.DocumentId + "/Document.xml";
        if (package.SourceEnvelope is { } source && source.DocumentPath != documentPath) throw new NotSupportedException("Cannot remap imported document source path.");
        var body = root.Elements(ns + "DocBody").FirstOrDefault(node => node.Element(ns + "DocRoot") is XElement path && material.ResolvePath("OFD.xml", path.Value, token) == documentPath);
        if (body is null) throw new InvalidDataException("Selected source DocBody is missing.");
        var created = fresh.Root.Element(ns + "DocBody")!;
        Preserve(body, package.SourceEnvelope?.Body, package.PreservedDocBodyElements, token);
        Scalar(body, "DocRoot", documentPath);
        var info = Single(body, "DocInfo");
        if (info is null) { info = new XElement(ns + "DocInfo"); body.AddFirst(info); }
        var generatedInfo = created.Element(ns + "DocInfo")!;
        foreach (var name in new[] { "Title", "Author", "Subject", "Keywords", "Creator", "CreationDate", "ModDate" })
        {
            var value = generatedInfo.Element(ns + name)?.Value;
            if (package.SourceEnvelope?.Metadata.TryGetValue(name, out var baseline) == true && baseline == value) continue;
            if (package.SourceEnvelope is null && value is null) continue;
            Scalar(info, name, value);
        }
        root.SetAttributeValue("DocType", package.Options.DocType);
        return original;
    }

    internal static XDocument Document(OfdDocumentPackage package, XDocument fresh, IReadOnlyList<OfdPage> pages,
        IReadOnlyDictionary<OfdPage, string> ids, OfdModelMaterialization material, CancellationToken token)
    {
        var path = package.DocumentEntryPath;
        if (path is null || !package.PreservedEntries.TryGetValue(path, out var bytes)) return fresh;
        var original = material.ReadXml(path, bytes, token); var root = original.Root!; var ns = fresh.Root!.Name.Namespace;
        if (root.Name != fresh.Root.Name) throw new NotSupportedException("Cannot migrate Document source namespace.");
        Preserve(root, package.SourceEnvelope?.Document, package.PreservedDocumentElements, token);
        var common = Single(root, "CommonData");
        if (common is null) { common = new XElement(ns + "CommonData"); root.AddFirst(common); }
        Preserve(common, package.SourceEnvelope?.Common, package.PreservedCommonDataElements, token);
        var generated = fresh.Root.Element(ns + "CommonData")!;
        foreach (var name in new[] { "MaxUnitID", "PublicRes", "DocumentRes" }) Scalar(common, name, generated.Element(ns + name)?.Value);
        var area = Single(common, "PageArea");
        if (area is null) { area = new XElement(ns + "PageArea"); common.Add(area); }
        Scalar(area, "PhysicalBox", generated.Element(ns + "PageArea")!.Element(ns + "PhysicalBox")!.Value);
        var pageRoot = Single(root, "Pages");
        if (pageRoot is null) { pageRoot = new XElement(ns + "Pages"); common.AddAfterSelf(pageRoot); }
        var originalPages = pageRoot.Elements(ns + "Page").Where(node => node.Attribute("ID") is not null)
            .ToDictionary(node => node.Attribute("ID")!.Value, StringComparer.Ordinal);
        var generatedPages = fresh.Root.Element(ns + "Pages")!.Elements(ns + "Page").ToDictionary(node => node.Attribute("ID")!.Value, StringComparer.Ordinal);
        OfdXmlEnvelope.Children(pageRoot, package.SourceEnvelope?.Pages, pages, page =>
        {
            token.ThrowIfCancellationRequested(); var id = ids[page];
            var declaration = originalPages.TryGetValue(id, out var existing) ? new XElement(existing) : new XElement(ns + "Page");
            declaration.SetAttributeValue("ID", id); declaration.SetAttributeValue("BaseLoc", generatedPages[id].Attribute("BaseLoc")!.Value); return declaration;
        }, token);
        foreach (var name in new[] { "Outlines", "Actions", "Bookmarks", "Attachments" })
            OfdXmlEnvelope.Singleton(root, name, fresh.Root.Element(ns + name));
        var originalTags = root.Elements(ns + "CustomTags").ToArray();
        var generatedTags = fresh.Root.Elements(ns + "CustomTags").ToArray();
        if (package.CustomTagSource is null) OfdXmlEnvelope.Singleton(root, "CustomTags", generatedTags.SingleOrDefault());
        else
        {
            // Imported declarations retain their attributes, child nodes and positions.
            if (originalTags.Length == 0 && generatedTags.Length == 1) OfdXmlEnvelope.Singleton(root, "CustomTags", generatedTags[0]);
            else foreach (var declaration in generatedTags.Skip(originalTags.Length)) root.Add(new XElement(declaration));
        }
        return original;
    }

    private static void Preserve(XElement root, OfdNavigationSource? source, IReadOnlyList<string> values, CancellationToken token)
    {
        if (source is not null) OfdXmlEnvelope.Children(root, source, values, OfdXmlEnvelope.Parse, token);
        else foreach (var value in values) root.Add(OfdXmlEnvelope.Parse(value));
    }
    private static XElement? Single(XElement parent, string local)
    {
        var nodes = parent.Elements(parent.Name.Namespace + local).ToList();
        if (nodes.Count > 1) throw new InvalidDataException("Duplicate owned document singleton.");
        return nodes.SingleOrDefault();
    }
    private static void Scalar(XElement parent, string local, string? value)
    {
        var node = Single(parent, local);
        if (node is null) { if (value is not null) OfdXmlEnvelope.Singleton(parent, local, new XElement(parent.Name.Namespace + local, value)); return; }
        var literal = string.Concat(node.Nodes().OfType<XText>().Select(text => text.Value));
        if (literal == value) return;
        if (node.HasElements || node.Nodes().Any(child => child is XProcessingInstruction)) throw new NotSupportedException("Cannot overwrite an opaque mixed scalar envelope.");
        node.Nodes().OfType<XText>().Remove();
        if (value is not null) node.AddFirst(new XText(value));
        else if (!node.HasAttributes && !node.Nodes().Any()) node.Remove();
    }
}
