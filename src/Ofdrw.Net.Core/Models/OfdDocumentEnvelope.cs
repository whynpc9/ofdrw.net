using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Models;

// Compact original slot identities; source XML stays in PreservedEntries.
internal sealed class OfdDocumentEnvelope
{
    internal string DocumentPath { get; private set; } = string.Empty;
    internal readonly OfdNavigationSource Body = new();
    internal readonly OfdNavigationSource Common = new();
    internal readonly OfdNavigationSource Document = new();
    internal readonly OfdNavigationSource Pages = new();
    internal readonly Dictionary<string, string?> Metadata = new(StringComparer.Ordinal);

    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("envelopePath", DocumentPath); Source(Body); Source(Common); Source(Document); Source(Pages);
        sink.Item("metadataBaseline", Metadata, true); sink.Item("count", Metadata.Count);
        foreach (var item in Metadata) { sink.Item("key", item.Key); sink.Item("value", item.Value); }
        void Source(OfdNavigationSource source)
        {
            sink.Item("envelopeSource", source, true); sink.Item("xml", source.Xml); sink.Item("owner", source.DocumentOwner, true);
            sink.Item("opaque", source.HasOpaqueLocalContent); sink.Item("name", source.OriginalName);
            sink.Item("slots", source.Slots, true); sink.Item("count", source.Slots.Count);
            foreach (var slot in source.Slots) { sink.Item("ordinal", slot.Key); sink.Item("slot", slot.Value, slot.Value is not string); }
        }
    }

    internal static OfdDocumentEnvelope Capture(OfdDocumentPackage package, XElement? body, XElement document,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(4096L + package.Pages.Count * 256L, token);
        var result = new OfdDocumentEnvelope { DocumentPath = package.DocumentEntryPath! };
        var ns = document.Name.Namespace;
        CaptureSlots(body, result.Body, package.PreservedDocBodyElements, new[] { "DocInfo", "DocRoot" });
        CaptureSlots(document.Element(ns + "CommonData"), result.Common, package.PreservedCommonDataElements,
            new[] { "MaxUnitID", "PageArea", "PublicRes", "DocumentRes" });
        CaptureSlots(document, result.Document, package.PreservedDocumentElements,
            new[] { "CommonData", "Pages", "Attachments", "CustomTags", "Outlines", "Actions", "Bookmarks" });
        var pageIndices = package.Pages.ToDictionary(page => page.Index);
        var pageOrdinal = 0;
        var ordinal = 0;
        foreach (var page in document.Element(ns + "Pages")?.Elements() ?? Enumerable.Empty<XElement>())
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (page.Name == ns + "Page")
            {
                if (page.Attribute("BaseLoc") is not null && pageIndices.TryGetValue(pageOrdinal, out var model)) result.Pages.Slots.Add(ordinal, model);
                pageOrdinal++;
            }
            ordinal++;
        }
        var metadata = package.Options.Metadata;
        result.Metadata.Add("Title", metadata.Title); result.Metadata.Add("Author", metadata.Author);
        result.Metadata.Add("Subject", metadata.Subject); result.Metadata.Add("Keywords", metadata.Keywords); result.Metadata.Add("Creator", metadata.Creator);
        result.Metadata.Add("CreationDate", metadata.CreationDate?.ToString("O")); result.Metadata.Add("ModDate", metadata.ModificationDate?.ToString("O"));
        result.Metadata.Add("DocUsage", metadata.DocUsage?.ToString()); result.Metadata.Add("Cover", metadata.Cover);
        result.Metadata.Add("CreatorVersion", metadata.CreatorVersion);
        return result;

        void CaptureSlots(XElement? root, OfdNavigationSource slots, List<string> values, string[] owned)
        {
            if (root is null) return;
            var next = 0; var position = 0;
            foreach (var child in root.Elements())
            {
                token.ThrowIfCancellationRequested(); material.Structure.Reference();
                if (child.Name.Namespace != root.Name.Namespace || !owned.Contains(child.Name.LocalName))
                {
                    material.ReserveReadModel(160, token);
                    if (next >= values.Count) throw new InvalidOperationException("Preserved envelope slots do not match reader projections.");
                    slots.Slots.Add(position, values[next++]);
                }
                position++;
            }
            if (next != values.Count) throw new InvalidOperationException("Preserved envelope slot count is inconsistent.");
        }
    }

    internal static void ValidateOwnedFragments(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        var ns = XNamespace.Get(package.Options.Namespace);
        Check(package.PreservedDocBodyElements, new[] { "DocInfo", "DocRoot" });
        Check(package.PreservedCommonDataElements, new[] { "MaxUnitID", "PageArea", "PublicRes", "DocumentRes" });
        var documentOwned = new List<string> { "CommonData", "Pages", "Attachments", "CustomTags" };
        if (package.Outlines is not null || package.OutlineOwnership.OriginalXml is not null) documentOwned.Add("Outlines");
        if (package.Bookmarks is not null || package.BookmarkOwnership.OriginalXml is not null) documentOwned.Add("Bookmarks");
        if (package.Actions is not null || package.ActionOwnership.OriginalXml is not null) documentOwned.Add("Actions");
        Check(package.PreservedDocumentElements, documentOwned);
        foreach (var page in package.Pages) Check(page.PreservedPageElements, new[] { "Area", "Content" });
        void Check(IEnumerable<string> fragments, IEnumerable<string> owned)
        {
            var names = new HashSet<XName>(owned.Select(name => ns + name));
            var navigation = new HashSet<XName>();
            foreach (var xml in fragments)
            {
                material.ChargeCopy(2L * xml.Length, token); var name = OfdXmlEnvelope.Parse(xml).Name;
                if (names.Contains(name)) throw new System.IO.InvalidDataException("Preserved XML conflicts with an owned document node.");
                if (name.Namespace == ns && name.LocalName is "Outlines" or "Actions" or "Bookmarks" && !navigation.Add(name))
                    throw new System.IO.InvalidDataException("Duplicate preserved navigation singleton.");
            }
        }
    }

    internal void ValidatePageOrder(OfdDocumentPackage package, IReadOnlyList<OfdPage> pages,
        OfdModelMaterialization material, CancellationToken token)
    {
        if (!package.PreservedEntries.TryGetValue(DocumentPath, out var bytes))
            throw new System.IO.InvalidDataException("Imported document envelope is missing.");
        var root = material.ReadXml(DocumentPath, bytes, token).Root!;
        var pageRoot = root.Element(root.Name.Namespace + "Pages");
        if (pageRoot is null) throw new System.IO.InvalidDataException("Imported page declarations are missing.");
        OfdXmlEnvelope.ValidateChildrenOrder(pageRoot, Pages, pages, token);
    }

    internal OfdDocumentEnvelope Clone(IReadOnlyDictionary<object, object> mapping, OfdModelMaterialization material, CancellationToken token, bool requireClosedMap = false)
    {
        var copy = new OfdDocumentEnvelope { DocumentPath = DocumentPath };
        Copy(Body, copy.Body); Copy(Common, copy.Common); Copy(Document, copy.Document); Copy(Pages, copy.Pages);
        foreach (var item in Metadata) { material.Structure.Reference(); copy.Metadata.Add(item.Key, item.Value); }
        return copy;
        void Copy(OfdNavigationSource source, OfdNavigationSource destination)
        {
            destination.Xml = source.Xml; destination.OriginalName = source.OriginalName; destination.HasOpaqueLocalContent = source.HasOpaqueLocalContent;
            if (source.DocumentOwner is not null && !mapping.TryGetValue(source.DocumentOwner, out _))
            { if (requireClosedMap) throw new InvalidOperationException("Envelope document owner mapping is incomplete."); destination.DocumentOwner = source.DocumentOwner; }
            else if (source.DocumentOwner is not null) destination.DocumentOwner = mapping[source.DocumentOwner];
            foreach (var slot in source.Slots)
            {
                token.ThrowIfCancellationRequested(); material.Structure.Reference();
                if (!mapping.TryGetValue(slot.Value, out var mapped))
                {
                    if (requireClosedMap && slot.Value is not string) throw new InvalidOperationException("Envelope owner mapping is incomplete.");
                    mapped = slot.Value;
                }
                material.ChargeCopy(64, token); destination.Slots.Add(slot.Key, mapped);
            }
        }
    }
}
