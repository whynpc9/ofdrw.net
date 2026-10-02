using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A source document and zero-based page position for mixing.</summary>
public sealed class OfdMixSource
{
    /// <summary>Creates a page selection; the source remains unchanged.</summary>
    public OfdMixSource(OfdDocumentPackage package, int pageIndex) { Package = package ?? throw new ArgumentNullException(nameof(package)); PageIndex = pageIndex; }
    /// <summary>Source document.</summary>
    public OfdDocumentPackage Package { get; }
    /// <summary>Zero-based page position.</summary>
    public int PageIndex { get; }
}

/// <summary>Overlays pages in input order without scaling, using the first source PhysicalBox.</summary>
public static class OfdDocumentMixer
{
    /// <summary>Flattens each source's background templates, body, foreground templates and visible annotation appearances, then overlays the next source.
    /// Unsupported raw/referencing content fails. Source packages remain unchanged; existing signatures are not copied.</summary>
    public static OfdDocumentPackage Mix(IEnumerable<OfdMixSource> sources, int maxSourceCount = 1_000,
        long maxExpandedBytes = 512L * 1024 * 1024, int maxObjectCount = 100_000, CancellationToken cancellationToken = default)
    {
        if (sources is null) throw new ArgumentNullException(nameof(sources));
        if (maxSourceCount <= 0 || maxExpandedBytes <= 0 || maxObjectCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxSourceCount));
        var selected = new List<OfdDocumentPackage>();
        var customTags = new Dictionary<string, string>(StringComparer.Ordinal);
        long bytes = 0;
        int objects = 0;
        foreach (var item in sources)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item is null || selected.Count >= maxSourceCount) throw new ArgumentException("Mix source limit exceeded or null source.");
            OfdDocumentSplitter.ValidateSingleDocument(item.Package);
            OfdDocumentSplitter.ValidatePages(item.Package, new[] { item.PageIndex });
            OfdPageXmlContract.ValidateDocumentArea(item.Package, cancellationToken);
            ValidateCustomTags(item.Package, cancellationToken);
            foreach (var tag in item.Package.CustomTags)
            {
                if (customTags.TryGetValue(tag.Key, out var value) && !string.Equals(value, tag.Value, StringComparison.Ordinal))
                    throw new NotSupportedException($"Mix CustomTags conflict for key '{tag.Key}' in source {selected.Count + 1}.");
                customTags[tag.Key] = tag.Value;
            }
            var rootNamespace = EntryNamespace(item.Package, "OFD.xml");
            if (item.Package.PreservedDocBodyElements.Any(xml => XElement.Parse(xml).Name != rootNamespace + "Signatures"))
                throw new NotSupportedException("Mix cannot safely remap DocBody extensions.");
            var documentNamespace = EntryNamespace(item.Package, item.Package.DocumentEntryPath);
            if (item.Package.PreservedDocumentElements.Any(xml => !SupportedDeclaration(XElement.Parse(xml), documentNamespace, false)) ||
                item.Package.PreservedCommonDataElements.Any(xml => !SupportedDeclaration(XElement.Parse(xml), documentNamespace, true)))
                throw new NotSupportedException("Mix cannot safely remap document extensions or shared drawing resources.");
            var page = item.Package.Pages[item.PageIndex];
            var pageNamespace = EntryNamespace(item.Package, page.SourceEntryPath);
            ValidatePageContainers(item.Package, page.SourceEntryPath, cancellationToken, true);
            foreach (var template in page.Templates)
            {
                if (template.ZOrder is not ("Background" or "Foreground"))
                    throw new NotSupportedException("Mix supports only Background/Foreground template ordering.");
                if (template.BaseLocation is not null)
                    ValidatePageContainers(item.Package, OfdPackagePath.Resolve(item.Package.DocumentEntryPath ??
                        item.Package.Options.DocumentId + "/Document.xml", template.BaseLocation), cancellationToken);
            }
            foreach (var reference in page.PreservedPageElements.Select(XElement.Parse).Where(node => node.Name == pageNamespace + "Template"))
            {
                if (!SupportedTemplateOrder(reference) || reference.HasElements || reference.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)) ||
                    reference.Nodes().Any(node => node is not XText and not XComment) ||
                    reference.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                        (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("TemplateID" or "ZOrder"))))
                    throw new NotSupportedException("Mix cannot flatten an unmodeled template reference wrapper.");
                if (!page.Templates.Any(template => template.TemplateId == reference.Attribute("TemplateID")?.Value))
                    throw new NotSupportedException("Mix cannot flatten an unresolved template reference.");
            }
            objects = checked(objects + page.Elements.Count + page.Templates.Sum(template => template.Elements.Count) + page.AnnotationAppearances.Count);
            bytes = checked(bytes + item.Package.PreservedEntries.Values.Sum(data => (long)data.Length) + item.Package.Fonts.Sum(font => (long)font.Data.Length) + item.Package.Attachments.Sum(attachment => (long)attachment.Data.Length)
                + item.Package.CustomTags.Sum(tag => 2L * (tag.Key.Length + tag.Value.Length))
                + page.Elements.Concat(page.Templates.SelectMany(template => template.Elements)).Concat(page.AnnotationAppearances)
                    .OfType<OfdImageElement>().Sum(image => (long)image.Data.Length));
            if (bytes > maxExpandedBytes || objects > maxObjectCount) throw new ArgumentException("Mix expanded-byte/object budget exceeded.");
            if (page.PreservedPageElements.Any(xml => XElement.Parse(xml).Name != pageNamespace + "Template"))
                throw new NotSupportedException("Mix cannot safely remap page extensions/actions.");
            var package = new OfdDocumentPackage { Options = item.Package.Options };
            package.Pages.Add(page); package.Fonts.AddRange(item.Package.Fonts);
            package.Attachments.AddRange(item.Package.Attachments);
            selected.Add(package);
        }
        if (selected.Count == 0) throw new ArgumentException("At least one source page is required.", nameof(sources));
        var merged = OfdDocumentMerger.Merge(selected, new OfdDocumentMergeOptions { RequireKnownAttributes = true }, cancellationToken);
        foreach (var tag in customTags) merged.CustomTags.Add(tag.Key, tag.Value);
        var first = merged.Pages[0];
        var target = new OfdPage { WidthMillimeters = first.WidthMillimeters, HeightMillimeters = first.HeightMillimeters,
            XMillimeters = first.XMillimeters, YMillimeters = first.YMillimeters };
        foreach (var page in merged.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var element in page.Elements)
            {
                // Per-source layer identity preserves overlay order despite writer grouping.
                element.LayerId = $"mix-{page.Index}-{element.LayerId}";
                element.LayerType = "Body";
                target.Elements.Add(element);
            }
        }
        merged.Pages.Clear(); merged.Pages.Add(target);
        return merged;
    }

    private static bool SupportedDeclaration(XElement node, XNamespace ns, bool template)
    {
        if (node.Name != ns + (template ? "TemplatePage" : "Annotations") || node.HasElements ||
            node.Nodes().Any(child => child is not XText and not XComment)) return false;
        if (!template) return node.Attributes().All(attribute => attribute.IsNamespaceDeclaration) && !string.IsNullOrWhiteSpace(node.Value);
        return !node.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)) &&
            !string.IsNullOrWhiteSpace(node.Attribute("ID")?.Value) && !string.IsNullOrWhiteSpace(node.Attribute("BaseLoc")?.Value) &&
            SupportedTemplateOrder(node) &&
            node.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.None &&
                (attribute.Name.LocalName is "ID" or "BaseLoc" or "Name" or "ZOrder"));
    }

    private static bool SupportedTemplateOrder(XElement node) => node.Attribute("ZOrder") is null ||
        node.Attribute("ZOrder")!.Value is "Background" or "Foreground";

    private static void ValidatePageContainers(OfdDocumentPackage package, string? path, CancellationToken cancellationToken, bool allowTemplateReferences = false)
    {
        if (path is null) return; // Newly constructed model pages have no source XML.
        cancellationToken.ThrowIfCancellationRequested();
        if (!package.PreservedEntries.TryGetValue(path, out var bytes))
            throw new NotSupportedException("Mix cannot inspect missing page/template XML.");
        using var input = new MemoryStream(bytes, false);
        var xml = XDocument.Load(input); var root = xml.Root;
        bool Plain(XElement node, params string[] attributes) => node.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attributes.Contains(attribute.Name.LocalName)) &&
            node.Nodes().All(child => child is XElement or XComment || child is XText text && string.IsNullOrWhiteSpace(text.Value));
        bool supported = root is not null && root.Name.LocalName == "Page" &&
            (root.Name.NamespaceName == package.Options.Namespace || root.Name.NamespaceName == OfdConstants.Namespace || root.Name.NamespaceName == OfdConstants.StandardNamespace) &&
            Plain(root) && xml.Nodes().All(node => node is XElement or XComment);
        if (supported)
        {
            var ns = root!.Name.Namespace;
            supported = root.Elements(ns + "Area").Count() <= 1 && root.Elements(ns + "Content").Count() <= 1;
            foreach (var child in root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (allowTemplateReferences && child.Name == ns + "Template") continue; // Validated by the reference preflight.
                if (child.Name == ns + "Area")
                {
                    supported &= OfdPageXmlContract.IsRepresentedArea(child, ns);
                    continue;
                }
                if (child.Name == ns + "Content")
                {
                    supported &= Plain(child);
                    foreach (var layer in child.Elements())
                    {
                        supported &= layer.Name == ns + "Layer" && Plain(layer, "ID", "Type") &&
                            (layer.Attribute("Type") is null || layer.Attribute("Type")!.Value is "Background" or "Body" or "Foreground");
                        foreach (var drawing in layer.Elements()) supported &= drawing.Name.Namespace == ns &&
                            drawing.Name.LocalName is "TextObject" or "PathObject" or "ImageObject";
                    }
                    continue;
                }
                supported = false;
            }
        }
        if (!supported) throw new NotSupportedException("Mix cannot flatten unmodeled page/template containers, area boxes or drawing resources.");
    }

    private static void ValidateCustomTags(OfdDocumentPackage package, CancellationToken cancellationToken)
    {
        var documentPath = package.DocumentEntryPath ?? package.Options.DocumentId + "/Document.xml";
        var detailPaths = new List<string>();
        var inspected = new Dictionary<string, string>(StringComparer.Ordinal);
        XDocument Read(string path)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!package.PreservedEntries.TryGetValue(path, out var data)) throw new NotSupportedException("Mix cannot preserve missing CustomTags XML.");
            using var input = new MemoryStream(data, false); var xml = XDocument.Load(input);
            if (xml.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                throw new NotSupportedException("Mix cannot preserve unmodeled CustomTags XML nodes.");
            return xml;
        }
        bool Plain(XElement node, params string[] attributes) => !node.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            (attribute.Name.Namespace != XNamespace.None || !attributes.Contains(attribute.Name.LocalName))) &&
            !node.Nodes().Any(child => child is not XElement and not XComment && (child is not XText text || !string.IsNullOrWhiteSpace(text.Value)));
        bool Literal(XElement node) => !node.HasElements && node.Nodes().All(child => child is XText or XComment) &&
            node.Attributes().All(attribute => attribute.IsNamespaceDeclaration);
        bool Supported(XElement node, string name, XNamespace ns) => node.Name.LocalName == name &&
            (node.Name.Namespace == ns || node.Name.NamespaceName == OfdConstants.Namespace || node.Name.NamespaceName == OfdConstants.StandardNamespace);
        if (package.PreservedEntries.ContainsKey(documentPath))
        {
            var document = Read(documentPath); var ns = document.Root!.Name.Namespace;
            var declarations = document.Root.Elements().Where(node => node.Name.LocalName == "CustomTags").ToList();
            if (declarations.Count > 1) throw new NotSupportedException("Mix cannot preserve multiple CustomTags declarations.");
            foreach (var declaration in declarations)
            {
                if (declaration.Name != ns + "CustomTags" || !Literal(declaration))
                    throw new NotSupportedException("Mix cannot preserve unmodeled CustomTags declarations.");
                var listPath = OfdPackagePath.Resolve(documentPath, declaration.Value); var list = Read(listPath);
                if (list.Root is null || !Supported(list.Root, "CustomTags", ns) || !Plain(list.Root)) throw new NotSupportedException("Mix supports only flat EMR CustomTags.");
                foreach (var record in list.Root.Elements())
                {
                    if (record.Name != list.Root.Name.Namespace + "CustomTag" || !Plain(record, "TypeID", "NameSpace") ||
                        record.Attribute("TypeID")?.Value != "EMR" || record.Attribute("NameSpace")?.Value != "urn:ofdrw-net:custom-tags:emr" || record.Elements().Count() != 1)
                        throw new NotSupportedException("Mix supports only flat EMR CustomTags.");
                    var location = record.Elements().Single();
                    if (location.Name != record.Name.Namespace + "FileLoc" || !Literal(location))
                        throw new NotSupportedException("Mix cannot preserve unmodeled CustomTags file references.");
                    detailPaths.Add(OfdPackagePath.Resolve(listPath, location.Value));
                }
            }
            if (declarations.Count == 0)
            {
                var legacy = package.Options.DocumentId + "/Tags/CustomTag_EMR.xml";
                if (package.PreservedEntries.ContainsKey(legacy)) detailPaths.Add(legacy);
            }
            foreach (var path in detailPaths)
            {
                var detail = Read(path);
                if (detail.Root is null || !Supported(detail.Root, "EMRTags", ns) || !Plain(detail.Root)) throw new NotSupportedException("Mix supports only flat EMR CustomTags.");
                foreach (var tag in detail.Root.Elements())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var key = tag.Attribute("Key")?.Value; var value = tag.Attribute("Value")?.Value;
                    if (tag.Name != detail.Root.Name.Namespace + "Tag" || tag.HasElements || !Plain(tag, "Key", "Value") || string.IsNullOrWhiteSpace(key) || value is null)
                        throw new NotSupportedException("Mix cannot preserve unmodeled CustomTags records.");
                    if (inspected.TryGetValue(key!, out var prior) && prior != value) throw new NotSupportedException($"Mix CustomTags conflict for key '{key}' in source XML.");
                    inspected[key!] = value;
                }
            }
        }
        foreach (var tag in package.CustomTags)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(tag.Key) || tag.Value is null) throw new NotSupportedException("Mix requires scalar CustomTags keys and values.");
            if (package.Attachments.Any(attachment => !string.IsNullOrEmpty(attachment.Id) && attachment.Id == tag.Value))
                throw new NotSupportedException($"Mix cannot remap the attachment reference in CustomTags key '{tag.Key}'.");
            foreach (var basis in detailPaths.Concat(new[] { documentPath, "OFD.xml" }))
            {
                string path; try { path = OfdPackagePath.Resolve(basis, tag.Value); } catch (InvalidDataException) { continue; }
                if (package.PreservedEntries.ContainsKey(path)) throw new NotSupportedException($"Mix cannot remap the package reference in CustomTags key '{tag.Key}'.");
            }
        }
    }
    private static XNamespace EntryNamespace(OfdDocumentPackage package, string? path)
    {
        if (path is not null && package.PreservedEntries.TryGetValue(path, out var bytes))
        {
            using var input = new MemoryStream(bytes, false);
            return XDocument.Load(input).Root!.Name.Namespace;
        }
        return XNamespace.Get(package.Options.Namespace);
    }

}
