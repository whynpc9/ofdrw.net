using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

// These inspections use expanded XML names and package-owned references. No source XML is rewritten.
// An opaque edge prevents Pass/NotApplicable, but a known violation still takes precedence.
internal sealed class ArchivalStructuralCoverage
{
    private readonly OfdPackageArchive archive;
    private readonly XElement root;
    private readonly ArchivalPreflightOptions options;
    private readonly CancellationToken token;
    private readonly Observation[] observations = Enumerable.Range(3, 9).Select(i => new Observation(i)).ToArray();
    private readonly Dictionary<string, XElement?> xml = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<XElement, int> siblingOrdinals = new();
    private readonly Dictionary<string, XElement> bookmarks = new(StringComparer.Ordinal);
    private readonly HashSet<string> ambiguousBookmarkNames = new(StringComparer.Ordinal);
    private readonly HashSet<long> pageIds = new();
    private readonly HashSet<long> ambiguousPageIds = new();
    private readonly HashSet<long> templateIds = new();
    private readonly HashSet<string> inspectedPageEntries = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<XElement> ordinaryPageTrees = new();
    private readonly Dictionary<long, XElement> templatePageTrees = new();
    private readonly HashSet<long> ambiguousTemplateIds = new();
    private readonly Dictionary<XElement, string> pageTreeEntries = new();
    private readonly Dictionary<XElement, bool> graphicPositions = new();
    private long xmlBytes;
    private int xmlNodes;
    private int references;
    private bool pageInventoryUnknown;
    private bool bookmarkInventoryUnknown;
    private string documentEntry = "<document>";
    private XNamespace documentNamespace = XNamespace.None;

    internal ArchivalStructuralCoverage(OfdPackageArchive archive, XElement root, ArchivalPreflightOptions options,
        CancellationToken token)
    { this.archive = archive; this.root = root; this.options = options; this.token = token; }

    internal IReadOnlyList<ArchivalCoverageRuleResult> Check()
    {
        var document = ResolveDocument();
        if (document is null)
        {
            UnknownAll("DocumentScopeUnavailable", "<document>");
            return Finish();
        }
        observations[6].Unknown("SourceActionProvenanceUnavailable", documentEntry);
        var ns = document.Name.Namespace;
        documentNamespace = ns;
        var body = root.Elements(root.Name.Namespace + "DocBody").Single();
        if (root.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("DocType" or "Version"))) ||
            OpaqueNodes(root) || root.Elements().Any(x => x.Name != root.Name.Namespace + "DocBody") ||
            body.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(body) ||
            body.Elements().Any(x => x.Name.Namespace != root.Name.Namespace || x.Name.LocalName is not
                ("DocInfo" or "DocRoot" or "Versions" or "Signatures") ||
                OpaqueDocBodyChild(x, root.Name.Namespace)))
            UnknownAll("OpaqueOfdEnvelope", documentEntry);
        if (document.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(document) ||
            document.Elements().Any(x => x.Name.Namespace != ns || x.Name.LocalName is not
                ("CommonData" or "Pages" or "Outlines" or "Permissions" or "Actions" or "VPreferences" or
                 "Bookmarks" or "Annotations" or "CustomTags" or "Attachments") ||
                OpaqueDocumentAuxiliary(x, ns)))
            UnknownAll("OpaqueDocumentEnvelope", documentEntry);
        if (MisplacedDocumentScope(document, ns))
            UnknownAll("MisplacedDocumentScope", documentEntry);
        if (ns != root.Name.Namespace) UnknownAll("MixedOfdNamespaces", documentEntry);
        if (root.Elements().Any(x => x.Name.LocalName == "DocBody" && x.Name.Namespace != root.Name.Namespace) ||
            root.Elements(root.Name.Namespace + "DocBody").Single().Elements().Any(x =>
                x.Name.LocalName == "DocRoot" && x.Name.Namespace != root.Name.Namespace))
            UnknownAll("ForeignDocumentDeclaration", documentEntry);

        ReadPageInventory(document, ns);
        ReadBookmarkInventory(document, ns);
        ReadOutlines(document, ns);
        ReadAnnotations(document, ns);
        return Finish();
    }

    private XElement? ResolveDocument()
    {
        var bodies = root.Elements(root.Name.Namespace + "DocBody").ToArray();
        if (bodies.Length != 1) return null;
        var roots = bodies[0].Elements(root.Name.Namespace + "DocRoot").ToArray();
        if (roots.Length != 1 || roots[0].HasElements || roots[0].Attributes().Any(x => !x.IsNamespaceDeclaration) ||
            roots[0].Nodes().Any(x => x is not XText) ||
            !TryEntry("OFD.xml", roots[0].Value, out documentEntry)) return null;
        var document = Read(documentEntry, "Document");
        return document;
    }

    private void ReadPageInventory(XElement document, XNamespace ns)
    {
        var common = document.Elements(ns + "CommonData").ToArray();
        foreach (var template in common.SelectMany(x => x.Elements(ns + "TemplatePage")))
        {
            var id = (string?)template.Attribute("ID");
            if (!Id(id, out var number))
            { UnknownPageScope("TemplateIdAmbiguous", documentEntry); continue; }
            if (!templateIds.Add(number) || ForeignAttribute(template, "ID"))
            { ambiguousTemplateIds.Add(number);
                UnknownPageScope("TemplateIdAmbiguous", documentEntry);
            }
        }
        var pages = document.Elements(ns + "Pages").ToArray();
        if (pages.Length != 1 || Foreign(document, "Pages", ns))
        {
            UnknownPageScope("PageTreeAmbiguous", documentEntry);
            pageInventoryUnknown = true;
        }
        foreach (var pageTree in pages)
        {
            if (pageTree.Elements().Any(x => x.Name != ns + "Page") ||
                pageTree.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(pageTree))
            { UnknownPageScope("OpaquePageTree", documentEntry); pageInventoryUnknown = true; }
            if (Foreign(pageTree, "Page", ns))
            { UnknownPageScope("ForeignPageDeclaration", documentEntry); pageInventoryUnknown = true; }
            foreach (var page in pageTree.Elements(ns + "Page"))
            {
                token.ThrowIfCancellationRequested();
                var id = (string?)page.Attribute("ID");
                if (!Id(id, out var number)) pageInventoryUnknown = true;
                else if (!pageIds.Add(number) || ForeignAttribute(page, "ID"))
                    ambiguousPageIds.Add(number);
                ReadPage(page, documentEntry, false);
            }
        }
        if (pageIds.Count == 0)
        { pageInventoryUnknown = true; UnknownPageScope("PageTreeEmpty", documentEntry); }
        if (common.Length > 1 || Foreign(document, "CommonData", ns))
            UnknownPageScope("CommonDataAmbiguous", documentEntry);
        if (common.Any(x => x.Attributes().Any(a => !a.IsNamespaceDeclaration) || OpaqueNodes(x) ||
            x.Elements().Any(y => y.Name.Namespace != ns || y.Name.LocalName is not
                ("MaxUnitID" or "PageArea" or "PublicRes" or "DocumentRes" or "TemplatePage" or "DefaultCS") ||
                OpaqueCommonChild(y, ns))))
            UnknownPageScope("OpaqueCommonData", documentEntry);
        if (common.Any(x => x.Descendants(ns + "TemplatePage").Any(y => y.Parent != x)))
            UnknownPageScope("NestedTemplatePage", documentEntry);
        foreach (var template in common.SelectMany(x => x.Elements(ns + "TemplatePage")))
            ReadPage(template, documentEntry, true);
        if (common.Any(x => Foreign(x, "TemplatePage", ns)))
            UnknownPageScope("ForeignTemplatePage", documentEntry);
        CheckEffectiveLayerNames();
    }

    private void ReadPage(XElement declaration, string containing, bool template)
    {
        if (declaration.HasElements || OpaqueNodes(declaration) ||
            declaration.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("ID" or "BaseLoc" or "Name" or "ZOrder") ||
                 !template && x.Name.LocalName is ("Name" or "ZOrder"))))
        {
            UnknownPageScope("OpaquePageDeclaration", containing);
            if (!template) pageInventoryUnknown = true;
        }
        if (ForeignAttribute(declaration, "BaseLoc") ||
            !TryEntry(containing, (string?)declaration.Attribute("BaseLoc"), out var entry))
        {
            observations[0].Unknown("PageReferenceUnavailable", containing);
            UnknownGraphics("PageReferenceUnavailable", containing);
            observations[7].Unknown("PageReferenceUnavailable", containing);
            if (!template) MarkPageAmbiguous(declaration);
            return;
        }
        var page = Read(entry, "Page");
        if (page is null)
        {
            observations[0].Unknown("PageContentUnavailable", entry);
            UnknownGraphics("PageContentUnavailable", entry);
            observations[7].Unknown("PageContentUnavailable", entry);
            if (!template) MarkPageAmbiguous(declaration);
            return;
        }
        if (!pageTreeEntries.ContainsKey(page)) pageTreeEntries.Add(page, entry);
        if (template)
        {
            if (Id((string?)declaration.Attribute("ID"), out var templateId) &&
                !ambiguousTemplateIds.Contains(templateId))
                if (!templatePageTrees.ContainsKey(templateId)) templatePageTrees.Add(templateId, page);
        }
        else ordinaryPageTrees.Add(page);
        // Declaration and reference checks still run for each use, but a cached page tree is inspected once.
        if (!inspectedPageEntries.Add(entry)) return;
        if (page.Name.Namespace != documentNamespace)
            UnknownPageScope("MixedPageNamespace", entry);
        if (page.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(page) ||
            page.Elements().Any(x => x.Name.Namespace != page.Name.Namespace || x.Name.LocalName is not
                ("Area" or "Template" or "PageRes" or "Content" or "Actions") ||
                OpaquePageChild(x, page.Name.Namespace)))
            UnknownPageScope("OpaquePageRoot", entry);
        if (page.Elements(page.Name.Namespace + "Template").Any(x =>
                !Id((string?)x.Attribute("TemplateID"), out var number) || !templateIds.Contains(number)))
            UnknownPageScope("TemplateReferenceUnavailable", entry);
        InspectActions(page, page.Name.Namespace, observations[0], null, entry);
        if (page.Elements(page.Name.Namespace + "Content").Count() != 1 ||
            Foreign(page, "Content", page.Name.Namespace))
        { UnknownGraphics("PageContentAmbiguous", entry); observations[7].Unknown("PageContentAmbiguous", entry); }
        InspectLayers(page, entry);
        InspectGraphics(page, entry);
    }

    private void MarkPageAmbiguous(XElement declaration)
    {
        if (Id((string?)declaration.Attribute("ID"), out var number)) ambiguousPageIds.Add(number);
        else pageInventoryUnknown = true;
    }

    private void CheckEffectiveLayerNames()
    {
        var work = 0;
        foreach (var page in ordinaryPageTrees)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            var active = new HashSet<long>();
            var pending = new Stack<(XElement Page, long? TemplateId, bool Exit)>();
            pending.Push((page, null, false));
            while (pending.Count != 0)
            {
                token.ThrowIfCancellationRequested();
                if (work >= options.MaxXmlNodes)
                { observations[7].Unknown("LayerCompositionBudgetExceeded", documentEntry); return; }
                work++;
                var (tree, templateId, exit) = pending.Pop();
                if (exit)
                { if (templateId.HasValue) active.Remove(templateId.Value); continue; }
                if (templateId.HasValue && !active.Add(templateId.Value))
                { UnknownPageScope("TemplateCycle", documentEntry); continue; }
                var ns = tree.Name.Namespace;
                var entry = pageTreeEntries.TryGetValue(tree, out var path) ? path : documentEntry;
                foreach (var layer in tree.Elements(ns + "Content").SelectMany(x => x.Elements(ns + "Layer")))
                {
                    if (work >= options.MaxXmlNodes)
                    { observations[7].Unknown("LayerCompositionBudgetExceeded", entry); return; }
                    work++;
                    var name = (string?)layer.Attribute("Name");
                    if (name is not null && !names.Add(name))
                        observations[7].Fail("DuplicateEffectiveLayerName", entry, layer, this);
                }
                pending.Push((tree, templateId, true));
                foreach (var reference in tree.Elements(ns + "Template").Reverse())
                {
                    if (work >= options.MaxXmlNodes)
                    { observations[7].Unknown("LayerCompositionBudgetExceeded", entry); return; }
                    work++;
                    if (!Id((string?)reference.Attribute("TemplateID"), out var number) ||
                        !templatePageTrees.TryGetValue(number, out var nested))
                    { UnknownPageScope("TemplateReferenceUnavailable", entry); continue; }
                    pending.Push((nested, number, false));
                }
            }
        }
    }

    private void ReadBookmarkInventory(XElement document, XNamespace ns)
    {
        var containers = document.Elements(ns + "Bookmarks").ToArray();
        if (containers.Length > 1 || Foreign(document, "Bookmarks", ns)) bookmarkInventoryUnknown = true;
        foreach (var container in containers)
        {
            if (container.Elements().Any(x => x.Name != ns + "Bookmark"))
            { bookmarkInventoryUnknown = true; }
            if (Foreign(container, "Bookmark", ns))
            { bookmarkInventoryUnknown = true; }
            if (container.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(container))
            { bookmarkInventoryUnknown = true; }
            foreach (var bookmark in container.Elements(ns + "Bookmark"))
            {
                if (bookmark.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Name") ||
                    OpaqueNodes(bookmark) || bookmark.Elements().Any(x => x.Name != ns + "Dest"))
                { bookmarkInventoryUnknown = true; }
                if (bookmark.Elements(ns + "Dest").Any(x => x.HasElements || OpaqueNodes(x)))
                { bookmarkInventoryUnknown = true; }
                var name = (string?)bookmark.Attribute("Name");
                if (string.IsNullOrWhiteSpace(name))
                { bookmarkInventoryUnknown = true; continue; }
                if (ForeignAttribute(bookmark, "Name") || ambiguousBookmarkNames.Contains(name) ||
                    bookmarks.ContainsKey(name))
                {
                    ambiguousBookmarkNames.Add(name);
                    bookmarks.Remove(name);
                }
                else bookmarks.Add(name, bookmark);
            }
        }
    }

    private void ReadOutlines(XElement document, XNamespace ns)
    {
        var containers = document.Elements(ns + "Outlines").ToArray();
        if (containers.Length > 1 || Foreign(document, "Outlines", ns))
            UnknownOutline("OutlineScopeAmbiguous", documentEntry);
        foreach (var container in containers)
        {
            if (container.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(container))
                UnknownOutline("OpaqueOutlineContainer", documentEntry);
            if (Foreign(container, "OutlineElem", ns))
                UnknownOutline("ForeignOutline", documentEntry);
            if (container.Elements().Any(x => x.Name != ns + "OutlineElem"))
                UnknownOutline("OpaqueOutlineContent", documentEntry);
            foreach (var node in container.Elements(ns + "OutlineElem")) ReadOutlineTree(node, ns);
        }
    }

    private void ReadOutlineTree(XElement first, XNamespace ns)
    {
        var pending = new Stack<(XElement Node, int Depth)>();
        pending.Push((first, 0));
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var (node, depth) = pending.Pop();
            if (depth > options.MaxXmlDepth)
            { UnknownOutline("OutlineDepthExceeded", documentEntry); continue; }
            var expanded = (string?)node.Attribute("Expanded");
            var count = (string?)node.Attribute("Count");
            if (node.Attribute("Title") is null ||
                node.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name.LocalName is not ("Title" or "Expanded" or "Count")) ||
                node.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name.Namespace != XNamespace.None) ||
                expanded is not (null or "true" or "false" or "1" or "0") ||
                count is not null && (!int.TryParse(count, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value < 0) ||
                OpaqueNodes(node))
                UnknownOutline("OpaqueOutlineNode", documentEntry);
            InspectActions(node, ns, observations[1], observations[2], documentEntry);
            if (Foreign(node, "OutlineElem", ns)) UnknownOutline("ForeignOutline", documentEntry);
            if (node.Elements().Any(x => x.Name != ns + "OutlineElem" && x.Name != ns + "Actions"))
                UnknownOutline("OpaqueOutlineContent", documentEntry);
            foreach (var child in node.Elements(ns + "OutlineElem").Reverse())
                pending.Push((child, depth + 1));
        }
    }

    private void ReadAnnotations(XElement document, XNamespace ns)
    {
        var declarations = document.Elements(ns + "Annotations").ToArray();
        var inspectedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (declarations.Length > 1 || Foreign(document, "Annotations", ns))
            UnknownAnnotationScope("AnnotationScopeAmbiguous", documentEntry);
        foreach (var declaration in declarations)
        {
            if (declaration.HasElements || declaration.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                declaration.Nodes().Any(x => x is not XText) ||
                !TryEntry(documentEntry, declaration.Value, out var indexEntry))
            { UnknownAnnotationScope("AnnotationIndexUnavailable", documentEntry); continue; }
            var index = Read(indexEntry, "Annotations");
            if (index is null)
            { UnknownAnnotationScope("AnnotationIndexUnavailable", indexEntry); continue; }
            if (index.Name.Namespace != documentNamespace)
                UnknownAnnotationScope("MixedAnnotationNamespace", indexEntry);
            if (index.Elements().Any(x => x.Name != index.Name.Namespace + "Page") ||
                index.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(index))
                UnknownAnnotationScope("OpaqueAnnotationIndex", indexEntry);
            if (Foreign(index, "Page", index.Name.Namespace))
            { UnknownAnnotationScope("ForeignAnnotationPage", indexEntry); }
            foreach (var page in index.Elements(index.Name.Namespace + "Page"))
            {
                if (page.Elements().Any(x => x.Name != index.Name.Namespace + "FileLoc") ||
                    page.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "PageID") ||
                    OpaqueNodes(page))
                    UnknownAnnotationScope("OpaqueAnnotationPage", indexEntry);
                if (ForeignAttribute(page, "PageID") ||
                    !Id((string?)page.Attribute("PageID"), out var pageNumber) || !pageIds.Contains(pageNumber))
                    UnknownAnnotationScope("AnnotationPageOwnerAmbiguous", indexEntry);
                var files = page.Elements(index.Name.Namespace + "FileLoc").ToArray();
                if (files.Length != 1 || Foreign(page, "FileLoc", index.Name.Namespace) ||
                    files[0].HasElements || files[0].Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                    files[0].Nodes().Any(x => x is not XText) ||
                    !TryEntry(indexEntry, files[0].Value, out var entry))
                { UnknownAnnotationScope("AnnotationPageUnavailable", indexEntry); continue; }
                // The same PageAnnot can be referenced by many index records; inspect its objects once.
                if (!inspectedEntries.Add(entry)) continue;
                var content = Read(entry, "PageAnnot");
                if (content is null)
                { UnknownAnnotationScope("AnnotationPageUnavailable", entry); continue; }
                if (content.Name.Namespace != documentNamespace)
                    UnknownAnnotationScope("MixedAnnotationNamespace", entry);
                if (content.Elements().Any(x => x.Name != content.Name.Namespace + "Annot") ||
                    content.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(content))
                    UnknownAnnotationScope("OpaqueAnnotationContent", entry);
                if (Foreign(content, "Annot", content.Name.Namespace))
                { UnknownAnnotationScope("ForeignAnnotation", entry); }
                foreach (var annot in content.Elements(content.Name.Namespace + "Annot"))
                {
                    if (OpaqueAnnotationWrapper(annot))
                        UnknownAnnotationScope("OpaqueAnnotationWrapper", entry);
                    observations[5].Seen(entry, annot, this);
                    foreach (var flag in new[] { "ReadOnly", "NoZoom", "NoRotate" })
                    {
                        var value = (string?)annot.Attribute(flag);
                        if (value is null || value is "false" or "0")
                            observations[5].Fail("RetainedAnnotationFlagNotTrue", entry, annot, this);
                        else if (value is not ("true" or "1") || ForeignAttribute(annot, flag))
                            observations[5].Unknown("AnnotationFlagAmbiguous", entry, annot, this);
                    }
                    foreach (var appearance in annot.Elements(content.Name.Namespace + "Appearance"))
                    {
                        InspectAnnotationBlocks(appearance, entry);
                        InspectGraphics(appearance, entry);
                    }
                    if (Foreign(annot, "Appearance", content.Name.Namespace))
                        UnknownAnnotationScope("ForeignAnnotationAppearance", entry);
                }
            }
        }
    }

    private void InspectLayers(XElement page, string entry)
    {
        var ns = page.Name.Namespace;
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var content in page.Elements(ns + "Content"))
        {
            if (content.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(content))
                observations[7].Unknown("OpaqueLayerScope", entry);
            if (content.Elements().Any(x => x.Name != ns + "Layer"))
                observations[7].Unknown("OpaqueLayerScope", entry);
            if (Foreign(content, "Layer", ns)) observations[7].Unknown("ForeignLayer", entry);
            foreach (var layer in content.Elements(ns + "Layer"))
            {
                observations[7].Seen(entry, layer, this);
                if (OpaqueGraphicContainerAttributes(layer))
                    observations[7].Unknown("OpaqueLayerAttributes", entry, layer, this);
                if (OpaqueGraphicContainers(layer, ns))
                    observations[7].Unknown("OpaqueLayerPayload", entry, layer, this);
                var name = (string?)layer.Attribute("Name");
                if (ForeignAttribute(layer, "Name")) observations[7].Unknown("LayerNameAmbiguous", entry, layer, this);
                if (name is not null && !names.Add(name))
                    observations[7].Fail("DuplicateLayerName", entry, layer, this);
            }
        }
    }

    private void InspectAnnotationBlocks(XElement appearance, string entry)
    {
        var ns = appearance.Name.Namespace;
        if (OpaqueAppearance(appearance))
            observations[8].Unknown("OpaqueAnnotationAppearance", entry, appearance, this);
        var pending = new Stack<(XElement Block, int Depth)>();
        foreach (var block in appearance.Elements(ns + "PageBlock").Reverse()) pending.Push((block, 0));
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var (block, depth) = pending.Pop();
            observations[8].Seen(entry, block, this);
            if (depth > 0)
                observations[8].Fail("NestedAnnotationPageBlock", entry, block, this);
            foreach (var child in block.Elements(ns + "PageBlock").Reverse())
                pending.Push((child, depth + 1));
        }
    }

    private void InspectGraphics(XElement source, string entry)
    {
        var ns = source.Name.Namespace;
        if (source.Name == ns + "Appearance" && OpaqueAppearance(source))
            UnknownGraphics("OpaqueAnnotationAppearance", entry);
        if (source.Descendants().Any(x => x.Name.Namespace != ns))
            UnknownGraphics("ForeignGraphicContent", entry);
        if (OpaqueGraphicContainers(source, ns))
            UnknownGraphics("OpaqueGraphicContainer", entry);
        if (source.DescendantsAndSelf().Any(x =>
                (x.Name.LocalName == "ImageObject" || x.Name.LocalName == "TextObject") && x.Name.Namespace != ns) ||
            source.DescendantsAndSelf().Any(x => x.Name == ns + "CompositeObject"))
            UnknownGraphics("OpaqueGraphicScope", entry);
        foreach (var image in source.Descendants(ns + "ImageObject").Where(x =>
                     RecognizedGraphicPosition(x, source, ns)))
        {
            var value = (string?)image.Attribute("Interpolate");
            if (value is null) { if (ForeignAttribute(image, "Interpolate")) observations[3].Unknown("InterpolateAmbiguous", entry, image, this); continue; }
            observations[3].Seen(entry, image, this);
            if (value is "true" or "1") observations[3].Fail("InterpolateEnabled", entry, image, this);
            else if (value is not ("false" or "0") || ForeignAttribute(image, "Interpolate"))
                observations[3].Unknown("InterpolateAmbiguous", entry, image, this);
        }
        foreach (var text in source.Descendants(ns + "TextObject").Where(x =>
                     RecognizedGraphicPosition(x, source, ns)))
        {
            observations[4].Seen(entry, text, this);
            var value = (string?)text.Attribute("Size");
            if (value is null || !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) ||
                double.IsNaN(size) || double.IsInfinity(size) || size <= 0)
                observations[4].Fail("TextSizeUnavailable", entry, text, this);
            else if (ForeignAttribute(text, "Size"))
                observations[4].Unknown("TextSizeAmbiguous", entry, text, this);
        }
        foreach (var graphic in source.Descendants().Where(x => x.Name.Namespace == ns &&
                     (x.Name.LocalName == "TextObject" || x.Name.LocalName == "PathObject" ||
                      x.Name.LocalName == "ImageObject" || x.Name.LocalName == "CompositeObject") &&
                      RecognizedGraphicPosition(x, source, ns)))
        {
            if (graphic.Elements(ns + "Actions").Any(x => x.Elements(ns + "Action").Any()))
                observations[6].Fail("ObjectActionRetained", entry, graphic, this);
            else if (graphic.Elements().Any(x => x.Name.LocalName is ("Actions" or "Action") &&
                         (x.Name != ns + "Actions" || x.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                          OpaqueNodes(x) || x.Elements().Any(a => a.Name != ns + "Action"))))
                observations[6].Unknown("ObjectActionAmbiguous", entry, graphic, this);
        }
    }

    private bool RecognizedGraphicPosition(XElement node, XElement source, XNamespace ns)
    {
        var pending = new Stack<XElement>();
        var current = node;
        while (!ReferenceEquals(current, source) && !graphicPositions.ContainsKey(current))
        {
            pending.Push(current);
            if (current.Parent is null) return false;
            current = current.Parent;
        }
        var recognized = ReferenceEquals(current, source) || graphicPositions[current];
        while (pending.Count != 0)
        {
            var child = pending.Pop();
            var parent = child.Parent!;
            recognized &= child.Name.Namespace == ns && (child.Name.LocalName switch
            {
                "TextObject" or "PathObject" or "ImageObject" or "CompositeObject" =>
                    parent.Name == ns + "Layer" || parent.Name == ns + "PageBlock" || parent.Name == ns + "Appearance",
                "PageBlock" => parent.Name == ns + "Layer" || parent.Name == ns + "PageBlock" ||
                               parent.Name == ns + "Appearance",
                "Layer" => parent.Name == ns + "Content",
                "Content" => parent.Name == ns + "Page",
                _ => false
            });
            graphicPositions[child] = recognized;
        }
        return recognized;
    }

    private static bool OpaqueGraphicContainers(XElement source, XNamespace ns)
    {
        foreach (var container in source.DescendantsAndSelf())
        {
            if (container.Name == ns + "Content")
            {
                if (container.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(container) ||
                    container.Elements().Any(x => x.Name != ns + "Layer")) return true;
            }
            else if (container.Name == ns + "Layer" || container.Name == ns + "PageBlock" ||
                     container.Name == ns + "Appearance")
            {
                if (OpaqueGraphicContainerAttributes(container) || OpaqueNodes(container) ||
                    container.Elements().Any(x => x.Name.Namespace != ns || x.Name.LocalName is not
                    ("PageBlock" or "TextObject" or "PathObject" or "ImageObject" or "CompositeObject"))) return true;
            }
            else if (OpaqueGraphicChild(container, ns)) return true;
        }
        return false;
    }

    private static bool OpaqueGraphicContainerAttributes(XElement node) =>
        node.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
            (attribute.Name.Namespace != XNamespace.None || (node.Name.LocalName switch
            {
                "Layer" => attribute.Name.LocalName is not ("ID" or "Type" or "Name" or "DrawParam") ||
                           attribute.Name.LocalName == "DrawParam" && !Id(attribute.Value, out _),
                "PageBlock" => attribute.Name.LocalName is not ("ID" or "Boundary" or "CTM"),
                _ => false
            })));

    private static bool OpaqueGraphicChild(XElement node, XNamespace ns)
    {
        var objectActions = node.AncestorsAndSelf(ns + "Actions").FirstOrDefault(x =>
            x.Parent?.Name.Namespace == ns && x.Parent.Name.LocalName is
                ("TextObject" or "PathObject" or "ImageObject" or "CompositeObject"));
        if (objectActions is not null)
            return ReferenceEquals(node, objectActions) && OpaqueActionContainer(objectActions, ns);
        var parent = node.Name.LocalName;
        if (node.Name.Namespace != ns || !node.AncestorsAndSelf().Any(x => x.Name.Namespace == ns &&
            x.Name.LocalName is ("TextObject" or "PathObject" or "ImageObject" or "CompositeObject"))) return false;
        if (node.Attributes().Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute) &&
            !(parent == "ImageObject" && attribute.Name == "Interpolate"))) return true;
        if (parent is "TextCode" or "AbbreviatedData" or "Glyphs")
            return node.Nodes().Any(x => x is not XText);
        if (parent is "FillColor" or "StrokeColor")
            return node.HasElements || OpaqueNodes(node);
        if (parent is not ("TextObject" or "PathObject" or "ImageObject" or "CompositeObject" or
            "Clips" or "Clip" or "Area" or "Path" or "CGTransform")) return false;
        if (OpaqueNodes(node)) return true;
        return node.Elements().Any(child => child.Name.Namespace != ns || !(parent switch
        {
            "TextObject" => child.Name.LocalName is "Clips" or "FillColor" or "StrokeColor" or "CGTransform" or "TextCode" or "Actions",
            "PathObject" => child.Name.LocalName is "Clips" or "FillColor" or "StrokeColor" or "AbbreviatedData" or "Actions",
            "ImageObject" => child.Name.LocalName is "Clips" or "Actions",
            "CompositeObject" => child.Name.LocalName is "Actions",
            "Clips" => child.Name.LocalName is "Clip",
            "Clip" => child.Name.LocalName is "Area",
            "Area" => child.Name.LocalName is "Path",
            "Path" => child.Name.LocalName is "AbbreviatedData",
            "CGTransform" => child.Name.LocalName is "Glyphs",
            _ => false
        }));
    }

    private static bool OpaquePageChild(XElement child, XNamespace ns)
    {
        if (child.Name == ns + "Area")
            return child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child) ||
                   child.Elements().Any(box => box.Name.Namespace != ns || box.Name.LocalName is not
                       ("PhysicalBox" or "ApplicationBox" or "ContentBox" or "BleedBox") ||
                       box.HasElements || box.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                       box.Nodes().Any(x => x is not XText));
        if (child.Name == ns + "Template")
            return child.HasElements || OpaqueNodes(child) ||
                   string.IsNullOrWhiteSpace((string?)child.Attribute("TemplateID")) ||
                   (string?)child.Attribute("ZOrder") is not (null or "Background" or "Foreground") ||
                   child.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                       (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("TemplateID" or "ZOrder")));
        if (child.Name == ns + "PageRes")
            return child.HasElements || child.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                   child.Nodes().Any(x => x is not XText);
        return false;
    }

    private static bool OpaqueCommonChild(XElement child, XNamespace ns)
    {
        if (child.Name == ns + "PageArea")
            return child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child) ||
                   child.Elements().Any(box => box.Name.Namespace != ns || box.Name.LocalName is not
                       ("PhysicalBox" or "ApplicationBox" or "ContentBox" or "BleedBox") ||
                       box.HasElements || box.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                       box.Nodes().Any(x => x is not XText));
        if (child.Name == ns + "TemplatePage")
            return child.HasElements || OpaqueNodes(child) ||
                   child.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                       (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("ID" or "BaseLoc" or "Name" or "ZOrder")));
        return child.HasElements || child.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
               child.Nodes().Any(x => x is not XText);
    }

    private static bool OpaqueDocumentAuxiliary(XElement child, XNamespace ns)
    {
        if (child.Name == ns + "CustomTags" || child.Name == ns + "Attachments")
            return child.HasElements || child.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                   child.Nodes().Any(x => x is not XText);
        if (child.Name == ns + "Permissions" || child.Name == ns + "VPreferences")
        {
            if (child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child)) return true;
            return child.Elements().Any(field =>
            {
                var name = field.Name.LocalName;
                var recognized = child.Name == ns + "Permissions"
                    ? name is "Edit" or "Annot" or "Export" or "Signature" or "Watermark" or "PrintScreen" or "Print" or "ValidPeriod"
                    : name is "PageMode" or "PageLayout" or "TabDisplay" or "HideToolbar" or "HideMenubar" or "HideWindowUI" or "ZoomMode" or "Zoom";
                if (field.Name.Namespace != ns || !recognized || field.HasElements ||
                    field.Nodes().Any(x => x is not XText)) return true;
                return field.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    (attribute.Name.Namespace != XNamespace.None || !(name switch
                    {
                        "Print" => attribute.Name.LocalName is "Printable" or "Copies",
                        "ValidPeriod" => attribute.Name.LocalName is "StartDate" or "EndDate",
                        _ => false
                    })));
            });
        }
        return child.Name == ns + "Actions" && OpaqueActionContainer(child, ns);
    }

    private static bool OpaqueActionContainer(XElement child, XNamespace ns)
    {
        if (child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child)) return true;
        return child.Elements().Any(action => action.Name != ns + "Action" ||
            action.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Event") ||
            OpaqueNodes(action) || action.Elements().Any(kind =>
                kind.Name == ns + "Region" ? !StandardActionRegion(kind, ns) :
                kind.Name == ns + "Goto" ? GotoOpaque(kind, ns) :
                kind.Name.Namespace != ns || kind.Name.LocalName is not ("URI" or "GotoA" or "Sound" or "Movie") ||
                kind.HasElements || OpaqueNodes(kind)));
    }

    private static bool OpaqueDocBodyChild(XElement child, XNamespace ns)
    {
        if (child.Name == ns + "Signatures")
            return child.HasElements || child.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                   child.Nodes().Any(x => x is not XText);
        if (child.Name == ns + "Versions")
            return child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child) ||
                   child.Elements().Any(version => version.Name != ns + "Version" || version.HasElements ||
                       version.Nodes().Any(x => x is not XText text || !string.IsNullOrWhiteSpace(text.Value)) ||
                       version.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                           (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("ID" or "Index" or "Current" or "BaseLoc"))));
        if (child.Name != ns + "DocInfo") return false;
        if (child.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(child)) return true;
        return child.Elements().Any(field =>
        {
            if (field.Name == ns + "CustomDatas")
                return field.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(field) ||
                       field.Elements().Any(data => data.Name != ns + "CustomData" || data.HasElements ||
                           data.Nodes().Any(x => x is not XText) || data.Attributes().Any(x =>
                               !x.IsNamespaceDeclaration && x.Name != "Name"));
            return field.Name.Namespace != ns || field.Name.LocalName is not
                ("DocID" or "Title" or "Author" or "Subject" or "Abstract" or "CreationDate" or
                 "ModDate" or "DocUsage" or "Cover" or "Keywords" or "Creator" or "CreatorVersion") ||
                field.HasElements || field.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                field.Nodes().Any(x => x is not XText);
        });
    }

    private static bool MisplacedDocumentScope(XElement document, XNamespace ns) =>
        document.Descendants().Any(node => node.Name.LocalName switch
        {
            "CommonData" or "Pages" or "Outlines" or "Permissions" or "VPreferences" or
                "Bookmarks" or "Annotations" or "CustomTags" or "Attachments" =>
                node.Parent != document || node.Name.Namespace != ns,
            "TemplatePage" => node.Parent?.Name != ns + "CommonData" || node.Parent.Parent != document,
            "Page" => node.Parent?.Name != ns + "Pages" || node.Parent.Parent != document,
            "OutlineElem" => node.Parent?.Name != ns + "Outlines" && node.Parent?.Name != ns + "OutlineElem",
            "Layer" or "ImageObject" or "TextObject" or "PathObject" or "CompositeObject" => true,
            _ => false
        });

    private static bool OpaqueAppearance(XElement appearance)
    {
        var ns = appearance.Name.Namespace;
        return appearance.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                   (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not ("ID" or "Boundary" or "CTM"))) ||
               OpaqueNodes(appearance) || OpaqueGraphicContainers(appearance, ns) ||
               appearance.Descendants().Any(x => x.Name.Namespace != ns);
    }

    private static bool OpaqueAnnotationWrapper(XElement annot)
    {
        var ns = annot.Name.Namespace;
        if (annot.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not
                    ("ID" or "Type" or "Creator" or "LastModDate" or "Subtype" or "Visible" or "Print" or
                     "NoZoom" or "NoRotate" or "ReadOnly"))) || OpaqueNodes(annot) ||
            annot.Elements(ns + "Appearance").Count() != 1 ||
            annot.Elements(ns + "Remark").Count() > 1 ||
            annot.Elements(ns + "Parameters").Count() > 1) return true;
        foreach (var child in annot.Elements())
        {
            if (child.Name == ns + "Appearance") continue;
            if (child.Name == ns + "Remark" && !child.HasElements &&
                child.Attributes().All(x => x.IsNamespaceDeclaration) &&
                child.Nodes().All(x => x is XText)) continue;
            if (child.Name == ns + "Parameters" &&
                child.Attributes().All(x => x.IsNamespaceDeclaration) && !OpaqueNodes(child) &&
                child.Elements().All(parameter => parameter.Name == ns + "Parameter" && !parameter.HasElements &&
                    parameter.Attributes().All(x => x.IsNamespaceDeclaration || x.Name == "Name") &&
                    parameter.Nodes().All(x => x is XText))) continue;
            return true;
        }
        return false;
    }

    private void InspectActions(XElement owner, XNamespace ns, Observation type, Observation? target, string entry)
    {
        var containers = owner.Elements(ns + "Actions").ToArray();
        if (owner.Elements().Any(x => x.Name.LocalName == "Action"))
        { type.Unknown("ActionOutsideContainer", entry); target?.Unknown("ActionOutsideContainer", entry); }
        if (containers.Length > 1 || Foreign(owner, "Actions", ns))
        { type.Unknown("ActionsAmbiguous", entry); target?.Unknown("ActionsAmbiguous", entry); }
        foreach (var container in containers)
        {
            if (container.Elements().Any(x => x.Name != ns + "Action"))
            { type.Unknown("OpaqueActionsContent", entry); target?.Unknown("OpaqueActionsContent", entry); }
            if (container.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                container.Nodes().Any(x => x is XText text && !string.IsNullOrWhiteSpace(text.Value) ||
                    x is not XElement && x is not XText))
            { type.Unknown("ActionsStructureAmbiguous", entry); target?.Unknown("ActionsStructureAmbiguous", entry); }
            if (Foreign(container, "Action", ns))
            { type.Unknown("ForeignAction", entry); target?.Unknown("ForeignAction", entry); }
            foreach (var action in container.Elements(ns + "Action"))
            {
                token.ThrowIfCancellationRequested();
                type.Seen(entry, action, this);
                var kinds = action.Elements().Where(x => x.Name != ns + "Region").ToArray();
                var actionOpaque = action.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Event") ||
                    (string?)action.Attribute("Event") is not (null or "DO" or "PO" or "CLICK") ||
                    OpaqueNodes(action) || action.Elements(ns + "Region").Any(x => !StandardActionRegion(x, ns));
                if (target is not null)
                    foreach (var go in kinds.Where(x => x.Name == ns + "Goto"))
                    {
                        target.Seen(entry, action, this);
                        if (!actionOpaque && !GotoOpaque(go, ns) &&
                            go.Elements().Count(x => x.Name == ns + "Dest" || x.Name == ns + "Bookmark") == 1)
                            CheckTarget(go, ns, target, entry, action);
                        else target.Unknown("GotoTargetAmbiguous", entry, action, this);
                    }
                if (kinds.Any(x => x.Name.Namespace == ns && x.Name.LocalName is "URI" or "GotoA" or "Sound" or "Movie"))
                    type.Fail("NonGotoAction", entry, action, this);
                else if (kinds.Length != 1 || kinds[0].Name != ns + "Goto" ||
                    action.Elements(ns + "Region").Count() > 1 ||
                    kinds[0].Elements().Count(x => x.Name == ns + "Dest" || x.Name == ns + "Bookmark") != 1 ||
                    kinds[0].Elements().Any(x => x.Name != ns + "Dest" && x.Name != ns + "Bookmark") ||
                    actionOpaque || GotoOpaque(kinds[0], ns))
                { type.Unknown("UnsupportedAction", entry, action, this); target?.Unknown("UnsupportedAction", entry, action, this); }
            }
        }
    }

    private static bool StandardActionRegion(XElement region, XNamespace ns)
    {
        if (region.Parent?.Name != ns + "Action" ||
            !ReferenceEquals(region.Parent.Elements().FirstOrDefault(), region) ||
            region.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(region)) return false;
        var areas = region.Elements().ToArray();
        if (areas.Length == 0) return false;
        foreach (var area in areas)
        {
            if (area.Name != ns + "Area" || string.IsNullOrWhiteSpace((string?)area.Attribute("Start")) ||
                area.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Start") ||
                OpaqueNodes(area)) return false;
            var commands = area.Elements().ToArray();
            if (commands.Length == 0) return false;
            foreach (var command in commands)
            {
                if (command.HasElements || OpaqueNodes(command)) return false;
                string[]? names = command.Name == ns + "Move" || command.Name == ns + "Line" ? new[] { "Point1" } :
                    command.Name == ns + "QuadraticBezier" ? new[] { "Point1", "Point2" } :
                    command.Name == ns + "CubicBezier" ? new[] { "Point1", "Point2", "Point3" } :
                    command.Name == ns + "Arc" ? new[] { "SweepDirection", "LargeArc", "RotationAngle", "EllipseSize", "EndPoint" } :
                    command.Name == ns + "Close" ? Array.Empty<string>() : null;
                if (names is null || command.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                    (x.Name.Namespace != XNamespace.None || !names.Contains(x.Name.LocalName))))
                    return false;
                if (command.Name != ns + "Close" && names.Any(name => string.IsNullOrWhiteSpace((string?)command.Attribute(name))))
                    return false;
            }
        }
        return true;
    }

    private static bool GotoOpaque(XElement go, XNamespace ns)
    {
        if (go.Attributes().Any(x => !x.IsNamespaceDeclaration) || OpaqueNodes(go)) return true;
        foreach (var child in go.Elements())
        {
            if (child.Name == ns + "Dest")
            {
                if (DestinationOpaque(child)) return true;
            }
            else if (child.Name == ns + "Bookmark")
            {
                if (child.HasElements || OpaqueNodes(child) || string.IsNullOrWhiteSpace((string?)child.Attribute("Name")) ||
                    child.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Name"))
                    return true;
            }
            else return true;
        }
        return false;
    }

    private static bool OpaqueNodes(XElement element) => element.Nodes().Any(x =>
        x is XText text && !string.IsNullOrWhiteSpace(text.Value) || x is not XElement && x is not XText);

    private static bool DestinationOpaque(XElement dest)
    {
        if (dest.HasElements || OpaqueNodes(dest) ||
            (string?)dest.Attribute("Type") is not ("XYZ" or "Fit" or "FitH" or "FitV" or "FitR") ||
            !Id((string?)dest.Attribute("PageID"), out _) ||
            dest.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                (x.Name.Namespace != XNamespace.None || x.Name.LocalName is not
                    ("Type" or "PageID" or "Left" or "Top" or "Right" or "Bottom" or "Zoom")))) return true;
        foreach (var name in new[] { "Left", "Top", "Right", "Bottom", "Zoom" })
        {
            var attribute = dest.Attribute(name);
            if (attribute is null) continue;
            if (!double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ||
                double.IsNaN(number) || double.IsInfinity(number) ||
                name == "Zoom" && number != 0 && (number < .1 || number > 64)) return true;
        }
        return false;
    }

    private void CheckTarget(XElement gotoElement, XNamespace ns, Observation target, string entry, XElement action)
    {
        var dest = gotoElement.Element(ns + "Dest");
        if (dest is null)
        {
            var bookmark = gotoElement.Element(ns + "Bookmark");
            var name = (string?)bookmark?.Attribute("Name");
            if (string.IsNullOrWhiteSpace(name) || bookmark is null || ForeignAttribute(bookmark, "Name"))
            { target.Unknown("BookmarkTargetAmbiguous", entry, action, this); return; }
            if (ambiguousBookmarkNames.Contains(name!) || bookmarkInventoryUnknown)
            { target.Unknown("BookmarkInventoryAmbiguous", entry, action, this); return; }
            if (!bookmarks.TryGetValue(name!, out var declared))
            { target.Fail("BookmarkTargetMissing", entry, action, this); return; }
            var destinations = declared.Elements(ns + "Dest").ToArray();
            if (destinations.Length != 1 || Foreign(declared, "Dest", ns) || OpaqueNodes(declared) ||
                declared.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Name") ||
                declared.Elements().Any(x => x.Name != ns + "Dest"))
            { target.Unknown("BookmarkDestinationAmbiguous", entry, action, this); return; }
            dest = destinations[0];
        }
        var pageId = (string?)dest.Attribute("PageID");
        if (DestinationOpaque(dest))
        { target.Unknown("PageTargetAmbiguous", entry, action, this); return; }
        if (!Id(pageId, out var number) || ForeignAttribute(dest, "PageID"))
        { target.Unknown("PageTargetAmbiguous", entry, action, this); return; }
        if (ambiguousPageIds.Contains(number))
        { target.Unknown("PageInventoryAmbiguous", entry, action, this); return; }
        if (!pageIds.Contains(number))
        {
            if (pageInventoryUnknown)
                target.Unknown("PageInventoryAmbiguous", entry, action, this);
            else target.Fail("PageTargetMissing", entry, action, this);
        }
    }

    private XElement? Read(string entry, string rootName)
    {
        token.ThrowIfCancellationRequested();
        if (xml.TryGetValue(entry, out var cached))
            return cached?.Name.LocalName == rootName && Supported(cached.Name.NamespaceName) ? cached : null;
        if (!archive.TryGetBytes(entry, out var bytes) || bytes.Length > options.MaxSingleXmlBytes ||
            xmlBytes > options.MaxXmlBytes - bytes.Length) return null;
        xmlBytes += bytes.Length;
        XElement? value;
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = options.MaxSingleXmlBytes, MaxCharactersFromEntities = 1 };
        try
        {
            // Charge nodes and depth while streaming, before an untrusted tree is allocated.
            using (var source = new MemoryStream(bytes, false))
            using (var reader = XmlReader.Create(source, settings))
                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();
                    if (xmlNodes >= options.MaxXmlNodes ||
                        reader.AttributeCount > options.MaxXmlNodes - xmlNodes - 1 ||
                        reader.Depth > options.MaxXmlDepth)
                    { xml[entry] = null; return null; }
                    xmlNodes += 1 + reader.AttributeCount;
                }
            using var materialization = new MemoryStream(bytes, false);
            using var treeReader = XmlReader.Create(materialization, settings);
            value = XDocument.Load(treeReader, LoadOptions.PreserveWhitespace).Root;
        }
        catch (XmlException) { xml[entry] = null; return null; }
        if (value?.Name.LocalName != rootName || !Supported(value.Name.NamespaceName))
        { xml[entry] = null; return null; }
        IndexOrdinals(value);
        xml[entry] = value;
        return value;
    }

    private void IndexOrdinals(XElement value)
    {
        siblingOrdinals[value] = 1;
        var pending = new Stack<XElement>();
        pending.Push(value);
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var parent = pending.Pop();
            var counts = new Dictionary<XName, int>();
            foreach (var child in parent.Elements())
            {
                counts.TryGetValue(child.Name, out var ordinal);
                siblingOrdinals[child] = counts[child.Name] = ordinal + 1;
                pending.Push(child);
            }
        }
    }

    private bool TryEntry(string containing, string? path, out string entry)
    {
        entry = string.Empty;
        if (++references > options.MaxReferences) return false;
        path = path?.Trim();
        if (path is null || path.Length == 0 || path.Contains(":") || path.Contains("?") || path.Contains("#") ||
            path.Contains("%") || path.Contains("\\") || path.StartsWith("//", StringComparison.Ordinal) ||
            path.Any(char.IsControl)) return false;
        try { entry = OfdPackagePath.Resolve(containing, path); return true; }
        catch (InvalidDataException) { return false; }
    }

    private void UnknownAll(string reason, string entry) { foreach (var item in observations) item.Unknown(reason, entry); }
    private void UnknownPageScope(string reason, string entry)
    { observations[0].Unknown(reason, entry); observations[7].Unknown(reason, entry); UnknownGraphics(reason, entry); }
    private void UnknownOutline(string reason, string entry) { observations[1].Unknown(reason, entry); observations[2].Unknown(reason, entry); }
    private void UnknownGraphics(string reason, string entry) { observations[3].Unknown(reason, entry); observations[4].Unknown(reason, entry); }
    private void UnknownAnnotationScope(string reason, string entry)
    { observations[5].Unknown(reason, entry); observations[8].Unknown(reason, entry); UnknownGraphics(reason, entry); }
    private IReadOnlyList<ArchivalCoverageRuleResult> Finish() => Array.AsReadOnly(observations.Select(x => x.Finish()).ToArray());
    private static bool Supported(string ns) => ns is "http://www.ofdspec.org" or "http://www.ofdspec.org/2016";
    private static bool Id(string? value, out long number) =>
        long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out number) && number > 0;
    private static bool Foreign(XElement parent, string local, XNamespace ns) =>
        parent.Elements().Any(x => x.Name.LocalName == local && x.Name.Namespace != ns);
    private static bool ForeignAttribute(XElement parent, string local) =>
        parent.Attributes().Any(x => !x.IsNamespaceDeclaration &&
            x.Name.LocalName == local && x.Name.Namespace != XNamespace.None);
    private string Selector(XElement node)
    {
        var value = "/" + string.Join("/", node.AncestorsAndSelf().Reverse().Select(x =>
            "{" + (Supported(x.Name.NamespaceName) ? x.Name.NamespaceName : "<unsupported-namespace>") + "}" +
            x.Name.LocalName + "[" + (siblingOrdinals.TryGetValue(x, out var ordinal) ? ordinal : 1) + "]"));
        return value.Length <= 8192 ? value : "<selector exceeds metadata limit>";
    }

    private sealed class Observation
    {
        private readonly int index;
        private int seen;
        private bool uncertain;
        private bool failed;
        private string reason = "";
        private string entry = "<document>";
        private string selector = "<scope>";
        internal Observation(int index) { this.index = index; }
        internal void Seen(string at, XElement node, ArchivalStructuralCoverage owner)
        { if (++seen == 1 && !failed && !uncertain) { entry = at; selector = owner.Selector(node); } }
        internal void Unknown(string why, string at, string where = "<scope>")
        { if (failed) return; if (!uncertain) { reason = why; entry = at; selector = where; } uncertain = true; }
        internal void Unknown(string why, string at, XElement node, ArchivalStructuralCoverage owner)
        { if (failed) return; if (!uncertain) { reason = why; entry = at; selector = owner.Selector(node); } uncertain = true; }
        internal void Fail(string why, string at, XElement node, ArchivalStructuralCoverage owner)
        { if (failed) return; failed = true; reason = why; entry = at; selector = owner.Selector(node); }
        internal ArchivalCoverageRuleResult Finish()
        {
            var rule = ArchivalCoverageRuleCatalog.Rules[index];
            var outcome = failed ? PreflightOutcome.Fail : uncertain ? PreflightOutcome.Unknown :
                seen == 0 ? PreflightOutcome.NotApplicable : PreflightOutcome.Pass;
            var code = failed || uncertain ? reason : seen == 0 ? "NoApplicableObject" : "ConditionSatisfied";
            return new ArchivalCoverageRuleResult(rule, outcome, code, entry, selector,
                "Recognized objects=" + seen.ToString(CultureInfo.InvariantCulture) + "; " + code);
        }
    }
}
