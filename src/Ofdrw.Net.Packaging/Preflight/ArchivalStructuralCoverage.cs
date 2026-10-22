using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
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
    private readonly Dictionary<string, XElement> bookmarks = new(StringComparer.Ordinal);
    private readonly HashSet<string> pageIds = new(StringComparer.Ordinal);
    private long xmlBytes;
    private int xmlNodes;
    private int references;
    private bool targetInventoryUnknown;
    private string documentEntry = "<document>";

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
            !TryEntry("OFD.xml", roots[0].Value, out documentEntry)) return null;
        var document = Read(documentEntry, "Document");
        return document;
    }

    private void ReadPageInventory(XElement document, XNamespace ns)
    {
        var pages = document.Elements(ns + "Pages").ToArray();
        if (pages.Length != 1 || Foreign(document, "Pages", ns))
        {
            UnknownPageScope("PageTreeAmbiguous", documentEntry);
            targetInventoryUnknown = true;
        }
        foreach (var pageTree in pages)
        {
            if (Foreign(pageTree, "Page", ns))
            { UnknownPageScope("ForeignPageDeclaration", documentEntry); targetInventoryUnknown = true; }
            foreach (var page in pageTree.Elements(ns + "Page"))
            {
                token.ThrowIfCancellationRequested();
                var id = (string?)page.Attribute("ID");
                if (string.IsNullOrWhiteSpace(id) || !pageIds.Add(id!) || ForeignAttribute(page, "ID"))
                    targetInventoryUnknown = true;
                ReadPage(page, documentEntry, false);
            }
        }
        if (pageIds.Count == 0)
        { targetInventoryUnknown = true; UnknownPageScope("PageTreeEmpty", documentEntry); }
        var common = document.Elements(ns + "CommonData").ToArray();
        if (common.Length > 1 || Foreign(document, "CommonData", ns))
        { UnknownGraphics("CommonDataAmbiguous", documentEntry); observations[7].Unknown("CommonDataAmbiguous", documentEntry); }
        foreach (var template in common.SelectMany(x => x.Elements(ns + "TemplatePage")))
            ReadPage(template, documentEntry, true);
        if (common.Any(x => Foreign(x, "TemplatePage", ns)))
        { UnknownGraphics("ForeignTemplatePage", documentEntry); observations[7].Unknown("ForeignTemplatePage", documentEntry); }
    }

    private void ReadPage(XElement declaration, string containing, bool template)
    {
        if (ForeignAttribute(declaration, "BaseLoc") ||
            !TryEntry(containing, (string?)declaration.Attribute("BaseLoc"), out var entry))
        {
            if (!template) observations[0].Unknown("PageReferenceUnavailable", containing);
            UnknownGraphics("PageReferenceUnavailable", containing);
            observations[7].Unknown("PageReferenceUnavailable", containing);
            if (!template) targetInventoryUnknown = true;
            return;
        }
        var page = Read(entry, "Page");
        if (page is null)
        {
            if (!template) observations[0].Unknown("PageContentUnavailable", entry);
            UnknownGraphics("PageContentUnavailable", entry);
            observations[7].Unknown("PageContentUnavailable", entry);
            if (!template) targetInventoryUnknown = true;
            return;
        }
        if (!template) InspectActions(page, page.Name.Namespace, observations[0], null, entry);
        if (page.Elements(page.Name.Namespace + "Content").Count() != 1 ||
            Foreign(page, "Content", page.Name.Namespace))
        { UnknownGraphics("PageContentAmbiguous", entry); observations[7].Unknown("PageContentAmbiguous", entry); }
        InspectLayers(page, entry);
        InspectGraphics(page, entry);
    }

    private void ReadBookmarkInventory(XElement document, XNamespace ns)
    {
        var containers = document.Elements(ns + "Bookmarks").ToArray();
        if (containers.Length > 1 || Foreign(document, "Bookmarks", ns)) targetInventoryUnknown = true;
        foreach (var container in containers)
        {
            if (Foreign(container, "Bookmark", ns)) targetInventoryUnknown = true;
            foreach (var bookmark in container.Elements(ns + "Bookmark"))
            {
                var name = (string?)bookmark.Attribute("Name");
                if (string.IsNullOrWhiteSpace(name) || ForeignAttribute(bookmark, "Name") ||
                    bookmarks.ContainsKey(name!)) targetInventoryUnknown = true;
                else bookmarks.Add(name!, bookmark);
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
            if (Foreign(container, "OutlineElem", ns))
                UnknownOutline("ForeignOutline", documentEntry);
            if (container.Elements().Any(x => x.Name != ns + "OutlineElem"))
                UnknownOutline("OpaqueOutlineContent", documentEntry);
            foreach (var node in container.Elements(ns + "OutlineElem")) ReadOutlineNode(node, ns, 0);
        }
        if (targetInventoryUnknown) observations[2].Unknown("TargetInventoryAmbiguous", documentEntry);
    }

    private void ReadOutlineNode(XElement node, XNamespace ns, int depth)
    {
        token.ThrowIfCancellationRequested();
        if (depth > options.MaxXmlDepth)
        { UnknownOutline("OutlineDepthExceeded", documentEntry); return; }
        InspectActions(node, ns, observations[1], observations[2], documentEntry);
        if (Foreign(node, "OutlineElem", ns)) UnknownOutline("ForeignOutline", documentEntry);
        if (node.Elements().Any(x => x.Name != ns + "OutlineElem" && x.Name != ns + "Actions"))
            UnknownOutline("OpaqueOutlineContent", documentEntry);
        foreach (var child in node.Elements(ns + "OutlineElem")) ReadOutlineNode(child, ns, depth + 1);
    }

    private void ReadAnnotations(XElement document, XNamespace ns)
    {
        var declarations = document.Elements(ns + "Annotations").ToArray();
        if (declarations.Length > 1 || Foreign(document, "Annotations", ns))
            UnknownAnnotationScope("AnnotationScopeAmbiguous", documentEntry);
        foreach (var declaration in declarations)
        {
            if (declaration.HasElements || !TryEntry(documentEntry, declaration.Value, out var indexEntry))
            { UnknownAnnotationScope("AnnotationIndexUnavailable", documentEntry); continue; }
            var index = Read(indexEntry, "Annotations");
            if (index is null)
            { UnknownAnnotationScope("AnnotationIndexUnavailable", indexEntry); continue; }
            if (index.Elements().Any(x => x.Name != index.Name.Namespace + "Page"))
                UnknownAnnotationScope("OpaqueAnnotationIndex", indexEntry);
            if (Foreign(index, "Page", index.Name.Namespace))
            { UnknownAnnotationScope("ForeignAnnotationPage", indexEntry); }
            foreach (var page in index.Elements(index.Name.Namespace + "Page"))
            {
                if (ForeignAttribute(page, "PageID") || !pageIds.Contains((string?)page.Attribute("PageID") ?? ""))
                    UnknownAnnotationScope("AnnotationPageOwnerAmbiguous", indexEntry);
                var files = page.Elements(index.Name.Namespace + "FileLoc").ToArray();
                if (files.Length != 1 || Foreign(page, "FileLoc", index.Name.Namespace) ||
                    files[0].HasElements || !TryEntry(indexEntry, files[0].Value, out var entry))
                { UnknownAnnotationScope("AnnotationPageUnavailable", indexEntry); continue; }
                var content = Read(entry, "PageAnnot");
                if (content is null)
                { UnknownAnnotationScope("AnnotationPageUnavailable", entry); continue; }
                if (content.Elements().Any(x => x.Name != content.Name.Namespace + "Annot"))
                    UnknownAnnotationScope("OpaqueAnnotationContent", entry);
                if (Foreign(content, "Annot", content.Name.Namespace))
                { UnknownAnnotationScope("ForeignAnnotation", entry); }
                foreach (var annot in content.Elements(content.Name.Namespace + "Annot"))
                {
                    observations[5].Seen(entry, Selector(annot));
                    foreach (var flag in new[] { "ReadOnly", "NoZoom", "NoRotate" })
                    {
                        var value = (string?)annot.Attribute(flag);
                        if (value is null || value is "false" or "0")
                            observations[5].Fail("RetainedAnnotationFlagNotTrue", entry, Selector(annot));
                        else if (value is not ("true" or "1") || ForeignAttribute(annot, flag))
                            observations[5].Unknown("AnnotationFlagAmbiguous", entry, Selector(annot));
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
            if (content.Elements().Any(x => x.Name != ns + "Layer"))
                observations[7].Unknown("OpaqueLayerScope", entry);
            if (Foreign(content, "Layer", ns)) observations[7].Unknown("ForeignLayer", entry);
            foreach (var layer in content.Elements(ns + "Layer"))
            {
                observations[7].Seen(entry, Selector(layer));
                var name = (string?)layer.Attribute("Name");
                if (ForeignAttribute(layer, "Name")) observations[7].Unknown("LayerNameAmbiguous", entry, Selector(layer));
                if (name is not null && !names.Add(name))
                    observations[7].Fail("DuplicateLayerName", entry, Selector(layer));
            }
        }
    }

    private void InspectAnnotationBlocks(XElement appearance, string entry)
    {
        var ns = appearance.Name.Namespace;
        if (appearance.Descendants().Any(x => x.Name.LocalName == "PageBlock" && x.Name.Namespace != ns) ||
            appearance.Descendants().Any(x => x.Name.Namespace != ns))
            observations[8].Unknown("OpaqueAnnotationAppearance", entry, Selector(appearance));
        foreach (var block in appearance.Descendants(ns + "PageBlock"))
        {
            observations[8].Seen(entry, Selector(block));
            if (block.Ancestors().TakeWhile(x => x != appearance).Any(x => x.Name == ns + "PageBlock"))
                observations[8].Fail("NestedAnnotationPageBlock", entry, Selector(block));
        }
    }

    private void InspectGraphics(XElement source, string entry)
    {
        var ns = source.Name.Namespace;
        if (source.Descendants().Any(x => x.Name.Namespace != ns))
            UnknownGraphics("ForeignGraphicContent", entry);
        if (source.DescendantsAndSelf().Any(x =>
                (x.Name.LocalName == "ImageObject" || x.Name.LocalName == "TextObject") && x.Name.Namespace != ns) ||
            source.DescendantsAndSelf().Any(x => x.Name == ns + "CompositeObject"))
            UnknownGraphics("OpaqueGraphicScope", entry);
        foreach (var image in source.Descendants(ns + "ImageObject"))
        {
            var value = (string?)image.Attribute("Interpolate");
            if (value is null) { if (ForeignAttribute(image, "Interpolate")) observations[3].Unknown("InterpolateAmbiguous", entry, Selector(image)); continue; }
            observations[3].Seen(entry, Selector(image));
            if (value is "true" or "1") observations[3].Fail("InterpolateEnabled", entry, Selector(image));
            else if (value is not ("false" or "0") || ForeignAttribute(image, "Interpolate"))
                observations[3].Unknown("InterpolateAmbiguous", entry, Selector(image));
        }
        foreach (var text in source.Descendants(ns + "TextObject"))
        {
            observations[4].Seen(entry, Selector(text));
            var value = (string?)text.Attribute("Size");
            if (value is null || !decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) || size <= 0)
                observations[4].Fail("TextSizeUnavailable", entry, Selector(text));
            else if (ForeignAttribute(text, "Size"))
                observations[4].Unknown("TextSizeAmbiguous", entry, Selector(text));
        }
        foreach (var graphic in source.Descendants().Where(x => x.Name.Namespace == ns &&
                     (x.Name.LocalName == "TextObject" || x.Name.LocalName == "PathObject" ||
                      x.Name.LocalName == "ImageObject" || x.Name.LocalName == "CompositeObject")))
        {
            if (graphic.Elements().Any(x => x.Name == ns + "Actions" || x.Name == ns + "Action"))
                observations[6].Fail("ObjectActionRetained", entry, Selector(graphic));
            else if (graphic.Elements().Any(x =>
                         (x.Name.LocalName == "Actions" || x.Name.LocalName == "Action") && x.Name.Namespace != ns))
                observations[6].Unknown("ForeignObjectAction", entry, Selector(graphic));
        }
    }

    private void InspectActions(XElement owner, XNamespace ns, Observation type, Observation? target, string entry)
    {
        var containers = owner.Elements(ns + "Actions").ToArray();
        if (containers.Length > 1 || Foreign(owner, "Actions", ns))
        { type.Unknown("ActionsAmbiguous", entry); target?.Unknown("ActionsAmbiguous", entry); }
        foreach (var container in containers)
        {
            if (container.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
                container.Nodes().Any(x => x is XText text && !string.IsNullOrWhiteSpace(text.Value) ||
                    x is not XElement && x is not XText))
            { type.Unknown("ActionsStructureAmbiguous", entry); target?.Unknown("ActionsStructureAmbiguous", entry); }
            if (Foreign(container, "Action", ns))
            { type.Unknown("ForeignAction", entry); target?.Unknown("ForeignAction", entry); }
            foreach (var action in container.Elements(ns + "Action"))
            {
                token.ThrowIfCancellationRequested();
                var location = Selector(action);
                type.Seen(entry, location);
                var kinds = action.Elements().Where(x => x.Name != ns + "Region").ToArray();
                if (target is not null)
                    foreach (var go in kinds.Where(x => x.Name == ns + "Goto"))
                    {
                        target.Seen(entry, location);
                        if (go.Elements().Count(x => x.Name == ns + "Dest" || x.Name == ns + "Bookmark") == 1)
                            CheckTarget(go, ns, target, entry, location);
                        else target.Unknown("GotoTargetAmbiguous", entry, location);
                    }
                if (kinds.Any(x => x.Name.Namespace == ns && x.Name.LocalName is "URI" or "GotoA" or "Sound" or "Movie"))
                    type.Fail("NonGotoAction", entry, location);
                else if (kinds.Length != 1 || kinds[0].Name != ns + "Goto" ||
                    action.Elements(ns + "Region").Count() > 1 ||
                    kinds[0].Elements().Count(x => x.Name == ns + "Dest" || x.Name == ns + "Bookmark") != 1 ||
                    kinds[0].Elements().Any(x => x.Name != ns + "Dest" && x.Name != ns + "Bookmark") ||
                    action.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "Event"))
                { type.Unknown("UnsupportedAction", entry, location); target?.Unknown("UnsupportedAction", entry, location); }
                else if (target is not null && kinds[0].Elements().Any(x => x.Name != ns + "Dest" && x.Name != ns + "Bookmark"))
                    target.Unknown("GotoTargetAmbiguous", entry, location);
            }
        }
    }

    private void CheckTarget(XElement gotoElement, XNamespace ns, Observation target, string entry, string location)
    {
        var dest = gotoElement.Element(ns + "Dest");
        if (dest is null)
        {
            var bookmark = gotoElement.Element(ns + "Bookmark");
            var name = (string?)bookmark?.Attribute("Name");
            if (string.IsNullOrWhiteSpace(name) || bookmark is null || ForeignAttribute(bookmark, "Name"))
            { target.Unknown("BookmarkTargetAmbiguous", entry, location); return; }
            if (!bookmarks.TryGetValue(name!, out var declared))
            {
                if (targetInventoryUnknown) target.Unknown("BookmarkInventoryAmbiguous", entry, location);
                else target.Fail("BookmarkTargetMissing", entry, location);
                return;
            }
            var destinations = declared.Elements(ns + "Dest").ToArray();
            if (destinations.Length != 1 || Foreign(declared, "Dest", ns))
            { target.Unknown("BookmarkDestinationAmbiguous", entry, location); return; }
            dest = destinations[0];
        }
        var pageId = (string?)dest.Attribute("PageID");
        if (dest.HasElements || dest.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                x.Name.Namespace != XNamespace.None))
        { target.Unknown("PageTargetAmbiguous", entry, location); return; }
        if (string.IsNullOrWhiteSpace(pageId) || ForeignAttribute(dest, "PageID"))
        { target.Unknown("PageTargetAmbiguous", entry, location); return; }
        if (!pageIds.Contains(pageId!))
        {
            if (targetInventoryUnknown) target.Unknown("PageInventoryAmbiguous", entry, location);
            else target.Fail("PageTargetMissing", entry, location);
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
        var value = ArchivalRuleService.Read(archive, entry, options)?.Root;
        if (value?.Name.LocalName != rootName || !Supported(value.Name.NamespaceName))
        { xml[entry] = null; return null; }
        var stack = new Stack<(XElement Node, int Depth)>();
        stack.Push((value, 1));
        while (stack.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var current = stack.Pop();
            if (++xmlNodes > options.MaxXmlNodes || current.Depth > options.MaxXmlDepth) return null;
            foreach (var child in current.Node.Elements()) stack.Push((child, current.Depth + 1));
        }
        xml[entry] = value;
        return value;
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
    private static bool Foreign(XElement parent, string local, XNamespace ns) =>
        parent.Elements().Any(x => x.Name.LocalName == local && x.Name.Namespace != ns);
    private static bool ForeignAttribute(XElement parent, string local) =>
        parent.Attributes().Any(x => x.Name.LocalName == local && x.Name.Namespace != XNamespace.None);
    private static string Selector(XElement node)
    {
        var value = "/" + string.Join("/", node.AncestorsAndSelf().Reverse().Select(x =>
            "{" + (Supported(x.Name.NamespaceName) ? x.Name.NamespaceName : "<unsupported-namespace>") + "}" +
            x.Name.LocalName + "[" + (x.ElementsBeforeSelf(x.Name).Count() + 1) + "]"));
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
        internal void Seen(string at, string where) { seen++; if (!failed && !uncertain) { entry = at; selector = where; } }
        internal void Unknown(string why, string at, string where = "<scope>")
        { if (failed) return; if (!uncertain) { reason = why; entry = at; selector = where; } uncertain = true; }
        internal void Fail(string why, string at, string where)
        { if (failed) return; failed = true; reason = why; entry = at; selector = where; }
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
