using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>Kinds of package locations the scope walk records.</summary>
internal enum OfdScopeLocationKind
{
    Document, PublicRes, DocumentRes, Page, PageRes, Template, Signatures, Signature, Attachments, Attachment,
    CustomTags, CustomTag, Annotations, PageAnnotation, Extensions, ResourcePayload
}

/// <summary>
/// The parts of one document whose completeness a read-only check can ask about. A part is complete only when every
/// declaration that could contribute to it resolved and was readable.
/// </summary>
internal enum OfdScopeArea
{
    /// <summary>OFD.xml, the selected DocBody and its Document.xml.</summary>
    Document,
    /// <summary>Page and template declarations and their content.</summary>
    Pages,
    /// <summary>PublicRes, DocumentRes and every page/template PageRes (declared or the implicit sibling).</summary>
    Resources,
    /// <summary>The Signatures index and each Signature file location.</summary>
    Signatures,
    /// <summary>Annotation indexes and their page annotation file locations.</summary>
    Annotations,
    /// <summary>CustomTags indexes and their tag payload locations.</summary>
    CustomTags,
    /// <summary>Attachments indexes and their payload locations.</summary>
    Attachments
}

/// <summary>A reference declared in the package and where it resolves.</summary>
internal sealed class OfdScopeLocation
{
    internal OfdScopeLocation(OfdScopeLocationKind kind, string declaredIn, string selector, string declared, string? resolved)
    { Kind = kind; DeclaredIn = declaredIn; Selector = selector; Declared = declared; Resolved = resolved; }
    internal OfdScopeLocationKind Kind { get; }
    internal string DeclaredIn { get; }
    internal string Selector { get; }
    internal string Declared { get; }
    /// <summary>Package path, or null when the declared value is not a safe package-relative location.</summary>
    internal string? Resolved { get; }
}

/// <summary>Why part of the scope is not known: the location (resolved path, or the declared value) and its kind.</summary>
internal sealed class OfdScopeGap
{
    internal OfdScopeGap(OfdScopeArea area, OfdScopeLocationKind kind, string location, string? format = null)
    { Area = area; Kind = kind; Location = location; Format = format; }
    internal OfdScopeArea Area { get; }
    internal OfdScopeLocationKind Kind { get; }
    internal string Location { get; }
    /// <summary>The declared attachment Format, when the gap is an attachment payload.</summary>
    internal string? Format { get; }
}

/// <summary>Whether a set of scope areas is complete, with the gaps that make it otherwise.</summary>
internal sealed class OfdScopeVerdict
{
    internal OfdScopeVerdict(IReadOnlyList<OfdScopeGap> gaps) { Gaps = gaps; }
    internal IReadOnlyList<OfdScopeGap> Gaps { get; }
    internal bool Complete => Gaps.Count == 0;
    /// <summary>Evidence location for an incomplete verdict; null when complete.</summary>
    internal string? FirstGap => Gaps.Count == 0 ? null : Gaps[0].Location;
    /// <summary>Distinct gap locations in walk order.</summary>
    internal IReadOnlyList<string> Locations => Gaps.Select(x => x.Location).Distinct(StringComparer.Ordinal).ToList();
}

internal enum OfdResourceScope { Public, Document, Page }

/// <summary>A resource table (PublicRes, DocumentRes or PageRes) and its resolved payload base.</summary>
internal sealed class OfdScopeResourceTable
{
    internal OfdScopeResourceTable(OfdResourceScope scope, string path, XElement root, string? baseDirectory)
    { Scope = scope; Path = path; Root = root; BaseDirectory = baseDirectory; }
    internal OfdResourceScope Scope { get; }
    internal string Path { get; }
    internal XElement Root { get; }
    /// <summary>Directory payload references resolve against; null when BaseLoc is unsafe.</summary>
    internal string? BaseDirectory { get; }
}

/// <summary>A MultiMedia resource: Image, Audio, Video or another declared type.</summary>
internal sealed class OfdScopeMedia
{
    internal OfdScopeMedia(string type, OfdScopeResourceTable table, string selector, string declared, string? resolved)
    { Type = type; Table = table; Selector = selector; Declared = declared; Resolved = resolved; }
    internal string Type { get; }
    internal OfdScopeResourceTable Table { get; }
    internal string Selector { get; }
    internal string Declared { get; }
    internal string? Resolved { get; }
}

/// <summary>A page or template declaration and, when it resolved and parsed as a Page, its content.</summary>
internal sealed class OfdScopePage
{
    internal OfdScopePage(XElement declaration, bool template, string? path, XElement? root)
    { Declaration = declaration; IsTemplate = template; Path = path; Root = root; }
    internal XElement Declaration { get; }
    internal bool IsTemplate { get; }
    internal string? Path { get; }
    internal XElement? Root { get; }
}

/// <summary>
/// One bounded walk over a package's declared reference graph: OFD.xml, the first DocBody's Document.xml, its
/// resource tables, pages and templates, signatures, attachments, custom tags and annotations. Every unresolved,
/// missing or unreadable declaration is recorded as a gap in the areas it could contribute to, so a check asks
/// <see cref="Verdict"/> instead of rebuilding "is my scope complete?" itself. The walk charges the caller's
/// <see cref="BoundedPackageXml"/>, so each consumer keeps its own budget count.
/// </summary>
internal sealed class OfdPackageScope
{
    private static readonly OfdScopeArea[] DocumentAreas =
    { OfdScopeArea.Document, OfdScopeArea.Pages, OfdScopeArea.Resources, OfdScopeArea.Annotations, OfdScopeArea.CustomTags, OfdScopeArea.Attachments };

    private readonly List<OfdScopeGap> gaps = new List<OfdScopeGap>();

    private OfdPackageScope() { }

    /// <summary>The parsed OFD.xml, or null when missing or not well-formed.</summary>
    internal XDocument? Root { get; private set; }
    /// <summary>Namespace of the OFD root, whichever it is; check support separately.</summary>
    internal XNamespace? Namespace { get; private set; }
    internal IReadOnlyList<XElement> DocBodies { get; private set; } = Array.Empty<XElement>();
    /// <summary>The single DocRoot element of the first DocBody, when there is exactly one.</summary>
    internal XElement? DocRoot { get; private set; }
    internal string? DocumentPath { get; private set; }
    /// <summary>The parsed Document.xml when its root is a <c>Document</c> element; otherwise null.</summary>
    internal XDocument? Document { get; private set; }
    internal List<OfdScopeLocation> Locations { get; } = new List<OfdScopeLocation>();
    internal List<OfdScopeResourceTable> Tables { get; } = new List<OfdScopeResourceTable>();
    internal List<OfdScopeMedia> Media { get; } = new List<OfdScopeMedia>();
    /// <summary>Pages in document order, then templates; including declarations that did not resolve.</summary>
    internal List<OfdScopePage> Pages { get; } = new List<OfdScopePage>();
    /// <summary>Resolved ordinary page content paths in document order.</summary>
    internal IReadOnlyList<string> PagePaths => Pages.Where(x => !x.IsTemplate && x.Path != null).Select(x => x.Path!).ToList();
    internal string? SignaturesPath { get; private set; }
    /// <summary>True when OFD.xml declares a Signatures index, whether or not it resolves.</summary>
    internal bool SignaturesDeclared { get; private set; }
    internal List<string> SignaturePaths { get; } = new List<string>();
    /// <summary>Every declared signature in index order; null where its BaseLoc does not resolve.</summary>
    internal List<string?> SignatureDeclarations { get; } = new List<string?>();
    internal List<string> AnnotationIndexPaths { get; } = new List<string>();
    internal List<string> AnnotationPaths { get; } = new List<string>();
    internal List<(string Path, string? Format)> InternalAttachments { get; } = new List<(string, string?)>();
    internal List<string> TagPayloads { get; } = new List<string>();

    /// <summary>True when Document.xml could not be resolved or parsed: everything it declares is unknown, not absent.</summary>
    internal bool DocumentUnavailable => Document?.Root == null;

    /// <summary>Completeness of the union of the given areas; with no area given, of the whole walk.</summary>
    internal OfdScopeVerdict Verdict(params OfdScopeArea[] areas) =>
        new OfdScopeVerdict(areas.Length == 0 ? gaps.ToList() : gaps.Where(x => areas.Contains(x.Area)).ToList());

    /// <summary>Walks the package. Budget overruns propagate as <see cref="PackageBudgetExceededException"/>.</summary>
    internal static OfdPackageScope Walk(BoundedPackageXml xml)
    {
        var scope = new OfdPackageScope();
        scope.WalkPackage(xml);
        return scope;
    }

    private void WalkPackage(BoundedPackageXml xml)
    {
        Root = xml.Read("OFD.xml");
        var root = Root?.Root;
        if (root == null || root.Name.LocalName != "OFD") { GapAll("OFD.xml"); return; }
        var ns = Namespace = root.Name.Namespace;
        DocBodies = root.Elements(ns + "DocBody").ToList();
        if (DocBodies.Count == 0) { GapAll("OFD.xml"); return; }
        var body = DocBodies[0];

        var signatures = body.Element(ns + "Signatures");
        if (signatures != null) WalkSignatures(xml, signatures.Value);

        var roots = body.Elements(ns + "DocRoot").ToList();
        if (roots.Count != 1 || roots[0].HasElements) { GapDocument("OFD.xml"); return; }
        DocRoot = roots[0];
        var docRoot = roots[0].Value;
        var documentPath = Add(OfdScopeLocationKind.Document, "OFD.xml", "DocBody/DocRoot", docRoot);
        if (documentPath == null) { GapDocument(string.IsNullOrWhiteSpace(docRoot) ? "<missing DocRoot>" : docRoot); return; }
        DocumentPath = documentPath;
        xml.Visit();
        var parsed = xml.Read(documentPath);
        if (parsed?.Root == null || parsed.Root.Name.LocalName != "Document") { GapDocument(documentPath); return; }
        Document = parsed;
        WalkDocument(xml, documentPath, parsed.Root);
    }

    private void WalkDocument(BoundedPackageXml xml, string docPath, XElement document)
    {
        var ns = document.Name.Namespace;
        var common = document.Element(ns + "CommonData");
        foreach (var element in common?.Elements(ns + "PublicRes") ?? Enumerable.Empty<XElement>())
            AddTable(xml, OfdResourceScope.Public, OfdScopeLocationKind.PublicRes, docPath, "CommonData/PublicRes", element.Value);
        foreach (var element in common?.Elements(ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
            AddTable(xml, OfdResourceScope.Document, OfdScopeLocationKind.DocumentRes, docPath, "CommonData/DocumentRes", element.Value);

        foreach (var page in document.Elements(ns + "Pages").SelectMany(x => x.Elements(ns + "Page")))
            WalkPage(xml, docPath, page, false);
        // Template pages carry their own PageRes; their fonts and media are page-scoped resources too.
        foreach (var template in common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>())
            WalkPage(xml, docPath, template, true);

        // Every index declaration counts; a document may declare more than one.
        foreach (var attachments in document.Elements(ns + "Attachments").Select(x => x.Value))
            WalkAttachments(xml, docPath, attachments);
        foreach (var tags in document.Elements(ns + "CustomTags").Select(x => x.Value))
            WalkTags(xml, docPath, tags);
        foreach (var annotations in document.Elements(ns + "Annotations").Select(x => x.Value))
            WalkAnnotations(xml, docPath, annotations);

        var extensions = document.Element(ns + "Extensions")?.Value;
        if (extensions != null) Add(OfdScopeLocationKind.Extensions, docPath, "Document/Extensions", extensions);
    }

    private void WalkPage(BoundedPackageXml xml, string docPath, XElement declaration, bool template)
    {
        xml.Visit();
        var kind = template ? OfdScopeLocationKind.Template : OfdScopeLocationKind.Page;
        var location = (string?)declaration.Attribute("BaseLoc");
        var path = Add(kind, docPath, template ? "CommonData/TemplatePage/@BaseLoc" : "Pages/Page/@BaseLoc", location);
        // An unresolved or unreadable page may hold PageRes tables: the resource scope is incomplete, not empty.
        if (path == null)
        {
            Pages.Add(new OfdScopePage(declaration, template, null, null));
            GapPage(kind, location ?? (template ? "<missing TemplatePage BaseLoc>" : "<missing Page BaseLoc>"));
            return;
        }
        var content = xml.Read(path)?.Root;
        if (content == null || content.Name.LocalName != "Page")
        {
            Pages.Add(new OfdScopePage(declaration, template, path, null));
            GapPage(kind, path);
            return;
        }
        Pages.Add(new OfdScopePage(declaration, template, path, content));
        var ns = content.Name.Namespace;
        // Readers treat {pageDirectory}/PageRes.xml as a fixed page resource even without a PageRes link.
        var sibling = path.Substring(0, path.LastIndexOf('/') + 1) + "PageRes.xml";
        var declared = content.Elements(ns + "PageRes").Select(x => x.Value).ToList();
        if (xml.Exists(sibling) && !declared.Any(x => OfdPackageReference.TryResolve(path, x, out var target) &&
                string.Equals(target, sibling, StringComparison.OrdinalIgnoreCase)))
            AddTable(xml, OfdResourceScope.Page, OfdScopeLocationKind.PageRes, path, "Page/<implicit PageRes.xml>", "PageRes.xml");
        foreach (var pageRes in declared)
            AddTable(xml, OfdResourceScope.Page, OfdScopeLocationKind.PageRes, path, template ? "Template/PageRes" : "Page/PageRes", pageRes);
    }

    private void WalkSignatures(BoundedPackageXml xml, string declared)
    {
        SignaturesDeclared = true;
        var path = SignaturesPath = Add(OfdScopeLocationKind.Signatures, "OFD.xml", "DocBody/Signatures", declared);
        if (path == null) { Gap(OfdScopeArea.Signatures, OfdScopeLocationKind.Signatures, declared); return; }
        var list = xml.Read(path)?.Root;
        if (list == null) { Gap(OfdScopeArea.Signatures, OfdScopeLocationKind.Signatures, path); return; }
        foreach (var item in list.Elements(list.Name.Namespace + "Signature"))
        {
            xml.Visit();
            var location = (string?)item.Attribute("BaseLoc");
            var signature = Add(OfdScopeLocationKind.Signature, path, "Signature/@BaseLoc", location);
            SignatureDeclarations.Add(signature);
            if (signature != null) SignaturePaths.Add(signature);
            else Gap(OfdScopeArea.Signatures, OfdScopeLocationKind.Signature, location ?? "<missing Signature BaseLoc>");
        }
    }

    private void WalkAttachments(BoundedPackageXml xml, string docPath, string declared)
    {
        var path = Add(OfdScopeLocationKind.Attachments, docPath, "Document/Attachments", declared);
        var list = path == null ? null : xml.Read(path)?.Root;
        // The attachments an unreadable index lists are unknown.
        if (list == null) { Gap(OfdScopeArea.Attachments, OfdScopeLocationKind.Attachments, path ?? declared); return; }
        var ns = list.Name.Namespace;
        foreach (var item in list.Elements(ns + "Attachment"))
        {
            xml.Visit();
            // FileLoc may be an element or, as OfdReader and archival preflight accept, an attribute.
            var element = item.Element(ns + "FileLoc")?.Value; var attribute = (string?)item.Attribute("FileLoc");
            var format = (string?)item.Attribute("Format");
            if (element != null && attribute != null && element.Trim() != attribute.Trim())
            {
                Locations.Add(new OfdScopeLocation(OfdScopeLocationKind.Attachment, path!, "Attachment/FileLoc", "ambiguous FileLoc", null));
                Gap(OfdScopeArea.Attachments, OfdScopeLocationKind.Attachments, "ambiguous FileLoc in " + path);
                continue;
            }
            var location = element ?? attribute;
            // A location with a URI scheme is an external attachment; it is not a package entry.
            if (location != null && location.Contains("://")) continue;
            var payload = Add(OfdScopeLocationKind.Attachment, path!, "Attachment/FileLoc", location);
            if (payload != null) InternalAttachments.Add((payload, format));
            else Gap(OfdScopeArea.Attachments, OfdScopeLocationKind.Attachment, location ?? "<missing FileLoc>", format);
        }
    }

    private void WalkTags(BoundedPackageXml xml, string docPath, string declared)
    {
        var path = Add(OfdScopeLocationKind.CustomTags, docPath, "Document/CustomTags", declared);
        var list = path == null ? null : xml.Read(path)?.Root;
        if (list == null) { Gap(OfdScopeArea.CustomTags, OfdScopeLocationKind.CustomTags, path ?? declared); return; }
        var ns = list.Name.Namespace;
        foreach (var item in list.Elements(ns + "CustomTag"))
        {
            xml.Visit();
            var declaredPayload = item.Element(ns + "FileLoc")?.Value;
            var payload = Add(OfdScopeLocationKind.CustomTag, path!, "CustomTag/FileLoc", declaredPayload);
            if (payload != null) TagPayloads.Add(payload);
            else Gap(OfdScopeArea.CustomTags, OfdScopeLocationKind.CustomTag, declaredPayload ?? "<missing FileLoc>");
            var schema = item.Element(ns + "SchemaLoc")?.Value;
            if (schema != null) Add(OfdScopeLocationKind.CustomTag, path!, "CustomTag/SchemaLoc", schema);
        }
    }

    private void WalkAnnotations(BoundedPackageXml xml, string docPath, string declared)
    {
        var path = Add(OfdScopeLocationKind.Annotations, docPath, "Document/Annotations", declared);
        if (path == null) { Gap(OfdScopeArea.Annotations, OfdScopeLocationKind.Annotations, declared); return; }
        AnnotationIndexPaths.Add(path);
        var list = xml.Read(path)?.Root;
        if (list == null) { Gap(OfdScopeArea.Annotations, OfdScopeLocationKind.Annotations, path); return; }
        var ns = list.Name.Namespace;
        foreach (var item in list.Elements(ns + "Page"))
        {
            xml.Visit();
            var location = item.Element(ns + "FileLoc")?.Value;
            var file = Add(OfdScopeLocationKind.PageAnnotation, path, "Page/FileLoc", location);
            if (file != null) AnnotationPaths.Add(file);
            else Gap(OfdScopeArea.Annotations, OfdScopeLocationKind.PageAnnotation, location ?? "<missing FileLoc>");
        }
    }

    private void AddTable(BoundedPackageXml xml, OfdResourceScope scope, OfdScopeLocationKind kind, string from, string selector, string declared)
    {
        xml.Visit();
        var path = Add(kind, from, selector, declared);
        if (path == null) { Gap(OfdScopeArea.Resources, kind, declared); return; }
        var root = xml.Read(path)?.Root;
        if (root == null) { Gap(OfdScopeArea.Resources, kind, path); return; }
        var slash = path.LastIndexOf('/');
        string? baseDirectory = slash < 0 ? string.Empty : path.Substring(0, slash);
        var baseLoc = (string?)root.Attribute("BaseLoc");
        if (baseLoc != null)
            baseDirectory = OfdPackageReference.TryResolveDirectory(path, baseLoc, out var resolvedBase) ? resolvedBase : null;
        var table = new OfdScopeResourceTable(scope, path, root, baseDirectory);
        Tables.Add(table);
        var ns = root.Name.Namespace;
        foreach (var media in root.Descendants(ns + "MultiMedia"))
        {
            xml.Visit();
            var file = media.Element(ns + "MediaFile")?.Value ?? string.Empty;
            var resolved = ResolvePayload(table, file);
            Media.Add(new OfdScopeMedia((string?)media.Attribute("Type") ?? string.Empty, table, "MultiMedia[@ID='" + (string?)media.Attribute("ID") + "']/MediaFile", file, resolved));
            Locations.Add(new OfdScopeLocation(OfdScopeLocationKind.ResourcePayload, path, "MultiMedia/MediaFile", file, resolved));
        }
        foreach (var font in root.Descendants(ns + "Font"))
        {
            var file = font.Element(ns + "FontFile")?.Value;
            if (file == null) continue;
            xml.Visit();
            Locations.Add(new OfdScopeLocation(OfdScopeLocationKind.ResourcePayload, path, "Font/FontFile", file, ResolvePayload(table, file)));
        }
    }

    /// <summary>Resolves a FontFile or MediaFile against its table's BaseLoc; a leading slash is the package root.</summary>
    internal static string? ResolvePayload(OfdScopeResourceTable table, string declared)
    {
        if (declared.StartsWith("/", StringComparison.Ordinal))
            return OfdPackageReference.TryResolve(table.Path, declared, out var absolute) ? absolute : null;
        if (table.BaseDirectory == null) return null;
        return OfdPackageReference.TryResolve(table.BaseDirectory.Length == 0 ? "x" : table.BaseDirectory + "/x", declared, out var resolved) ? resolved : null;
    }

    private string? Add(OfdScopeLocationKind kind, string from, string selector, string? declared)
    {
        var resolved = OfdPackageReference.TryResolve(from, declared, out var path) ? path : null;
        Locations.Add(new OfdScopeLocation(kind, from, selector, declared ?? string.Empty, resolved));
        return resolved;
    }

    private void Gap(OfdScopeArea area, OfdScopeLocationKind kind, string location, string? format = null) =>
        gaps.Add(new OfdScopeGap(area, kind, location, format));

    private void GapPage(OfdScopeLocationKind kind, string location)
    {
        Gap(OfdScopeArea.Pages, kind, location);
        Gap(OfdScopeArea.Resources, kind, location);
    }

    private void GapDocument(string location)
    {
        foreach (var area in DocumentAreas) Gap(area, OfdScopeLocationKind.Document, location);
    }

    private void GapAll(string location)
    {
        GapDocument(location);
        Gap(OfdScopeArea.Signatures, OfdScopeLocationKind.Document, location);
    }
}
