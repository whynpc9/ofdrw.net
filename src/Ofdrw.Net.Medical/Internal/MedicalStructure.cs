using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Medical.Internal;

/// <summary>Kinds of package locations the structure walk records.</summary>
internal enum MedicalLocationKind
{
    Document, PublicRes, DocumentRes, Page, PageRes, Template, Signatures, Signature, Attachments, Attachment,
    CustomTags, CustomTag, Annotations, PageAnnotation, Extensions, ResourcePayload
}

/// <summary>A reference declared in the package and where it resolves.</summary>
internal sealed class MedicalLocation
{
    internal MedicalLocation(MedicalLocationKind kind, string declaredIn, string selector, string declared, string? resolved)
    { Kind = kind; DeclaredIn = declaredIn; Selector = selector; Declared = declared; Resolved = resolved; }
    internal MedicalLocationKind Kind { get; }
    internal string DeclaredIn { get; }
    internal string Selector { get; }
    internal string Declared { get; }
    /// <summary>Package path, or null when the declared value is not a safe package-relative location.</summary>
    internal string? Resolved { get; }
}

internal enum MedicalResourceScope { Public, Document, Page }

/// <summary>A resource table (PublicRes, DocumentRes or PageRes) and its resolved payload base.</summary>
internal sealed class MedicalResourceTable
{
    internal MedicalResourceTable(MedicalResourceScope scope, string path, XElement root, string? baseDirectory)
    { Scope = scope; Path = path; Root = root; BaseDirectory = baseDirectory; }
    internal MedicalResourceScope Scope { get; }
    internal string Path { get; }
    internal XElement Root { get; }
    /// <summary>Directory payload references resolve against; null when BaseLoc is unsafe.</summary>
    internal string? BaseDirectory { get; }
}

/// <summary>A MultiMedia resource: Image, Audio, Video or another declared type.</summary>
internal sealed class MedicalMedia
{
    internal MedicalMedia(string type, MedicalResourceTable table, string selector, string declared, string? resolved)
    { Type = type; Table = table; Selector = selector; Declared = declared; Resolved = resolved; }
    internal string Type { get; }
    internal MedicalResourceTable Table { get; }
    internal string Selector { get; }
    internal string Declared { get; }
    internal string? Resolved { get; }
}

/// <summary>
/// One walk over OFD.xml, the selected document, its resource tables, pages, signatures, attachments, tags and
/// annotations. Every unreadable or unsafe reference is recorded instead of guessed.
/// </summary>
internal sealed class MedicalStructure
{
    internal XDocument? Root;
    internal XNamespace? Namespace;
    internal List<XElement> DocBodies = new List<XElement>();
    internal string? DocumentPath;
    internal XDocument? Document;
    internal List<MedicalLocation> Locations = new List<MedicalLocation>();
    internal List<MedicalResourceTable> Tables = new List<MedicalResourceTable>();
    internal List<MedicalMedia> Media = new List<MedicalMedia>();
    internal List<string> PagePaths = new List<string>();
    internal List<string> Unreadable = new List<string>();
    internal string? SignaturesPath;
    /// <summary>True when OFD.xml declares a Signatures index, whether or not it resolves.</summary>
    internal bool SignaturesDeclared;
    internal List<string> SignaturePaths = new List<string>();
    /// <summary>Every declared signature in index order; null where its BaseLoc does not resolve.</summary>
    internal List<string?> SignatureDeclarations = new List<string?>();
    /// <summary>Resolved annotation index paths; a document may declare more than one.</summary>
    internal List<string> AnnotationIndexPaths = new List<string>();
    /// <summary>True when an annotation index is declared but its location does not resolve.</summary>
    internal bool AnnotationIndexUnresolved;
    /// <summary>Declared machine-readable payloads (or their indexes) that are unresolved, missing or unreadable.</summary>
    internal List<string> MachineReadableUnavailable = new List<string>();
    internal List<string> AnnotationPaths = new List<string>();
    internal List<(string Path, string? Format)> InternalAttachments = new List<(string, string?)>();
    internal List<string> TagPayloads = new List<string>();

    /// <summary>
    /// True when Document.xml could not be resolved or parsed: its tags, attachments, annotations and resources are then
    /// unknown, not absent.
    /// </summary>
    internal bool DocumentUnavailable => Document?.Root == null;

    internal static MedicalStructure Walk(MedicalPackage package)
    {
        var result = new MedicalStructure();
        result.Root = package.Xml("OFD.xml");
        var root = result.Root?.Root;
        if (root == null || root.Name.LocalName != "OFD") return result;
        var ns = result.Namespace = root.Name.Namespace;
        result.DocBodies = root.Elements(ns + "DocBody").ToList();
        if (result.DocBodies.Count == 0) return result;
        var body = result.DocBodies[0];

        var docRoot = body.Element(ns + "DocRoot")?.Value;
        if (MedicalPackage.TryResolve("OFD.xml", docRoot, out var documentPath))
        {
            result.DocumentPath = documentPath;
            result.Locations.Add(new MedicalLocation(MedicalLocationKind.Document, "OFD.xml", "DocBody/DocRoot", docRoot!, documentPath));
            result.Document = package.Xml(documentPath);
            if (result.Document == null) result.Unreadable.Add(documentPath);
        }
        var signatures = body.Element(ns + "Signatures")?.Value;
        if (signatures != null) result.WalkSignatures(package, signatures);

        var document = result.Document?.Root;
        if (document == null || result.DocumentPath == null) return result;
        var docPath = result.DocumentPath;
        var common = document.Element(ns + "CommonData");
        foreach (var element in common?.Elements(ns + "PublicRes") ?? Enumerable.Empty<XElement>())
            result.AddTable(package, MedicalResourceScope.Public, MedicalLocationKind.PublicRes, docPath, "CommonData/PublicRes", element.Value);
        foreach (var element in common?.Elements(ns + "DocumentRes") ?? Enumerable.Empty<XElement>())
            result.AddTable(package, MedicalResourceScope.Document, MedicalLocationKind.DocumentRes, docPath, "CommonData/DocumentRes", element.Value);
        foreach (var template in common?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>())
        {
            package.Visit();
            var templatePath = result.Add(MedicalLocationKind.Template, docPath, "CommonData/TemplatePage/@BaseLoc", (string?)template.Attribute("BaseLoc"));
            if (templatePath == null) continue;
            // Template pages carry their own PageRes; their fonts and media are page-scoped resources too.
            var content = package.Xml(templatePath)?.Root;
            if (content == null) { result.Unreadable.Add(templatePath); continue; }
            foreach (var pageRes in content.Elements(ns + "PageRes"))
                result.AddTable(package, MedicalResourceScope.Page, MedicalLocationKind.PageRes, templatePath, "Template/PageRes", pageRes.Value);
        }

        foreach (var page in document.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
        {
            package.Visit();
            var pagePath = result.Add(MedicalLocationKind.Page, docPath, "Pages/Page/@BaseLoc", (string?)page.Attribute("BaseLoc"));
            if (pagePath == null) continue;
            result.PagePaths.Add(pagePath);
            var content = package.Xml(pagePath)?.Root;
            if (content == null) { result.Unreadable.Add(pagePath); continue; }
            foreach (var pageRes in content.Elements(ns + "PageRes"))
                result.AddTable(package, MedicalResourceScope.Page, MedicalLocationKind.PageRes, pagePath, "Page/PageRes", pageRes.Value);
        }

        // Every index declaration counts; a document may declare more than one.
        foreach (var attachments in document.Elements(ns + "Attachments").Select(x => x.Value))
        {
            var path = result.Add(MedicalLocationKind.Attachments, docPath, "Document/Attachments", attachments);
            var list = path == null ? null : package.Xml(path)?.Root;
            if (path != null && list == null) result.Unreadable.Add(path);
            if (list == null) result.MachineReadableUnavailable.Add(path ?? attachments); // attachments it lists are unknown
            foreach (var item in list?.Elements(ns + "Attachment") ?? Enumerable.Empty<XElement>())
            {
                package.Visit();
                // FileLoc may be an element or, as OfdReader and archival preflight accept, an attribute.
                var element = item.Element(ns + "FileLoc")?.Value; var attribute = (string?)item.Attribute("FileLoc");
                if (element != null && attribute != null && element.Trim() != attribute.Trim())
                {
                    result.Locations.Add(new MedicalLocation(MedicalLocationKind.Attachment, path!, "Attachment/FileLoc", "ambiguous FileLoc", null));
                    result.MachineReadableUnavailable.Add("ambiguous FileLoc in " + path);
                    continue;
                }
                var location = element ?? attribute;
                // A location with a URI scheme is an external attachment (7.4(b)); it is not a package entry.
                if (location != null && location.Contains("://")) continue;
                var payload = result.Add(MedicalLocationKind.Attachment, path!, "Attachment/FileLoc", location);
                var format = (string?)item.Attribute("Format");
                if (payload != null) result.InternalAttachments.Add((payload, format));
                else if (MedicalRules.MachineReadableKind(location ?? string.Empty, format) != null) result.MachineReadableUnavailable.Add(location ?? "<missing FileLoc>");
            }
        }

        foreach (var tags in document.Elements(ns + "CustomTags").Select(x => x.Value))
        {
            var path = result.Add(MedicalLocationKind.CustomTags, docPath, "Document/CustomTags", tags);
            var list = path == null ? null : package.Xml(path)?.Root;
            if (path != null && list == null) result.Unreadable.Add(path);
            if (list == null) result.MachineReadableUnavailable.Add(path ?? tags);
            foreach (var item in list?.Elements(ns + "CustomTag") ?? Enumerable.Empty<XElement>())
            {
                package.Visit();
                var declaredPayload = item.Element(ns + "FileLoc")?.Value;
                var payload = result.Add(MedicalLocationKind.CustomTag, path!, "CustomTag/FileLoc", declaredPayload);
                if (payload != null) result.TagPayloads.Add(payload);
                else result.MachineReadableUnavailable.Add(declaredPayload ?? "<missing FileLoc>");
                var schema = item.Element(ns + "SchemaLoc")?.Value;
                if (schema != null) result.Add(MedicalLocationKind.CustomTag, path!, "CustomTag/SchemaLoc", schema);
            }
        }

        foreach (var annotations in document.Elements(ns + "Annotations").Select(x => x.Value))
        {
            var path = result.Add(MedicalLocationKind.Annotations, docPath, "Document/Annotations", annotations);
            if (path == null) { result.AnnotationIndexUnresolved = true; continue; }
            result.AnnotationIndexPaths.Add(path);
            var list = path == null ? null : package.Xml(path)?.Root;
            if (path != null && list == null) result.Unreadable.Add(path);
            foreach (var item in list?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>())
            {
                package.Visit();
                var file = result.Add(MedicalLocationKind.PageAnnotation, path!, "Page/FileLoc", item.Element(ns + "FileLoc")?.Value);
                if (file != null) result.AnnotationPaths.Add(file);
            }
        }

        var extensions = document.Element(ns + "Extensions")?.Value;
        if (extensions != null) result.Add(MedicalLocationKind.Extensions, docPath, "Document/Extensions", extensions);
        return result;
    }

    private void WalkSignatures(MedicalPackage package, string declared)
    {
        SignaturesDeclared = true;
        var path = SignaturesPath = Add(MedicalLocationKind.Signatures, "OFD.xml", "DocBody/Signatures", declared);
        if (path == null) return;
        var list = package.Xml(path)?.Root;
        if (list == null) { Unreadable.Add(path); return; }
        foreach (var item in list.Elements(list.Name.Namespace + "Signature"))
        {
            package.Visit();
            var signature = Add(MedicalLocationKind.Signature, path, "Signature/@BaseLoc", (string?)item.Attribute("BaseLoc"));
            SignatureDeclarations.Add(signature);
            if (signature != null) SignaturePaths.Add(signature);
        }
    }

    private void AddTable(MedicalPackage package, MedicalResourceScope scope, MedicalLocationKind kind, string from, string selector, string declared)
    {
        package.Visit();
        var path = Add(kind, from, selector, declared);
        if (path == null) return;
        var root = package.Xml(path)?.Root;
        if (root == null) { Unreadable.Add(path); return; }
        var slash = path.LastIndexOf('/');
        var directory = slash < 0 ? string.Empty : path.Substring(0, slash);
        string? baseDirectory = directory;
        var baseLoc = (string?)root.Attribute("BaseLoc");
        if (baseLoc != null)
            baseDirectory = MedicalPackage.TryResolveDirectory(path, baseLoc, out var resolvedBase) ? resolvedBase : null;
        var table = new MedicalResourceTable(scope, path, root, baseDirectory);
        Tables.Add(table);
        var ns = root.Name.Namespace;
        foreach (var media in root.Descendants(ns + "MultiMedia"))
        {
            package.Visit();
            var file = media.Element(ns + "MediaFile")?.Value ?? string.Empty;
            var resolved = ResolvePayload(table, file);
            Media.Add(new MedicalMedia((string?)media.Attribute("Type") ?? string.Empty, table, "MultiMedia[@ID='" + (string?)media.Attribute("ID") + "']/MediaFile", file, resolved));
            Locations.Add(new MedicalLocation(MedicalLocationKind.ResourcePayload, path, "MultiMedia/MediaFile", file, resolved));
        }
        foreach (var font in root.Descendants(ns + "Font"))
        {
            var file = font.Element(ns + "FontFile")?.Value;
            if (file == null) continue;
            package.Visit();
            Locations.Add(new MedicalLocation(MedicalLocationKind.ResourcePayload, path, "Font/FontFile", file, ResolvePayload(table, file)));
        }
    }

    internal static string? ResolvePayload(MedicalResourceTable table, string declared)
    {
        if (declared.StartsWith("/", StringComparison.Ordinal))
            return MedicalPackage.TryResolve(table.Path, declared, out var absolute) ? absolute : null;
        if (table.BaseDirectory == null) return null;
        return MedicalPackage.TryResolve(table.BaseDirectory.Length == 0 ? "x" : table.BaseDirectory + "/x", declared, out var resolved) ? resolved : null;
    }

    private string? Add(MedicalLocationKind kind, string from, string selector, string? declared)
    {
        var resolved = MedicalPackage.TryResolve(from, declared, out var path) ? path : null;
        Locations.Add(new MedicalLocation(kind, from, selector, declared ?? string.Empty, resolved));
        return resolved;
    }
}
