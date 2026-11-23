using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Core.Models;

/// <summary>
/// The one catalogue of how GB/T 33190 locates package entries: which Document children exist, which elements and
/// attributes carry package paths, the default document entry and the implicit page resource table. Writers,
/// pruning, attachment cleanup, signing selection, mixing and read-only assessment ask it instead of keeping their
/// own name lists, so a newly modelled reference kind is added once.
/// </summary>
internal static class OfdReferenceCatalog
{
    /// <summary>Children of <c>Document</c> in schema order.</summary>
    internal static readonly IReadOnlyList<string> DocumentChildren = Array.AsReadOnly(new[]
    {
        "CommonData", "Pages", "Outlines", "Permissions", "Actions", "VPreferences", "Bookmarks",
        "Annotations", "CustomTags", "Attachments", "Extensions"
    });

    /// <summary><c>Document</c> children whose text locates an index XML entry (annotations, tags, attachments, extensions).</summary>
    internal static readonly IReadOnlyCollection<string> DocumentIndexDeclarations =
        new HashSet<string>(StringComparer.Ordinal) { "Annotations", "CustomTags", "Attachments", "Extensions" };

    /// <summary>Elements whose text locates control XML: the document, resource tables and index files.</summary>
    internal static readonly IReadOnlyCollection<string> ControlLocationElements = new HashSet<string>(StringComparer.Ordinal)
    {
        "DocRoot", "PublicRes", "DocumentRes", "PageRes", "Signatures", "Attachments", "Annotations", "CustomTags", "Extensions"
    };

    /// <summary>Elements whose text may be a package path (control XML or payload).</summary>
    internal static readonly IReadOnlyCollection<string> ScalarPathElements = new HashSet<string>(StringComparer.Ordinal)
    {
        "SignedValue", "FileRef", "FileLoc", "FontFile", "MediaFile", "DocRoot", "PublicRes", "DocumentRes", "PageRes",
        "File", "BaseLoc", "Cover"
    };

    /// <summary>Attributes that always carry a package path. <c>FileLoc</c> and <c>URI</c> depend on their owner.</summary>
    internal static readonly IReadOnlyCollection<string> PathAttributes =
        new HashSet<string>(StringComparer.Ordinal) { "BaseLoc", "FileRef", "MediaFile" };

    /// <summary>Owners whose <c>BaseLoc</c> attribute locates control XML.</summary>
    internal static readonly IReadOnlyCollection<string> ControlBaseLocOwners =
        new HashSet<string>(StringComparer.Ordinal) { "Page", "TemplatePage", "Version", "Signature" };

    /// <summary>The document entry a package writes to: its source path, else <c>{DocumentId}/Document.xml</c>.</summary>
    internal static string DocumentPath(OfdDocumentPackage package) =>
        package.DocumentEntryPath ?? DefaultDocumentPath(package.Options.DocumentId);

    /// <summary>The conventional document entry for a document folder.</summary>
    internal static string DefaultDocumentPath(string documentId) => documentId + "/Document.xml";

    /// <summary>
    /// The fixed <c>PageRes.xml</c> beside a page or template content entry. Readers treat it as that page's resource
    /// table even when the page has no <c>PageRes</c> link.
    /// </summary>
    internal static string ImplicitPageResPath(string pageContentPath)
    {
        var slash = pageContentPath.LastIndexOf('/');
        return (slash < 0 ? string.Empty : pageContentPath.Substring(0, slash + 1)) + "PageRes.xml";
    }

    internal static bool IsDocumentChild(string localName) => DocumentChildren.Contains(localName);
}
