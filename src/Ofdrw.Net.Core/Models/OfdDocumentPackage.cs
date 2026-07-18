using System;
using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdDocumentPackage
{
    public OfdDocumentOptions Options { get; set; } = new OfdDocumentOptions();

    public List<OfdPage> Pages { get; } = [];

    public List<OfdAttachment> Attachments { get; } = [];

    public List<OfdFontResource> Fonts { get; } = [];

    public Dictionary<string, string> CustomTags { get; } = new Dictionary<string, string>();

    /// <summary>
    /// Original package entries retained while reading an existing OFD package.
    /// Writers copy these entries first and replace only entries they regenerate,
    /// which prevents unrelated resources and extension files from being discarded.
    /// </summary>
    public Dictionary<string, byte[]> PreservedEntries { get; } =
        new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// CommonData children not owned by the current strongly typed model, such as
    /// template page declarations and default color spaces.
    /// </summary>
    public List<string> PreservedCommonDataElements { get; } = [];

    /// <summary>
    /// Document children not owned by the current strongly typed model, such as
    /// annotations, outlines, permissions, actions, bookmarks and extensions.
    /// </summary>
    public List<string> PreservedDocumentElements { get; } = [];

    /// <summary>
    /// DocBody children other than DocInfo and DocRoot, including version and
    /// signature declarations.
    /// </summary>
    public List<string> PreservedDocBodyElements { get; } = [];

    public string? PublicResourceLocation { get; set; }

    public string? DocumentResourceLocation { get; set; }
}
