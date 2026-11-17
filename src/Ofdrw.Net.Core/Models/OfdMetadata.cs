using System;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdMetadata
{
    private OfdCustomDatas? customDatas;

    /// <summary>
    /// Optional GB/T 33190 <c>DocInfo/DocID</c>, the document's unique identifier. It is unrelated to
    /// <see cref="OfdDocumentOptions.DocumentId"/>, which names the package folder. Reading fills it from a plain-text
    /// <c>DocID</c>; setting null before saving removes that value. Merging documents does not copy it.
    /// </summary>
    public string? DocId { get; set; }

    public string? Title { get; set; }

    public string? Author { get; set; }

    public string? Subject { get; set; }

    public string? Keywords { get; set; }

    public string? Creator { get; set; }

    public DateTimeOffset? CreationDate { get; set; }

    public DateTimeOffset? ModificationDate { get; set; }

    /// <summary>Optional document usage; omission retains the standard's Normal default.</summary>
    public OfdDocUsage? DocUsage { get; set; }

    /// <summary>Package-relative cover image location; the library does not fetch it.</summary>
    public string? Cover { get; set; }

    /// <summary>Version string of the application that created this document.</summary>
    public string? CreatorVersion { get; set; }

    /// <summary>Ordered custom metadata. Null preserves an imported source; withdrawing edits fails.</summary>
    public OfdCustomDatas? CustomDatas
    {
        get => customDatas;
        set
        {
            if (value is null && customDatas?.IsDirty == true)
                throw new InvalidOperationException("Cannot withdraw changed custom metadata.");
            customDatas = value;
        }
    }
}
