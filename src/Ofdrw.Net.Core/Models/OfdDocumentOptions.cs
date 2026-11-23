using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdDocumentOptions
{
    public string DocType { get; set; } = OfdConstants.DefaultDocType;

    public string Namespace { get; set; } = OfdConstants.Namespace;

    public string DocumentId { get; set; } = OfdConstants.DefaultDocId;

    public bool EnableDeflateCompression { get; set; } = true;

    /// <summary>
    /// Opt-in linearized ZIP entry order: <c>OFD.xml</c>, the document entry, its resource tables, then each page's
    /// entries in page order, followed by the remaining entries. False (the default) keeps the ordinal path order.
    /// Only freshly written packages are affected; it does not reorder entries rewritten by signing.
    /// </summary>
    public bool LinearizedEntryOrder { get; set; }

    /// <summary>Package-wide embedded font subsetting and resource budgets.</summary>
    public OfdFontEmbeddingOptions FontEmbedding { get; set; } = new();

    public double DefaultPageWidthMillimeters { get; set; } = 210d;

    public double DefaultPageHeightMillimeters { get; set; } = 297d;

    public OfdMetadata Metadata { get; set; } = new OfdMetadata();

    /// <summary>
    /// Copies every option, the font-embedding budgets and the scalar metadata. This is the one place a package-level
    /// option is copied: a new option added here is carried by controlled edits, merge/split and converters alike.
    /// <see cref="OfdMetadata.CustomDatas"/> is not copied; it is owned XML whose clone needs the caller's budget.
    /// </summary>
    internal OfdDocumentOptions Copy() => new OfdDocumentOptions
    {
        DocType = DocType, Namespace = Namespace, DocumentId = DocumentId,
        EnableDeflateCompression = EnableDeflateCompression, LinearizedEntryOrder = LinearizedEntryOrder,
        FontEmbedding = FontEmbedding.Copy(),
        DefaultPageWidthMillimeters = DefaultPageWidthMillimeters, DefaultPageHeightMillimeters = DefaultPageHeightMillimeters,
        Metadata = Metadata.CopyScalars()
    };

    /// <summary>
    /// <see cref="Copy"/> for a document assembled from others (merge and split): the source document's identity is
    /// not carried. <see cref="DocumentId"/> returns to the default and the GB/T 33190 <c>DocID</c> is not copied.
    /// </summary>
    internal OfdDocumentOptions CopyForNewDocument()
    {
        var copy = Copy();
        copy.DocumentId = OfdConstants.DefaultDocId;
        copy.Metadata.DocId = null;
        return copy;
    }
}
