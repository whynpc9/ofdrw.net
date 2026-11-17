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
}
