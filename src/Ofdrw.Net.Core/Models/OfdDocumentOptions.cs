using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdDocumentOptions
{
    public string DocType { get; set; } = OfdConstants.DefaultDocType;

    public string Namespace { get; set; } = OfdConstants.Namespace;

    public string DocumentId { get; set; } = OfdConstants.DefaultDocId;

    public bool EnableDeflateCompression { get; set; } = true;

    public double DefaultPageWidthMillimeters { get; set; } = 210d;

    public double DefaultPageHeightMillimeters { get; set; } = 297d;

    public OfdMetadata Metadata { get; set; } = new OfdMetadata();
}
