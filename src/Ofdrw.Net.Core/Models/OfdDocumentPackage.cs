using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdDocumentPackage
{
    public OfdDocumentOptions Options { get; set; } = new OfdDocumentOptions();

    public List<OfdPage> Pages { get; } = [];

    public List<OfdAttachment> Attachments { get; } = [];

    public Dictionary<string, string> CustomTags { get; } = new Dictionary<string, string>();
}
