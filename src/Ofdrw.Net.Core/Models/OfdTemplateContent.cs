using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdTemplateContent
{
    public string TemplateId { get; set; } = string.Empty;

    public string ZOrder { get; set; } = "Background";

    public string? BaseLocation { get; set; }

    public List<OfdElement> Elements { get; } = [];
}
