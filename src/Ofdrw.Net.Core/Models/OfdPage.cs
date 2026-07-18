using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdPage
{
    public string? Id { get; set; }

    public int Index { get; set; }

    public double XMillimeters { get; set; }

    public double YMillimeters { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }

    public List<OfdElement> Elements { get; } = [];

    public List<OfdTemplateContent> Templates { get; } = [];

    /// <summary>
    /// Page-level XML children not yet represented by a strongly typed model,
    /// such as template references and page actions.
    /// </summary>
    public List<string> PreservedPageElements { get; } = [];
}
