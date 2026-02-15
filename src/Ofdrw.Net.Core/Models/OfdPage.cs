using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

public sealed class OfdPage
{
    public int Index { get; set; }

    public double WidthMillimeters { get; set; }

    public double HeightMillimeters { get; set; }

    public List<OfdElement> Elements { get; } = [];
}
