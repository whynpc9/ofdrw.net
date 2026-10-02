using System;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdPathStyle
{
    internal static bool EvenOdd(OfdPathElement path) => !string.IsNullOrWhiteSpace(path.SourceXml) &&
        string.Equals(XElement.Parse(path.SourceXml!).Attribute("Rule")?.Value, "Even-Odd", StringComparison.OrdinalIgnoreCase);
}
