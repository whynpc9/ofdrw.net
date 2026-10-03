using System;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdPathStyle
{
    internal static string? Attribute(OfdPathElement path, string name) => string.IsNullOrWhiteSpace(path.SourceXml)
        ? null : XElement.Parse(path.SourceXml!).Attribute(name)?.Value;
    internal static bool EvenOdd(OfdPathElement path) => !string.IsNullOrWhiteSpace(path.SourceXml) &&
        string.Equals(XElement.Parse(path.SourceXml!).Attribute("Rule")?.Value, "Even-Odd", StringComparison.OrdinalIgnoreCase);
}
