using System;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// The five resource tables are the only Res children admitted for PageBlock
// resolution. Unknown envelope semantics must not be lost while selecting a
// known definition from the same resource file.
internal static class OfdResourceXmlContract
{
    internal static bool IsPlainRoot(XDocument xml) => xml.Root is { } root &&
        root.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attribute.Name.LocalName == "BaseLoc") &&
        root.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
        xml.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
        root.Elements().All(child => child.Name.Namespace == root.Name.Namespace &&
            child.Name.LocalName is "Fonts" or "MultiMedias" or "CompositeGraphicUnits" or "DrawParams" or "ColorSpaces");
}
