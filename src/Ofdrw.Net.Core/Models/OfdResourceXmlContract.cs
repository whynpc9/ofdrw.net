using System;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Preserve ordinary same-namespace resource extensions already supported by
// the Reader. Foreign children and root-level semantics cannot be discarded
// while selecting a known PageBlock resource definition.
internal static class OfdResourceXmlContract
{
    internal static bool IsPlainRoot(XDocument xml) => xml.Root is { } root &&
        root.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attribute.Name.LocalName == "BaseLoc") &&
        root.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
        xml.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
        root.Elements().All(child => child.Name.Namespace == root.Name.Namespace);
}
