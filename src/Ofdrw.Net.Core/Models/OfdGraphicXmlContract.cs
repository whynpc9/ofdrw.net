using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Core.Models;

// The subset whose XML can be copied without an unknown reference-bearing extension.
internal static class OfdGraphicXmlContract
{
    internal static bool HasKnownChildren(XElement root)
    {
        var ns = root.Name.Namespace;
        if (ns != XNamespace.None && ns != OfdConstants.Namespace && ns != OfdConstants.StandardNamespace) return false;
        foreach (var node in root.Descendants())
        {
            if (node.Name.Namespace != ns) return false;
            var child = node.Name.LocalName;
            var allowed = node.Parent!.Name.LocalName switch
            {
                "TextObject" => child is "Clips" or "FillColor" or "StrokeColor" or "CGTransform" or "TextCode",
                "PathObject" => child is "Clips" or "FillColor" or "StrokeColor" or "AbbreviatedData",
                "ImageObject" => child is "Clips",
                "Clips" => child == "Clip",
                "Clip" => child == "Area",
                "Area" => child == "Path",
                "Path" => child == "AbbreviatedData",
                "CGTransform" => child == "Glyphs",
                _ => false
            };
            if (!allowed) return false;
        }
        return true;
    }

    internal static bool IsKnownAttribute(XAttribute attribute)
    {
        if (attribute.IsNamespaceDeclaration) return true;
        var parent = attribute.Parent!.Name.LocalName;
        if (parent == "TextObject" && attribute.Name == OfdTextEmphasis.FauxItalicFactor) return true;
        if (attribute.Name.Namespace != XNamespace.None) return false;
        var name = attribute.Name.LocalName;
        if (parent is "TextObject" or "ImageObject" or "PathObject" or "Path" && name is "ID" or "Name" or "Visible" or "Boundary" or "CTM" or "Alpha" or "LineWidth" or "Cap" or "Join" or "MiterLimit" or "DashOffset" or "DashPattern") return true;
        return parent switch
        {
            "TextObject" => name is "Font" or "Size" or "Stroke" or "Fill" or "HScale" or "ReadDirection" or "CharDirection" or "Weight" or "Italic",
            "ImageObject" => name is "ResourceID" or "Substitution",
            "PathObject" => name is "Stroke" or "Fill" or "Rule" or "LineWidth" or "Cap" or "Join" or "MiterLimit" or "DashOffset" or "DashPattern",
            "FillColor" or "StrokeColor" => name is "Value" or "Alpha",
            "TextCode" => name is "X" or "Y" or "DeltaX" or "DeltaY",
            "CGTransform" => name is "CodePosition" or "CodeCount" or "GlyphCount",
            "Path" => name is "Stroke" or "Fill" or "Rule",
            "Area" => name is "CTM" or "Start",
            _ => false
        };
    }

    internal static bool HasUnsupportedReferences(XElement root) => root.DescendantsAndSelf().Attributes().Any(attribute =>
        !attribute.IsNamespaceDeclaration && attribute.Name.Namespace == XNamespace.None &&
        (attribute.Name.LocalName is "DrawParam" or "ColorSpace" or "RefID" or "ObjectRef" or "TemplateID" or "PageID" or "Substitution" or "ImageMask" ||
         attribute.Name.LocalName == "Font" && attribute.Parent != root ||
         attribute.Name.LocalName == "ResourceID" && attribute.Parent != root));

    internal static bool IsVisible(OfdElement element)
    {
        var xml = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
        if (string.IsNullOrWhiteSpace(xml)) return true;
        return XElement.Parse(xml!).Attribute("Visible")?.Value.ToLowerInvariant() is not ("false" or "0");
    }
}
