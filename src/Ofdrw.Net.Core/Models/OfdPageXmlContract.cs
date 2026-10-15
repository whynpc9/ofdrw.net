using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Core.Models;

// Containers rebuilt by Writer must be fully represented by OfdPage. Preserved
// top-level page children and raw drawing objects remain archive-owned XML.
internal static class OfdPageXmlContract
{
    internal static bool IsRepresentedArea(XElement area, XNamespace ns) =>
        area.Name.Namespace == ns && area.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
        area.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value)) &&
        area.Elements().Count() == 1 && area.Elements().All(box => box.Name == ns + "PhysicalBox" &&
            box.Attributes().All(attribute => attribute.IsNamespaceDeclaration) && !box.HasElements &&
            box.Nodes().All(node => node is XText or XComment) && OfdBoxParser.TryParse(box.Value, out var bounds) &&
            OfdBoxParser.IsWritablePositiveSide(bounds.w) && OfdBoxParser.IsWritablePositiveSide(bounds.h));

    internal static void ValidateDocumentArea(OfdDocumentPackage package, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var path = package.DocumentEntryPath ?? package.Options.DocumentId + "/Document.xml";
        if (!package.PreservedEntries.TryGetValue(path, out var data)) return;
        using var input = new MemoryStream(data, false);
        var xml = XDocument.Load(input, LoadOptions.PreserveWhitespace); var ns = xml.Root!.Name.Namespace;
        var areas = xml.Root.Element(ns + "CommonData")?.Elements().Where(node => node.Name.LocalName == "PageArea").ToList();
        if (areas is not null && (areas.Count > 1 || areas.Any(area => !IsRepresentedArea(area, ns))))
            throw new NotSupportedException("Document PageArea cannot preserve unmodeled boxes or malformed PhysicalBox values.");
    }

    internal static void ValidateForRewrite(OfdDocumentPackage package, OfdPage page, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        ValidateWritableDimensions(page.WidthMillimeters, page.HeightMillimeters);
        if (page.SourceEntryPath is null || !package.PreservedEntries.TryGetValue(page.SourceEntryPath, out var data)) return;
        using var input = new MemoryStream(data, false);
        var xml = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        var root = xml.Root;
        bool Plain(XElement node, params string[] attributes) => node.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attributes.Contains(attribute.Name.LocalName)) &&
            node.Nodes().All(child => child is XElement or XComment || child is XText text && string.IsNullOrWhiteSpace(text.Value));
        bool valid = root is not null && root.Name.LocalName == "Page" &&
            (root.Name.NamespaceName == package.Options.Namespace || root.Name.NamespaceName == OfdConstants.Namespace || root.Name.NamespaceName == OfdConstants.StandardNamespace) &&
            Plain(root) && xml.Nodes().All(node => node is XElement or XComment || node is XText text && string.IsNullOrWhiteSpace(text.Value));
        if (valid)
        {
            var ns = root!.Name.Namespace;
            valid = root.Elements(ns + "Area").Count() <= 1 && root.Elements(ns + "Content").Count() <= 1;
            foreach (var child in root.Elements())
            {
                token.ThrowIfCancellationRequested();
                // Reader consumes these by local name instead of preserving them.
                if (child.Name.LocalName is "Area" or "Content" && child.Name.Namespace != ns) valid = false;
                if (child.Name == ns + "Area")
                {
                    valid &= IsRepresentedArea(child, ns);
                }
                if (child.Name == ns + "Content")
                {
                    valid &= Plain(child);
                    foreach (var layer in child.Elements()) valid &= layer.Name == ns + "Layer" && Plain(layer, "ID", "Type");
                }
            }
        }
        if (!valid) throw new NotSupportedException("Page rewrite cannot preserve unmodeled Page/Area/Content/Layer containers.");
    }

    internal static void ValidateWritableDimensions(double width, double height)
    {
        // Preserve the existing unspecified/nonpositive model-side behavior,
        // but never let a positive size collapse to zero at Writer precision.
        if (double.IsNaN(width) || double.IsNaN(height) || double.IsInfinity(width) || double.IsInfinity(height) ||
            width > 0 && !OfdBoxParser.IsWritablePositiveSide(width) || height > 0 && !OfdBoxParser.IsWritablePositiveSide(height))
            throw new NotSupportedException("Page dimensions cannot be represented at OFD writer precision.");
    }
}
