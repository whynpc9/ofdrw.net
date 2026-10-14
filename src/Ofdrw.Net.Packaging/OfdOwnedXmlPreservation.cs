using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

// Conservative identity profile for the extension surface of actual rewrites.
// Unknown nodes/attributes are compared in place. Unsupported relocation is
// rejected; mere membership in the Writer's output map grants no permission.
internal static class OfdOwnedXmlPreservation
{
    private static readonly HashSet<string> knownNodes = new(StringComparer.Ordinal)
    {
        "OFD", "DocBody", "DocInfo", "DocID", "Title", "Author", "Subject", "Keywords", "Creator", "CreatorVersion",
        "CreationDate", "ModDate", "DocRoot", "Document", "CommonData", "MaxUnitID", "PageArea", "PhysicalBox", "ApplicationBox",
        "ContentBox", "BleedBox", "PublicRes", "DocumentRes", "Pages", "Page", "Content", "Layer", "TextObject", "TextCode",
        "ImageObject", "PathObject", "FillColor", "StrokeColor", "Color", "AbbreviatedData", "Res", "Fonts", "Font", "FontFile",
        "MultiMedias", "MultiMedia", "MediaFile"
    };
    private static readonly HashSet<string> knownAttributes = new(StringComparer.Ordinal)
    {
        "ID", "Version", "DocType", "BaseLoc", "Type", "Font", "FontName", "FamilyName", "Charset", "Bold", "Italic", "Size",
        "Boundary", "CTM", "Weight", "HScale", "ReadDirection", "CharDirection", "X", "Y", "DeltaX", "DeltaY", "Value",
        "ColorSpace", "Alpha", "Rule", "Stroke", "Fill", "LineWidth", "ResourceID", "Format", "MediaFile"
    };
    internal static void RequirePreserved(string entry, byte[] original, byte[] rewritten,
        OfdOwnedPruneDispositions pruning, OfdModelMaterialization material, CancellationToken token)
    {
        var source = material.ReadXml(entry, original, token);
        var output = material.ReadXml(entry, rewritten, token);
        if (source.Root is null || output.Root is null || source.Root.Name != output.Root.Name)
            throw new InvalidDataException("Owned rewrite changed the source root authority.");
        Compare(source.Root, output.Root, source.Root.Name.Namespace, pruning, material, token);
    }
    private static void Compare(XElement source, XElement output, XNamespace ns,
        OfdOwnedPruneDispositions pruning, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeWork(64, token);
        if (source.Name.Namespace != ns || !knownNodes.Contains(source.Name.LocalName))
        {
            ChargeTree(source, material, token); ChargeTree(output, material, token);
            if (!XNode.DeepEquals(source, output)) throw new InvalidDataException("Owned rewrite changed an opaque subtree.");
            return;
        }
        foreach (var attribute in source.Attributes())
        {
            material.ChargeWork(checked(32L + 2L * attribute.Value.Length), token);
            if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.None && knownAttributes.Contains(attribute.Name.LocalName)) continue;
            if (output.Attribute(attribute.Name)?.Value != attribute.Value)
                throw new InvalidDataException("Owned rewrite changed an opaque attribute.");
        }
        material.ChargeCopy(checked(512L + 256L * source.Elements().Count()), token);
        var ordinals = new Dictionary<XName, int>();
        foreach (var child in source.Elements())
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(16, token);
            ordinals.TryGetValue(child.Name, out var ordinal); ordinals[child.Name] = ordinal + 1;
            var match = output.Elements(child.Name).Skip(ordinal).FirstOrDefault();
            if (match is null)
            {
                if (pruning.MatchesActualRemovedSignature(child, token)) continue;
                throw new InvalidDataException("Owned rewrite removed an unpermitted source child.");
            }
            Compare(child, match, ns, pruning, material, token);
        }
        // Unknown output children must also have a source position. Writer cannot
        // synthesize extension authority by adding an opaque subtree.
        foreach (var child in output.Elements().Where(child => child.Name.Namespace != ns || !knownNodes.Contains(child.Name.LocalName)))
        {
            material.ChargeWork(16, token);
            if (source.Elements(child.Name).Count() != output.Elements(child.Name).Count())
                throw new InvalidDataException("Owned rewrite introduced an opaque subtree.");
        }
    }
    private static void ChargeTree(XElement element, OfdModelMaterialization material, CancellationToken token)
    {
        foreach (var node in element.DescendantsAndSelf())
        {
            material.ChargeWork(64, token);
            foreach (var attribute in node.Attributes()) material.ChargeWork(2L * attribute.Value.Length + 32, token);
            foreach (var text in node.Nodes().OfType<XText>()) material.ChargeWork(2L * text.Value.Length + 16, token);
        }
    }
}
