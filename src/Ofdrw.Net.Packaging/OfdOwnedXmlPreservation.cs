using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static class OfdOwnedXmlPreservation
{
    private static readonly HashSet<string> structuralParents = new(StringComparer.Ordinal)
    { "OFD", "DocBody", "DocInfo", "Document", "CommonData", "PageArea", "Pages", "Page", "Content", "Layer", "Res", "Fonts", "Font", "MultiMedias", "MultiMedia" };

    // Expected is the actual source DOM with only explicit candidate/plan/pruner
    // permissions applied by the private Writer. No global name/attribute allowlist
    // grants permission to the actual output. All opaque nodes and order remain.
    internal static void RequireExactClosedRewrite(string entry, XDocument independentlyExpected,
        byte[] actualRewrite, OfdModelMaterialization material, CancellationToken token)
    {
        var actual = material.ReadXml(entry, actualRewrite, token);
        if (independentlyExpected.Root is null || actual.Root is null || independentlyExpected.Root.Name != actual.Root.Name)
            throw new InvalidDataException("Owned rewrite changed the source root authority.");
        NormalizeStructuralFormatting(independentlyExpected.Root, independentlyExpected.Root.Name.Namespace, material, token);
        NormalizeStructuralFormatting(actual.Root, independentlyExpected.Root.Name.Namespace, material, token);
        if (!OrderedEqual(independentlyExpected, actual, material, token))
            throw new InvalidDataException("Owned rewrite differs from its exact source/candidate/plan/pruner permissions.");
    }
    private static void NormalizeStructuralFormatting(XElement root, XNamespace ns,
        OfdModelMaterialization material, CancellationToken token)
    {
        foreach (var element in root.DescendantsAndSelf())
        {
            material.ChargeWork(64, token);
            if (element.Name.Namespace != ns || !structuralParents.Contains(element.Name.LocalName) ||
                element.AncestorsAndSelf().Any(node => node.Attribute(XNamespace.Xml + "space")?.Value == "preserve")) continue;
            material.ChargeCopy(checked(256L + 32L * element.Nodes().Count()), token);
            foreach (var text in element.Nodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).ToArray())
                text.Remove();
        }
    }
    private static bool OrderedEqual(XNode expected, XNode actual,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeWork(64, token); material.ChargeCopy(128, token);
        if (expected.NodeType != actual.NodeType) return false;
        if (expected is XElement source && actual is XElement output)
        {
            if (source.Name != output.Name) return false;
            using var left = source.Attributes().GetEnumerator(); using var right = output.Attributes().GetEnumerator();
            while (true)
            {
                var a = left.MoveNext(); var b = right.MoveNext();
                if (a != b) return false; if (!a) break;
                material.ChargeWork(checked(32L + 2L * (left.Current.Value.Length + (long)right.Current.Value.Length)), token);
                if (left.Current.Name != right.Current.Name || left.Current.Value != right.Current.Value) return false;
            }
        }
        if (expected is XText sourceText && actual is XText actualText)
        { material.ChargeWork(checked(16L + 2L * (sourceText.Value.Length + (long)actualText.Value.Length)), token); return sourceText.Value == actualText.Value; }
        if (expected is XComment sourceComment && actual is XComment actualComment)
        { material.ChargeWork(checked(16L + 2L * (sourceComment.Value.Length + (long)actualComment.Value.Length)), token); return sourceComment.Value == actualComment.Value; }
        if (expected is XProcessingInstruction sourcePi && actual is XProcessingInstruction actualPi)
        {
            material.ChargeWork(checked(32L + 2L * (sourcePi.Target.Length + (long)actualPi.Target.Length + sourcePi.Data.Length + actualPi.Data.Length)), token);
            return sourcePi.Target == actualPi.Target && sourcePi.Data == actualPi.Data;
        }
        if (expected is XContainer sourceContainer && actual is XContainer actualContainer)
        {
            using var left = sourceContainer.Nodes().GetEnumerator(); using var right = actualContainer.Nodes().GetEnumerator();
            while (true)
            {
                var a = left.MoveNext(); var b = right.MoveNext();
                if (a != b) return false; if (!a) return true;
                if (!OrderedEqual(left.Current, right.Current, material, token)) return false;
            }
        }
        // Controlled parsers reject DTDs; unsupported node kinds are never an
        // inferred permission to rewrite a source representation.
        return false;
    }
}
