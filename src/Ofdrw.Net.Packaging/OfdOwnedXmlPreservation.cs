using System;
using System.IO;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static class OfdOwnedXmlPreservation
{
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
        long position = 0;
        if (!OrderedEqual(independentlyExpected, actual, material, token, ref position, out var reason))
        {
            material.ChargeCopy(checked(2048L + 2L * entry.Length), token);
            throw new InvalidDataException($"Owned rewrite differs from its exact source/candidate/plan/pruner permissions: {entry}, node {position}, {reason}.");
        }
    }
    private static void NormalizeStructuralFormatting(XElement root, XNamespace ns,
        OfdModelMaterialization material, CancellationToken token)
    {
        if (root.Name.LocalName is "OFD" or "Document" or "Page" or "Res") Visit(root, false);
        void Visit(XElement element, bool preserve)
        {
            material.ChargeWork(64, token);
            preserve |= element.Attribute(XNamespace.Xml + "space")?.Value == "preserve";
            for (var node = element.FirstNode; node is not null;)
            {
                var next = node.NextNode;
                material.ChargeWork(32, token);
                // XCData derives from XText. Its whitespace remains exact, as
                // does ordinary whitespace anywhere behind an opaque boundary.
                if (!preserve && node is XText text && node is not XCData)
                {
                    material.ChargeWork(checked(16L + 2L * text.Value.Length), token);
                    if (string.IsNullOrWhiteSpace(text.Value)) text.Remove();
                }
                else if (node is XElement child && child.Name.Namespace == ns &&
                    IsStructuralChild(element, child.Name.LocalName)) Visit(child, preserve);
                node = next;
            }
        }
    }
    private static bool IsStructuralChild(XElement parent, string child) => parent.Name.LocalName switch
    {
        "OFD" => child == "DocBody",
        "DocBody" => child == "DocInfo",
        "Document" => child is "CommonData" or "Pages",
        "CommonData" => child == "PageArea",
        "Pages" => child == "Page",
        "Page" => parent.Parent is null && (child is "Area" or "Content"),
        "Content" => child == "Layer",
        "Res" => child is "Fonts" or "MultiMedias",
        "Fonts" => child == "Font",
        "MultiMedias" => child == "MultiMedia",
        _ => false
    };
    internal static void RemoveReifiedInheritedNamespaces(XElement importedLiteral, XElement originalRoot,
        OfdModelMaterialization material, CancellationToken token)
    {
        // An independently parsed SourceXml fragment reifies inherited xmlns
        // declarations. Remove only declarations absent on the actual source
        // root and exactly bound by that root's ancestors. Original declarations
        // and newly introduced namespace bindings remain subject to exact proof.
        for (var attribute = importedLiteral.FirstAttribute; attribute is not null;)
        {
            var next = attribute.NextAttribute;
            material.ChargeWork(checked(32L + 2L * (attribute.Value.Length + (long)attribute.Name.LocalName.Length)), token);
            if (attribute.IsNamespaceDeclaration)
            {
                string? inherited = null; var originallyDeclared = false;
                for (var ancestor = originalRoot; ancestor is not null; ancestor = ancestor.Parent)
                {
                    material.ChargeWork(64, token);
                    foreach (var declaration in ancestor.Attributes())
                    {
                        material.ChargeWork(checked(32L + 2L * (declaration.Value.Length + (long)declaration.Name.LocalName.Length)), token);
                        if (declaration.Name != attribute.Name) continue;
                        originallyDeclared = ReferenceEquals(ancestor, originalRoot);
                        inherited = declaration.Value;
                        break;
                    }
                    if (inherited is not null) break;
                }
                if (!originallyDeclared && inherited == attribute.Value) attribute.Remove();
            }
            attribute = next;
        }
    }
    private static bool OrderedEqual(XNode expected, XNode actual,
        OfdModelMaterialization material, CancellationToken token, ref long position, out string reason)
    {
        position++; reason = "node kind";
        material.ChargeWork(64, token); material.ChargeCopy(128, token);
        if (expected.NodeType != actual.NodeType) return false;
        if (expected is XElement source && actual is XElement output)
        {
            reason = "element name";
            if (source.Name != output.Name) return false;
            using var left = source.Attributes().GetEnumerator(); using var right = output.Attributes().GetEnumerator();
            while (true)
            {
                var a = left.MoveNext(); var b = right.MoveNext();
                reason = "attribute count";
                if (a != b) return false; if (!a) break;
                material.ChargeWork(checked(32L + 2L * (left.Current.Value.Length + (long)right.Current.Value.Length)), token);
                reason = "ordered attribute";
                if (left.Current.Name != right.Current.Name || left.Current.Value != right.Current.Value) return false;
            }
        }
        if (expected is XText sourceText && actual is XText actualText)
        { reason = "text"; material.ChargeWork(checked(16L + 2L * (sourceText.Value.Length + (long)actualText.Value.Length)), token); return sourceText.Value == actualText.Value; }
        if (expected is XComment sourceComment && actual is XComment actualComment)
        { reason = "comment"; material.ChargeWork(checked(16L + 2L * (sourceComment.Value.Length + (long)actualComment.Value.Length)), token); return sourceComment.Value == actualComment.Value; }
        if (expected is XProcessingInstruction sourcePi && actual is XProcessingInstruction actualPi)
        {
            reason = "processing instruction";
            material.ChargeWork(checked(32L + 2L * (sourcePi.Target.Length + (long)actualPi.Target.Length + sourcePi.Data.Length + actualPi.Data.Length)), token);
            return sourcePi.Target == actualPi.Target && sourcePi.Data == actualPi.Data;
        }
        if (expected is XContainer sourceContainer && actual is XContainer actualContainer)
        {
            using var left = sourceContainer.Nodes().GetEnumerator(); using var right = actualContainer.Nodes().GetEnumerator();
            while (true)
            {
                var a = left.MoveNext(); var b = right.MoveNext();
                reason = "child count";
                if (a != b) return false; if (!a) return true;
                if (!OrderedEqual(left.Current, right.Current, material, token, ref position, out reason)) return false;
            }
        }
        // Controlled parsers reject DTDs; unsupported node kinds are never an
        // inferred permission to rewrite a source representation.
        return false;
    }
}
