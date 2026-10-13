using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Shared source-envelope operations; no second OFD object parser. Inherited
// prefixes matter even when used only inside unknown QName-valued content.
internal static class OfdXmlEnvelope
{
    internal static string Capture(XElement source, OfdModelMaterialization budget, CancellationToken token)
    {
        foreach (var node in source.DescendantNodesAndSelf()) { token.ThrowIfCancellationRequested(); budget.Structure.Node(1); }
        var bindings = new Dictionary<XName, string>();
        foreach (var ancestor in source.AncestorsAndSelf())
            foreach (var attribute in ancestor.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                token.ThrowIfCancellationRequested(); budget.Structure.Reference();
                if (!bindings.ContainsKey(attribute.Name)) bindings.Add(attribute.Name, attribute.Value);
            }
        var copy = new XElement(source);
        foreach (var binding in bindings) if (copy.Attribute(binding.Key) is null) copy.SetAttributeValue(binding.Key, binding.Value);
        using (var counter = new CountingWriter(budget, token))
        using (var writer = XmlWriter.Create(counter, new XmlWriterSettings { OmitXmlDeclaration = true, Indent = false, ConformanceLevel = ConformanceLevel.Fragment }))
            copy.WriteTo(writer);
        token.ThrowIfCancellationRequested();
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    internal static XElement Parse(string xml)
    {
        using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = new OfdModelLimits().MaxSourceXmlBytes
        });
        return XElement.Load(reader, LoadOptions.PreserveWhitespace);
    }

    internal static XElement Template(OfdNavigationSource? source, string localName, XNamespace ns)
    {
        if (source?.Xml is null) return new XElement(ns + localName);
        var element = Parse(source.Xml);
        if (element.Name != ns + localName) throw new NotSupportedException("Source envelope cannot change namespace or owned node kind.");
        return element;
    }

    internal static void Singleton(XElement parent, string localName, XElement? replacement)
    {
        var name = parent.Name.Namespace + localName;
        var originals = parent.Elements(name).ToList();
        if (originals.Count > 1) throw new InvalidDataException("Duplicate owned XML singleton.");
        if (originals.Count == 1)
        {
            if (replacement is null) originals[0].Remove(); else originals[0].ReplaceWith(replacement);
        }
        else if (replacement is not null)
        {
            // Modeled document order; existing unknown siblings retain their
            // relative positions. Outline actions precede nested OutlineElem.
            var next = parent.Elements().FirstOrDefault(child => child.Name.Namespace == parent.Name.Namespace && Order(parent.Name.LocalName, child.Name.LocalName) > Order(parent.Name.LocalName, localName));
            if (next is null) parent.Add(replacement); else next.AddBeforeSelf(replacement);
        }
    }

    internal static void Children<T>(XElement parent, OfdNavigationSource? source, IReadOnlyList<T> current,
        Func<T, XElement> write, CancellationToken token) where T : class
    {
        var original = parent.Elements().ToList();
        var slots = source?.Slots ?? new Dictionary<int, object>();
        var reordered = ValidateChildrenOrder(parent, source, current, token);
        if (reordered)
        {
            // Reordering is safe only after the opaque-anchor check above.
            var index = 0; XNode? anchor = null;
            for (var ordinal = 0; ordinal < original.Count; ordinal++)
            {
                token.ThrowIfCancellationRequested(); if (!slots.ContainsKey(ordinal)) continue;
                if (index < current.Count) { var replacement = write(current[index++]); original[ordinal].ReplaceWith(replacement); anchor = replacement; }
                else original[ordinal].Remove();
            }
            while (index < current.Count) { var added = write(current[index++]); if (anchor is null) parent.Add(added); else anchor.AddAfterSelf(added); anchor = added; }
            return;
        }
        var retained = new HashSet<object>(current.Cast<object>(), ReferenceComparer.Instance);
        // Surviving source identities keep their exact source slots. New items
        // are inserted before the next surviving identity, or at the end.
        var next = 0;
        for (var ordinal = 0; ordinal < original.Count; ordinal++)
        {
            token.ThrowIfCancellationRequested();
            if (!slots.TryGetValue(ordinal, out var identity)) continue;
            if (!retained.Contains(identity)) { original[ordinal].Remove(); continue; }
            while (next < current.Count && !ReferenceEquals(current[next], identity))
                original[ordinal].AddBeforeSelf(write(current[next++]));
            if (next >= current.Count) throw new InvalidDataException("Navigation source-slot order is inconsistent.");
            original[ordinal].ReplaceWith(write(current[next++]));
        }
        while (next < current.Count) { token.ThrowIfCancellationRequested(); parent.Add(write(current[next++])); }

    }

    internal static bool ValidateChildrenOrder<T>(XElement parent, OfdNavigationSource? source, IReadOnlyList<T> current,
        CancellationToken token) where T : class
    {
        var original = parent.Elements().ToList();
        var slots = source?.Slots ?? new Dictionary<int, object>();
        var positions = new Dictionary<object, int>(ReferenceComparer.Instance);
        foreach (var slot in slots) positions[slot.Value] = slot.Key;
        var last = -1; var reordered = false;
        var retained = new HashSet<object>(ReferenceComparer.Instance);
        foreach (var item in current)
        {
            token.ThrowIfCancellationRequested();
            if (!retained.Add(item)) throw new InvalidDataException("An owned node occurs more than once in a collection.");
            if (positions.TryGetValue(item, out var position)) { reordered |= position < last; last = position; }
        }
        if (reordered)
        {
            var ordinal = 0;
            foreach (var node in parent.Nodes())
            {
                token.ThrowIfCancellationRequested();
                var opaque = node is XProcessingInstruction || node is XText text && !string.IsNullOrWhiteSpace(text.Value);
                if (node is XElement) opaque |= !slots.ContainsKey(ordinal++);
                if (opaque) throw new NotSupportedException("Cannot reorder owned nodes across opaque mixed-content anchors.");
            }
        }
        return reordered;
    }

    private static int Order(string parent, string name)
    {
        if (parent == "CommonData") return name switch { "MaxUnitID" => 0, "PageArea" => 1, "PublicRes" => 2, "DocumentRes" => 3, "TemplatePage" => 4, "DefaultCS" => 5, _ => int.MaxValue };
        if (parent == "DocBody") return name switch { "DocInfo" => 0, "DocRoot" => 1, "Versions" => 2, "Signatures" => 3, _ => int.MaxValue };
        if (parent == "PageArea") return name switch { "PhysicalBox" => 0, "ApplicationBox" => 1, "ContentBox" => 2, "BleedBox" => 3, _ => int.MaxValue };
        return name switch
    {
        "CommonData" => 0, "Pages" => 1, "Outlines" => 2, "Permissions" => 3,
        "Actions" => 4, "VPreferences" => 5, "Bookmarks" => 6, "Annotations" => 7,
        "CustomTags" => 8, "Attachments" => 9, "OutlineElem" => 10, _ => int.MaxValue
        };
    }

    private sealed class CountingWriter(OfdModelMaterialization budget, CancellationToken token) : TextWriter
    {
        public override Encoding Encoding => Encoding.Unicode;
        public override void Write(char value) => budget.ChargeCopy(sizeof(char), token);
        public override void Write(char[] buffer, int index, int count) => budget.ChargeCopy(2L * count, token);
        public override void Write(string? value) { if (value is not null) budget.ChargeCopy(2L * value.Length, token); }
    }
    private sealed class ReferenceComparer : IEqualityComparer<object>
    {
        internal static readonly ReferenceComparer Instance = new();
        public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
        public int GetHashCode(object value) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value);
    }
}
