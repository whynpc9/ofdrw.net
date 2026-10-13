using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// This certificate reads only control envelopes and already enumerated direct-child
// descriptors. It never emits graphics or descends into Raw to recover text.
internal sealed class OfdKeywordSourceAdmission(OfdDocumentPackage package, OfdKeywordSourceSnapshot snapshot,
    OfdKeywordLocatorBudget budget)
{
    private readonly Dictionary<string, XElement> xml = new(StringComparer.Ordinal);
    private readonly Dictionary<string, FontDefinition> fonts = new(StringComparer.Ordinal);
    private readonly Dictionary<int, string> pageEntries = new();
    private readonly HashSet<string> resourceEntries = new(StringComparer.Ordinal);
    private readonly Dictionary<string, OfdFontResource?> modelFonts = new(StringComparer.Ordinal);
    private double[]? documentBox;
    private XNamespace ns = XNamespace.None;
    private bool generated;

    internal static bool KnownNamespace(XNamespace value) => value.NamespaceName is OfdConstants.Namespace or OfdConstants.StandardNamespace;
    internal void PrepareDocument()
    {
        foreach (var font in snapshot.Fonts)
        {
            budget.Reference();
            if (string.IsNullOrWhiteSpace(font.Id)) throw new InvalidDataException("Invalid source font identity.");
            if (modelFonts.ContainsKey(font.Id)) modelFonts[font.Id] = null;
            else modelFonts.Add(font.Id, font);
        }
        generated = package.DocumentEntryPath is null;
        if (generated)
        {
            if (package.PreservedEntries.Count != 0 || package.PreservedCommonDataElements.Count != 0 ||
                package.PreservedDocumentElements.Count != 0 || package.PreservedDocBodyElements.Count != 0 ||
                package.Pages.Any(p => p.SourceEntryPath is not null))
                throw new NotSupportedException("UnprovedGeneratedModelAuthority");
            return;
        }
        var ofd = Control("OFD.xml");
        ns = ofd.Name.Namespace;
        if (ofd.Name != ns + "OFD" || !KnownNamespace(ns)) throw new NotSupportedException("UnknownDocumentEnvelope");
        OfdStrictTextGeometry.Attributes(ofd, "Version", "DocType");
        var bodies = ofd.Elements(ns + "DocBody").ToArray();
        if (bodies.Length != 1 || ofd.Elements().Any(e => e.Name != ns + "DocBody"))
            throw new NotSupportedException("MultipleOrUnknownDocumentScope");
        var roots = bodies[0].Elements(ns + "DocRoot").ToArray();
        if (roots.Length != 1 || Resolve("OFD.xml", roots[0].Value) != package.DocumentEntryPath)
            throw new NotSupportedException("UnprovedDocumentSourceBinding");
        var document = Control(package.DocumentEntryPath!);
        if (document.Name != ns + "Document") throw new NotSupportedException("UnknownDocumentEnvelope");
        OfdStrictTextGeometry.Attributes(document);
        var common = One(document, "CommonData", required: true)!;
        OfdStrictTextGeometry.Attributes(common);
        foreach (var child in common.Elements())
        {
            budget.Check();
            if (child.Name.Namespace != ns || child.Name.LocalName is not ("MaxUnitID" or "PageArea" or "PublicRes" or "DocumentRes" or "TemplatePage"))
                throw new NotSupportedException("UnknownInheritedDocumentState");
        }
        if (One(common, "PageArea", false) is { } area) documentBox = Box(area);
        foreach (var name in new[] { "PublicRes", "DocumentRes" })
        {
            if (One(common, name, false) is { } declaration)
            {
                OfdStrictTextGeometry.Attributes(declaration);
                if (declaration.HasElements) throw new NotSupportedException("UnknownResourceDeclaration");
                ReadResource(Resolve(package.DocumentEntryPath!, declaration.Value));
            }
        }
        var pages = One(document, "Pages", true)!; OfdStrictTextGeometry.Attributes(pages);
        var refs = pages.Elements().ToArray();
        if (refs.Length != package.Pages.Count) throw new NotSupportedException("PageSourceOwnershipMismatch");
        var unique = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 0; i < refs.Length; i++)
        {
            budget.Check(); var reference = refs[i];
            if (reference.Name != ns + "Page" || reference.HasElements) throw new NotSupportedException("UnknownPageReference");
            OfdStrictTextGeometry.Attributes(reference, "ID", "BaseLoc");
            var entry = Resolve(package.DocumentEntryPath!, OfdStrictTextGeometry.Required(reference, "BaseLoc"));
            if (!unique.Add(entry) || entry != package.Pages[i].SourceEntryPath || reference.Attribute("ID")?.Value != package.Pages[i].Id)
                throw new NotSupportedException("PageSourceOwnershipMismatch");
            pageEntries.Add(i, entry);
            // Reader's fixed PageRes lookup is audited across every page: later
            // duplicate IDs were discarded by Reader and must not escape here.
            ReadOptionalFixedResource(entry);
        }
        foreach (var template in common.Elements(ns + "TemplatePage"))
        {
            OfdStrictTextGeometry.Attributes(template, "ID", "BaseLoc", "ZOrder");
            ReadOptionalFixedResource(Resolve(package.DocumentEntryPath!, OfdStrictTextGeometry.Required(template, "BaseLoc")));
        }
    }

    internal Dictionary<int, (int Layer, string Path)> Page(OfdPage page, int position)
    {
        budget.Check();
        if (new[] { page.XMillimeters, page.YMillimeters, page.WidthMillimeters, page.HeightMillimeters }.Any(v => !OfdGlyphBox.Finite(v)) ||
            page.WidthMillimeters <= 0 || page.HeightMillimeters <= 0) throw new InvalidDataException("Invalid finite page PhysicalBox.");
        var paths = new Dictionary<int, (int, string)>();
        if (generated)
        {
            if (page.SourceEntryPath is not null || page.PreservedPageElements.Count != 0)
                throw new NotSupportedException("UnprovedGeneratedPageEnvelope");
            var layer = -1; string? lastId = null; string? lastType = null;
            for (var i = 0; i < page.Elements.Count; i++)
            {
                budget.Check(); var element = page.Elements[i];
                if (i == 0 || element.LayerId != lastId || element.LayerType != lastType) layer++;
                lastId = element.LayerId; lastType = element.LayerType;
                paths.Add(i, (layer, "model:Pages[" + position + "]/model:Elements[" + i + "]"));
            }
            return paths;
        }
        if (!pageEntries.TryGetValue(position, out var entry)) throw new NotSupportedException("MissingPageSourceEnvelope");
        var root = Control(entry);
        if (root.Name != ns + "Page") throw new NotSupportedException("UnknownPageEnvelope");
        OfdStrictTextGeometry.Attributes(root); OfdStrictTextGeometry.PlainTextNodes(root, false);
        foreach (var child in root.Elements())
        {
            budget.Check();
            if (child.Name.Namespace != ns || child.Name.LocalName is not ("Area" or "Content" or "Template"))
                throw new NotSupportedException("UnknownPageEnvelope");
            if (child.Name == ns + "Template")
            {
                OfdStrictTextGeometry.Attributes(child, "TemplateID", "ZOrder");
                if (child.HasElements) throw new NotSupportedException("UnknownTemplateEnvelope");
            }
        }
        foreach (var preserved in page.PreservedPageElements)
        {
            var element = budget.Parse(preserved);
            if (element.Name != ns + "Template") throw new NotSupportedException("UnknownPreservedPageContent");
            OfdStrictTextGeometry.Attributes(element, "TemplateID", "ZOrder");
        }
        var ownArea = One(root, "Area", false);
        if (ownArea is null && root.Elements(ns + "Template").Any())
            throw new NotSupportedException("UnprovedTemplatePageAreaInheritance");
        var box = ownArea is null ? documentBox ?? throw new NotSupportedException("MissingPagePhysicalBox") : Box(ownArea);
        if (!box.SequenceEqual(new[] { page.XMillimeters, page.YMillimeters, page.WidthMillimeters, page.HeightMillimeters }))
            throw new NotSupportedException("SourcePageBoxConflict");
        var content = One(root, "Content", false);
        if (content is null)
        {
            if (page.Elements.Count != 0) throw new NotSupportedException("SourceElementOrderMismatch");
            return paths;
        }
        OfdStrictTextGeometry.Attributes(content); OfdStrictTextGeometry.PlainTextNodes(content, false);
        var index = 0; var layerOrdinal = 0;
        foreach (var layer in content.Elements())
        {
            budget.Check();
            if (layer.Name != ns + "Layer") throw new NotSupportedException("UnknownContentContainer");
            OfdStrictTextGeometry.Attributes(layer, "ID", "Type"); OfdStrictTextGeometry.PlainTextNodes(layer, false);
            var layerType = layer.Attribute("Type")?.Value ?? "Body";
            if (layerType is not ("Body" or "Background" or "Foreground")) throw new NotSupportedException("UnsupportedLayerType");
            var objectOrdinal = 0;
            foreach (var child in layer.Elements())
            {
                budget.Check();
                if (index >= page.Elements.Count) throw new NotSupportedException("SourceElementOrderMismatch");
                var model = page.Elements[index];
                var expected = model switch { OfdTextElement t => t.SourceXml, OfdPathElement p => p.SourceXml,
                    OfdImageElement image => image.SourceXml, OfdRawElement raw => raw.Xml, _ => null };
                // Standalone serialization is exactly what Reader retained. We do
                // not enumerate this object's descendants or create another model.
                if (expected is null || !MatchesOriginalElement(child, expected) ||
                    model.LayerId != layer.Attribute("ID")?.Value || model.LayerType != layerType || child.Name.Namespace != ns)
                    throw new NotSupportedException("SourceElementOrderMismatch");
                var path = "/" + root.Name + "[0]/" + content.Name + "[0]/" + layer.Name + "[" + layerOrdinal + "]/" + child.Name + "[" + objectOrdinal + "]";
                paths.Add(index++, (layerOrdinal, path)); objectOrdinal++;
            }
            layerOrdinal++;
        }
        if (index != page.Elements.Count) throw new NotSupportedException("SourceElementOrderMismatch");
        return paths;
    }

    internal (OfdFontResource Resource, string? Entry) ResolveFont(string id)
    {
        if (!modelFonts.TryGetValue(id, out var source) || source is null) throw new NotSupportedException("MissingOrAmbiguousExactFont");
        if (generated) return (source, null);
        if (!fonts.TryGetValue(id, out var definition) || definition.File is null)
            throw new NotSupportedException("MissingExactOriginalFontFile");
        if (source.Bold != definition.Bold || source.Italic != definition.Italic || source.FileName != definition.FileName)
            throw new NotSupportedException("SourceFontBindingConflict");
        return (source, definition.File);
    }
    // Reader drops element-only indentation; Writer emits it by default. Only
    // the admitted TextObject container can ignore ordinary XML
    // whitespace. TextCode values, CDATA, comments and child order stay exact.
    private bool MatchesOriginalElement(XElement original, string expected)
    {
        if (original.Name != ns + "TextObject")
            return string.Equals(original.ToString(SaveOptions.DisableFormatting), expected, StringComparison.Ordinal);
        var projected = budget.Parse(expected);
        CheckTextShape(original); CheckTextShape(projected);
        return SameElement(original, projected);
    }

    private void CheckTextShape(XElement root)
    {
        budget.Field();
        if (root.Name != ns + "TextObject") throw new NotSupportedException("UnsupportedTextNamespace");
        OfdStrictTextGeometry.Attributes(root, "ID", "Name", "Visible", "Boundary", "CTM", "Font", "Size", "Weight", "Italic",
            "Stroke", "Fill", "HScale", "ReadDirection", "CharDirection", "Alpha");
        OfdStrictTextGeometry.PlainTextNodes(root, false);
        foreach (var child in root.Elements())
        {
            budget.Field();
            if (child.Name == ns + "TextCode")
            {
                OfdStrictTextGeometry.Attributes(child, "X", "Y", "DeltaX", "DeltaY");
                OfdStrictTextGeometry.PlainTextNodes(child, true);
            }
            else if (child.Name == ns + "FillColor")
                OfdStrictTextGeometry.Attributes(child, "Value", "Alpha");
            else throw new NotSupportedException("UnsupportedTextChild");
            if (child.HasElements) throw new NotSupportedException("UnsupportedTextChild");
        }
    }

    private bool SameElement(XElement left, XElement right)
    {
        budget.Field();
        if (left.Name != right.Name) return false;
        // This admitted profile has expanded element names and unqualified
        // attributes only, with no QName-valued fields. Redundant declarations
        // inserted by standalone serialization carry no additional semantics.
        using var la = ComparableAttributes(left).GetEnumerator();
        using var ra = ComparableAttributes(right).GetEnumerator();
        while (true)
        {
            budget.Field(); var l = la.MoveNext(); var r = ra.MoveNext();
            if (l != r) return false;
            if (!l) break;
            if (la.Current.Name != ra.Current.Name || !SameValue(la.Current.Value, ra.Current.Value)) return false;
        }
        using var ln = ComparableNodes(left).GetEnumerator();
        using var rn = ComparableNodes(right).GetEnumerator();
        while (true)
        {
            budget.Field(); var l = ln.MoveNext(); var r = rn.MoveNext();
            if (l != r) return false;
            if (!l) return true;
            var a = ln.Current; var b = rn.Current;
            if (a.NodeType != b.NodeType) return false;
            if (a is XElement ae && b is XElement be) { if (!SameElement(ae, be)) return false; }
            else if (a is XText at && b is XText bt) { if (!SameValue(at.Value, bt.Value)) return false; }
            else if (a is XComment ac && b is XComment bc) { if (!SameValue(ac.Value, bc.Value)) return false; }
            else return false;
        }
    }

    private IEnumerable<XAttribute> ComparableAttributes(XElement element)
    {
        foreach (var attribute in element.Attributes())
        {
            budget.Field();
            if (!attribute.IsNamespaceDeclaration) yield return attribute;
        }
    }

    private IEnumerable<XNode> ComparableNodes(XElement parent)
    {
        var container = parent.Name == ns + "TextObject";
        foreach (var node in parent.Nodes())
        {
            budget.Field();
            // XCData derives from XText: explicitly retain it, even when blank.
            if (container && node is XText text && node is not XCData && IsIndentation(text.Value)) continue;
            yield return node;
        }
    }

    private bool IsIndentation(string value)
    {
        budget.Characters(value.Length);
        for (var i = 0; i < value.Length; i++)
        {
            if ((i & 255) == 0) budget.Check();
            if (value[i] is not (' ' or '\t' or '\r' or '\n')) return false;
        }
        return true;
    }

    private bool SameValue(string left, string right)
    {
        budget.Characters(left.Length); budget.Characters(right.Length);
        return string.Equals(left, right, StringComparison.Ordinal);
    }

    private void ReadOptionalFixedResource(string pagePath)
    {
        var split = pagePath.LastIndexOf('/'); var path = (split < 0 ? "" : pagePath.Substring(0, split + 1)) + "PageRes.xml";
        if (snapshot.HasExactEntry(path)) ReadResource(path);
    }
    private void ReadResource(string path)
    {
        if (!resourceEntries.Add(path)) return;
        var root = Control(path);
        if (root.Name != ns + "Res") throw new NotSupportedException("UnknownFontResourceEnvelope");
        OfdStrictTextGeometry.Attributes(root, "BaseLoc");
        foreach (var child in root.Elements())
            if (child.Name.Namespace != ns || child.Name.LocalName is not ("Fonts" or "MultiMedias" or "DrawParams" or "ColorSpaces"))
                throw new NotSupportedException("UnknownFontResourceEnvelope");
        // Legacy Reader.ReadFonts recognizes every descendant LocalName=Font,
        // including foreign lookalikes inside nonfont containers. Certify that
        // none can pollute or duplicate the direct declarations audited below.
        foreach (var node in root.Descendants())
        {
            budget.Reference();
            if (node.Name.LocalName == "Font" && (node.Name != ns + "Font" || node.Parent?.Name != ns + "Fonts" ||
                !ReferenceEquals(node.Parent.Parent, root)))
                throw new NotSupportedException("UnexpectedFontDeclarationLocation");
        }
        if (root.Elements(ns + "Fonts").Count() > 1) throw new NotSupportedException("AmbiguousFontResourceEnvelope");
        if (root.Element(ns + "Fonts") is not { } collection) return;
        OfdStrictTextGeometry.Attributes(collection);
        foreach (var font in collection.Elements())
        {
            budget.Check();
            if (font.Name != ns + "Font") throw new NotSupportedException("UnknownFontDeclaration");
            OfdStrictTextGeometry.Attributes(font, "ID", "FontName", "FamilyName", "Charset", "Bold", "Italic", "Serif", "FixedWidth");
            if (font.Elements().Any(e => e.Name != ns + "FontFile") || font.Elements(ns + "FontFile").Count() > 1)
                throw new NotSupportedException("UnknownFontDeclaration");
            var id = OfdStrictTextGeometry.Required(font, "ID");
            var fileName = font.Element(ns + "FontFile")?.Value;
            string? file = null;
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                var baseLocation = root.Attribute("BaseLoc")?.Value;
                var reference = string.IsNullOrWhiteSpace(baseLocation) || fileName!.StartsWith("/", StringComparison.Ordinal)
                    ? fileName! : baseLocation!.TrimEnd('/') + "/" + fileName;
                file = Resolve(path, reference);
            }
            if (fonts.ContainsKey(id)) throw new NotSupportedException("DuplicateOriginalFontId");
            fonts.Add(id, new FontDefinition(fileName, file, OfdStrictTextGeometry.Boolean(font.Attribute("Bold")?.Value, false),
                OfdStrictTextGeometry.Boolean(font.Attribute("Italic")?.Value, false)));
        }
    }
    private XElement Control(string path)
    {
        if (xml.TryGetValue(path, out var cached)) return cached;
        var bytes = snapshot.ControlEntry(path, budget);
        string value;
        try { value = new UTF8Encoding(false, true).GetString(bytes); }
        catch (DecoderFallbackException exception) { throw new InvalidDataException("Control entry must be valid UTF8.", exception); }
        if (value.Length > 0 && value[0] == '\ufeff') value = value.Substring(1);
        var root = budget.Parse(value); xml.Add(path, root); return root;
    }
    private XElement? One(XElement parent, string name, bool required)
    {
        var children = parent.Elements(ns + name).ToArray();
        if (children.Length > 1 || required && children.Length != 1) throw new NotSupportedException("AmbiguousSourceContainer");
        return children.Length == 0 ? null : children[0];
    }
    private double[] Box(XElement area)
    {
        OfdStrictTextGeometry.Attributes(area); OfdStrictTextGeometry.PlainTextNodes(area, false);
        var boxes = area.Elements().ToArray();
        if (boxes.Length != 1 || boxes[0].Name != ns + "PhysicalBox") throw new NotSupportedException("UnsupportedPageArea");
        OfdStrictTextGeometry.Attributes(boxes[0]);
        if (boxes[0].HasElements) throw new NotSupportedException("UnknownPhysicalBox");
        var box = OfdStrictTextGeometry.Numbers(boxes[0].Value, 4);
        if (box[2] <= 0 || box[3] <= 0) throw new InvalidDataException("Page PhysicalBox must be positive.");
        return box;
    }
    private string Resolve(string relativeTo, string reference)
    {
        budget.Reference();
        if (reference.Length > budget.Limits.MaxSingleXmlBytes) throw new InvalidDataException("Source reference capacity exceeded.");
        for (var i = 0; i < reference.Length; i++) { if ((i & 255) == 0) budget.Check(); if (reference[i] is '/' or '\\') budget.Reference(); }
        return OfdPackagePath.Resolve(relativeTo, reference);
    }
    private sealed class FontDefinition(string? name, string? file, bool bold, bool italic)
    {
        internal readonly string? FileName = name, File = file;
        internal readonly bool Bold = bold, Italic = italic;
    }
}
