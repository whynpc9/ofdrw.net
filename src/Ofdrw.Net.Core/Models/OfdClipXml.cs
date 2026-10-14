using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdClipXml
{
    internal static void Load(OfdElement element, OfdModelMaterialization material, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(element.ClippingXml)) return;
        material.GuardFragmentXml(element.ClippingXml!, token);
        var root = OfdXmlEnvelope.Parse(element.ClippingXml!); var ns = root.Name.Namespace;
        if ((ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.Namespace && ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.StandardNamespace) || root.Name != ns + "Clips" || root.Attribute("TransFlag") is not null) return;
        var result = new OfdClipCollection { SourceXml = element.ClippingXml };
        var slot = 0;
        foreach (var node in root.Elements())
        {
            token.ThrowIfCancellationRequested(); material.ReserveReadModel(1024, token);
            if (node.Name != ns + "Clip") { slot++; continue; }
            var clip = new OfdClip { SourceXml = OfdXmlEnvelope.Capture(node, material, token), SourceSlot = slot++ };
            var areaSlot = 0;
            foreach (var area in node.Elements())
            {
                token.ThrowIfCancellationRequested(); material.ReserveReadModel(2048, token);
                if (area.Name != ns + "Area") { areaSlot++; continue; }
                var shape = area.Element(ns + "Path");
                if (shape is null || area.Elements().Count(n => n.Name.Namespace == ns) != 1 || area.Attribute("DrawParam") is not null ||
                    shape.Element(ns + "Clips") is not null || shape.Attribute("DrawParam") is not null ||
                    shape.Attribute("Stroke")?.Value is not ("false" or "0") || shape.Attribute("Fill")?.Value is not ("true" or "1")) return;
                var box = Numbers(shape.Attribute("Boundary")?.Value); if (box.Length != 0 && box.Length != 4) return;
                if (shape.Element(ns + "AbbreviatedData")?.HasElements == true) return;
                var rule = shape.Attribute("Rule")?.Value;
                if (rule is not null && rule is not ("NonZero" or "Even-Odd")) return;
                var path = new OfdPathElement
                {
                    AbbreviatedData = shape.Element(ns + "AbbreviatedData")?.Value ?? string.Empty,
                    XMillimeters = box.Length == 4 ? box[0] : 0, YMillimeters = box.Length == 4 ? box[1] : 0,
                    WidthMillimeters = box.Length == 4 ? box[2] : 0, HeightMillimeters = box.Length == 4 ? box[3] : 0,
                    Transform = Matrix(shape.Attribute("CTM")?.Value), Stroke = false, Fill = true,
                    FillRule = rule == "Even-Odd" ? OfdFillRule.EvenOdd : OfdFillRule.NonZero,
                    SourceXml = OfdXmlEnvelope.Capture(shape, material, token)
                };
                clip.Areas.Add(new OfdClipArea { Path = path, Transform = Matrix(area.Attribute("CTM")?.Value),
                    SourceXml = OfdXmlEnvelope.Capture(area, material, token), SourceSlot = areaSlot++ });
            }
            if (clip.Areas.Count == 0) return;
            result.Items.Add(clip);
        }
        if (result.Items.Count > 0) element.ClipOwnership.Load(result, element.ClippingXml!, material, token);
    }

    internal static XElement? Write(OfdElement element, XNamespace ns)
    {
        if (element.Clips is not { } collection) return string.IsNullOrWhiteSpace(element.ClippingXml) ? null : OfdXmlEnvelope.Parse(element.ClippingXml!);
        var root = collection.SourceXml is null ? new XElement(ns + "Clips") : OfdXmlEnvelope.Parse(collection.SourceXml);
        Replace(root, "Clip", collection.Items.Select(clip => (clip.SourceSlot, Emit(clip))));
        return root.HasElements || root.Attributes().Any(a => !a.IsNamespaceDeclaration) || root.Nodes().Any(n => n is not XText text || !string.IsNullOrWhiteSpace(text.Value)) ? root : null;
        XElement Emit(OfdClip clip)
        {
            if (clip.Areas.Count == 0) throw new InvalidDataException("A clip requires at least one Area.");
            var node = clip.SourceXml is null ? new XElement(ns + "Clip") : OfdXmlEnvelope.Parse(clip.SourceXml);
            Replace(node, "Area", clip.Areas.Select(area => (area.SourceSlot, EmitArea(area)))); return node;
        }
        XElement EmitArea(OfdClipArea area)
        {
            var path = area.Path ?? throw new InvalidDataException("Clip Area requires a Path.");
            if (path.ObjectId is not null || path.LayerId is not null || path.Clips is not null || path.ClippingXml is not null)
                throw new InvalidDataException("CT_Path is local geometry without unit ID, layer or nested clips.");
            var node = area.SourceXml is null ? new XElement(ns + "Area") : OfdXmlEnvelope.Parse(area.SourceXml);
            node.SetAttributeValue("CTM", Format(area.Transform));
            var shape = path.SourceXml is null ? new XElement(ns + "Path") : OfdXmlEnvelope.Parse(path.SourceXml);
            shape.SetAttributeValue("Boundary", string.Join(" ", new[] { path.XMillimeters, path.YMillimeters, path.WidthMillimeters, path.HeightMillimeters }.Select(Number)));
            shape.SetAttributeValue("CTM", Format(path.Transform)); shape.SetAttributeValue("Stroke", path.Stroke ? "true" : "false"); shape.SetAttributeValue("Fill", path.Fill ? "true" : "false");
            if (path.FillRule is { } rule) shape.SetAttributeValue("Rule", rule == OfdFillRule.EvenOdd ? "Even-Odd" : "NonZero");
            var data = shape.Element(shape.Name.Namespace + "AbbreviatedData");
            if (data is null) shape.Add(new XElement(shape.Name.Namespace + "AbbreviatedData", path.AbbreviatedData));
            else if (data.Value != path.AbbreviatedData)
            {
                if (data.HasElements || data.Nodes().Any(n => n is not XText)) throw new NotSupportedException("Mixed path data cannot be edited safely.");
                data.Value = path.AbbreviatedData;
            }
            var old = node.Element(node.Name.Namespace + "Path"); if (old is null) node.Add(shape); else old.ReplaceWith(shape); return node;
        }
    }
    private static void Replace(XElement parent, string name, System.Collections.Generic.IEnumerable<(int Slot, XElement Node)> replacements)
    {
        var old = parent.Elements().ToArray(); var emitted = replacements.ToArray();
        var previous = -1;
        foreach (var item in emitted.Where(i => i.Slot >= 0))
        { if (item.Slot <= previous || item.Slot >= old.Length || old[item.Slot].Name != parent.Name.Namespace + name) throw new NotSupportedException("Imported clip slot order cannot be changed."); previous = item.Slot; }
        foreach (var pair in old.Select((node, slot) => (node, slot)).Where(p => p.node.Name == parent.Name.Namespace + name))
        {
            var replacement = emitted.FirstOrDefault(i => i.Slot == pair.slot).Node;
            if (replacement is null) pair.node.Remove(); else pair.node.ReplaceWith(replacement);
        }
        foreach (var item in emitted.Where(i => i.Slot < 0)) parent.Add(item.Node);
    }
    internal static string? Format(double[]? matrix) => matrix is null ? null : string.Join(" ", matrix.Select(Number));
    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    internal static double[]? Matrix(string? literal)
    { var result = Numbers(literal); if (result.Length == 0) return null; if (result.Length != 6) throw new InvalidDataException("CTM requires six values."); return result; }
    private static double[] Numbers(string? literal) => string.IsNullOrWhiteSpace(literal) ? Array.Empty<double>() : literal!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(part =>
    { if (!double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Invalid clip coordinate."); return value; }).ToArray();
}
