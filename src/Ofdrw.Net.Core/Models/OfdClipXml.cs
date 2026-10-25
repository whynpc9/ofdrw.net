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
        if ((ns != XNamespace.None && ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.Namespace && ns.NamespaceName != Ofdrw.Net.Core.Constants.OfdConstants.StandardNamespace) || root.Name != ns + "Clips" || root.Attribute("TransFlag") is not null) return;
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
                var box = Numbers(shape.Attribute("Boundary")?.Value, 4, material, token); if (box.Length != 0 && box.Length != 4) return;
                if (shape.Element(ns + "AbbreviatedData")?.HasElements == true) return;
                var rule = shape.Attribute("Rule")?.Value;
                if (rule is not null && rule is not ("NonZero" or "Even-Odd")) return;
                var path = new OfdPathElement
                {
                    AbbreviatedData = shape.Element(ns + "AbbreviatedData")?.Value ?? string.Empty,
                    XMillimeters = box.Length == 4 ? box[0] : 0, YMillimeters = box.Length == 4 ? box[1] : 0,
                    WidthMillimeters = box.Length == 4 ? box[2] : 0, HeightMillimeters = box.Length == 4 ? box[3] : 0,
                    Transform = Matrix(shape.Attribute("CTM")?.Value, material, token), Stroke = false, Fill = true,
                    FillRule = rule == "Even-Odd" ? OfdFillRule.EvenOdd : OfdFillRule.NonZero,
                    SourceXml = OfdXmlEnvelope.Capture(shape, material, token)
                };
                clip.Areas.Add(new OfdClipArea { Path = path, OriginalPath = path, OriginalPathXml = path.SourceXml, Transform = Matrix(area.Attribute("CTM")?.Value, material, token),
                    SourceXml = OfdXmlEnvelope.Capture(area, material, token), SourceSlot = areaSlot++ });
            }
            if (clip.Areas.Count == 0) return;
            result.Items.Add(clip);
        }
        if (result.Items.Count > 0) element.ClipOwnership.Load(result, element.ClippingXml!, material, token);
    }

    internal static XElement? Write(OfdElement element, XNamespace ns, OfdModelMaterialization material, CancellationToken token)
    {
        element.ClipOwnership.Validate(element.ClippingXml, material, token);
        if (element.Clips is not { } collection) return string.IsNullOrWhiteSpace(element.ClippingXml) ? null : Parse(element.ClippingXml!);
        // The original envelope owns opaque wrapper content even if the caller
        // assigns a new empty view. Replacing nonempty imported owners is unsafe.
        var original = element.ClipOwnership.Original;
        if (original is not null && !ReferenceEquals(original, collection) && collection.Items.Count != 0)
            throw new NotSupportedException("Replacing imported clips requires authenticated slots.");
        material.ChargeCopy(8192L + 128L * collection.Items.Count, token);
        var source = original?.SourceXml ?? collection.SourceXml;
        var root = source is null ? new XElement(ns + "Clips") : Parse(source);
        Replace(root, "Clip", collection.Items.Select(clip => (clip.SourceSlot, Emit(clip))), material, token);
        return root.HasElements || root.Attributes().Any(a => !a.IsNamespaceDeclaration) || root.Nodes().Any(n => n is not XText text || !string.IsNullOrWhiteSpace(text.Value)) ? root : null;
        XElement Emit(OfdClip clip)
        {
            if (clip.Areas.Count == 0) throw new InvalidDataException("A clip requires at least one Area.");
            material.ChargeCopy(8192L + 128L * clip.Areas.Count, token);
            var node = clip.SourceXml is null ? new XElement(ns + "Clip") : Parse(clip.SourceXml);
            Replace(node, "Area", clip.Areas.Select(area => (area.SourceSlot, EmitArea(area))), material, token); return node;
        }
        XElement EmitArea(OfdClipArea area)
        {
            var path = area.Path ?? throw new InvalidDataException("Clip Area requires a Path.");
            if (path.ObjectId is not null || path.LayerId is not null || path.Clips is not null || path.ClippingXml is not null)
                throw new InvalidDataException("CT_Path is local geometry without unit ID, layer or nested clips.");
            material.ChargeCopy(8192L + 16L * path.AbbreviatedData.Length, token); material.ChargeWork(2L * path.AbbreviatedData.Length, token);
            var node = area.SourceXml is null ? new XElement(ns + "Area") : Parse(area.SourceXml);
            node.SetAttributeValue("CTM", Format(area.Transform, material, token));
            var sourcePath = node.Element(node.Name.Namespace + "Path");
            if (area.OriginalPath is not null && ReferenceEquals(path, area.OriginalPath) && path.SourceXml != area.OriginalPathXml)
                throw new NotSupportedException("Imported clip Path.SourceXml is reader-owned; edit the typed fields.");
            if (!ReferenceEquals(path, area.OriginalPath) && path.SourceXml is not null)
                throw new NotSupportedException("Fresh clip paths use typed fields without caller XML.");
            if (sourcePath is not null && !ReferenceEquals(path, area.OriginalPath) && !CanDelete(sourcePath, material, token))
                throw new NotSupportedException("Cannot replace a clip Path containing unclaimed XML.");
            var shape = sourcePath is null ? new XElement(node.Name.Namespace + "Path") : new XElement(sourcePath);
            shape.SetAttributeValue("Boundary", string.Join(" ", new[] { path.XMillimeters, path.YMillimeters, path.WidthMillimeters, path.HeightMillimeters }.Select(Number)));
            shape.SetAttributeValue("CTM", Format(path.Transform, material, token)); shape.SetAttributeValue("Stroke", path.Stroke ? "true" : "false"); shape.SetAttributeValue("Fill", path.Fill ? "true" : "false");
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
        XElement Parse(string xml)
        { material.GuardFragmentXml(xml, token); material.ChargeCopy(8192L + 64L * xml.Length, token); material.ChargeWork(2L * xml.Length, token); return OfdXmlEnvelope.Parse(xml); }
    }
    private static void Replace(XElement parent, string name, System.Collections.Generic.IEnumerable<(int Slot, XElement Node)> replacements,
        OfdModelMaterialization material, CancellationToken token)
    {
        var count = 0;
        foreach (var child in parent.Elements()) { token.ThrowIfCancellationRequested(); material.ChargeWork(64, token); count++; }
        material.ChargeCopy(256L + 32L * count, token);
        var old = parent.Elements().ToArray(); var emitted = new System.Collections.Generic.List<(int Slot, XElement Node)>();
        var previous = -1; var fresh = false;
        foreach (var item in replacements)
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(128, token); material.ChargeWork(64, token);
            if (item.Slot < 0) fresh = true;
            else
            {
                if (fresh || item.Slot <= previous || item.Slot >= old.Length || old[item.Slot].Name != parent.Name.Namespace + name)
                    throw new NotSupportedException("Imported clip slots retain their order; fresh slots append.");
                previous = item.Slot;
            }
            emitted.Add(item);
        }
        var next = 0;
        for (var slot = 0; slot < old.Length; slot++)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            if (old[slot].Name != parent.Name.Namespace + name) continue;
            if (next < emitted.Count && emitted[next].Slot == slot) old[slot].ReplaceWith(emitted[next++].Node);
            else
            {
                if (!CanDelete(old[slot], material, token)) throw new NotSupportedException("Cannot remove a clip slot containing unclaimed XML.");
                old[slot].Remove();
            }
        }
        for (; next < emitted.Count; next++)
        { token.ThrowIfCancellationRequested(); material.ChargeWork(64, token); parent.Add(emitted[next].Node); }
    }
    private static bool CanDelete(XElement root, OfdModelMaterialization material, CancellationToken token)
    {
        if (!OfdGraphicXmlContract.HasKnownChildren(root)) return false;
        foreach (var node in root.DescendantNodesAndSelf())
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            if (node is XComment or XProcessingInstruction) return false;
            if (node is not XElement element) continue;
            foreach (var attribute in element.Attributes())
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(32, token);
                if (!OfdGraphicXmlContract.IsKnownAttribute(attribute) || !attribute.IsNamespaceDeclaration &&
                    (attribute.Name == "ID" || attribute.Name == "Start")) return false;
            }
        }
        return true;
    }
    internal static string? Format(double[]? matrix, OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        if (matrix is null) return null;
        if (matrix.Length != 6 || matrix.Any(v => double.IsNaN(v) || double.IsInfinity(v))) throw new InvalidDataException("CTM requires six finite values.");
        token.ThrowIfCancellationRequested(); material?.ChargeCopy(2048, token); material?.ChargeWork(96, token);
        return string.Join(" ", matrix.Select(Number));
    }
    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    internal static double[]? Matrix(string? literal, OfdModelMaterialization material, CancellationToken token)
    { var result = Numbers(literal, 6, material, token); return result.Length == 0 ? null : result; }
    private static double[] Numbers(string? literal, int expected, OfdModelMaterialization material, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); if (literal is null) return Array.Empty<double>();
        material.ChargeCopy(256L + 2L * literal.Length + expected * sizeof(double), token);
        material.ChargeWork(4L * literal.Length, token);
        var result = new double[expected]; var count = 0; var index = 0; var nextCheck = 0;
        while (index < literal.Length)
        {
            OfdModelTraversal.CheckCancellationCheckpoint(index, ref nextCheck, token);
            if (char.IsWhiteSpace(literal[index])) { index++; continue; }
            if (count == expected) throw new InvalidDataException("Too many clip coordinates.");
            var start = index;
            while (index < literal.Length && !char.IsWhiteSpace(literal[index]))
            { OfdModelTraversal.CheckCancellationCheckpoint(index, ref nextCheck, token); index++; }
            token.ThrowIfCancellationRequested();
            if (!double.TryParse(literal.Substring(start, index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidDataException("Invalid clip coordinate.");
            result[count++] = value;
        }
        if (count == 0) return Array.Empty<double>();
        if (count != expected) throw new InvalidDataException("Clip coordinate cardinality is invalid.");
        return result;
    }
}
