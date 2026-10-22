using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdCompositeProjection
{
    internal static IReadOnlyList<OfdElement> Expand(IEnumerable<OfdElement> roots, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(512, token); var result = new List<OfdElement>();
        var active = new HashSet<OfdCompositeResource>();
        foreach (var root in roots) Visit(root, Identity(), new List<OfdPathElement>(), 1, false, false);
        return result;
        void Visit(OfdElement element, double[] parent, List<OfdPathElement> inherited, int depth, bool compositeChild, bool pageBlockChild)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Depth(depth); material.Structure.Node(depth); material.Structure.Result(false);
            material.ChargeCopy(1024L + 32L * inherited.Count, token); material.ChargeWork(128, token);
            if (!OfdGraphicXmlContract.IsVisible(element)) return;
            if (element is OfdRawElement { LocalName: "PageBlock" } block)
            {
                if (block.PageBlockChildren is null || block.PageBlockSourceXml != block.Xml)
                    throw new NotSupportedException("Opaque or modified PageBlock cannot be rendered.");
                material.GuardFragmentXml(block.Xml, token);
                var node = OfdXmlEnvelope.Parse(block.Xml);
                if (node.Name.LocalName != "PageBlock" ||
                    node.Name.NamespaceName is not ("" or Ofdrw.Net.Core.Constants.OfdConstants.Namespace or Ofdrw.Net.Core.Constants.OfdConstants.StandardNamespace) ||
                    node.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                        (a.Name.Namespace != XNamespace.None || a.Name.LocalName != "ID")) ||
                    node.Nodes().Any(n => n is XProcessingInstruction || n is XText t && !string.IsNullOrWhiteSpace(t.Value)))
                    throw new NotSupportedException("PageBlock has unsupported drawing attributes or content.");
                foreach (var child in block.PageBlockChildren) Visit(child, parent, inherited, depth + 1, compositeChild, true);
                return;
            }
            var origin = Multiply(parent, new[] { 1d, 0, 0, 1, element.XMillimeters, element.YMillimeters });
            if (element is OfdCompositeElement composite)
            {
                if (composite.Resource is null || !active.Add(composite.Resource)) throw new InvalidDataException("Cyclic or missing composite resource.");
                var placementError = AdmitComposite(composite, pageBlockChild, material, token);
                if (placementError is not null)
                {
                    // Extraction shares this projection but does not run render
                    // preflight. Keep the resource's TextCode reachable while
                    // every renderer rejects the placement before writing.
                    material.Structure.Result(true);
                    material.ChargeCopy(256, token);
                    result.Add(new OfdPathElement { RenderPaintError = placementError });
                }
                var local = Multiply(origin, composite.Transform ?? Identity());
                var clips = new List<OfdPathElement>(inherited);
                clips.Add(Rectangle(composite.WidthMillimeters, composite.HeightMillimeters, origin));
                clips.Add(Rectangle(composite.Resource.WidthMillimeters, composite.Resource.HeightMillimeters, local));
                AddClips(composite, origin, clips);
                foreach (var child in composite.Resource.Elements) Visit(child, local, clips, depth + 1, true, pageBlockChild);
                active.Remove(composite.Resource); return;
            }
            if (ReferenceEquals(parent, null)) throw new InvalidDataException("Missing parent transform.");
            if (element is OfdRawElement { LocalName: "CompositeObject" }) throw new NotSupportedException("Unresolved or opaque composite placement cannot be exported.");
            if (compositeChild && element is not (OfdTextElement or OfdImageElement or OfdPathElement))
                throw new NotSupportedException("Opaque composite children cannot be exported.");
            var source = element switch { OfdTextElement t => t.SourceXml, OfdImageElement i => i.SourceXml, OfdPathElement p => p.SourceXml, _ => null };
            if (element.RenderSourceXml is { } resolvedSource)
            {
                if (element.RenderSourceAuthority != source)
                    throw new NotSupportedException("Imported drawing changed after paint resolution.");
                if (pageBlockChild || compositeChild) source = resolvedSource;
            }
            string? scalarError = null;
            if (pageBlockChild && element is OfdImageElement { SourceMediaAmbiguous: true })
                scalarError = "PageBlock image has an ambiguous media declaration.";
            // This admission is for rendering. Keep a Reader paint failure on
            // the projected leaf so text extraction can still read TextCode;
            // every renderer rejects the leaf in OfdRenderPreflight.
            if ((compositeChild || pageBlockChild) && source is not null &&
                element.RenderPaintError is not { Length: > 0 })
            {
                // Paint normalization can replace the source XML, and Reader
                // applies defaults to malformed scalar attributes. Admit the
                // original lexemes before checking the normalized render XML.
                var original = element switch { OfdTextElement t => t.SourceXml, OfdImageElement i => i.SourceXml, OfdPathElement p => p.SourceXml, _ => null };
                if (original is not null)
                {
                    material.GuardFragmentXml(original, token);
                    var originalNode = OfdXmlEnvelope.Parse(original);
                    if (pageBlockChild &&
                        (!OfdBoxParser.TryParse(originalNode.Attribute("Boundary")?.Value, out var box) ||
                            box.w <= 0 || box.h <= 0))
                        scalarError = "Malformed PageBlock drawing boundary cannot be exported.";
                    if (scalarError is null && !HasWellFormedScalars(originalNode, material, token))
                        scalarError = "Malformed PageBlock or composite child attribute cannot be exported.";
                    if (scalarError is null && element is OfdTextElement &&
                        originalNode.DescendantsAndSelf().Any(node => node.Name.LocalName == "CGTransform"))
                        scalarError = "PageBlock or composite text with explicit glyph substitutions cannot be exported.";
                    if (scalarError is null && originalNode.Name.LocalName == "PathObject" &&
                        !OfdPathDataSyntax.HasSingleValidDrawing(originalNode, token))
                        scalarError = "Malformed PageBlock or composite path data cannot be exported.";
                }
                if (scalarError is null)
                {
                    material.GuardFragmentXml(source, token); var node = OfdXmlEnvelope.Parse(source);
                    if (!OfdGraphicXmlContract.HasKnownChildren(node) || OfdGraphicXmlContract.HasUnsupportedReferences(node) || node.DescendantsAndSelf().Attributes().Any(a => !OfdGraphicXmlContract.IsKnownAttribute(a)))
                        throw new NotSupportedException("Unsupported PageBlock or composite child style cannot be exported.");
                }
            }
            if (!compositeChild)
            {
                if (scalarError is null) result.Add(element);
                else
                {
                    var rejected = OfdModelCloner.CloneElement(element, clonePayloads: false, material: material, cancellationToken: token);
                    rejected.RenderPaintError = scalarError; result.Add(rejected);
                }
                return;
            }
            var renderSource = scalarError is null && element is OfdTextElement sourceText
                ? OfdTextRenderProjection.Create(sourceText, material, token) : element;
            var projected = OfdModelCloner.CloneElement(renderSource, clonePayloads: false, material: material, cancellationToken: token);
            if (scalarError is not null) projected.RenderPaintError = scalarError;
            var inner = element switch { OfdTextElement t => t.Transform, OfdImageElement i => i.Transform ?? new[] { i.WidthMillimeters, 0, 0, i.HeightMillimeters, 0, 0 }, OfdPathElement p => p.Transform, _ => null } ?? Identity();
            var transformed = Multiply(origin, inner);
            var clipsForLeaf = inherited.Select(p => new OfdPathElement { AbbreviatedData = p.AbbreviatedData, Transform = p.Transform, FillRule = p.FillRule, Fill = true, Stroke = false }).ToList(); AddClips(element, origin, clipsForLeaf);
            if (element is OfdImageElement) clipsForLeaf.Add(Rectangle(element.WidthMillimeters, element.HeightMillimeters, origin));
            var x0 = origin[4] + Math.Min(0, origin[0] * element.WidthMillimeters) + Math.Min(0, origin[2] * element.HeightMillimeters);
            var y0 = origin[5] + Math.Min(0, origin[1] * element.WidthMillimeters) + Math.Min(0, origin[3] * element.HeightMillimeters);
            var shift = new[] { 1d, 0, 0, 1, -x0, -y0 };
            transformed = Multiply(shift, transformed);
            foreach (var clip in clipsForLeaf) clip.Transform = Multiply(shift, clip.Transform ?? Identity());
            projected.XMillimeters = x0; projected.YMillimeters = y0;
            projected.WidthMillimeters = Math.Abs(origin[0]) * element.WidthMillimeters + Math.Abs(origin[2]) * element.HeightMillimeters;
            projected.HeightMillimeters = Math.Abs(origin[1]) * element.WidthMillimeters + Math.Abs(origin[3]) * element.HeightMillimeters;
            material.StoreMatrix(transformed);
            if (projected is OfdTextElement text) { text.Transform = transformed; text.SourceXml = null; }
            if (projected is OfdPathElement path) { if (scalarError is null) OfdPathStyle.CaptureProjection((OfdPathElement)element, path); path.Transform = transformed; path.SourceXml = null; }
            if (projected is OfdImageElement image) { image.Transform = transformed; image.SourceXml = null; }
            // Projection consumes the original drawing envelope. Its resolved
            // paint now belongs to this transient leaf, whose SourceXml is null.
            projected.RenderSourceAuthority = null;
            projected.ClipOwnership.Value = projected.ClipOwnership.Original = null; projected.ClipOwnership.Baseline = null;
            projected.ClippingXml = Encode(clipsForLeaf, material, token); result.Add(projected);
        }
        void AddClips(OfdElement owner, double[] origin, List<OfdPathElement> clips)
        {
            var xml = OfdClipXml.Write(owner, Ofdrw.Net.Core.Constants.OfdConstants.Namespace, material, token);
            if (xml is null) return;
            foreach (var region in OfdClipGeometry.Read(xml, material, token, requireClosedProfile: true))
                foreach (var path in region.Paths)
                { path.Transform = Multiply(origin, path.Transform ?? Identity()); path.FillRule = region.EvenOdd ? OfdFillRule.EvenOdd : OfdFillRule.NonZero; clips.Add(path); }
        }
    }
    private static string? AdmitComposite(OfdCompositeElement composite, bool pageBlockChild, OfdModelMaterialization material, CancellationToken token)
    {
        string? placementError = pageBlockChild && composite.SourceXml is null
            ? "PageBlock composite placement has no original drawing attributes." : null;
        if (pageBlockChild && !composite.Resource.SourceContainerSupported)
            placementError ??= "PageBlock composite resource container cannot be rendered.";
        if (composite.SourceXml is { } source)
        {
            material.GuardFragmentXml(source, token); var node = OfdXmlEnvelope.Parse(source);
            if (pageBlockChild &&
                (!OfdBoxParser.TryParse(node.Attribute("Boundary")?.Value, out var box) || box.w <= 0 || box.h <= 0 ||
                    !HasWellFormedScalars(node, material, token)))
                placementError = "Malformed PageBlock composite placement cannot be exported.";
            if (node.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Boundary" or "CTM" or "ResourceID" or "Visible"))) || node.Elements().Any(n => n.Name != node.Name.Namespace + "Clips") || node.Nodes().Any(n => n is XProcessingInstruction || n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Composite opacity, drawing parameters and extensions are outside the render profile.");
        }
        if (composite.Resource.SourceXml is { } resourceSource)
        {
            material.GuardFragmentXml(resourceSource, token); var node = OfdXmlEnvelope.Parse(resourceSource);
            if (node.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Width" or "Height"))) || node.Nodes().Any(n => n is XProcessingInstruction || n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Composite resource parameters are outside the render profile.");
        }
        return placementError;
    }
    private static OfdPathElement Rectangle(double w, double h, double[] matrix) => new() { Stroke = false, Fill = true,
        AbbreviatedData = $"M 0 0 L {Number(w)} 0 L {Number(w)} {Number(h)} L 0 {Number(h)} C", Transform = matrix };
    private static string Encode(IReadOnlyList<OfdPathElement> paths, OfdModelMaterialization material, CancellationToken token)
    {
        XNamespace ns = Ofdrw.Net.Core.Constants.OfdConstants.Namespace;
        long length = 0;
        foreach (var path in paths)
        { token.ThrowIfCancellationRequested(); length = checked(length + path.AbbreviatedData.Length); }
        // Repeated ancestor strings are a new materialization for each leaf.
        material.ChargeCopy(checked(8192L + 64L * length + 8192L * paths.Count), token);
        material.ChargeWork(checked(4L * length + 256L * paths.Count), token);
        var root = new XElement(ns + "Clips");
        foreach (var path in paths)
        {
            token.ThrowIfCancellationRequested();
            root.Add(new XElement(ns + "Clip", new XElement(ns + "Area", new XElement(ns + "Path",
                new XAttribute("CTM", OfdClipXml.Format(path.Transform, material, token)!),
                new XAttribute("Rule", path.FillRule == OfdFillRule.EvenOdd ? "Even-Odd" : "NonZero"),
                new XElement(ns + "AbbreviatedData", path.AbbreviatedData)))));
        }
        var encoded = OfdXmlEnvelope.Capture(root, material, token); material.StoreString(encoded); return encoded;
    }
    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
    private static bool HasWellFormedScalars(XElement root, OfdModelMaterialization material, CancellationToken token)
    {
        foreach (var attribute in root.DescendantsAndSelf().Attributes())
        {
            token.ThrowIfCancellationRequested();
            if (attribute.IsNamespaceDeclaration || attribute.Name.Namespace != XNamespace.None) continue;
            var name = attribute.Name.LocalName;
            var value = attribute.Value;
            if (name is "Visible" or "Stroke" or "Fill" or "Italic")
            {
                var normalized = value.Trim();
                if (normalized is not ("0" or "1") && !bool.TryParse(normalized, out _)) return false;
                if (name == "Visible" && attribute.Parent!.Name.LocalName == "Path" &&
                    (normalized == "0" || string.Equals(normalized, "false", StringComparison.OrdinalIgnoreCase))) return false;
            }
            else if (name == "Boundary")
            {
                if (!OfdBoxParser.TryParse(value, out _)) return false;
            }
            else if (name == "CTM")
            {
                if (!FiniteNumbers(value, 6, 6)) return false;
            }
            else if (name == "Start")
            {
                var parts = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length != 2 || !FiniteNumber(parts[0], out var x) || !FiniteNumber(parts[1], out var y) || x != 0 || y != 0) return false;
            }
            else if (name is "Alpha" or "Weight" or "ReadDirection" or "CharDirection" or "CodePosition" or "CodeCount" or "GlyphCount")
            {
                if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ||
                    (name == "Alpha" && n is < 0 or > 255) ||
                    (name == "Weight" && (n is < 100 or > 900 || n % 100 != 0)) ||
                    (name is "CodePosition" or "CodeCount" or "GlyphCount") && n < 0)
                    return false;
            }
            else if (name is "Size" or "LineWidth" or "HScale" or "MiterLimit" or "DashOffset" or "X" or "Y")
            {
                if (!FiniteNumber(value, out var n) ||
                    (name is "Size" or "HScale" or "MiterLimit") && n <= 0 ||
                    (name is "LineWidth" or "DashOffset") && n < 0) return false;
            }
            else if (name == "DashPattern")
            {
                if (!FiniteNumbers(value, 1, int.MaxValue)) return false;
            }
            else if (name is "DeltaX" or "DeltaY")
            {
                if (!WellFormedDeltas(value, material, token)) return false;
            }
            else if (name == "Rule" && value is not ("NonZero" or "Even-Odd")) return false;
            else if (name == "Cap" && value.ToLowerInvariant() is not ("butt" or "round" or "square")) return false;
            else if (name == "Join" && value.ToLowerInvariant() is not ("miter" or "round" or "bevel")) return false;
        }
        return true;
    }
    private static bool WellFormedDeltas(string value, OfdModelMaterialization material, CancellationToken token)
    {
        // ExpandDeltas stops as soon as it has enough glyph advances. Validate
        // the entire expression before that renderer can discard a bad suffix.
        var index = 0;
        while (NextDeltaToken(value, ref index, material, token, out var part))
        {
            token.ThrowIfCancellationRequested();
            if (string.Equals(part, "g", StringComparison.OrdinalIgnoreCase))
            {
                if (!NextDeltaToken(value, ref index, material, token, out var repeat) ||
                    !int.TryParse(repeat, NumberStyles.Integer, CultureInfo.InvariantCulture, out var count) || count < 0 ||
                    !NextDeltaToken(value, ref index, material, token, out part))
                    return false;
            }
            if (!FiniteNumber(part, out _)) return false;
        }
        return true;
    }
    private static bool NextDeltaToken(string value, ref int index, OfdModelMaterialization material,
        CancellationToken token, out string part)
    {
        while (index < value.Length && value[index] is ' ' or '\t' or '\r' or '\n')
        { if ((index & 1023) == 0) token.ThrowIfCancellationRequested(); index++; }
        var start = index;
        while (index < value.Length && value[index] is not (' ' or '\t' or '\r' or '\n'))
        { if ((index & 1023) == 0) token.ThrowIfCancellationRequested(); index++; }
        if (start == index) { part = string.Empty; return false; }
        material.ChargeCopy(checked(64L + 2L * (index - start)), token);
        part = value.Substring(start, index - start);
        return true;
    }
    private static bool FiniteNumber(string value, out double number) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
        !double.IsNaN(number) && !double.IsInfinity(number);
    private static bool FiniteNumbers(string value, int minimum, int maximum)
    {
        var parts = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length >= minimum && parts.Length <= maximum && parts.All(part => FiniteNumber(part, out _));
    }
    internal static double[] Identity() => new[] { 1d, 0, 0, 1, 0, 0 };
    internal static double[] Multiply(double[] a, double[] b) => new[] { a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1], a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3], a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5] };
}
