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
        foreach (var root in roots) Visit(root, Identity(), new List<OfdPathElement>(), 1);
        return result;
        void Visit(OfdElement element, double[] parent, List<OfdPathElement> inherited, int depth)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Depth(depth); material.Structure.Node(depth); material.Structure.Result(false);
            material.ChargeCopy(1024L + 32L * inherited.Count, token); material.ChargeWork(128, token);
            if (!OfdGraphicXmlContract.IsVisible(element)) return;
            var origin = Multiply(parent, new[] { 1d, 0, 0, 1, element.XMillimeters, element.YMillimeters });
            if (element is OfdCompositeElement composite)
            {
                if (composite.Resource is null || !active.Add(composite.Resource)) throw new InvalidDataException("Cyclic or missing composite resource.");
                AdmitComposite(composite, material, token);
                var local = Multiply(origin, composite.Transform ?? Identity());
                var clips = new List<OfdPathElement>(inherited);
                clips.Add(Rectangle(composite.WidthMillimeters, composite.HeightMillimeters, origin));
                clips.Add(Rectangle(composite.Resource.WidthMillimeters, composite.Resource.HeightMillimeters, local));
                AddClips(composite, origin, clips);
                foreach (var child in composite.Resource.Elements) Visit(child, local, clips, depth + 1);
                active.Remove(composite.Resource); return;
            }
            if (ReferenceEquals(parent, null)) throw new InvalidDataException("Missing parent transform.");
            if (element is OfdRawElement { LocalName: "CompositeObject" }) throw new NotSupportedException("Unresolved or opaque composite placement cannot be exported.");
            if (depth == 1) { result.Add(element); return; }
            if (element is not (OfdTextElement or OfdImageElement or OfdPathElement)) throw new NotSupportedException("Opaque composite children cannot be exported.");
            var source = element switch { OfdTextElement t => t.SourceXml, OfdImageElement i => i.SourceXml, OfdPathElement p => p.SourceXml, _ => null };
            if (source is not null)
            {
                material.GuardFragmentXml(source, token); var node = OfdXmlEnvelope.Parse(source);
                if (!OfdGraphicXmlContract.HasKnownChildren(node) || OfdGraphicXmlContract.HasUnsupportedReferences(node) || node.DescendantsAndSelf().Attributes().Any(a => !OfdGraphicXmlContract.IsKnownAttribute(a)))
                    throw new NotSupportedException("Unsupported composite child style cannot be exported.");
            }
            var renderSource = element is OfdTextElement sourceText ? OfdTextRenderProjection.Create(sourceText, material, token) : element;
            var projected = OfdModelCloner.CloneElement(renderSource, clonePayloads: false, material: material, cancellationToken: token);
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
            if (projected is OfdPathElement path) { path.Transform = transformed; path.SourceXml = null; }
            if (projected is OfdImageElement image) { image.Transform = transformed; image.SourceXml = null; }
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
    private static void AdmitComposite(OfdCompositeElement composite, OfdModelMaterialization material, CancellationToken token)
    {
        if (composite.SourceXml is { } source)
        {
            material.GuardFragmentXml(source, token); var node = OfdXmlEnvelope.Parse(source);
            if (node.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Boundary" or "CTM" or "ResourceID" or "Visible"))) || node.Elements().Any(n => n.Name != node.Name.Namespace + "Clips") || node.Nodes().Any(n => n is XProcessingInstruction || n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Composite opacity, drawing parameters and extensions are outside the render profile.");
        }
        if (composite.Resource.SourceXml is { } resourceSource)
        {
            material.GuardFragmentXml(resourceSource, token); var node = OfdXmlEnvelope.Parse(resourceSource);
            if (node.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || a.Name.LocalName is not ("ID" or "Width" or "Height"))) || node.Nodes().Any(n => n is XProcessingInstruction || n is XText text && !string.IsNullOrWhiteSpace(text.Value)))
                throw new NotSupportedException("Composite resource parameters are outside the render profile.");
        }
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
    internal static double[] Identity() => new[] { 1d, 0, 0, 1, 0, 0 };
    internal static double[] Multiply(double[] a, double[] b) => new[] { a[0]*b[0]+a[2]*b[1], a[1]*b[0]+a[3]*b[1], a[0]*b[2]+a[2]*b[3], a[1]*b[2]+a[3]*b[3], a[0]*b[4]+a[2]*b[5]+a[4], a[1]*b[4]+a[3]*b[5]+a[5] };
}
