using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// The one local-Appearance to page projection, shared by Reader and consumers.
internal static class OfdAnnotationProjection
{
    internal sealed class Occurrence(OfdElement? definition, OfdElement render, double[]? parent, string? entry, int annotation, int element, bool visible = true)
    {
        internal OfdElement? Definition { get; } = definition;
        internal OfdElement RenderElement { get; } = render;
        internal double[]? ParentToPage { get; } = parent;
        internal string? Entry { get; } = entry;
        internal int Annotation { get; } = annotation;
        internal int Element { get; } = element;
        internal bool Visible { get; } = visible;
    }
    internal static System.Collections.Generic.List<Occurrence> Resolve(OfdPage page, OfdModelMaterialization material, CancellationToken token = default, bool includeHidden = false, string? documentNamespace = null)
    {
        page.AnnotationOwnership.Validate(page, material, token);
        material.ChargeCopy(128, token);
        var result = new System.Collections.Generic.List<Occurrence>();
        var collection = page.Annotations;
        if (collection is null)
        {
            foreach (var element in page.AnnotationAppearances) Add(new Occurrence(null, element, null, null, -1, result.Count));
            return result;
        }
        material.ChargeCopy(256L + collection.Items.Count * 128L, token);
        var current = new System.Collections.Generic.Dictionary<OfdAnnotation, int>();
        for (var i = 0; i < collection.Items.Count; i++) { token.ThrowIfCancellationRequested(); current.Add(collection.Items[i], i); }
        var original = page.AnnotationOwnership.Original;
        var slots = new System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<OfdAnnotation>>();
        if (original is not null) foreach (var annotation in original.Items)
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(192, token);
            if (annotation.Source is not { } source) continue;
            if (!slots.TryGetValue(source.ProjectionStart, out var at)) slots.Add(source.ProjectionStart, at = new());
            at.Add(annotation);
        }
        for (var index = 0; index <= page.AnnotationAppearances.Count; index++)
        {
            token.ThrowIfCancellationRequested(); var skip = 0;
            if (slots.TryGetValue(index, out var at)) foreach (var annotation in at)
            {
                if (current.TryGetValue(annotation, out var ordinal)) Emit(annotation, ordinal);
                skip += annotation.Source!.ProjectionCount;
            }
            if (skip > 0) { index += skip - 1; continue; }
            if (index < page.AnnotationAppearances.Count) Add(new Occurrence(null, page.AnnotationAppearances[index], null, null, -1, index));
        }
        for (var ordinal = 0; ordinal < collection.Items.Count; ordinal++)
            if (collection.Items[ordinal].Source is null) Emit(collection.Items[ordinal], ordinal);
        return result;
        void Add(Occurrence value) { token.ThrowIfCancellationRequested(); material.Structure.Reference(); material.ChargeCopy(128, token); result.Add(value); }
        void Emit(OfdAnnotation annotation, int ordinal)
        {
            if (annotation.Visible == false && !includeHidden) return;
            var appearance = annotation.Appearance;
            material.ChargeCopy(192, token);
            var matrix = appearance.Transform ?? new double[] { 1, 0, 0, 1, 0, 0 };
            var parent = new[] { matrix[0], matrix[1], matrix[2], matrix[3], appearance.XMillimeters + matrix[4], appearance.YMillimeters + matrix[5] };
            material.ChargeCopy(96, token);
            var ns = XNamespace.Get(annotation.Source?.Namespace ?? documentNamespace ?? Ofdrw.Net.Core.Constants.OfdConstants.Namespace);
            for (var index = 0; index < appearance.Elements.Count; index++)
            {
                var element = appearance.Elements[index];
                Add(new Occurrence(element, Project(element, appearance, ns, material, token), parent, annotation.Source?.Entry, ordinal, index, annotation.Visible != false));
            }
        }
    }

    internal static OfdElement Project(OfdElement definition, OfdAnnotationAppearance appearance,
        XNamespace ns, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        material.ReserveReadModel(16384, cancellationToken);
        material.ChargeCopy(512L + (definition is OfdTextElement textDefinition ? textDefinition.Runs.Count * 128L : 0), cancellationToken);
        var element = OfdModelCloner.CloneElement(definition, clonePayloads: false, material: material, cancellationToken: cancellationToken);
        if (element.Clips is not null)
        {
            element.ClippingXml = OfdClipXml.Write(element, ns, material, cancellationToken) is { } typedClip ? OfdXmlEnvelope.Capture(typedClip, material, cancellationToken) : null;
            element.ClipOwnership.Value = element.ClipOwnership.Original = null; element.ClipOwnership.Baseline = null;
        }
        var appearanceTransform = appearance.Transform;
        var box = (x: appearance.XMillimeters, y: appearance.YMillimeters, w: appearance.WidthMillimeters, h: appearance.HeightMillimeters);
        var transform = appearanceTransform ?? new double[] { 1, 0, 0, 1, 0, 0 };
        var x = element.XMillimeters; var y = element.YMillimeters;
        var originalWidth = element.WidthMillimeters; var originalHeight = element.HeightMillimeters;
        var minX = Math.Min(0, transform[0] * originalWidth) + Math.Min(0, transform[2] * originalHeight);
        var minY = Math.Min(0, transform[1] * originalWidth) + Math.Min(0, transform[3] * originalHeight);
        element.XMillimeters = box.x + transform[0] * x + transform[2] * y + transform[4] + minX;
        element.YMillimeters = box.y + transform[1] * x + transform[3] * y + transform[5] + minY;
        element.WidthMillimeters = Math.Abs(transform[0]) * originalWidth + Math.Abs(transform[2]) * originalHeight;
        element.HeightMillimeters = Math.Abs(transform[1]) * originalWidth + Math.Abs(transform[3]) * originalHeight;
        if (appearanceTransform is not null)
        {
            var inner = element switch
            {
                OfdTextElement text => text.Transform,
                OfdPathElement path => path.Transform,
                OfdImageElement image => image.Transform ?? new double[] { originalWidth, 0, 0, originalHeight, 0, 0 },
                _ => null
            } ?? new double[] { 1, 0, 0, 1, 0, 0 };
            var combined = new double[]
            {
                transform[0] * inner[0] + transform[2] * inner[1], transform[1] * inner[0] + transform[3] * inner[1],
                transform[0] * inner[2] + transform[2] * inner[3], transform[1] * inner[2] + transform[3] * inner[3],
                transform[0] * inner[4] + transform[2] * inner[5] - minX, transform[1] * inner[4] + transform[3] * inner[5] - minY
            };
            if (element is OfdTextElement transformedText) transformedText.Transform = combined;
            if (element is OfdImageElement transformedImage) transformedImage.Transform = combined;
            if (element is OfdPathElement transformedPath) transformedPath.Transform = combined;
        }
        element.ClippingXml = TransformAnnotationClips(element.ClippingXml, ns, transform, minX, minY,
            box.x + transform[4] - element.XMillimeters, box.y + transform[5] - element.YMillimeters, box.w, box.h,
            element is OfdImageElement ? originalWidth : 0, element is OfdImageElement ? originalHeight : 0, material, cancellationToken);
        // Keep the serialized boundary consistent with the translated typed geometry.
        var source = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
        if (source is not null)
        {
            material.GuardFragmentXml(source, cancellationToken); material.ChargeCopy(2L * source.Length, cancellationToken); material.ChargeWork(2L * source.Length, cancellationToken);
            material.PrepareReadXmlParser(checked(2L * source.Length), cancellationToken);
            material.ReserveReadModel(16384L + 8L * source.Length, cancellationToken);
            var node = OfdXmlEnvelope.Parse(source);
            node.SetAttributeValue("Boundary", string.Join(" ", new[] { element.XMillimeters, element.YMillimeters, element.WidthMillimeters, element.HeightMillimeters }.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
            if (appearanceTransform is not null)
            {
                var objectTransform = element switch { OfdTextElement matrixText => matrixText.Transform, OfdImageElement matrixImage => matrixImage.Transform, OfdPathElement matrixPath => matrixPath.Transform, _ => null };
                if (objectTransform is not null) node.SetAttributeValue("CTM", string.Join(" ", objectTransform.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
            }
            source = OfdXmlEnvelope.Capture(node, material, cancellationToken);
            if (element is OfdTextElement text) text.SourceXml = source;
            if (element is OfdImageElement image) image.SourceXml = source;
            if (element is OfdPathElement path) path.SourceXml = source;
        }
        return element;
    }
    private static string TransformAnnotationClips(string? original, XNamespace ns, double[] outer,
        double minX, double minY, double outerX, double outerY, double width, double height, double imageWidth, double imageHeight,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(16384L + 16L * ns.NamespaceName.Length, token);
        var clips = new XElement(ns + "Clips");
        if (!string.IsNullOrWhiteSpace(original))
        {
            material.GuardFragmentXml(original!, token); material.ChargeCopy(24L * original!.Length, token); material.ChargeWork(2L * original.Length, token);
            material.PrepareReadXmlParser(checked(2L * original!.Length), token);
            material.ReserveReadModel(32768L + 32L * original.Length, token);
            var source = OfdXmlEnvelope.Parse(original!);
            if (source.Name.Namespace != ns || !OfdGraphicXmlContract.HasKnownChildren(source) || OfdGraphicXmlContract.HasUnsupportedReferences(source) ||
                source.DescendantsAndSelf().Attributes().Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                throw new NotSupportedException("Unsupported annotation clip cannot be flattened safely.");
            foreach (var region in OfdClipGeometry.Read(source, material, token))
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(512, token);
                var area = new XElement(ns + "Area");
                foreach (var path in region.Paths)
                {
                    token.ThrowIfCancellationRequested(); material.ChargeCopy(16384L + 16L * path.AbbreviatedData.Length, token);
                    var inner = path.Transform ?? new double[] { 1, 0, 0, 1, 0, 0 };
                    var matrix = new[] { outer[0] * inner[0] + outer[2] * inner[1], outer[1] * inner[0] + outer[3] * inner[1],
                        outer[0] * inner[2] + outer[2] * inner[3], outer[1] * inner[2] + outer[3] * inner[3],
                        outer[0] * inner[4] + outer[2] * inner[5] - minX, outer[1] * inner[4] + outer[3] * inner[5] - minY };
                    area.Add(ClipPath(path.AbbreviatedData, matrix, region.EvenOdd));
                }
                clips.Add(new XElement(ns + "Clip", area));
            }
        }
        if (imageWidth > 0 && imageHeight > 0)
        {
            var imageRectangle = $"M 0 0 L {Number(imageWidth)} 0 L {Number(imageWidth)} {Number(imageHeight)} L 0 {Number(imageHeight)} C";
            clips.Add(new XElement(ns + "Clip", new XElement(ns + "Area", ClipPath(imageRectangle,
                new[] { outer[0], outer[1], outer[2], outer[3], -minX, -minY }, false))));
        }
        // Separate Clip means intersection with existing clipping, not union.
        var rectangle = $"M 0 0 L {Number(width)} 0 L {Number(width)} {Number(height)} L 0 {Number(height)} C";
        clips.Add(new XElement(ns + "Clip", new XElement(ns + "Area", ClipPath(rectangle,
            new[] { outer[0], outer[1], outer[2], outer[3], outerX, outerY }, false))));
        return OfdXmlEnvelope.Capture(clips, material, token);

        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        XElement ClipPath(string data, double[] matrix, bool evenOdd) => new(ns + "Path",
            new XAttribute("ID", "annotation-clip"), new XAttribute("CTM", string.Join(" ", matrix.Select(Number))),
            evenOdd ? new XAttribute("Rule", "Even-Odd") : null, new XElement(ns + "AbbreviatedData", data));
    }

}
