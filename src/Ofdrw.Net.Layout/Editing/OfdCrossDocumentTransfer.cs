using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

// A new context is required for every occurrence of a source document. Nothing
// reader-owned (paths, slots, IDs or projection baselines) crosses this boundary.
internal sealed class OfdCrossDocumentTransfer
{
    private readonly OfdDocumentPackage source;
    private readonly OfdDocumentPackage destination;
    private readonly IReadOnlyDictionary<OfdFontResource, OfdFontResource> fonts;
    private readonly IReadOnlyList<OfdPage> pages;
    private readonly bool includeTemplates;
    private readonly OfdModelMaterialization material;
    private readonly CancellationToken token;
    private readonly Dictionary<object, object> map = new(OfdReferenceIdentity.Instance);

    internal OfdCrossDocumentTransfer(OfdDocumentPackage source, OfdDocumentPackage destination,
        IReadOnlyDictionary<OfdFontResource, OfdFontResource> fonts, IReadOnlyList<OfdPage> pages,
        bool includeTemplates, OfdModelMaterialization material, CancellationToken token)
    {
        this.source = source; this.destination = destination; this.fonts = fonts;
        this.pages = pages; this.includeTemplates = includeTemplates; this.material = material; this.token = token;
    }

    internal void CopyResources()
    {
        var available = source.CompositeResources?.Items ?? new List<OfdCompositeResource>();
        var owned = new HashSet<OfdCompositeResource>(available);
        var reachable = new HashSet<OfdCompositeResource>();
        var pending = new Stack<OfdCompositeResource>();
        foreach (var page in pages)
        {
            token.ThrowIfCancellationRequested();
            foreach (var element in page.Elements) Add(element);
            if (includeTemplates) foreach (var template in page.Templates) foreach (var element in template.Elements) Add(element);
            foreach (var annotation in page.Annotations?.Items ?? Enumerable.Empty<OfdAnnotation>())
                foreach (var element in annotation.Appearance.Elements) Add(element);
        }
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            var resource = pending.Pop();
            foreach (var child in resource.Elements) Add(child);
        }
        if (reachable.Count == 0) return;
        destination.CompositeResources ??= new OfdCompositeResourceCollection();
        foreach (var resource in available.Where(reachable.Contains))
        {
            ValidateResource(resource);
            material.ChargeCopy(512L + 32L * resource.Elements.Count, token);
            var copy = new OfdCompositeResource { WidthMillimeters = resource.WidthMillimeters,
                HeightMillimeters = resource.HeightMillimeters };
            destination.CompositeResources.Items.Add(copy);
            map.Add(resource, copy);
        }
        foreach (var resource in available.Where(reachable.Contains))
        {
            var copy = (OfdCompositeResource)map[resource];
            foreach (var child in resource.Elements) copy.Elements.Add(CopyElement(child, true));
        }
        void Add(OfdElement element)
        {
            if (element is not OfdCompositeElement placement) return;
            if (placement.Resource is null || !owned.Contains(placement.Resource))
                throw new InvalidDataException("Composite placement is outside the source document.");
            if (reachable.Add(placement.Resource)) pending.Push(placement.Resource);
        }
    }

    internal OfdElement CopyElement(OfdElement element, bool requireKnownAttributes)
    {
        token.ThrowIfCancellationRequested();
        var closedTransfer = requireKnownAttributes || element.Clips is not null || element is OfdCompositeElement;
        if (closedTransfer && (element.RenderPaintError is not null || element is OfdImageElement { SourceMediaAmbiguous: true }))
            throw new NotSupportedException("Unresolved source paint or image media cannot be transferred.");
        if (closedTransfer && element is OfdTextElement sourceText && sourceText.FontResourceId is not null &&
            OfdFontSelection.Resolve(source.Fonts, sourceText) is null)
            throw new NotSupportedException("Unresolved text font cannot be transferred.");
        if (element is OfdCompositeElement placement)
        {
            ValidatePlacement(placement);
            OfdDocumentMerger.ValidateStandaloneElement(element, true);
        }
        else OfdDocumentMerger.ValidateStandaloneElement(element, requireKnownAttributes || element.Clips is not null);
        ValidateClips(element);
        var copy = OfdModelCloner.CloneElement(element, destination.Options.Namespace,
            material: material, cancellationToken: token, mapping: map);
        copy.ObjectId = null;
        copy.RenderSourceXml = null; copy.RenderSourceAuthority = null; copy.RenderPaintError = null;
        if (copy is OfdCompositeElement composite) composite.SourceXml = null;
        else ResetPrimitiveId(copy);
        if (copy is OfdPathElement path) path.ReaderAttachmentActionId = null;
        if (copy is OfdTextElement text)
        {
            var original = OfdFontSelection.Resolve(source.Fonts, (OfdTextElement)element);
            if (original is null) text.FontResourceId = null;
            else { text.FontResourceId = fonts[original].Id; text.FontName = fonts[original].FontName; }
        }
        if (element.Clips is { } clips)
        {
            var fresh = new OfdClipCollection();
            foreach (var clip in clips.Items)
            {
                var next = new OfdClip();
                foreach (var area in clip.Areas)
                {
                    var clipPath = (OfdPathElement)OfdModelCloner.CloneElement(area.Path, destination.Options.Namespace,
                        material: material, cancellationToken: token);
                    clipPath.ObjectId = null; clipPath.SourceXml = null;
                    clipPath.RenderSourceXml = null; clipPath.RenderSourceAuthority = null; clipPath.RenderPaintError = null;
                    next.Areas.Add(new OfdClipArea { Transform = area.Transform?.ToArray(), Path = clipPath });
                }
                fresh.Items.Add(next);
            }
            copy.ClippingXml = null;
            copy.ClipOwnership.Value = fresh; copy.ClipOwnership.Original = null;
            copy.ClipOwnership.Baseline = null; copy.ClipOwnership.Raw = null;
        }
        return copy;
    }

    internal void CopyAnnotations(OfdPage from, OfdPage to)
    {
        if (from.Annotations is not { } collection) return;
        if (from.AnnotationOwnership.HasUnclaimedSource || from.AnnotationOwnership.UnsafeDeletionEntries.Count != 0)
            throw new NotSupportedException("Unclaimed annotation source slots cannot be transferred.");
        material.ChargeCopy(from.AnnotationAppearances.Count, token);
        var projected = new bool[from.AnnotationAppearances.Count];
        foreach (var annotation in collection.Items)
        {
            if (annotation.Source is not { } owner) continue;
            if (owner.ProjectionStart < 0 || owner.ProjectionCount < 0 ||
                owner.ProjectionStart > projected.Length - owner.ProjectionCount)
                throw new InvalidDataException("Annotation projection range is invalid.");
            for (var index = owner.ProjectionStart; index < owner.ProjectionStart + owner.ProjectionCount; index++)
            {
                if (projected[index]) throw new InvalidDataException("Annotation projections overlap.");
                projected[index] = true;
            }
        }
        if (projected.Any(covered => !covered))
            throw new NotSupportedException("Unclaimed annotation projection cannot be transferred.");
        to.Annotations = new OfdAnnotationCollection();
        foreach (var annotation in collection.Items)
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(512L + annotation.Appearance.Elements.Count * 128L, token);
            if (annotation.AttachmentId is not null || annotation.Type == OfdAnnotationType.Link)
                throw new NotSupportedException("Annotation attachment and navigation actions require a separate remap.");
            ValidateAnnotation(annotation);
            var appearance = annotation.Appearance;
            var copy = new OfdAnnotation
            {
                Type = annotation.Type, Creator = annotation.Creator, LastModifiedDate = annotation.LastModifiedDate,
                Subtype = annotation.Subtype, Visible = annotation.Visible, Print = annotation.Print,
                NoZoom = annotation.NoZoom, NoRotate = annotation.NoRotate, ReadOnly = annotation.ReadOnly,
                Remark = annotation.Remark,
                Appearance = new OfdAnnotationAppearance { XMillimeters = appearance.XMillimeters,
                    YMillimeters = appearance.YMillimeters, WidthMillimeters = appearance.WidthMillimeters,
                    HeightMillimeters = appearance.HeightMillimeters, Transform = appearance.Transform?.ToArray() }
            };
            foreach (var element in appearance.Elements)
            {
                ValidateFreshAnnotationPrimitive(element);
                var primitive = CopyElement(element, true);
                // A new PageAnnot owns typed local definitions, never imported XML.
                if (primitive is OfdTextElement text) text.SourceXml = null;
                if (primitive is OfdImageElement image) image.SourceXml = null;
                if (primitive is OfdPathElement path) path.SourceXml = null;
                copy.Appearance.Elements.Add(primitive);
            }
            to.Annotations.Items.Add(copy);
        }
    }

    private void ValidateResource(OfdCompositeResource resource)
    {
        if (!resource.SourceContainerSupported) throw new NotSupportedException("Unsupported composite source container.");
        if (resource.SourceXml is null) return;
        var root = StrictXml(resource.SourceXml);
        if (root.Name.LocalName != "CompositeGraphicUnit" || !OnlyAttributes(root, "ID", "Width", "Height") ||
            root.Elements().Count() != resource.Elements.Count ||
            root.Elements().Any(child => child.Name.Namespace != root.Name.Namespace ||
                child.Name.LocalName is not ("TextObject" or "ImageObject" or "PathObject" or "CompositeObject")))
            throw new NotSupportedException("Composite source XML has unmodeled content.");
    }

    private void ValidatePlacement(OfdCompositeElement placement)
    {
        if (!map.ContainsKey(placement.Resource)) throw new InvalidDataException("Composite resource closure was not copied.");
        if (placement.SourceXml is null) return;
        var root = StrictXml(placement.SourceXml);
        if (root.Name.LocalName != "CompositeObject" || !OnlyAttributes(root, "ID", "ResourceID", "Boundary", "CTM") ||
            root.Elements().Any(child => child.Name != root.Name.Namespace + "Clips"))
            throw new NotSupportedException("Composite placement has unmodeled source XML.");
    }

    private void ValidateClips(OfdElement element)
    {
        if (element.Clips is not { } clips) return;
        if (clips.Items.Count == 0) throw new NotSupportedException("Empty typed clip cannot be transferred.");
        if (element.ClippingXml is not null) ValidateClipXml(element.ClippingXml);
        foreach (var clip in clips.Items)
        {
            if (clip.Areas.Count == 0) throw new NotSupportedException("Empty clip cannot be transferred.");
            foreach (var area in clip.Areas)
            {
                if (area.Path is null || area.Path.Clips is not null || area.Path.ClippingXml is not null ||
                    area.Path.ObjectId is not null || area.Path.LayerId is not null || !area.Path.Fill || area.Path.Stroke)
                    throw new NotSupportedException("Clip Area must have one local filled path.");
                if (area.Path.SourceXml is not null) ValidateClipPathXml(area.Path.SourceXml);
            }
        }
    }

    private static void ValidateClipXml(string xml)
    {
        var root = StrictXml(xml);
        if (root.Name.LocalName != "Clips" || !OnlyAttributes(root) || root.Elements().Any(clip =>
            clip.Name != root.Name.Namespace + "Clip" || !OnlyAttributes(clip) || clip.Elements().Any(area =>
                area.Name != root.Name.Namespace + "Area" || !OnlyAttributes(area, "CTM") ||
                area.Elements().Count() != 1 || !ValidClipPath(area.Elements().Single()))))
            throw new NotSupportedException("Clip source XML has unmodeled content.");
    }

    private static void ValidateClipPathXml(string xml)
    {
        if (!ValidClipPath(StrictXml(xml))) throw new NotSupportedException("Clip Path has unmodeled source XML.");
    }

    private static bool ValidClipPath(XElement path) => path.Name.LocalName == "Path" &&
        OnlyAttributes(path, "ID", "Boundary", "CTM", "Stroke", "Fill", "Rule") &&
        path.Elements().Count() <= 1 && path.Elements().All(node => node.Name == path.Name.Namespace + "AbbreviatedData" &&
            node.Attributes().All(attribute => attribute.IsNamespaceDeclaration) && !node.HasElements);

    private static void ValidateAnnotation(OfdAnnotation annotation)
    {
        if (annotation.SourceXml is null) return;
        var root = StrictXml(annotation.SourceXml);
        if (root.Name.LocalName != "Annot" || !OnlyAttributes(root, "ID", "Type", "Creator", "LastModDate", "Subtype", "Visible", "Print", "NoZoom", "NoRotate", "ReadOnly") ||
            root.Elements().Count(node => node.Name.LocalName == "Appearance") != 1 ||
            root.Elements().Any(node => node.Name.Namespace != root.Name.Namespace || node.Name.LocalName is not ("Appearance" or "Remark")) ||
            root.Elements().Count(node => node.Name.LocalName == "Remark") > 1)
            throw new NotSupportedException("Annotation source XML has unmodeled content or actions.");
        var appearance = root.Elements().Single(node => node.Name.LocalName == "Appearance");
        if (!OnlyAttributes(appearance, "Boundary", "CTM") ||
            appearance.Elements().Count() != annotation.Appearance.Elements.Count ||
            appearance.Elements().Any(node => node.Name.Namespace != root.Name.Namespace || node.Name.LocalName is not ("TextObject" or "ImageObject" or "PathObject")))
            throw new NotSupportedException("Annotation Appearance has unmodeled content.");
        var remark = root.Elements().FirstOrDefault(node => node.Name.LocalName == "Remark");
        if (remark is not null && (remark.HasElements || remark.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)))
            throw new NotSupportedException("Annotation Remark has unmodeled content.");
    }

    private static void ValidateFreshAnnotationPrimitive(OfdElement element)
    {
        var xml = element switch
        {
            OfdTextElement text => text.SourceXml,
            OfdImageElement image => image.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        if (xml is null) return;
        var root = StrictXml(xml);
        var admitted = element switch
        {
            OfdTextElement text => !text.SourceInheritedTextCodeAxes && root.Name.LocalName == "TextObject" &&
                OnlyAttributes(root, "ID", "Boundary", "Font", "Size", "Weight", "Italic", "CTM") &&
                root.Elements(root.Name.Namespace + "TextCode").Count() == (text.Runs.Count == 0 ? 1 : text.Runs.Count) &&
                root.Elements().All(child => child.Name == root.Name.Namespace + "TextCode" ?
                    child.Attributes().All(attribute => attribute.IsNamespaceDeclaration || attribute.Name.Namespace == XNamespace.None &&
                        attribute.Name.LocalName is "X" or "Y" or "DeltaX" or "DeltaY") && !child.HasElements :
                    child.Name == root.Name.Namespace + "Clips" ? text.Clips is not null && ValidClips(child) :
                    child.Name == root.Name.Namespace + "FillColor" && OnlyAttributes(child, "Value", "Alpha")),
            OfdImageElement image => !image.SourceMediaAmbiguous && root.Name.LocalName == "ImageObject" &&
                OnlyAttributes(root, "ID", "Boundary", "CTM", "ResourceID", "Alpha") &&
                root.Elements().All(child => child.Name == root.Name.Namespace + "Clips" && image.Clips is not null && ValidClips(child)),
            OfdPathElement => root.Name.LocalName == "PathObject" &&
                OnlyAttributes(root, "ID", "Boundary", "CTM", "LineWidth", "Stroke", "Fill", "Rule") &&
                root.Elements().All(child => child.Name == root.Name.Namespace + "AbbreviatedData" ?
                    child.Attributes().All(attribute => attribute.IsNamespaceDeclaration) && !child.HasElements :
                    child.Name == root.Name.Namespace + "Clips" ? element.Clips is not null && ValidClips(child) :
                    child.Name == root.Name.Namespace + "StrokeColor" || child.Name == root.Name.Namespace + "FillColor" ?
                        OnlyAttributes(child, "Value", "Alpha") : false),
            _ => false
        };
        if (!admitted) throw new NotSupportedException("Annotation primitive has source paint that cannot be reconstructed from typed fields.");
    }

    private static bool ValidClips(XElement root)
    {
        try { ValidateClipXml(root.ToString(SaveOptions.DisableFormatting)); return true; }
        catch (NotSupportedException) { return false; }
    }

    private static XElement StrictXml(string xml)
    {
        var root = OfdXmlEnvelope.Parse(xml);
        if (root.DescendantNodesAndSelf().Any(node => node is XComment or XProcessingInstruction) ||
            root.DescendantsAndSelf().Any(node => node.Name.Namespace != root.Name.Namespace) ||
            root.DescendantsAndSelf().Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.Namespace != XNamespace.None))
            throw new NotSupportedException("Foreign XML or comments cannot be transferred into a fresh owner.");
        return root;
    }

    private static bool OnlyAttributes(XElement node, params string[] names) =>
        node.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && names.Contains(attribute.Name.LocalName)) &&
        node.Nodes().All(child => child is XElement || child is XText text && string.IsNullOrWhiteSpace(text.Value));

    private void ResetPrimitiveId(OfdElement copy)
    {
        var xml = copy switch
        {
            OfdTextElement text => text.SourceXml,
            OfdImageElement image => image.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        if (xml is null) return;
        var root = OfdXmlEnvelope.Parse(xml); root.SetAttributeValue("ID", null);
        var rebuilt = OfdXmlEnvelope.Capture(root, material, token);
        if (copy is OfdTextElement t) t.SourceXml = rebuilt;
        if (copy is OfdImageElement i) i.SourceXml = rebuilt;
        if (copy is OfdPathElement p) p.SourceXml = rebuilt;
    }
}
