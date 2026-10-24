using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdModelCloner
{
    internal static OfdPage ClonePage(OfdPage source, bool keepSourcePath = false, bool clonePayloads = true,
        OfdModelMaterialization? material = null, CancellationToken cancellationToken = default, IDictionary<object, object>? mapping = null,
        OfdPayloadCloneSession? payloads = null)
    {
        if (material is null)
        {
            material = new OfdModelMaterialization(new OfdModelLimits());
            var input = new OfdDocumentPackage(); input.Pages.Add(source); material.InspectStorage(input, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (mapping is null) material.ChargeCopy(256, cancellationToken);
        var cloneMap = mapping ?? new Dictionary<object, object>();
        material.Structure.Node(1);
        material.ChargeCopy(1024L + (long)(source.Elements.Count + source.AnnotationAppearances.Count + source.Templates.Count + source.PreservedPageElements.Count) * IntPtr.Size * 2, cancellationToken);
        var page = new OfdPage
        {
            Id = keepSourcePath ? source.Id : null,
            SourceEntryPath = keepSourcePath ? source.SourceEntryPath : null,
            Index = source.Index, XMillimeters = source.XMillimeters, YMillimeters = source.YMillimeters,
            WidthMillimeters = source.WidthMillimeters, HeightMillimeters = source.HeightMillimeters
        };
        foreach (var element in source.Elements) page.Elements.Add(CloneElement(element, clonePayloads: clonePayloads, material: material, cancellationToken: cancellationToken, mapping: cloneMap, payloads: payloads));
        foreach (var template in source.Templates)
        {
            cancellationToken.ThrowIfCancellationRequested(); material.Structure.Node(2);
            material.ChargeCopy(256L + (long)template.Elements.Count * IntPtr.Size * 2, cancellationToken);
            var clone = new OfdTemplateContent
            {
                TemplateId = template.TemplateId, ZOrder = template.ZOrder, BaseLocation = template.BaseLocation
            };
            foreach (var element in template.Elements) clone.Elements.Add(CloneElement(element, clonePayloads: clonePayloads, material: material, cancellationToken: cancellationToken, mapping: cloneMap, payloads: payloads));
            page.Templates.Add(clone);
            { cloneMap.Add(template, clone); cloneMap.Add(template.Elements, clone.Elements); }
        }
        foreach (var element in source.AnnotationAppearances) page.AnnotationAppearances.Add(CloneElement(element, clonePayloads: clonePayloads, material: material, cancellationToken: cancellationToken, mapping: cloneMap, payloads: payloads));
        page.PreservedPageElements.AddRange(source.PreservedPageElements);
        {
            cloneMap.Add(source, page); cloneMap.Add(source.Elements, page.Elements); cloneMap.Add(source.AnnotationAppearances, page.AnnotationAppearances);
            cloneMap.Add(source.Templates, page.Templates); cloneMap.Add(source.PreservedPageElements, page.PreservedPageElements);
        }
        if (keepSourcePath || source.Annotations is not null || source.AnnotationOwnership.Original is not null)
            OfdAnnotationCloner.Clone(source, page, cloneMap, clonePayloads, material, cancellationToken, payloads);
        else source.AnnotationOwnership.Validate(source, material, cancellationToken);
        return page;
    }

    internal static OfdElement CloneElement(OfdElement source, string? targetNamespace = null, bool clonePayloads = true,
        OfdModelMaterialization? material = null, CancellationToken cancellationToken = default, IDictionary<object, object>? mapping = null,
        OfdPayloadCloneSession? payloads = null)
    {
        if (material is null)
        {
            material = new OfdModelMaterialization(new OfdModelLimits());
            var input = new OfdDocumentPackage(); var page = new OfdPage(); page.Elements.Add(source); input.Pages.Add(page);
            material.InspectStorage(input, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested(); material.Structure.Node(1);
        if (mapping is not null) material.ChargeCopy(256, cancellationToken);
        OfdElement result;
        switch (source)
        {
            case OfdTextElement text:
                if (mapping is not null) material.ChargeCopy(128L + (long)text.Runs.Count * (128 + IntPtr.Size * 2), cancellationToken);
                var clone = new OfdTextElement
                {
                    Text = text.Text, FontName = text.FontName, FontResourceId = text.FontResourceId,
                    FontSizeMillimeters = text.FontSizeMillimeters, Transform = Matrix(text.Transform),
                    EmitSyntheticBoldStroke = text.EmitSyntheticBoldStroke,
                    SourceInheritedTextCodeAxes = text.SourceInheritedTextCodeAxes,
                    Weight = text.Weight, Italic = text.Italic,
                    FillColor = text.FillColor, SourceXml = CloneXml(text.SourceXml, targetNamespace, material, cancellationToken)
                };
                foreach (var run in text.Runs)
                {
                    cancellationToken.ThrowIfCancellationRequested(); material.Structure.Node(2);
                    var clonedRun = new OfdTextRun
                    {
                        Text = run.Text, XMillimeters = run.XMillimeters, YMillimeters = run.YMillimeters,
                        DeltaX = run.DeltaX, DeltaY = run.DeltaY
                    };
                    clone.Runs.Add(clonedRun); mapping?.Add(run, clonedRun);
                }
                mapping?.Add(text.Runs, clone.Runs);
                result = clone;
                break;
            case OfdImageElement image:
                result = new OfdImageElement
                {
                    ResourceId = image.ResourceId, FileName = image.FileName, MediaType = image.MediaType,
                    Data = clonePayloads ? Payload(image.Data) : image.Data, Transform = Matrix(image.Transform), Alpha = image.Alpha,
                    ClipsXml = CloneXml(image.ClipsXml, targetNamespace, material, cancellationToken), SourceXml = CloneXml(image.SourceXml, targetNamespace, material, cancellationToken),
                    SourceMediaAmbiguous = image.SourceMediaAmbiguous
                };
                break;
            case OfdPathElement path:
                result = new OfdPathElement
                {
                    AbbreviatedData = path.AbbreviatedData, Transform = Matrix(path.Transform),
                    FillRule = path.FillRule, LineWidthMillimeters = path.LineWidthMillimeters, Stroke = path.Stroke, Fill = path.Fill,
                    StrokeColor = path.StrokeColor, FillColor = path.FillColor,
                    SourceXml = CloneXml(path.SourceXml, targetNamespace, material, cancellationToken)
                };
                break;
            case OfdCompositeElement composite:
                if (mapping is null || !mapping.TryGetValue(composite.Resource, out var resource))
                    throw new NotSupportedException("Composite cloning requires the package resource closure.");
                result = new OfdCompositeElement { Resource = (OfdCompositeResource)resource, Transform = Matrix(composite.Transform), SourceXml = composite.SourceXml };
                break;
            case OfdRawElement raw:
                result = new OfdRawElement { LocalName = raw.LocalName, Xml = CloneXml(raw.Xml, targetNamespace, material, cancellationToken) ?? string.Empty };
                if (raw.ReaderAnnotationMetadataXml is not null && raw.ReaderAnnotationMetadataXml == raw.Xml)
                    ((OfdRawElement)result).ReaderAnnotationMetadataXml = ((OfdRawElement)result).Xml;
                break;
            default:
                throw new NotSupportedException($"Cannot clone OFD element '{source.GetType().Name}'.");
        }
        result.ClippingXml = CloneXml(source.ClippingXml, targetNamespace, material, cancellationToken);
        result.RenderSourceXml = CloneXml(source.RenderSourceXml, targetNamespace, material, cancellationToken);
        var originalSource = source switch
        {
            OfdTextElement text => text.SourceXml,
            OfdImageElement image => image.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        var clonedSource = result switch
        {
            OfdTextElement text => text.SourceXml,
            OfdImageElement image => image.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        result.RenderSourceAuthority = targetNamespace is not null && source.RenderSourceAuthority is not null &&
            source.RenderSourceAuthority == originalSource ? clonedSource : source.RenderSourceAuthority;
        result.RenderPaintError = source.RenderPaintError;
        result.ObjectId = source.ObjectId;
        result.LayerId = source.LayerId;
        result.LayerType = source.LayerType;
        result.XMillimeters = source.XMillimeters;
        result.YMillimeters = source.YMillimeters;
        result.WidthMillimeters = source.WidthMillimeters;
        result.HeightMillimeters = source.HeightMillimeters;
        mapping?.Add(source, result);
        if (source is OfdRawElement { LocalName: "PageBlock" } sourceBlock &&
            result is OfdRawElement clonedBlock && sourceBlock.PageBlockChildren is { } children &&
            sourceBlock.PageBlockSourceXml == sourceBlock.Xml)
        {
            material.ChargeCopy(256L + (long)children.Count * IntPtr.Size * 2, cancellationToken);
            clonedBlock.PageBlockSourceXml = clonedBlock.Xml;
            clonedBlock.PageBlockChildren = new List<OfdElement>(children.Count);
            mapping?.Add(children, clonedBlock.PageBlockChildren);
            foreach (var child in children)
                clonedBlock.PageBlockChildren.Add(CloneElement(child, targetNamespace, clonePayloads, material,
                    cancellationToken, mapping, payloads));
        }
        OfdClipCloner.Clone(source, result, mapping, material, cancellationToken, clonePayloads, payloads);
        return result;
        double[]? Matrix(double[]? value)
        {
            if (value is null) return null;
            material.ChargeCopy(value.LongLength * sizeof(double), cancellationToken); return value.ToArray();
        }
        byte[] Payload(byte[] value)
        {
            if (payloads is not null) return payloads.Copy(value, material, cancellationToken);
            material.ChargeCopy(value.LongLength, cancellationToken); return value.ToArray();
        }
    }

    private static string? CloneXml(string? xml, string? targetNamespace, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(xml) || targetNamespace is null) return xml;
        material.StoreString(targetNamespace); material.GuardFragmentXml(xml!, cancellationToken);
        var root = OfdXmlEnvelope.Parse(xml!);
        var originalNamespace = root.Name.Namespace;
        var destinationNamespace = XNamespace.Get(targetNamespace);
        foreach (var node in root.DescendantsAndSelf())
        {
            if (node.Name.Namespace == originalNamespace) node.Name = destinationNamespace + node.Name.LocalName;
            node.Attributes().Where(attribute => attribute.IsNamespaceDeclaration && attribute.Value == originalNamespace.NamespaceName).Remove();
        }
        return OfdXmlEnvelope.Capture(root, material, cancellationToken);
    }
}
