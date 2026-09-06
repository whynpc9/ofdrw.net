using System;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdModelCloner
{
    internal static OfdPage ClonePage(OfdPage source, bool keepSourcePath = false, bool clonePayloads = true)
    {
        var page = new OfdPage
        {
            Id = keepSourcePath ? source.Id : null,
            SourceEntryPath = keepSourcePath ? source.SourceEntryPath : null,
            Index = source.Index, XMillimeters = source.XMillimeters, YMillimeters = source.YMillimeters,
            WidthMillimeters = source.WidthMillimeters, HeightMillimeters = source.HeightMillimeters
        };
        foreach (var element in source.Elements) page.Elements.Add(CloneElement(element, clonePayloads: clonePayloads));
        foreach (var template in source.Templates)
        {
            var clone = new OfdTemplateContent
            {
                TemplateId = template.TemplateId, ZOrder = template.ZOrder, BaseLocation = template.BaseLocation
            };
            foreach (var element in template.Elements) clone.Elements.Add(CloneElement(element, clonePayloads: clonePayloads));
            page.Templates.Add(clone);
        }
        page.PreservedPageElements.AddRange(source.PreservedPageElements);
        return page;
    }

    internal static OfdElement CloneElement(OfdElement source, string? targetNamespace = null, bool clonePayloads = true)
    {
        OfdElement result;
        switch (source)
        {
            case OfdTextElement text:
                var clone = new OfdTextElement
                {
                    Text = text.Text, FontName = text.FontName, FontResourceId = text.FontResourceId,
                    FontSizeMillimeters = text.FontSizeMillimeters, Transform = text.Transform?.ToArray(),
                    FillColor = text.FillColor, SourceXml = CloneXml(text.SourceXml, targetNamespace)
                };
                foreach (var run in text.Runs) clone.Runs.Add(new OfdTextRun
                {
                    Text = run.Text, XMillimeters = run.XMillimeters, YMillimeters = run.YMillimeters,
                    DeltaX = run.DeltaX, DeltaY = run.DeltaY
                });
                result = clone;
                break;
            case OfdImageElement image:
                result = new OfdImageElement
                {
                    ResourceId = image.ResourceId, FileName = image.FileName, MediaType = image.MediaType,
                    Data = clonePayloads ? image.Data.ToArray() : image.Data, Transform = image.Transform?.ToArray(), Alpha = image.Alpha,
                    ClipsXml = CloneXml(image.ClipsXml, targetNamespace), SourceXml = CloneXml(image.SourceXml, targetNamespace)
                };
                break;
            case OfdPathElement path:
                result = new OfdPathElement
                {
                    AbbreviatedData = path.AbbreviatedData, Transform = path.Transform?.ToArray(),
                    LineWidthMillimeters = path.LineWidthMillimeters, Stroke = path.Stroke, Fill = path.Fill,
                    StrokeColor = path.StrokeColor, FillColor = path.FillColor,
                    SourceXml = CloneXml(path.SourceXml, targetNamespace)
                };
                break;
            case OfdRawElement raw:
                result = new OfdRawElement { LocalName = raw.LocalName, Xml = CloneXml(raw.Xml, targetNamespace) ?? string.Empty };
                break;
            default:
                throw new NotSupportedException($"Cannot clone OFD element '{source.GetType().Name}'.");
        }
        result.ObjectId = source.ObjectId;
        result.LayerId = source.LayerId;
        result.LayerType = source.LayerType;
        result.XMillimeters = source.XMillimeters;
        result.YMillimeters = source.YMillimeters;
        result.WidthMillimeters = source.WidthMillimeters;
        result.HeightMillimeters = source.HeightMillimeters;
        return result;
    }

    private static string? CloneXml(string? xml, string? targetNamespace)
    {
        if (string.IsNullOrWhiteSpace(xml) || targetNamespace is null) return xml;
        var root = XElement.Parse(xml!, LoadOptions.PreserveWhitespace);
        var originalNamespace = root.Name.Namespace;
        foreach (var node in root.DescendantsAndSelf())
        {
            if (node.Name.Namespace == originalNamespace) node.Name = XName.Get(node.Name.LocalName, targetNamespace);
            node.Attributes().Where(attribute => attribute.IsNamespaceDeclaration && attribute.Value == originalNamespace.NamespaceName).Remove();
        }
        return root.ToString(SaveOptions.DisableFormatting);
    }
}
