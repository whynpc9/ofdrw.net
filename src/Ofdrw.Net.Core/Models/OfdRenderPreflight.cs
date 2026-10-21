using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Shared admission for the page renderers. Imported paint is resolved against
// its original XML, while opaque top-level drawing is never silently skipped.
internal static class OfdRenderPreflight
{
    private const string PaintFailureKey = "Ofdrw.Net.RenderPaintFailure";

    internal static bool IsPaintFailure(Exception exception) => exception.Data.Contains(PaintFailureKey);

    internal static NotSupportedException PaintFailure(string message)
    {
        var exception = new NotSupportedException(message);
        exception.Data[PaintFailureKey] = true;
        return exception;
    }

    internal static void Validate(IReadOnlyList<OfdElement> elements)
    {
        foreach (var element in elements)
        {
            if (!OfdGraphicXmlContract.IsVisible(element)) continue;
            if (element is OfdRawElement { LocalName: "UnmodeledAnnotationMetadata" } metadata &&
                metadata.ReaderAnnotationMetadataXml is not null &&
                metadata.ReaderAnnotationMetadataXml == metadata.Xml) continue;
            // The reader retains opaque annotation index records in this
            // synthetic metadata envelope. They are not page drawings; an
            // unsupported annotation appearance has its own failure marker.
            if (element is OfdRawElement { LocalName: "Annotations", Xml: var annotationXml } annotationIndex &&
                annotationIndex.ReaderAnnotationMetadataXml == annotationXml &&
                annotationXml.StartsWith("<UnmodeledAnnotations>", StringComparison.Ordinal) &&
                annotationXml.EndsWith("</UnmodeledAnnotations>", StringComparison.Ordinal)) continue;
            if (element is OfdRawElement raw)
                throw new NotSupportedException($"Opaque OFD drawing '{raw.LocalName}' cannot be rendered.");
            if (element is not (OfdTextElement or OfdImageElement or OfdPathElement))
                throw new NotSupportedException("Unknown OFD drawing cannot be rendered.");
            if (element is OfdImageElement unresolvedImage && unresolvedImage.Data is not { Length: > 0 })
                throw PaintFailure("OFD image resource has no resolved payload.");
            var source = element switch
            {
                OfdTextElement sourceText => sourceText.SourceXml,
                OfdImageElement image => image.SourceXml,
                OfdPathElement path => path.SourceXml,
                _ => null
            };
            if (element.RenderPaintError is { Length: > 0 } error)
                throw PaintFailure(error);
            if (element.RenderSourceXml is not null)
            {
                if (element.RenderSourceAuthority != source)
                    throw new NotSupportedException("Imported drawing changed after paint resolution.");
                source = element.RenderSourceXml;
            }
            if (string.IsNullOrWhiteSpace(source)) continue;
            if (element is OfdPathElement)
            {
                try
                {
                    var root = OfdXmlEnvelope.Parse(source!);
                    // Annotation artwork can retain unrelated extension metadata,
                    // while its source paint must be resolved and well formed.
                    if (OfdGraphicXmlContract.HasUnsupportedReferences(root))
                        throw new NotSupportedException("OFD path paint has an unresolved resource reference.");
                    foreach (var color in root.Elements().Where(node => node.Name.LocalName is "FillColor" or "StrokeColor"))
                        _ = OfdTextPaint.Color(color, OfdColor.Black);
                }
                catch (Exception paintError) when (paintError is NotSupportedException or InvalidDataException)
                { throw PaintFailure(paintError.Message); }
            }
            if (element is OfdTextElement text)
            {
                try
                {
                    var root = OfdXmlEnvelope.Parse(source!);
                    _ = OfdTextPaint.Read(text);
                    if (root.Attribute("ReadDirection") is { Value: not "0" } ||
                        root.Attribute("CharDirection") is { Value: not "0" } ||
                        root.Attribute("Alpha") is { Value: not "255" })
                        throw new NotSupportedException("OFD text direction or object alpha is outside this render profile.");
                }
                catch (Exception paintError) when (paintError is NotSupportedException or InvalidDataException)
                { throw PaintFailure(paintError.Message); }
            }
        }
    }
}
