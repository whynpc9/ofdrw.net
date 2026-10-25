using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
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
            if (element is OfdPathElement drawingPath &&
                OfdPathDataSyntax.IsValidLiteral(drawingPath.AbbreviatedData, CancellationToken.None) &&
                !OfdPathDataSyntax.IsValidLiteral(drawingPath.AbbreviatedData, CancellationToken.None, requireDrawableSegment: true))
                throw PaintFailure("OFD path has no valid drawable segment.");
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
            if (element is OfdPathElement pathElement)
            {
                try
                {
                    var root = OfdXmlEnvelope.Parse(source!);
                    // Annotation artwork can retain unrelated extension metadata,
                    // while its source paint must be resolved and well formed.
                    var attachmentAction = pathElement.ReaderAttachmentActionId is { } attachmentId &&
                        HasClosedAttachmentAction(root, attachmentId);
                    var hasActionNodes = root.Descendants().Any(node => node.Name.LocalName is "Actions" or "Action" or "GotoA");
                    if (((hasActionNodes || pathElement.ReaderAttachmentActionId is not null) && !attachmentAction) ||
                        OfdGraphicXmlContract.HasUnsupportedReferences(root, attachmentAction))
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
                    if (!MatchesInteger(root.Attribute("ReadDirection"), 0) ||
                        !MatchesInteger(root.Attribute("CharDirection"), 0) ||
                        !MatchesInteger(root.Attribute("Alpha"), 255))
                        throw new NotSupportedException("OFD text direction or object alpha is outside this render profile.");
                }
                catch (Exception paintError) when (paintError is NotSupportedException or InvalidDataException)
                { throw PaintFailure(paintError.Message); }
            }
        }
    }

    internal static bool MatchesInteger(XAttribute? attribute, int expected) =>
        attribute is null || int.TryParse(attribute.Value, NumberStyles.Integer,
            CultureInfo.InvariantCulture, out var value) && value == expected;

    private static bool HasClosedAttachmentAction(XElement root, string attachmentId)
    {
        var ns = root.Name.Namespace;
        if (root.Name != ns + "PathObject" ||
            ns.NamespaceName is not (Ofdrw.Net.Core.Constants.OfdConstants.Namespace or Ofdrw.Net.Core.Constants.OfdConstants.StandardNamespace) ||
            !OfdGraphicXmlContract.HasKnownChildren(root) ||
            OfdSourceIdentity.ParseId(attachmentId) == 0) return false;
        var actions = root.Elements(ns + "Actions").ToArray();
        if (actions.Length != 1 || actions[0].Attributes().Any(a => !a.IsNamespaceDeclaration) ||
            actions[0].Elements().Count() != 1) return false;
        var action = actions[0].Element(ns + "Action");
        if (action?.Attribute("Event")?.Value != "CLICK" ||
            action.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "Event") ||
            action.Elements().Count() != 1) return false;
        var gotoAction = action.Element(ns + "GotoA");
        var newWindow = gotoAction?.Attribute("NewWindow")?.Value;
        return gotoAction?.Attribute("AttachID")?.Value == attachmentId &&
            !gotoAction.HasElements && !gotoAction.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value)) &&
            (newWindow is null or "true" or "false" or "1" or "0") &&
            gotoAction.Attributes().All(a => a.IsNamespaceDeclaration || a.Name == "AttachID" || a.Name == "NewWindow") &&
            root.DescendantsAndSelf().Attributes().Count(a => a.Name == "AttachID") == 1;
    }
}
