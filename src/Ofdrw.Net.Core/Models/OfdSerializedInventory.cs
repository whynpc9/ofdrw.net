using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// The shared field table for input accounting and callback-mutation guards.
// Positions are framed by role, container identity/count and ordered edges.
internal static class OfdSerializedInventory
{
    internal interface ISink
    {
        void Item(string role, object? value, bool identity = false);
        void Bytes(string role, byte[] value, object? owner = null);
        void Matrix(double[]? value);
        void Xml(string? value);
        void Path(string? value);
    }

    internal static void Walk(OfdDocumentPackage package, ISink sink, OfdModelMaterialization material,
        CancellationToken token, bool rejectRepeatedOwners = false, object? scope = null)
    {
        material.ChargeCopy(256, token);
        var owners = rejectRepeatedOwners ? new HashSet<object>(OfdReferenceIdentity.Instance) : null;
        var navigation = new HashSet<object>(OfdReferenceIdentity.Instance);
        var annotations = new HashSet<object>(OfdReferenceIdentity.Instance);
        if (scope is OfdCompositeResourceCollection resourceScope) { Resources(resourceScope); return; }
        if (scope is OfdClipCollection clipScope) { Clips(clipScope); return; }
        if (scope is List<OfdElement> graphics) { Graphics("graphics", graphics); return; }
        if (scope is OfdAnnotationCollection collection) { Annotations(collection); return; }
        Object("package", package); Object("options", package.Options); Object("metadata", package.Options.Metadata);
        S("documentPath", package.DocumentEntryPath); P(package.DocumentEntryPath);
        S("publicResource", package.PublicResourceLocation); P(package.PublicResourceLocation);
        S("documentResource", package.DocumentResourceLocation); P(package.DocumentResourceLocation);
        var options = package.Options; S("docType", options.DocType); S("namespace", options.Namespace);
        S("documentId", options.DocumentId); P(options.DocumentId); V("compression", options.EnableDeflateCompression);
        V("defaultWidth", options.DefaultPageWidthMillimeters); V("defaultHeight", options.DefaultPageHeightMillimeters);
        var metadata = options.Metadata;
        S("title", metadata.Title); S("author", metadata.Author); S("subject", metadata.Subject);
        S("keywords", metadata.Keywords); S("creator", metadata.Creator);
        V("creation", metadata.CreationDate); V("modified", metadata.ModificationDate);
        Container("entries", package.PreservedEntries, package.PreservedEntries.Count);
        foreach (var entry in package.PreservedEntries) { S("entryName", entry.Key); sink.Bytes("entryBytes", entry.Value); }
        Container("fonts", package.Fonts, package.Fonts.Count);
        foreach (var font in package.Fonts)
        {
            Own("font", font); S("fontId", font.Id); S("fontName", font.FontName); S("family", font.FamilyName);
            S("charset", font.Charset); S("fileName", font.FileName); P(font.FileName);
            V("bold", font.Bold); V("italic", font.Italic); sink.Bytes("fontBytes", font.Data);
        }
        Container("attachments", package.Attachments, package.Attachments.Count);
        foreach (var attachment in package.Attachments)
        {
            Own("attachment", attachment); S("attachmentId", attachment.Id); S("name", attachment.Name);
            S("mediaType", attachment.MediaType); S("format", attachment.Format); S("usage", attachment.Usage);
            V("creation", attachment.CreationDate); V("modified", attachment.ModificationDate);
            V("size", attachment.SizeKilobytes); V("visible", attachment.Visible); V("external", attachment.IsExternal);
            S("externalPath", attachment.ExternalPath); P(attachment.ExternalPath); sink.Bytes("attachmentBytes", attachment.Data, attachment);
        }
        Container("tags", package.CustomTags, package.CustomTags.Count);
        foreach (var tag in package.CustomTags) { S("tagKey", tag.Key); S("tagValue", tag.Value); }
        Fragments("common", package.PreservedCommonDataElements); Fragments("document", package.PreservedDocumentElements);
        Fragments("body", package.PreservedDocBodyElements);
        package.ValidateCompositeOwnership(material, token);
        V("originalCompositeOrder", package.OriginalCompositeOrder, true); V("originalCompositeCount", package.OriginalCompositeOrder.Length);
        foreach (var originalComposite in package.OriginalCompositeOrder) V("originalComposite", originalComposite, true);
        Resources(package.CompositeResources);
        if (!ReferenceEquals(package.CompositeResources, package.OriginalCompositeResources)) Resources(package.OriginalCompositeResources);
        package.CompositeBaseline?.VisitInventory(sink);
        Container("pages", package.Pages, package.Pages.Count);
        foreach (var page in package.Pages) Page(page);
        // Historical pages are read by clone/baseline reconciliation too.
        material.ChargeCopy(128L + (long)package.Pages.Count * 32, token);
        var currentPages = new HashSet<OfdPage>(package.Pages);
        foreach (var page in package.SourceIdentity?.OriginalPages ?? Array.Empty<OfdPage>())
            if (currentPages.Add(page)) { V("historicalPage", page, true); Page(page); }
        NavigationOwner("outlines", package.OutlineOwnership); NavigationOwner("bookmarks", package.BookmarkOwnership);
        NavigationOwner("actions", package.ActionOwnership);
        V("sourceEnvelope", package.SourceEnvelope, true); package.SourceEnvelope?.VisitInventory(sink);
        V("annotationIndexSource", package.AnnotationIndexSource, true); package.AnnotationIndexSource?.VisitInventory(sink);
        V("sourceIdentity", package.SourceIdentity, true); package.SourceIdentity?.VisitInventory(sink);
        token.ThrowIfCancellationRequested();

        void V(string role, object? value, bool identity = false) { token.ThrowIfCancellationRequested(); sink.Item(role, value, identity); }
        void S(string role, string? value) => V(role, value);
        void P(string? value) => sink.Path(value);
        void Object(string role, object? value)
        { if (value is null) throw new InvalidDataException("Null serialized owner: " + role); material.Structure.Node(1); V(role, value, true); }
        void Own(string role, object? value)
        { Object(role, value); if (owners is not null) { material.ChargeCopy(64, token); if (!owners.Add(value!)) throw new InvalidDataException("Repeated serialized ownership: " + role); } }
        void Container(string role, object value, int count) { V(role, value, true); V("count", count); }
        void Fragments(string role, List<string> values)
        { Container(role, values, values.Count); foreach (var value in values) { S("fragment", value); sink.Xml(value); } }
        void Page(OfdPage page)
        {
            Own("page", page); S("pageId", page.Id); S("sourcePath", page.SourceEntryPath); P(page.SourceEntryPath);
            V("index", page.Index); V("x", page.XMillimeters); V("y", page.YMillimeters);
            V("width", page.WidthMillimeters); V("height", page.HeightMillimeters);
            Graphics("graphics", page.Elements); Graphics("annotations", page.AnnotationAppearances);
            V("annotationOwnership", page.AnnotationOwnership, true);
            page.AnnotationOwnership.VisitInventory(sink);
            Annotations(page.Annotations);
            Annotations(page.AnnotationOwnership.Original, baselineOnly: true);
            Annotations(page.AnnotationOwnership.BaselineCollection, baselineOnly: true);
            Fragments("pageFragments", page.PreservedPageElements);
            Container("templates", page.Templates, page.Templates.Count);
            foreach (var template in page.Templates)
            {
                Own("template", template); S("templateId", template.TemplateId); S("zOrder", template.ZOrder);
                S("baseLocation", template.BaseLocation); P(template.BaseLocation); Graphics("templateGraphics", template.Elements);
            }
        }
        void Annotations(OfdAnnotationCollection? collection, bool baselineOnly = false)
        {
            V("annotationCollection", collection, true);
            if (collection is null) return;
            if (!annotations.Add(collection))
            {
                if (!baselineOnly && owners is not null) throw new InvalidDataException("Repeated serialized annotation collection ownership.");
                return;
            }
            Container("annotationItems", collection.Items, collection.Items.Count);
            foreach (var annotation in collection.Items)
            {
                V("annotationEdge", annotation, true);
                if (!annotations.Add(annotation))
                {
                    if (!baselineOnly && owners is not null) throw new InvalidDataException("Repeated serialized annotation ownership.");
                    continue;
                }
                Own("annotation", annotation); S("annotationId", annotation.Id); V("annotationType", annotation.Type);
                S("creator", annotation.Creator); V("lastModifiedDate", annotation.LastModifiedDate);
                S("subtype", annotation.Subtype); V("visible", annotation.Visible); V("print", annotation.Print);
                V("noZoom", annotation.NoZoom); V("noRotate", annotation.NoRotate); V("readOnly", annotation.ReadOnly);
                S("remark", annotation.Remark); Source(annotation.SourceXml);
                var appearance = annotation.Appearance;
                Own("appearance", appearance); V("x", appearance.XMillimeters); V("y", appearance.YMillimeters);
                V("width", appearance.WidthMillimeters); V("height", appearance.HeightMillimeters);
                sink.Matrix(appearance.Transform); Source(appearance.SourceXml);
                Graphics("localGraphics", appearance.Elements);
                if (scope is null) { V("annotationSource", annotation.Source, true); annotation.Source?.VisitInventory(sink); }
            }
        }
        void Graphics(string role, List<OfdElement> elements)
        {
            Container(role, elements, elements.Count);
            foreach (var element in elements) Graphic(element);
        }
        void Graphic(OfdElement element)
        {
                Own("graphic", element); V("kind", element.GetType()); V("type", element.Type);
                S("objectId", element.ObjectId); S("layerId", element.LayerId); S("layerType", element.LayerType);
                V("x", element.XMillimeters); V("y", element.YMillimeters); V("width", element.WidthMillimeters); V("height", element.HeightMillimeters);
                S("clips", element.ClippingXml); sink.Xml(element.ClippingXml);
                element.ClipOwnership.Validate(element.ClippingXml, material, token);
                V("clipOwner", element.ClipOwnership, true); S("originalRawClips", element.ClipOwnership.Raw);
                Clips(element.Clips);
                if (!ReferenceEquals(element.Clips, element.ClipOwnership.Original)) Clips(element.ClipOwnership.Original);
                element.ClipOwnership.Baseline?.VisitInventory(sink);
                switch (element)
                {
                    case OfdTextElement text:
                        S("text", text.Text); S("fontName", text.FontName); S("fontId", text.FontResourceId);
                        V("fontSize", text.FontSizeMillimeters); V("weight", text.Weight); V("italic", text.Italic); Color(text.FillColor);
                        sink.Matrix(text.Transform); Source(text.SourceXml); Container("runs", text.Runs, text.Runs.Count);
                        foreach (var run in text.Runs)
                        { Own("run", run); S("text", run.Text); V("x", run.XMillimeters); V("y", run.YMillimeters); S("deltaX", run.DeltaX); S("deltaY", run.DeltaY); }
                        break;
                    case OfdImageElement image:
                        S("resourceId", image.ResourceId); S("fileName", image.FileName); S("mediaType", image.MediaType);
                        sink.Bytes("imageBytes", image.Data, image); sink.Matrix(image.Transform); V("alpha", image.Alpha); Source(image.SourceXml); break;
                    case OfdPathElement path:
                        if (path.FillRule is { } rule && !Enum.IsDefined(typeof(OfdFillRule), rule)) throw new InvalidDataException("Unknown typed fill rule.");
                        S("data", path.AbbreviatedData); sink.Matrix(path.Transform); V("lineWidth", path.LineWidthMillimeters);
                        V("fillRule", path.FillRule); V("stroke", path.Stroke); V("fill", path.Fill); Color(path.StrokeColor); Color(path.FillColor); Source(path.SourceXml); break;
                    case OfdCompositeElement composite:
                        V("resourceEdge", composite.Resource, true); sink.Matrix(composite.Transform); Source(composite.SourceXml); break;
                    case OfdRawElement raw: S("localName", raw.LocalName); Source(raw.Xml); break;
                    default: throw new NotSupportedException("Unknown serialized graphic contract.");
                }
        }
        void Resources(OfdCompositeResourceCollection? resources)
        {
            V("compositeResources", resources, true); if (resources is null) return;
            Own("compositeCollection", resources); Container("compositeItems", resources.Items, resources.Items.Count);
            foreach (var resource in resources.Items)
            {
                Own("compositeDefinition", resource); S("compositeId", resource.Id);
                V("vectorWidth", resource.WidthMillimeters); V("vectorHeight", resource.HeightMillimeters);
                Source(resource.SourceXml); S("compositeEntry", resource.SourceEntryPath); V("compositeSlot", resource.SourceSlot);
                Graphics("compositeGraphics", resource.Elements);
            }
        }
        void Clips(OfdClipCollection? collection)
        {
            V("typedClips", collection, true); if (collection is null) return;
            Own("clipCollection", collection); Source(collection.SourceXml); Container("clipItems", collection.Items, collection.Items.Count);
            foreach (var clip in collection.Items)
            {
                Own("clip", clip); Source(clip.SourceXml); V("clipSlot", clip.SourceSlot); Container("areas", clip.Areas, clip.Areas.Count);
                foreach (var area in clip.Areas)
                {
                    Own("area", area); Source(area.SourceXml); V("areaSlot", area.SourceSlot); sink.Matrix(area.Transform);
                    if (area.Path.Clips is not null) throw new InvalidDataException("Clip paths cannot own nested typed clips.");
                    V("originalClipPath", area.OriginalPath, true); Source(area.OriginalPathXml);
                    Graphic(area.Path);
                    if (area.OriginalPath is not null && !ReferenceEquals(area.Path, area.OriginalPath)) Graphic(area.OriginalPath);
                }
            }
        }
        void Source(string? xml) { S("sourceXml", xml); sink.Xml(xml); }
        void Color(OfdColor? color)
        { V("colorPresent", color is not null); if (color is not null) { V("red", color.Red); V("green", color.Green); V("blue", color.Blue); V("alpha", color.Alpha); } }
        void NavigationSource(OfdNavigationSource? source, int depth)
        {
            V("navSource", source, true); if (source is null) return;
            Source(source.Xml); V("documentOwner", source.DocumentOwner, true); V("opaque", source.HasOpaqueLocalContent);
            S("originalName", source.OriginalName); Container("slots", source.Slots, source.Slots.Count);
            foreach (var slot in source.Slots)
            { V("ordinal", slot.Key); V("slotOwner", slot.Value, true); if (slot.Value is IOfdNavigationNode) Nav(slot.Value, depth + 1); }
        }
        void NavigationOwner<T>(string role, OfdNavigationOwnership<T> owner, int depth = 1) where T : class
        {
            V(role, owner, true); S("originalXml", owner.OriginalXml); sink.Xml(owner.OriginalXml);
            NavigationSource(owner.OriginalSource, depth); V("baseline", owner.BaselineSnapshot, true);
            owner.BaselineSnapshot?.VisitInventory(sink);
            if (owner.BaselineSnapshot is not null) foreach (var node in owner.BaselineSnapshot.ReferencedNodes) Nav(node, depth);
            Nav(owner.Value, depth);
        }
        void Nav(object? value, int depth)
        {
            V("navEdge", value, true); if (value is null) return;
            material.Structure.Depth(depth); if (navigation.Contains(value)) return;
            material.ChargeCopy(64, token); navigation.Add(value);
            Object("navigation", value); NavigationSource(((IOfdNavigationNode)value).Source, depth);
            switch (value)
            {
                case OfdOutlineCollection outlines: Container("outlineItems", outlines.Items, outlines.Items.Count); foreach (var item in outlines.Items) Nav(item, depth + 1); break;
                case OfdBookmarkCollection bookmarks: Container("bookmarkItems", bookmarks.Items, bookmarks.Items.Count); foreach (var item in bookmarks.Items) Nav(item, depth + 1); break;
                case OfdActionCollection actions: Container("actionItems", actions.Items, actions.Items.Count); foreach (var item in actions.Items) Nav(item, depth + 1); break;
                case OfdOutline outline:
                    S("title", outline.Title); V("expanded", outline.Expanded); V("countHint", outline.Count);
                    NavigationOwner("outlineActions", outline.ActionOwnership, depth + 1);
                    Container("children", outline.Children, outline.Children.Count); foreach (var item in outline.Children) Nav(item, depth + 1); break;
                case OfdBookmark bookmark: S("name", bookmark.Name); Nav(bookmark.Destination, depth + 1); break;
                case OfdDestination destination:
                    V("pageReference", destination.Page, true); V("destinationType", destination.Type);
                    V("left", destination.Left); V("top", destination.Top); V("right", destination.Right); V("bottom", destination.Bottom); V("zoom", destination.Zoom); break;
                case OfdGoToAction go: S("event", go.Event); Nav(go.Destination, depth + 1); Nav(go.Bookmark, depth + 1); break;
                case OfdRawAction raw: S("event", raw.Event); Source(raw.Xml); break;
                default: throw new NotSupportedException("Unknown serialized navigation contract.");
            }
        }
    }
}

internal sealed class OfdReferenceIdentity : IEqualityComparer<object>
{
    internal static readonly OfdReferenceIdentity Instance = new();
    public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
    public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
}
