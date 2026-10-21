using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// One navigation mapping contract shared by reader/writer/edit preflight.
// Legal unsupported forms remain opaque source slots or complete RawAction.
internal static class OfdNavigationXml
{
    internal static void Read(OfdDocumentPackage package, XElement document, OfdModelMaterialization budget, CancellationToken token)
    {
        var ns = document.Name.Namespace;
        var pages = new Dictionary<long, OfdPage>();
        foreach (var page in package.Pages)
            if (Id(page.Id, out var number))
            {
                budget.Structure.Reference();
                if (pages.ContainsKey(number)) throw new InvalidDataException("Ambiguous page ID in navigation source.");
                pages.Add(number, page);
            }
        var named = new Dictionary<string, OfdBookmark>(StringComparer.Ordinal);
        var sourceNames = new HashSet<string>(StringComparer.Ordinal);
        var bookmarkRoot = Single(document, "Bookmarks");
        if (bookmarkRoot is not null)
        {
            var collection = new OfdBookmarkCollection(); Capture(collection, bookmarkRoot);
            var ordinal = 0;
            foreach (var element in bookmarkRoot.Elements())
            {
                token.ThrowIfCancellationRequested();
                if (element.Name == ns + "Bookmark" && element.Attribute("Name") is XAttribute sourceName && !sourceNames.Add(sourceName.Value))
                    throw new InvalidDataException("Duplicate bookmark name in source.");
                if (element.Name == ns + "Bookmark" && element.Attribute("Name") is XAttribute name &&
                    element.Elements(ns + "Dest").Count() == 1 && Destination(element.Element(ns + "Dest")!) is OfdDestination destination)
                {
                    if (named.ContainsKey(name.Value)) throw new InvalidDataException("Duplicate bookmark name in source.");
                    var bookmark = new OfdBookmark { Name = name.Value, Destination = destination }; Capture(bookmark, element);
                    bookmark.Source.OriginalName = bookmark.Name;
                    collection.Items.Add(bookmark); named.Add(bookmark.Name, bookmark); Slot(collection.Source, ordinal, bookmark);
                }
                ordinal++;
            }
            FinishSlots(collection.Source, bookmarkRoot); package.BookmarkOwnership.Bind(collection, budget, token); RemovePreserved(package, bookmarkRoot);
        }
        var outlineRoot = Single(document, "Outlines");
        if (outlineRoot is not null)
        {
            var collection = new OfdOutlineCollection(); Capture(collection, outlineRoot);
            var ordinal = 0;
            foreach (var element in outlineRoot.Elements())
            {
                if (element.Name == ns + "OutlineElem" && Outline(element, 1) is OfdOutline outline)
                { collection.Items.Add(outline); Slot(collection.Source, ordinal, outline); }
                ordinal++;
            }
            FinishSlots(collection.Source, outlineRoot); package.OutlineOwnership.Bind(collection, budget, token); RemovePreserved(package, outlineRoot);
        }
        var actionRoot = Single(document, "Actions");
        if (actionRoot is not null)
        {
            package.ActionOwnership.Bind(Actions(actionRoot), budget, token); RemovePreserved(package, actionRoot);
        }

        OfdOutline? Outline(XElement element, int depth)
        {
            token.ThrowIfCancellationRequested(); budget.Structure.Node(depth);
            if (element.Attribute("Title") is not XAttribute title) return null;
            if (!Boolean(element.Attribute("Expanded")?.Value, out var expanded)) return null;
            int? count = null;
            if (element.Attribute("Count") is XAttribute countAttribute)
            {
                if (!int.TryParse(countAttribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var n) || n < 0) return null;
                count = n;
            }
            var result = new OfdOutline { Title = title.Value, Expanded = expanded, Count = count }; Capture(result, element);
            var actions = Single(element, "Actions");
            if (actions is not null) result.ActionOwnership.Bind(Actions(actions), budget, token);
            var ordinal = 0;
            foreach (var child in element.Elements())
            {
                if (child.Name == ns + "OutlineElem" && Outline(child, depth + 1) is OfdOutline nested)
                { result.Children.Add(nested); Slot(result.Source, ordinal, nested); }
                ordinal++;
            }
            FinishSlots(result.Source, element, "Actions"); return result;
        }

        OfdActionCollection Actions(XElement root)
        {
            var result = new OfdActionCollection(); Capture(result, root); var ordinal = 0;
            foreach (var element in root.Elements())
            {
                token.ThrowIfCancellationRequested();
                if (element.Name == ns + "Action")
                {
                    var eventValue = element.Attribute("Event")?.Value;
                    var gotoElements = element.Elements(ns + "Goto").ToList();
                    OfdAction? action = null;
                    if (gotoElements.Count == 1 && ValidEvent(eventValue) &&
                        !element.Elements().Any(child => child.Name.Namespace == ns && child.Name.LocalName is not "Goto"))
                    {
                        var go = gotoElements[0];
                        var destinations = go.Elements(ns + "Dest").ToList(); var bookmarks = go.Elements(ns + "Bookmark").ToList();
                        if (destinations.Count == 1 && bookmarks.Count == 0 && Destination(destinations[0]) is OfdDestination destination)
                            action = new OfdGoToAction { Destination = destination, Event = eventValue };
                        else if (destinations.Count == 0 && bookmarks.Count == 1 && bookmarks[0].Attribute("Name") is XAttribute name && named.TryGetValue(name.Value, out var bookmark))
                            action = new OfdGoToAction { Bookmark = bookmark, Event = eventValue };
                    }
                    if (action is null && ValidEvent(eventValue))
                    {
                        var kinds = element.Elements().Where(child => child.Name.Namespace == ns && child.Name.LocalName != "Region").ToArray();
                        if (kinds.Length == 1 && !kinds[0].HasElements)
                        {
                            var kind = kinds[0];
                            switch (kind.Name.LocalName)
                            {
                                case "URI" when kind.Attribute("URI") is XAttribute uri && !string.IsNullOrWhiteSpace(uri.Value):
                                    action = new OfdUriAction { Event = eventValue, Uri = uri.Value,
                                        Base = kind.Attribute("Base")?.Value, Target = kind.Attribute("Target")?.Value }; break;
                                case "GotoA" when kind.Attribute("AttachID") is XAttribute attachment && !string.IsNullOrWhiteSpace(attachment.Value) &&
                                    package.Attachments.Count(item => item.Id == attachment.Value) == 1 &&
                                    Boolean(kind.Attribute("NewWindow")?.Value, out var newWindow):
                                    action = new OfdAttachmentAction { Event = eventValue, AttachId = attachment.Value, NewWindow = newWindow }; break;
                                case "Sound" when Id(kind.Attribute("ResourceID")?.Value, out _) &&
                                    MediaResource(package, kind.Attribute("ResourceID")!.Value, "Audio", budget, token) &&
                                    Boolean(kind.Attribute("Repeat")?.Value, out var repeat) && Boolean(kind.Attribute("Synchronous")?.Value, out var synchronous) &&
                                    (kind.Attribute("Volume") is null || int.TryParse(kind.Attribute("Volume")!.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var volume) && volume >= 0 && volume <= 100):
                                    action = new OfdSoundAction { Event = eventValue, ResourceId = kind.Attribute("ResourceID")!.Value,
                                        Volume = kind.Attribute("Volume") is null ? null : int.Parse(kind.Attribute("Volume")!.Value, CultureInfo.InvariantCulture),
                                        Repeat = repeat, Synchronous = synchronous }; break;
                                case "Movie" when Id(kind.Attribute("ResourceID")?.Value, out _) &&
                                    MediaResource(package, kind.Attribute("ResourceID")!.Value, "Video", budget, token) &&
                                    (kind.Attribute("Operator") is null ||
                                        Enum.GetNames(typeof(OfdMovieOperator)).Contains(kind.Attribute("Operator")!.Value, StringComparer.Ordinal)):
                                    action = new OfdMovieAction { Event = eventValue, ResourceId = kind.Attribute("ResourceID")!.Value,
                                        Operator = kind.Attribute("Operator") is null ? null : (OfdMovieOperator)Enum.Parse(typeof(OfdMovieOperator), kind.Attribute("Operator")!.Value) }; break;
                            }
                        }
                    }
                    action ??= new OfdRawAction { Event = eventValue };
                    Capture(action, element); if (action is OfdRawAction raw) raw.Xml = raw.SourceXml!; result.Items.Add(action); Slot(result.Source, ordinal, action);
                }
                ordinal++;
            }
            FinishSlots(result.Source, root); return result;
        }

        OfdDestination? Destination(XElement element)
        {
            token.ThrowIfCancellationRequested();
            if ((element.Attribute("Type")?.Value is not ("XYZ" or "Fit" or "FitH" or "FitV" or "FitR")) ||
                !Enum.TryParse<OfdDestinationType>(element.Attribute("Type")?.Value, false, out var type) || !Enum.IsDefined(typeof(OfdDestinationType), type) ||
                !Id(element.Attribute("PageID")?.Value, out var number) || !pages.TryGetValue(number, out var page)) return null;
            var values = new double?[5]; var names = new[] { "Left", "Top", "Right", "Bottom", "Zoom" };
            for (var i = 0; i < names.Length; i++)
                if (element.Attribute(names[i]) is XAttribute attribute)
                {
                    if (!double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) || double.IsNaN(value) || double.IsInfinity(value)) return null;
                    values[i] = value;
                }
            if (values[4] is double zoom && zoom != 0 && (zoom < .1 || zoom > 64)) return null;
            var result = new OfdDestination { Page = page, Type = type, Left = values[0], Top = values[1], Right = values[2], Bottom = values[3], Zoom = values[4] };
            Capture(result, element); return result;
        }

        void Capture(IOfdNavigationNode node, XElement element)
        {
            token.ThrowIfCancellationRequested(); node.Source.Xml = OfdXmlEnvelope.Capture(element, budget, token); node.Source.DocumentOwner = package;
            node.Source.HasOpaqueLocalContent = node is OfdRawAction || HasOpaqueLocal(element);
        }
        void FinishSlots(OfdNavigationSource source, XElement root, string? independentlyOwned = null)
        {
            var ordinal = 0;
            foreach (var child in root.Elements())
            {
                if (!source.Slots.ContainsKey(ordinal) && (independentlyOwned is null || child.Name != ns + independentlyOwned)) source.HasOpaqueLocalContent = true;
                ordinal++;
            }
        }
        void Slot(OfdNavigationSource source, int ordinal, object item) { budget.Structure.Reference(); source.Slots.Add(ordinal, item); }
    }

    internal static IReadOnlyList<XElement> Write(OfdDocumentPackage package, IReadOnlyDictionary<OfdPage, string> pageIds,
        XNamespace ns, CancellationToken token, OfdModelMaterialization? material = null)
    {
        Validate(package, package.Pages, token, material);
        var result = new List<XElement>();
        var outlines = WriteOwned(package.OutlineOwnership, "Outlines", (root, source, value) =>
            OfdXmlEnvelope.Children(root, source, value.Items, WriteOutline, token));
        if (outlines is not null) result.Add(outlines);
        var actions = WriteActions(package.ActionOwnership); if (actions is not null) result.Add(actions);
        var bookmarks = WriteOwned(package.BookmarkOwnership, "Bookmarks", (root, source, value) =>
            OfdXmlEnvelope.Children(root, source, value.Items, WriteBookmark, token));
        if (bookmarks is not null) result.Add(bookmarks);
        return result;

        XElement? WriteOwned<T>(OfdNavigationOwnership<T> owner, string name,
            Action<XElement, OfdNavigationSource?, T> patch) where T : class, IOfdNavigationNode
        {
            token.ThrowIfCancellationRequested();
            if (owner.Value is null) return owner.OriginalXml is null ? null : OfdXmlEnvelope.Parse(owner.OriginalXml);
            var source = owner.OriginalSource ?? owner.Value.Source;
            var root = OfdXmlEnvelope.Template(source, name, ns); patch(root, source, owner.Value);
            // Optional collection wrappers require at least one standard item.
            // A clear with no retained extension metadata can omit the wrapper.
            if (!root.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) &&
                !root.Nodes().Any(node => node is not XText text || !string.IsNullOrWhiteSpace(text.Value))) return null;
            return root;
        }
        XElement WriteOutline(OfdOutline outline)
        {
            token.ThrowIfCancellationRequested();
            var root = OfdXmlEnvelope.Template(outline.Source, "OutlineElem", ns);
            root.SetAttributeValue("Title", outline.Title); root.SetAttributeValue("Expanded", outline.Expanded); root.SetAttributeValue("Count", outline.Count);
            OfdXmlEnvelope.Children(root, outline.Source, outline.Children, WriteOutline, token);
            OfdXmlEnvelope.Singleton(root, "Actions", WriteActions(outline.ActionOwnership));
            return root;
        }
        XElement? WriteActions(OfdNavigationOwnership<OfdActionCollection> owner)
            => WriteOwned(owner, "Actions", (root, source, value) => OfdXmlEnvelope.Children(root, source, value.Items, WriteAction, token));
        XElement WriteBookmark(OfdBookmark bookmark)
        {
            token.ThrowIfCancellationRequested(); var root = OfdXmlEnvelope.Template(bookmark.Source, "Bookmark", ns);
            root.SetAttributeValue("Name", bookmark.Name); OfdXmlEnvelope.Singleton(root, "Dest", WriteDestination(bookmark.Destination)); return root;
        }
        XElement WriteAction(OfdAction action)
        {
            token.ThrowIfCancellationRequested();
            if (action is OfdRawAction raw)
            {
                var preserved = OfdXmlEnvelope.Parse(raw.Xml);
                if (preserved.Name != ns + "Action") throw new ArgumentException("Raw action must be a complete Action in the document namespace.");
                preserved.SetAttributeValue("Event", raw.Event);
                return preserved;
            }
            if (action is not OfdGoToAction go)
            {
                var local = action switch
                {
                    OfdUriAction => "URI", OfdAttachmentAction => "GotoA",
                    OfdSoundAction => "Sound", OfdMovieAction => "Movie",
                    _ => throw new NotSupportedException("Unsupported typed action.")
                };
                var nodeRoot = OfdXmlEnvelope.Template(action.Source, "Action", ns); nodeRoot.SetAttributeValue("Event", action.Event);
                var child = nodeRoot.Element(ns + local) is XElement old ? new XElement(old) : new XElement(ns + local);
                switch (action)
                {
                    case OfdUriAction uri:
                        child.SetAttributeValue("URI", uri.Uri); child.SetAttributeValue("Base", uri.Base); child.SetAttributeValue("Target", uri.Target); break;
                    case OfdAttachmentAction attachment:
                        child.SetAttributeValue("AttachID", attachment.AttachId); child.SetAttributeValue("NewWindow", attachment.NewWindow?.ToString().ToLowerInvariant()); break;
                    case OfdSoundAction sound:
                        child.SetAttributeValue("ResourceID", sound.ResourceId); child.SetAttributeValue("Volume", sound.Volume);
                        child.SetAttributeValue("Repeat", sound.Repeat?.ToString().ToLowerInvariant());
                        child.SetAttributeValue("Synchronous", sound.Synchronous?.ToString().ToLowerInvariant()); break;
                    case OfdMovieAction movie:
                        child.SetAttributeValue("ResourceID", movie.ResourceId); child.SetAttributeValue("Operator", movie.Operator?.ToString()); break;
                }
                OfdXmlEnvelope.Singleton(nodeRoot, local, child); return nodeRoot;
            }
            var root = OfdXmlEnvelope.Template(action.Source, "Action", ns); root.SetAttributeValue("Event", action.Event);
            var gotoElement = root.Element(ns + "Goto") is XElement original ? new XElement(original) : new XElement(ns + "Goto");
            if (go.Destination is not null)
            {
                OfdXmlEnvelope.Singleton(gotoElement, "Bookmark", null);
                OfdXmlEnvelope.Singleton(gotoElement, "Dest", WriteDestination(go.Destination));
            }
            else
            {
                OfdXmlEnvelope.Singleton(gotoElement, "Dest", null);
                var target = gotoElement.Element(ns + "Bookmark") is XElement old ? new XElement(old) : new XElement(ns + "Bookmark");
                target.SetAttributeValue("Name", go.Bookmark!.Name); OfdXmlEnvelope.Singleton(gotoElement, "Bookmark", target);
            }
            OfdXmlEnvelope.Singleton(root, "Goto", gotoElement); return root;
        }
        XElement WriteDestination(OfdDestination destination)
        {
            token.ThrowIfCancellationRequested(); var root = OfdXmlEnvelope.Template(destination.Source, "Dest", ns);
            root.SetAttributeValue("Type", destination.Type); root.SetAttributeValue("PageID", pageIds[destination.Page]);
            Scalar("Left", destination.Left); Scalar("Top", destination.Top); Scalar("Right", destination.Right); Scalar("Bottom", destination.Bottom); Scalar("Zoom", destination.Zoom);
            return root;
            void Scalar(string name, double? value)
            {
                // The pinned Java implementation also emits scalar children;
                // those remain opaque rather than being silently normalized.
                if (root.Elements(ns + name).Any() && value.HasValue) throw new NotSupportedException("Cannot overlay a scalar attribute onto an opaque scalar child.");
                root.SetAttributeValue(name, value?.ToString("R", CultureInfo.InvariantCulture));
            }
        }
    }

    internal static void Validate(OfdDocumentPackage package, IReadOnlyCollection<OfdPage> retainedPages, CancellationToken token, OfdModelMaterialization? material = null)
    {
        token.ThrowIfCancellationRequested();
        var pages = new HashSet<OfdPage>(retainedPages);
        var names = new HashSet<string>(StringComparer.Ordinal); var bookmarks = new HashSet<OfdBookmark>();
        var active = new HashSet<object>(); var stack = new Stack<(object Node, int Depth, bool Exit)>();
        PushRoot(package.Bookmarks); PushRoot(package.Outlines); PushRoot(package.Actions);
        if (package.Bookmarks is null && package.BookmarkOwnership.OriginalXml is not null && package.BookmarkOwnership.BaselineSnapshot is OfdNavigationSnapshot originalBookmarks)
            foreach (var bookmark in originalBookmarks.ReferencedBookmarks) bookmarks.Add(bookmark);
        var effectiveBookmarkSource = package.BookmarkOwnership.OriginalSource ?? package.Bookmarks?.Source;
        if (effectiveBookmarkSource?.Xml is string sourceXml)
        {
            material?.ChargeCopy(2L * sourceXml.Length, token);
            var root = OfdXmlEnvelope.Parse(sourceXml); var ordinal = 0;
            foreach (var element in root.Elements())
            {
                token.ThrowIfCancellationRequested(); material?.Structure.Reference();
                if ((package.Bookmarks is null || !effectiveBookmarkSource.Slots.ContainsKey(ordinal)) &&
                    element.Name == root.Name.Namespace + "Bookmark" && element.Attribute("Name") is XAttribute name && !names.Add(name.Value))
                    throw new InvalidDataException("Duplicate effective bookmark name.");
                ordinal++;
            }
        }
        if (package.Bookmarks is not null)
            foreach (var bookmark in package.Bookmarks.Items)
            {
                token.ThrowIfCancellationRequested();
                if (bookmark is null || bookmark.Name is null || !names.Add(bookmark.Name)) throw new ArgumentException("Bookmark names must be unique and non-null.");
                bookmarks.Add(bookmark);
            }
        var deletingSourcePage = retainedPages.Count != package.Pages.Count ||
            package.SourceIdentity?.OriginalPages.Any(page => !pages.Contains(page)) == true;
        ValidateUnclaimed(package.OutlineOwnership); ValidateUnclaimed(package.BookmarkOwnership); ValidateUnclaimed(package.ActionOwnership);
        var work = material?.Structure ?? new OfdTraversalBudget(new OfdModelLimits());
        while (stack.Count > 0)
        {
            token.ThrowIfCancellationRequested(); var frame = stack.Pop();
            if (frame.Exit) { active.Remove(frame.Node); continue; }
            work.Node(frame.Depth); work.Reference();
            if (!active.Add(frame.Node)) throw new InvalidDataException("Navigation contains an ownership cycle.");
            stack.Push((frame.Node, frame.Depth, true));
            if (frame.Node is IOfdNavigationNode node && node.Source.DocumentOwner is object owner && !ReferenceEquals(owner, package))
                throw new NotSupportedException("Imported navigation must be cloned with its source ownership.");
            if (frame.Node is IOfdNavigationNode source && source.Source.HasOpaqueLocalContent && deletingSourcePage)
                throw new NotSupportedException("Opaque navigation prevents proving page-deletion reference closure.");
            switch (frame.Node)
            {
                case OfdOutlineCollection outlines: PushAll(outlines.Items, frame.Depth); break;
                case OfdBookmarkCollection collection: PushAll(collection.Items, frame.Depth); break;
                case OfdActionCollection actions: PushAll(actions.Items, frame.Depth); break;
                case OfdOutline outline:
                    if (outline.Title is null || outline.Count < 0) throw new ArgumentException("Outline title and count are invalid.");
                    PushAll(outline.Children, frame.Depth);
                    if (outline.Actions is not null) Push(outline.Actions, frame.Depth);
                    else ValidateUnclaimed(outline.ActionOwnership);
                    break;
                case OfdBookmark bookmark: Push(bookmark.Destination, frame.Depth); break;
                case OfdDestination destination:
                    if (destination.Page is null || !pages.Contains(destination.Page)) throw new InvalidOperationException("Navigation targets a page outside the retained document.");
                    if (!Enum.IsDefined(typeof(OfdDestinationType), destination.Type)) throw new ArgumentException("Invalid destination type.");
                    foreach (var number in new[] { destination.Left, destination.Top, destination.Right, destination.Bottom, destination.Zoom })
                        if (number.HasValue && (double.IsNaN(number.Value) || double.IsInfinity(number.Value))) throw new ArgumentException("Destination coordinates must be finite.");
                    if (destination.Zoom is double zoom && zoom != 0 && (zoom < .1 || zoom > 64)) throw new ArgumentException("Invalid destination zoom.");
                    break;
                case OfdGoToAction go:
                    if (!ValidEvent(go.Event) || (go.Destination is null) == (go.Bookmark is null)) throw new ArgumentException("GoTo requires a valid event and exactly one destination or bookmark.");
                    if (go.Bookmark is not null && !bookmarks.Contains(go.Bookmark)) throw new InvalidOperationException("GoTo bookmark does not belong to the document.");
                    if (go.Bookmark is not null && package.Bookmarks is null && go.Bookmark.Name != go.Bookmark.Source.OriginalName)
                        throw new InvalidOperationException("GoTo name disagrees with the unclaimed source bookmark.");
                    if (go.Destination is not null) Push(go.Destination, frame.Depth);
                    break;
                case OfdRawAction:
                    if (deletingSourcePage) throw new NotSupportedException("Opaque navigation prevents proving page-deletion reference closure.");
                    break;
                case OfdUriAction uri:
                    if (!ValidEvent(uri.Event) || string.IsNullOrWhiteSpace(uri.Uri)) throw new ArgumentException("URI action requires a valid event and target.");
                    break;
                case OfdAttachmentAction attachment:
                    if (!ValidEvent(attachment.Event) || string.IsNullOrWhiteSpace(attachment.AttachId) ||
                        package.Attachments.Count(item => item.Id == attachment.AttachId) != 1)
                        throw new NotSupportedException("Attachment action target is not a retained local attachment.");
                    break;
                case OfdSoundAction sound:
                    if (!ValidEvent(sound.Event) || !Id(sound.ResourceId, out _) || sound.Volume is < 0 or > 100)
                        throw new ArgumentException("Invalid Sound action.");
                    if (!MediaResource(package, sound.ResourceId, "Audio", material ?? new OfdModelMaterialization(new OfdModelLimits()), token))
                        throw new NotSupportedException("Sound action requires a retained audio resource.");
                    break;
                case OfdMovieAction movie:
                    if (!ValidEvent(movie.Event) || !Id(movie.ResourceId, out _) || movie.Operator.HasValue && !Enum.IsDefined(typeof(OfdMovieOperator), movie.Operator.Value))
                        throw new ArgumentException("Invalid Movie action.");
                    if (!MediaResource(package, movie.ResourceId, "Video", material ?? new OfdModelMaterialization(new OfdModelLimits()), token))
                        throw new NotSupportedException("Movie action requires a retained video resource.");
                    break;
                default: throw new NotSupportedException("Unsupported typed navigation node.");
            }
        }
        token.ThrowIfCancellationRequested();
        void PushRoot(object? node) { if (node is not null) stack.Push((node, 1, false)); }
        void ValidateUnclaimed<T>(OfdNavigationOwnership<T> owner) where T : class
        {
            if (owner.OriginalSource?.HasOpaqueLocalContent == true && deletingSourcePage)
                throw new NotSupportedException("Opaque effective source prevents proving page-deletion reference closure.");
            if (owner.Value is not null || owner.OriginalXml is null || owner.BaselineSnapshot is null) return;
            foreach (var target in owner.BaselineSnapshot.BookmarkTargets)
                if (!bookmarks.Contains(target.Bookmark) || target.Bookmark.Name != target.Name)
                    throw new InvalidOperationException("Unclaimed navigation has a removed or renamed bookmark target.");
            foreach (var page in owner.BaselineSnapshot.ReferencedPages)
                if (!pages.Contains(page)) throw new InvalidOperationException("Unclaimed navigation targets an excluded page.");
            if (owner.BaselineSnapshot.HasOpaqueSource && deletingSourcePage)
                throw new NotSupportedException("Opaque unclaimed navigation prevents proving page-deletion reference closure.");
        }
        void Push(object? node, int depth)
        {
            if (node is null) throw new ArgumentException("Null owned navigation child.");
            work.Reference(); work.Depth(depth + 1); stack.Push((node, depth + 1, false));
        }
        void PushAll<T>(List<T> values, int depth) where T : class { for (var i = values.Count - 1; i >= 0; i--) Push(values[i], depth); }
    }

    internal static bool HasTypedNavigation(OfdDocumentPackage package, OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        if (package.Outlines is not null || package.Bookmarks is not null || package.Actions is not null ||
            package.OutlineOwnership.OriginalXml is not null || package.BookmarkOwnership.OriginalXml is not null || package.ActionOwnership.OriginalXml is not null) return true;
        foreach (var xml in package.PreservedDocumentElements)
        {
            token.ThrowIfCancellationRequested(); material?.ChargeCopy(2L * xml.Length, token);
            var name = OfdXmlEnvelope.Parse(xml).Name;
            if (name.Namespace == package.Options.Namespace && name.LocalName is "Outlines" or "Actions" or "Bookmarks") return true;
        }
        return false;
    }

    private static XElement? Single(XElement parent, string name)
    {
        var elements = parent.Elements(parent.Name.Namespace + name).ToList();
        if (elements.Count > 1) throw new InvalidDataException("Duplicate navigation singleton.");
        return elements.SingleOrDefault();
    }
    private static bool Id(string? value, out long number) => long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out number) && number > 0;
    private static bool ValidEvent(string? value) => value is null or "DO" or "PO" or "CLICK";
    private static bool MediaResource(OfdDocumentPackage package, string id, string type, OfdModelMaterialization material, CancellationToken token)
    {
        var documentPath = package.DocumentEntryPath ?? package.Options.DocumentId + "/Document.xml";
        var found = 0; var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var location in new[] { package.PublicResourceLocation, package.DocumentResourceLocation })
        {
            token.ThrowIfCancellationRequested(); if (string.IsNullOrWhiteSpace(location)) continue;
            var path = material.ResolvePath(documentPath, location!, token);
            if (!paths.Add(path)) continue;
            if (!package.PreservedEntries.TryGetValue(path, out var bytes)) continue;
            var xml = material.ReadXml(path, bytes, token);
            if (xml.Root?.Name.LocalName != "Res") continue;
            var ns = xml.Root.Name.Namespace;
            foreach (var media in xml.Root.Descendants(ns + "MultiMedia"))
            {
                token.ThrowIfCancellationRequested();
                if (media.Attribute("ID")?.Value != id || media.Attribute("Type")?.Value != type) continue;
                var file = media.Attribute("MediaFile")?.Value ?? media.Element(ns + "MediaFile")?.Value;
                if (string.IsNullOrWhiteSpace(file)) continue;
                var baseLoc = xml.Root.Attribute("BaseLoc")?.Value;
                var mediaLoc = string.IsNullOrWhiteSpace(baseLoc) || file!.StartsWith("/", StringComparison.Ordinal)
                    ? file! : baseLoc!.TrimEnd('/') + "/" + file.TrimStart('/');
                var resolved = material.ResolvePath(path, mediaLoc, token);
                if (package.PreservedEntries.ContainsKey(resolved)) found++;
            }
        }
        return found == 1;
    }
    private static bool HasOpaqueLocal(XElement root)
    {
        var ns = root.Name.Namespace;
        bool Attribute(XAttribute attribute)
        {
            if (attribute.IsNamespaceDeclaration) return true;
            if (attribute.Name.Namespace != XNamespace.None) return false;
            var name = attribute.Name.LocalName;
            return root.Name.LocalName switch
            {
                "OutlineElem" => name is "Title" or "Expanded" or "Count",
                "Bookmark" => name == "Name", "Action" => name == "Event",
                "Dest" => name is "Type" or "PageID" or "Left" or "Top" or "Right" or "Bottom" or "Zoom",
                _ => false
            };
        }
        if (root.Attributes().Any(attribute => !Attribute(attribute)) || root.Nodes().Any(node => node is XProcessingInstruction || node is XText text && !string.IsNullOrWhiteSpace(text.Value))) return true;
        foreach (var child in root.Elements())
        {
            if (child.Name.Namespace != ns) return true;
            var allowed = root.Name.LocalName switch
            {
                "Outlines" => child.Name.LocalName == "OutlineElem", "OutlineElem" => child.Name.LocalName is "OutlineElem" or "Actions",
                "Bookmarks" => child.Name.LocalName == "Bookmark", "Bookmark" => child.Name.LocalName == "Dest",
                "Actions" => child.Name.LocalName == "Action", "Action" => child.Name.LocalName is "Goto" or "URI" or "GotoA" or "Sound" or "Movie", _ => false
            };
            if (!allowed) return true;
            if (child.Name == ns + "URI" && (child.Nodes().Any(OpaqueNode) ||
                child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("URI" or "Base" or "Target"))))) return true;
            if (child.Name == ns + "Goto" && (child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                child.Nodes().Any(OpaqueNode) || child.Elements().Any(target => target.Name != ns + "Dest" && target.Name != ns + "Bookmark" ||
                    target.Name == ns + "Bookmark" && (target.HasElements || target.Nodes().Any(OpaqueNode) || target.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "Name"))))) return true;
            if (child.Name.LocalName is "GotoA" or "Sound" or "Movie" &&
                (child.HasElements || child.Nodes().Any(OpaqueNode) || child.Attributes().Any(attribute =>
                    !attribute.IsNamespaceDeclaration && (attribute.Name.Namespace != XNamespace.None || (child.Name.LocalName switch
                    {
                        "GotoA" => attribute.Name.LocalName is not ("AttachID" or "NewWindow"),
                        "Sound" => attribute.Name.LocalName is not ("ResourceID" or "Volume" or "Repeat" or "Synchronous"),
                        "Movie" => attribute.Name.LocalName is not ("ResourceID" or "Operator"),
                        _ => true
                    }))))) return true;
        }
        return false;
        static bool OpaqueNode(XNode node) => node is XProcessingInstruction || node is XText text && !string.IsNullOrWhiteSpace(text.Value);
    }
    private static bool Boolean(string? value, out bool? result)
    {
        result = null; if (value is null) return true;
        if (value is "true" or "1") { result = true; return true; }
        if (value is "false" or "0") { result = false; return true; }
        return false;
    }
    private static void RemovePreserved(OfdDocumentPackage package, XElement owned)
    {
        for (var index = package.PreservedDocumentElements.Count - 1; index >= 0; index--)
            if (OfdXmlEnvelope.Parse(package.PreservedDocumentElements[index]).Name == owned.Name) package.PreservedDocumentElements.RemoveAt(index);
    }
}
