using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Annotations;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig.Outline.Destinations;
using UglyToad.PdfPig.Parser.Parts;
using UglyToad.PdfPig.Tokens;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>Copies only closed local PDF navigation and embedded-file references into a fresh OFD package.</summary>
internal static class PdfImportMetadataMigrator
{
    internal static bool HasOnlyFileAttachments(PdfDocument source, int zeroBasedPage, PdfToOfdOptions options,
        RawPageLookup pages, CancellationToken token)
    {
        if (!options.CopyAttachments) return false;
        _ = ReadPageAnnotations(source, zeroBasedPage, options, pages, token, out var onlyFileAttachments);
        return onlyFileAttachments;
    }

    private static Annotation[] ReadPageAnnotations(PdfDocument source, int zeroBasedPage, PdfToOfdOptions options,
        RawPageLookup pages, CancellationToken token, out bool onlyFileAttachments)
    {
        if (!ReferenceEquals(source, pages.Source)) throw new InvalidOperationException("PDF raw page lookup belongs to another document.");
        onlyFileAttachments = true;
        var page = pages.Get(zeroBasedPage, token);
        var acceptedPage = source.GetPage(zeroBasedPage + 1);
        if (!page.Equals(acceptedPage.Dictionary))
            throw new InvalidDataException("PDF raw page tree disagrees with the parser's accepted page.");
        if (!page.Data.TryGetValue("Annots", out var annotsToken)) return Array.Empty<Annotation>();
        var annots = Resolve<ArrayToken>(annotsToken, source) ?? throw new InvalidDataException("PDF annotation array cannot be resolved.");
        var rawFileCount = 0;
        foreach (var entry in annots.Data)
        {
            token.ThrowIfCancellationRequested();
            var annotation = Resolve<DictionaryToken>(entry, source) ?? throw new InvalidDataException("PDF page annotation cannot be resolved.");
            if (!annotation.Data.TryGetValue("Subtype", out var subtype))
                throw new InvalidDataException("PDF annotation subtype is missing.");
            if (AnnotationSubtype(subtype, source).Data == "FileAttachment")
            {
                if (++rawFileCount > options.MaxAttachmentCount)
                    throw new InvalidDataException("PDF page attachment count exceeds the configured limit.");
            }
            else onlyFileAttachments = false;
        }
        var parsed = acceptedPage.GetAnnotations().ToArray();
        if (parsed.Count(annotation => annotation.Type == AnnotationType.FileAttachment) != rawFileCount ||
            onlyFileAttachments && parsed.Length != annots.Data.Count)
            throw new InvalidDataException("PDF page attachment parser omitted an annotation.");
        return parsed;
    }

    private static NameToken AnnotationSubtype(IToken token, PdfDocument source)
    {
        var visited = new HashSet<string>(StringComparer.Ordinal);
        IToken? current = token;
        while (current is IndirectReferenceToken reference && visited.Count < 32)
        {
            if (!visited.Add(reference.Data.ToString())) break;
            try { current = source.Structure.TokenScanner.Get(reference.Data)?.Data; }
            catch (PdfDocumentFormatException error)
            {
                throw new InvalidDataException("PDF annotation subtype reference is invalid.", error);
            }
        }
        return current as NameToken ?? throw new InvalidDataException("PDF annotation subtype is not a local name.");
    }

    // The lookup belongs to one PdfDocument. It walks monotonically to the highest
    // requested ordinal, so duplicate and out-of-order selections reuse the same
    // raw page without weakening the per-occurrence attachment output budget.
    internal sealed class RawPageLookup
    {
        private readonly PdfDocument source;
        private readonly PdfToOfdOptions options;
        private readonly Stack<IToken> pending = new();
        private readonly List<DictionaryToken> indexed = new();
        internal PdfDocument Source => source;
        internal long VisitedNodes { get; private set; }

        internal RawPageLookup(PdfDocument source, PdfToOfdOptions options)
        {
            this.source = source;
            this.options = options;
            if (!source.Structure.Catalog.CatalogDictionary.Data.TryGetValue("Pages", out var root))
                throw new InvalidDataException("PDF page tree is missing.");
            pending.Push(root);
        }

        internal DictionaryToken Get(int zeroBasedPage, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (zeroBasedPage < 0) throw new ArgumentOutOfRangeException(nameof(zeroBasedPage));
            while (indexed.Count <= zeroBasedPage && pending.Count != 0)
            {
                token.ThrowIfCancellationRequested();
                if (++VisitedNodes > options.MaxPageCount * 2L)
                    throw new InvalidDataException("PDF page tree exceeds the configured limit.");
                var page = Resolve<DictionaryToken>(pending.Pop(), source) ?? throw new InvalidDataException("PDF page tree node cannot be resolved.");
                if (IsRawPageLeaf(page, source))
                {
                    indexed.Add(page);
                    continue;
                }
                if (!page.Data.TryGetValue("Kids", out var children) || Resolve<ArrayToken>(children, source) is not { } kids)
                    throw new InvalidDataException("PDF page tree has no children.");
                foreach (var child in kids.Data.Reverse()) pending.Push(child);
            }
            if (zeroBasedPage >= indexed.Count) throw new InvalidDataException("PDF selected page was not found in the page tree.");
            return indexed[zeroBasedPage];
        }
    }

    private static bool IsRawPageLeaf(DictionaryToken node, PdfDocument source)
    {
        var hasKids = node.Data.ContainsKey("Kids");
        if (!node.Data.TryGetValue("Type", out var rawType))
        {
            // PdfPig's lenient page tree admits a leaf with neither key. A
            // typeless node with Kids is a container only if the caller can
            // resolve a real children array; both walkers check that below.
            return !hasKids;
        }
        var type = Resolve<NameToken>(rawType, source)
            ?? throw new InvalidDataException("PDF page tree type is not a local name.");
        if (type.Data == "Page")
        {
            if (hasKids) throw new InvalidDataException("PDF page leaf also declares children.");
            return true;
        }
        if (type.Data == "Pages")
        {
            if (!hasKids) throw new InvalidDataException("PDF page tree container has no children.");
            return false;
        }
        throw new InvalidDataException("PDF page tree node has an unsupported type.");
    }

    internal static void Apply(PdfDocument source, OfdDocumentPackage output, IReadOnlyList<int> selected,
        PdfToOfdOptions options, CancellationToken token, IReadOnlyList<bool> drawAttachmentPins)
    {
        if (!options.CopyBookmarks && !options.CopyAttachments) return;
        if (selected.Count != output.Pages.Count || drawAttachmentPins.Count != selected.Count)
            throw new InvalidDataException("PDF metadata page selection disagrees with output pages.");
        var pages = new Dictionary<int, List<OfdPage>>();
        for (var i = 0; i < selected.Count; i++)
        {
            token.ThrowIfCancellationRequested();
            var pageNumber = checked(selected[i] + 1);
            if (!pages.TryGetValue(pageNumber, out var copies)) pages.Add(pageNumber, copies = new List<OfdPage>());
            copies.Add(output.Pages[i]);
        }
        if (options.CopyBookmarks) CopyBookmarks(source, output, pages, options, token);
        if (options.CopyAttachments) CopyAttachments(source, output, selected, drawAttachmentPins, options, token);
    }

    private static void CopyBookmarks(PdfDocument source, OfdDocumentPackage output,
        IReadOnlyDictionary<int, List<OfdPage>> pages, PdfToOfdOptions options, CancellationToken token)
    {
        var rawDestinations = ValidateRawOutlines(source, options, token);
        if (rawDestinations.Count == 0) return;
        if (!source.TryGetBookmarks(out var bookmarks, true))
            throw new InvalidDataException("PDF outline contains a destination the parser could not resolve.");
        var collection = new OfdOutlineCollection();
        var work = new Stack<(BookmarkNode Node, List<OfdOutline> Target, int Depth)>();
        foreach (var root in bookmarks.Roots.Reverse()) work.Push((root, collection.Items, 1));
        var count = 0;
        while (work.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var (node, target, depth) = work.Pop();
            if (depth > options.MaxBookmarkDepth || ++count > options.MaxBookmarkCount)
                throw new InvalidDataException("PDF outline exceeds the configured node or depth limit.");
            var outline = new OfdOutline { Title = node.Title ?? string.Empty };
            if (node is DocumentBookmarkNode local)
            {
                if (!rawDestinations[count - 1].HasDestination)
                    throw new InvalidDataException("PDF outline parser invented a destination for a grouping node.");
                if (local.PageNumber < 1 || !pages.TryGetValue(local.PageNumber, out var retained))
                    throw new InvalidDataException("PDF outline targets an invalid or excluded page.");
                outline.Actions = new OfdActionCollection();
                outline.Actions.Items.Add(new OfdGoToAction { Event = "CLICK",
                    Destination = Destination(source.GetPage(local.PageNumber), retained[0], local.Destination, rawDestinations[count - 1]) });
            }
            else if (node is ContainerBookmarkNode)
            {
                if (rawDestinations[count - 1].HasDestination)
                    throw new NotSupportedException("PDF outline parser could not preserve an explicit destination.");
            }
            else
                throw new NotSupportedException("PDF outline has an external or unsupported action.");
            target.Add(outline);
            foreach (var child in node.Children.Reverse()) work.Push((child, outline.Children, depth + 1));
        }
        if (count != rawDestinations.Count) throw new InvalidDataException("PDF outline contains an omitted or unsupported node.");
        if (collection.Items.Count != 0) output.Outlines = collection;
    }

    private static List<(bool HasDestination, double? Zoom, double? Left, double? Top, double[]? Rect)> ValidateRawOutlines(PdfDocument source, PdfToOfdOptions options, CancellationToken token)
    {
        var destinations = new List<(bool HasDestination, double? Zoom, double? Left, double? Top, double[]? Rect)>();
        var catalog = source.Structure.Catalog.CatalogDictionary;
        if (!catalog.Data.TryGetValue("Outlines", out var outlineToken)) return destinations;
        var root = Resolve<DictionaryToken>(outlineToken, source)
            ?? throw new InvalidDataException("PDF outline root cannot be resolved.");
        if (!root.Data.TryGetValue("First", out var first))
        {
            if (root.Data.ContainsKey("Last")) throw new InvalidDataException("PDF outline root has a last node but no first node.");
            if (root.Data.TryGetValue("Count", out var countToken) &&
                (Resolve<NumericToken>(countToken, source) is not { } rootCount || rootCount.Double != 0))
                throw new InvalidDataException("PDF outline root declares entries but has no first node.");
            return destinations;
        }
        if (!root.Data.TryGetValue("Last", out var last))
            throw new InvalidDataException("PDF outline root has a first node but no last node.");
        var pageReferences = new HashSet<string>(StringComparer.Ordinal);
        var pendingPages = new Stack<IToken>();
        if (!catalog.Data.TryGetValue("Pages", out var pageTree)) throw new InvalidDataException("PDF page tree is missing.");
        pendingPages.Push(pageTree);
        var pageNodes = 0;
        while (pendingPages.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            if (++pageNodes > options.MaxPageCount * 2L) throw new InvalidDataException("PDF page tree exceeds the configured limit.");
            var current = pendingPages.Pop();
            var node = Resolve<DictionaryToken>(current, source) ?? throw new InvalidDataException("PDF page tree node cannot be resolved.");
            if (IsRawPageLeaf(node, source))
            {
                if (current is not IndirectReferenceToken reference || !pageReferences.Add(reference.Data.ToString()))
                    throw new InvalidDataException("PDF page reference is missing or repeated.");
                continue;
            }
            if (!node.Data.TryGetValue("Kids", out var children) || Resolve<ArrayToken>(children, source) is not { } kids)
                throw new InvalidDataException("PDF page tree has no children.");
            foreach (var child in kids.Data) pendingPages.Push(child);
        }
        var pending = new Stack<(IToken Node, IToken Last, int Depth)>(); pending.Push((first, last, 1));
        var count = 0;
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var (link, expectedLast, depth) = pending.Pop();
            if (depth > options.MaxBookmarkDepth || ++count > options.MaxBookmarkCount)
                throw new InvalidDataException("PDF outline exceeds the configured node or depth limit.");
            destinations.Add((false, null, null, null, null));
            var item = Resolve<DictionaryToken>(link, source) ?? throw new InvalidDataException("PDF outline node cannot be resolved.");
            if (item.Data.TryGetValue("Next", out var next)) pending.Push((next, expectedLast, depth));
            else if (!SameOutlineLink(link, expectedLast))
                throw new InvalidDataException("PDF outline sibling chain does not reach its last node.");
            if (item.Data.TryGetValue("First", out var child))
            {
                if (!item.Data.TryGetValue("Last", out var childLast))
                    throw new InvalidDataException("PDF outline node has a first child but no last child.");
                pending.Push((child, childLast, depth + 1));
            }
            else if (item.Data.ContainsKey("Last"))
                throw new InvalidDataException("PDF outline node has a last child but no first child.");
            if (item.Data.ContainsKey("Dest") && item.Data.ContainsKey("A"))
                throw new NotSupportedException("PDF outline has conflicting destinations.");
            if (!item.Data.ContainsKey("Dest") && !item.Data.ContainsKey("A") && !item.Data.ContainsKey("First"))
                throw new NotSupportedException("PDF outline needs one local destination or child outline.");
            IToken? destination = null;
            if (item.Data.TryGetValue("A", out var actionToken))
            {
                var action = Resolve<DictionaryToken>(actionToken, source);
                if (action is null || !action.Data.TryGetValue("S", out var kind) || kind is not NameToken { Data: "GoTo" } ||
                    !action.Data.TryGetValue("D", out destination))
                    throw new NotSupportedException("PDF outline action is not a local GoTo.");
                if (action.Data.ContainsKey("Next"))
                    throw new NotSupportedException("PDF outline action chain is unsupported.");
            }
            else item.Data.TryGetValue("Dest", out destination);
            if (destination is null) continue;
            if (Resolve<ArrayToken>(destination, source) is not { } target || target.Data.Count < 2 ||
                target.Data[0] is not IndirectReferenceToken pageRef || !pageReferences.Contains(pageRef.Data.ToString()))
                throw new InvalidDataException("PDF outline targets an invalid or nonlocal page.");
            if (target.Data[1] is not NameToken destinationType || destinationType.Data is not
                ("Fit" or "XYZ" or "FitH" or "FitV" or "FitR"))
                throw new NotSupportedException("PDF outline destination type is unsupported.");
            destinations[count - 1] = (true, null, null, null, null);
            if (destinationType.Data == "XYZ")
            {
                // PdfPig's parsed destination coordinates omit zoom; a missing trailing operand retains the reader's value.
                if (target.Data.Count > 5) throw new InvalidDataException("PDF XYZ destination has invalid coordinates.");
                var left = CoordinateOperand(target, 2, source);
                var top = CoordinateOperand(target, 3, source);
                if (left.HasValue != top.HasValue)
                    throw new NotSupportedException("PDF partial XYZ destination cannot be represented without changing its position.");
                double? zoom = null;
                if (target.Data.Count == 5)
                {
                    zoom = CoordinateOperand(target, 4, source);
                    if (zoom.HasValue && (zoom.Value != 0 && (zoom.Value < .1 || zoom.Value > 64)))
                        throw new NotSupportedException("PDF XYZ destination zoom cannot be represented in OFD.");
                }
                destinations[count - 1] = (true, zoom, left, top, null);
            }
            else if (destinationType.Data is "FitH" or "FitV")
            {
                if (target.Data.Count > 3) throw new InvalidDataException("PDF fit destination has invalid coordinates.");
                var coordinate = CoordinateOperand(target, 2, source);
                destinations[count - 1] = destinationType.Data == "FitH"
                    ? (true, null, null, coordinate, null) : (true, null, coordinate, null, null);
            }
            else if (destinationType.Data == "FitR")
            {
                if (target.Data.Count != 6) throw new InvalidDataException("PDF rectangle destination is incomplete.");
                var rectangle = new double[4];
                for (var index = 0; index < rectangle.Length; index++)
                    rectangle[index] = CoordinateOperand(target, index + 2, source)
                        ?? throw new InvalidDataException("PDF rectangle destination is incomplete.");
                destinations[count - 1] = (true, null, null, null, rectangle);
            }
        }
        return destinations;
    }

    private static bool SameOutlineLink(IToken left, IToken right) =>
        left is IndirectReferenceToken leftRef && right is IndirectReferenceToken rightRef
            ? leftRef.Data.Equals(rightRef.Data)
            : ReferenceEquals(left, right);

    private static double? CoordinateOperand(ArrayToken destination, int index, PdfDocument source)
    {
        if (index >= destination.Data.Count || destination.Data[index] is NullToken)
            return null;
        double value;
        try
        {
            value = Resolve<NumericToken>(destination.Data[index], source)?.Double
                ?? throw new InvalidDataException("PDF outline destination coordinate is not numeric.");
        }
        catch (PdfDocumentFormatException error)
        {
            throw new InvalidDataException("PDF outline destination coordinate is not numeric.", error);
        }
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new InvalidDataException("PDF outline destination coordinate is not finite.");
        return value;
    }

    private static OfdDestination Destination(Page source, OfdPage page, ExplicitDestination pdf,
        (bool HasDestination, double? Zoom, double? Left, double? Top, double[]? Rect) raw)
    {
        var result = new OfdDestination { Page = page, Type = OfdDestinationType.Fit };
        switch (pdf.Type)
        {
            case ExplicitDestinationType.FitPage:
                return result;
            case ExplicitDestinationType.XyzCoordinates:
                result.Type = OfdDestinationType.XYZ;
                result.Zoom = raw.Zoom;
                if (raw.Left is double x && raw.Top is double y)
                {
                    var point = Point(source, x, y); result.Left = point.X; result.Top = point.Y;
                }
                else if (raw.Left.HasValue || raw.Top.HasValue)
                    throw new NotSupportedException("PDF partial XYZ destination cannot be represented without changing its position.");
                return result;
            case ExplicitDestinationType.FitHorizontally:
                result.Type = source.Rotation.Value is 90 or 270 ? OfdDestinationType.FitV : OfdDestinationType.FitH;
                if (raw.Top is double top)
                {
                    var horizontal = Point(source, source.CropBox.Bounds.Left, top);
                    if (result.Type == OfdDestinationType.FitH) result.Top = horizontal.Y; else result.Left = horizontal.X;
                }
                return result;
            case ExplicitDestinationType.FitVertically:
                result.Type = source.Rotation.Value is 90 or 270 ? OfdDestinationType.FitH : OfdDestinationType.FitV;
                if (raw.Left is double left)
                {
                    var vertical = Point(source, left, source.CropBox.Bounds.Bottom);
                    if (result.Type == OfdDestinationType.FitH) result.Top = vertical.Y; else result.Left = vertical.X;
                }
                return result;
            case ExplicitDestinationType.FitRectangle:
                if (raw.Rect is not { Length: 4 } rectangle)
                    throw new InvalidDataException("PDF rectangle destination is incomplete.");
                var rect = Rectangle(source, rectangle[0], rectangle[1], rectangle[2], rectangle[3]);
                result.Type = OfdDestinationType.FitR; result.Left = rect.X; result.Top = rect.Y;
                result.Right = rect.X + rect.Width; result.Bottom = rect.Y + rect.Height;
                return result;
            default: throw new NotSupportedException("PDF outline destination type is unsupported.");
        }
    }

    private static (double X, double Y) Point(Page page, double x, double y)
    {
        if (double.IsNaN(x) || double.IsInfinity(x) || double.IsNaN(y) || double.IsInfinity(y))
            throw new InvalidDataException("PDF destination or annotation has non-finite coordinates.");
        var crop = page.CropBox.Bounds;
        var rotation = page.Rotation.Value;
        var point = rotation switch
        {
            0 => (x - crop.Left, crop.Top - y),
            90 => (y - crop.Bottom, x - crop.Left),
            180 => (crop.Right - x, y - crop.Bottom),
            270 => (crop.Top - y, crop.Right - x),
            _ => throw new NotSupportedException("PDF metadata page rotation is unsupported.")
        };
        return (point.Item1 * 25.4 / 72, point.Item2 * 25.4 / 72);
    }

    private static (double X, double Y, double Width, double Height) Rectangle(Page page, double left, double bottom, double right, double top)
    {
        if (right == left || top == bottom) throw new InvalidDataException("PDF annotation rectangle is empty.");
        var corners = new[] { Point(page, left, bottom), Point(page, left, top), Point(page, right, bottom), Point(page, right, top) };
        var minX = corners.Min(p => p.X); var maxX = corners.Max(p => p.X);
        var minY = corners.Min(p => p.Y); var maxY = corners.Max(p => p.Y);
        if (!(maxX > minX) || !(maxY > minY)) throw new InvalidDataException("PDF annotation rectangle is not representable.");
        return (minX, minY, maxX - minX, maxY - minY);
    }

    private static void CopyAttachments(PdfDocument source, OfdDocumentPackage output,
        IReadOnlyList<int> selected, IReadOnlyList<bool> drawAttachmentPins, PdfToOfdOptions options, CancellationToken token)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var count = 0; long total = 0;
        var globalFiles = ReadEmbeddedFileNameTree(source, options, token);
        var rawPages = new RawPageLookup(source, options);
        var annotations = new Dictionary<int, Annotation[]>();
        var parsedPages = new Dictionary<int, Page>();
        var indirectPayloads = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var directPayloads = new List<(StreamToken Stream, byte[] Bytes)>();
        OfdAttachment Add(string name, IToken streamLink, StreamToken stream)
        {
            token.ThrowIfCancellationRequested();
            if (++count > options.MaxAttachmentCount) throw new InvalidDataException("PDF attachment count exceeds the configured limit.");
            var remaining = Math.Min(options.MaxAttachmentBytes, options.MaxTotalAttachmentBytes - total);
            if (remaining < 0) throw new InvalidDataException("PDF attachment byte limit exceeded.");
            byte[] bytes;
            if (streamLink is IndirectReferenceToken reference)
            {
                var key = reference.Data.ToString();
                if (indirectPayloads.TryGetValue(key, out var cached)) bytes = cached;
                else
                {
                    bytes = Decode(stream, remaining, options.MaxInputBytes, token);
                    indirectPayloads.Add(key, bytes);
                }
            }
            else
            {
                var existing = directPayloads.FindIndex(item => ReferenceEquals(item.Stream, stream));
                if (existing >= 0) bytes = directPayloads[existing].Bytes;
                else
                {
                    bytes = Decode(stream, remaining, options.MaxInputBytes, token);
                    directPayloads.Add((stream, bytes));
                }
            }
            if (bytes.LongLength > remaining) throw new InvalidDataException("PDF attachment decoded byte limit exceeded.");
            total = checked(total + bytes.LongLength);
            var display = Name(name, names);
            var attachment = new OfdAttachment { Id = "800000" + count, Name = display, Data = bytes,
                MediaType = "application/octet-stream", SizeKilobytes = bytes.LongLength / 1024d, Visible = true };
            output.Attachments.Add(attachment);
            return attachment;
        }

        foreach (var file in globalFiles) Add(file.Name, file.Link, file.Stream);

        for (var index = 0; index < selected.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            if (!annotations.TryGetValue(selected[index], out var pageAnnotations))
            {
                pageAnnotations = ReadPageAnnotations(source, selected[index], options, rawPages, token, out _);
                annotations.Add(selected[index], pageAnnotations);
            }
            foreach (var annotation in pageAnnotations)
            {
                token.ThrowIfCancellationRequested();
                if (annotation.Type != AnnotationType.FileAttachment) continue;
                if (!parsedPages.TryGetValue(selected[index], out var pdfPage))
                {
                    pdfPage = source.GetPage(selected[index] + 1);
                    parsedPages.Add(selected[index], pdfPage);
                }
                // A rasterized PDF icon cannot be separated from the page image later. Only the
                // visible, printable form can be preserved across all importer/rasterizer paths.
                var flags = annotation.AnnotationDictionary.Data.TryGetValue("F", out var flagsToken)
                    ? Resolve<NumericToken>(flagsToken, source)?.Double ?? throw new InvalidDataException("PDF FileAttachment flags are not numeric.")
                    : 0;
                if (flags != 4) throw new NotSupportedException("PDF FileAttachment visibility or print flags are unsupported.");
                var spec = Resolve<DictionaryToken>(annotation.AnnotationDictionary.Data.TryGetValue("FS", out var fs) ? fs : null, source);
                if (spec is null) throw new InvalidDataException("PDF FileAttachment has no embedded file specification.");
                var name = Text(spec.Data.TryGetValue("UF", out var unicode) ? unicode : spec.Data.TryGetValue("F", out var fallback) ? fallback : null, source)
                    ?? throw new InvalidDataException("PDF FileAttachment has no file name.");
                var ef = Resolve<DictionaryToken>(spec.Data.TryGetValue("EF", out var efToken) ? efToken : null, source);
                if (ef is null) throw new NotSupportedException("PDF external FileAttachment has no embedded payload.");
                var streamLink = ef.Data.TryGetValue("UF", out var data) ? data : ef.Data.TryGetValue("F", out var other) ? other : null;
                var payload = Resolve<StreamToken>(streamLink, source)
                    ?? throw new InvalidDataException("PDF FileAttachment has no embedded stream.");
                var attachment = Add(name, streamLink!, payload);
                // PdfPig may expose Annotation.Rectangle in displayed coordinates
                // for rotated pages. The source /Rect is always PDF user space.
                if (!annotation.AnnotationDictionary.Data.TryGetValue("Rect", out var rectToken) ||
                    Resolve<ArrayToken>(rectToken, source) is not { } rawRect || rawRect.Data.Count != 4)
                    throw new InvalidDataException("PDF FileAttachment has no four-coordinate rectangle.");
                var values = rawRect.Data.Select(item => Resolve<NumericToken>(item, source)?.Double ??
                    throw new InvalidDataException("PDF FileAttachment rectangle is not numeric.")).ToArray();
                var box = Rectangle(pdfPage, values[0], values[1], values[2], values[3]);
                var page = output.Pages[index]; page.Annotations ??= new OfdAnnotationCollection();
                var link = new OfdAnnotation { Type = OfdAnnotationType.Link, Creator = "Ofdrw.Net PdfToOfdConverter",
                    LastModifiedDate = new DateTime(2000, 1, 1), Visible = true, Print = true,
                    Remark = annotation.Content, AttachmentId = attachment.Id,
                    Appearance = new OfdAnnotationAppearance { XMillimeters = box.X, YMillimeters = box.Y,
                        WidthMillimeters = box.Width, HeightMillimeters = box.Height } };
                link.Appearance.Elements.Add(new OfdPathElement { XMillimeters = 0, YMillimeters = 0,
                    WidthMillimeters = box.Width, HeightMillimeters = box.Height,
                    // A raster page already paints the PDF icon. Native pages draw a generic PushPin.
                    // Both retain the action on this path's full annotation boundary.
                    AbbreviatedData = "M " + P(.55, .08) + " L " + P(.84, .38) + " L " + P(.69, .51) +
                        " L " + P(.79, .65) + " L " + P(.66, .78) + " L " + P(.52, .65) +
                        " L " + P(.16, .91) + " L " + P(.10, .85) + " L " + P(.36, .49) +
                        " L " + P(.22, .34) + " L " + P(.35, .21) + " L " + P(.48, .32) + " C",
                    Fill = drawAttachmentPins[index], Stroke = drawAttachmentPins[index],
                    FillColor = new OfdColor(190, 197, 187), StrokeColor = new OfdColor(110, 118, 107) });
                page.Annotations.Items.Add(link);
                string P(double x, double y) =>
                    (x * box.Width).ToString("R", System.Globalization.CultureInfo.InvariantCulture) + " " +
                    (y * box.Height).ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            }
        }
    }

    private static List<(string Name, IToken Link, StreamToken Stream)> ReadEmbeddedFileNameTree(PdfDocument source, PdfToOfdOptions options, CancellationToken token)
    {
        var result = new List<(string Name, IToken Link, StreamToken Stream)>();
        var catalog = source.Structure.Catalog.CatalogDictionary;
        if (!catalog.Data.TryGetValue("Names", out var namesToken)) return result;
        var names = Resolve<DictionaryToken>(namesToken, source) ?? throw new InvalidDataException("PDF names catalog cannot be resolved.");
        if (!names.Data.TryGetValue("EmbeddedFiles", out var treeToken)) return result;
        var pending = new Stack<IToken>(); pending.Push(treeToken);
        var nodes = 0;
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            if (++nodes > options.MaxAttachmentCount * 2L) throw new InvalidDataException("PDF embedded-file name tree exceeds the configured limit.");
            var node = Resolve<DictionaryToken>(pending.Pop(), source) ?? throw new InvalidDataException("PDF embedded-file name-tree node cannot be resolved.");
            if (node.Data.TryGetValue("Kids", out var kidsToken))
            {
                if (node.Data.ContainsKey("Names") || Resolve<ArrayToken>(kidsToken, source) is not { } kids || kids.Data.Count == 0)
                    throw new InvalidDataException("PDF embedded-file name tree mixes child and file entries.");
                foreach (var child in kids.Data.Reverse()) pending.Push(child);
            }
            else
            {
                if (!node.Data.TryGetValue("Names", out var filesToken) || Resolve<ArrayToken>(filesToken, source) is not { } files ||
                    files.Data.Count % 2 != 0)
                    throw new InvalidDataException("PDF embedded-file name tree has malformed file pairs.");
                for (var index = 0; index < files.Data.Count; index += 2)
                {
                    token.ThrowIfCancellationRequested();
                    if (result.Count >= options.MaxAttachmentCount || Text(files.Data[index], source) is not { } treeName ||
                        Resolve<DictionaryToken>(files.Data[index + 1], source) is not { } specification ||
                        !specification.Data.TryGetValue("EF", out var payloadToken) ||
                        Resolve<DictionaryToken>(payloadToken, source) is not { } payload ||
                        !payload.Data.TryGetValue("UF", out var streamToken) && !payload.Data.TryGetValue("F", out streamToken) ||
                        Resolve<StreamToken>(streamToken, source) is not { } stream)
                        throw new InvalidDataException("PDF embedded-file name tree has an invalid local payload.");
                    result.Add((treeName, streamToken, stream));
                }
            }
        }
        return result;
    }

    private static string Name(string raw, HashSet<string> used)
    {
        var normalized = raw.Replace('\\', '/');
        normalized = normalized.Substring(normalized.LastIndexOf('/') + 1).Trim();
        if (normalized.Length == 0 || normalized.Length > 255 || normalized.Any(char.IsControl))
            throw new InvalidDataException("PDF attachment name is empty, excessive or contains control characters.");
        if (used.Add(normalized)) return normalized;
        var existing = used.Count;
        for (var index = 2; index <= existing + 2; index++)
        {
            var suffix = " (" + index + ")";
            var candidate = normalized + suffix;
            if (candidate.Length > 255)
            {
                var dot = normalized.LastIndexOf('.');
                var extension = dot > 0 && normalized.Length - dot <= 32 ? normalized.Substring(dot) : string.Empty;
                var stem = normalized.Substring(0, normalized.Length - extension.Length);
                var available = 255 - suffix.Length - extension.Length;
                if (available <= 0) throw new InvalidDataException("PDF attachment name cannot fit a collision suffix.");
                var length = Math.Min(stem.Length, available);
                if (length < stem.Length && char.IsHighSurrogate(stem[length - 1]) && char.IsLowSurrogate(stem[length]))
                    length--;
                candidate = stem.Substring(0, length) + suffix + extension;
            }
            if (used.Add(candidate)) return candidate;
        }
        throw new InvalidDataException("PDF attachment names cannot be made unique within the configured limit.");
    }

    private static T? Resolve<T>(IToken? token, PdfDocument source) where T : class, IToken
        => token is null ? null : DirectObjectFinder.Get<T>(token, source.Structure.TokenScanner);

    private static string? Text(IToken? token, PdfDocument source)
    {
        if (token is null) return null;
        if (token is StringToken literal) return literal.Data;
        if (token is HexToken hex) return hex.Data;
        try
        {
            return Resolve<StringToken>(token, source)?.Data;
        }
        catch (PdfDocumentFormatException) { }
        try
        {
            return Resolve<HexToken>(token, source)?.Data;
        }
        catch (PdfDocumentFormatException) { return null; }
    }

    private static byte[] Decode(StreamToken stream, long decodedLimit, long encodedLimit, CancellationToken token)
    {
        if (stream.StreamDictionary.Data.ContainsKey("DecodeParms") || stream.StreamDictionary.Data.ContainsKey("F") ||
            stream.StreamDictionary.Data.ContainsKey("FFilter") || stream.StreamDictionary.Data.ContainsKey("FDecodeParms"))
            throw new NotSupportedException("PDF embedded-file filter parameters are unsupported.");
        var source = stream.Data.ToArray();
        if (source.LongLength > encodedLimit) throw new InvalidDataException("PDF attachment encoded byte limit exceeded.");
        if (!stream.StreamDictionary.Data.TryGetValue("Filter", out var filter))
        {
            if (source.LongLength > decodedLimit) throw new InvalidDataException("PDF attachment decoded byte limit exceeded.");
            return source;
        }
        if (filter is ArrayToken { Data.Count: 1 } singleFilter) filter = singleFilter.Data[0];
        if (filter is not NameToken { Data: "FlateDecode" } && filter is not NameToken { Data: "Fl" })
            throw new NotSupportedException("PDF embedded-file filter is unsupported.");
        if (source.Length < 6 || (source[0] & 15) != 8 || ((source[0] << 8) + source[1]) % 31 != 0 || (source[1] & 32) != 0)
            throw new InvalidDataException("PDF embedded-file zlib envelope is invalid.");
        using var compressed = new MemoryStream(source, 2, source.Length - 6, false);
        using var inflater = new DeflateStream(compressed, CompressionMode.Decompress);
        using var output = new MemoryStream();
        var buffer = new byte[8192];
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var length = inflater.Read(buffer, 0, buffer.Length);
            if (length == 0) break;
            if (length > decodedLimit - output.Length) throw new InvalidDataException("PDF attachment decoded byte limit exceeded.");
            output.Write(buffer, 0, length);
        }
        var bytes = output.ToArray();
        uint a = 1, b = 0;
        foreach (var value in bytes) { a = (a + value) % 65521; b = (b + a) % 65521; }
        var checksum = ((uint)source[source.Length - 4] << 24) | ((uint)source[source.Length - 3] << 16) |
            ((uint)source[source.Length - 2] << 8) | source[source.Length - 1];
        if (checksum != (b << 16 | a)) throw new InvalidDataException("PDF attachment checksum mismatch.");
        return bytes;
    }
}
