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
using UglyToad.PdfPig.Outline;
using UglyToad.PdfPig.Outline.Destinations;
using UglyToad.PdfPig.Parser.Parts;
using UglyToad.PdfPig.Tokens;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>Copies only closed local PDF navigation and embedded-file references into a fresh OFD package.</summary>
internal static class PdfImportMetadataMigrator
{
    internal static bool HasOnlyFileAttachments(PdfDocument source, int zeroBasedPage, PdfToOfdOptions options, CancellationToken token)
    {
        if (!options.CopyAttachments) return false;
        _ = ReadPageAnnotations(source, zeroBasedPage, options, token, out var onlyFileAttachments);
        return onlyFileAttachments;
    }

    private static Annotation[] ReadPageAnnotations(PdfDocument source, int zeroBasedPage, PdfToOfdOptions options,
        CancellationToken token, out bool onlyFileAttachments)
    {
        onlyFileAttachments = true;
        var page = RawPage(source, zeroBasedPage, options, token);
        if (!page.Data.TryGetValue("Annots", out var annotsToken)) return Array.Empty<Annotation>();
        var annots = Resolve<ArrayToken>(annotsToken, source) ?? throw new InvalidDataException("PDF annotation array cannot be resolved.");
        var rawFileCount = 0;
        foreach (var entry in annots.Data)
        {
            token.ThrowIfCancellationRequested();
            var annotation = Resolve<DictionaryToken>(entry, source) ?? throw new InvalidDataException("PDF page annotation cannot be resolved.");
            if (annotation.Data.TryGetValue("Subtype", out var subtype) && subtype is NameToken { Data: "FileAttachment" })
            {
                if (++rawFileCount > options.MaxAttachmentCount)
                    throw new InvalidDataException("PDF page attachment count exceeds the configured limit.");
            }
            else onlyFileAttachments = false;
        }
        var parsed = source.GetPage(zeroBasedPage + 1).GetAnnotations().ToArray();
        if (parsed.Count(annotation => annotation.Type == AnnotationType.FileAttachment) != rawFileCount ||
            onlyFileAttachments && parsed.Length != annots.Data.Count)
            throw new InvalidDataException("PDF page attachment parser omitted an annotation.");
        return parsed;
    }

    private static DictionaryToken RawPage(PdfDocument source, int zeroBasedPage, PdfToOfdOptions options, CancellationToken token)
    {
        if (!source.Structure.Catalog.CatalogDictionary.Data.TryGetValue("Pages", out var root))
            throw new InvalidDataException("PDF page tree is missing.");
        var pending = new Stack<IToken>(); pending.Push(root);
        var visits = 0; var ordinal = 0;
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            if (++visits > options.MaxPageCount * 2L) throw new InvalidDataException("PDF page tree exceeds the configured limit.");
            var page = Resolve<DictionaryToken>(pending.Pop(), source) ?? throw new InvalidDataException("PDF page tree node cannot be resolved.");
            if (page.Data.TryGetValue("Type", out var type) && type is NameToken { Data: "Page" })
            {
                if (ordinal++ == zeroBasedPage) return page;
                continue;
            }
            if (!page.Data.TryGetValue("Kids", out var children) || Resolve<ArrayToken>(children, source) is not { } kids)
                throw new InvalidDataException("PDF page tree has no children.");
            foreach (var child in kids.Data.Reverse()) pending.Push(child);
        }
        throw new InvalidDataException("PDF selected page was not found in the page tree.");
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
                    Destination = Destination(source.GetPage(local.PageNumber), retained[0], local.Destination, rawDestinations[count - 1].Zoom) });
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

    private static List<(bool HasDestination, double? Zoom)> ValidateRawOutlines(PdfDocument source, PdfToOfdOptions options, CancellationToken token)
    {
        var destinations = new List<(bool HasDestination, double? Zoom)>();
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
            if (node.Data.TryGetValue("Type", out var type) && type is NameToken { Data: "Page" })
            {
                if (current is not IndirectReferenceToken reference || !pageReferences.Add(reference.Data.ToString()))
                    throw new InvalidDataException("PDF page reference is missing or repeated.");
                continue;
            }
            if (!node.Data.TryGetValue("Kids", out var children) || Resolve<ArrayToken>(children, source) is not { } kids)
                throw new InvalidDataException("PDF page tree has no children.");
            foreach (var child in kids.Data) pendingPages.Push(child);
        }
        var pending = new Stack<(IToken Node, int Depth)>(); pending.Push((first, 1));
        var count = 0;
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            var (link, depth) = pending.Pop();
            if (depth > options.MaxBookmarkDepth || ++count > options.MaxBookmarkCount)
                throw new InvalidDataException("PDF outline exceeds the configured node or depth limit.");
            destinations.Add((false, null));
            var item = Resolve<DictionaryToken>(link, source) ?? throw new InvalidDataException("PDF outline node cannot be resolved.");
            if (item.Data.TryGetValue("Next", out var next)) pending.Push((next, depth));
            if (item.Data.TryGetValue("First", out var child)) pending.Push((child, depth + 1));
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
            destinations[count - 1] = (true, null);
            if (destinationType.Data == "XYZ")
            {
                // PdfPig's parsed destination coordinates omit zoom; a missing trailing operand retains the reader's value.
                if (target.Data.Count > 5) throw new InvalidDataException("PDF XYZ destination has invalid coordinates.");
                if (target.Data.Count == 5 && target.Data[4] is not NullToken)
                {
                    var zoom = Resolve<NumericToken>(target.Data[4], source)?.Double
                        ?? throw new InvalidDataException("PDF XYZ destination zoom is not numeric.");
                    if (double.IsNaN(zoom) || double.IsInfinity(zoom) || zoom != 0 && (zoom < .1 || zoom > 64))
                        throw new NotSupportedException("PDF XYZ destination zoom cannot be represented in OFD.");
                    destinations[count - 1] = (true, zoom);
                }
            }
        }
        return destinations;
    }

    private static OfdDestination Destination(Page source, OfdPage page, ExplicitDestination pdf, double? rawZoom)
    {
        var result = new OfdDestination { Page = page, Type = OfdDestinationType.Fit };
        var coordinates = pdf.Coordinates;
        switch (pdf.Type)
        {
            case ExplicitDestinationType.FitPage:
                return result;
            case ExplicitDestinationType.XyzCoordinates:
                result.Type = OfdDestinationType.XYZ;
                result.Zoom = rawZoom;
                if (coordinates.Left is double x && coordinates.Top is double y)
                {
                    var point = Point(source, x, y); result.Left = point.X; result.Top = point.Y;
                }
                else if (coordinates.Left.HasValue || coordinates.Top.HasValue)
                    throw new NotSupportedException("PDF partial XYZ destination cannot be represented without changing its position.");
                return result;
            case ExplicitDestinationType.FitHorizontally:
                result.Type = source.Rotation.Value is 90 or 270 ? OfdDestinationType.FitV : OfdDestinationType.FitH;
                if (coordinates.Top is double top)
                {
                    var horizontal = Point(source, source.CropBox.Bounds.Left, top);
                    if (result.Type == OfdDestinationType.FitH) result.Top = horizontal.Y; else result.Left = horizontal.X;
                }
                return result;
            case ExplicitDestinationType.FitVertically:
                result.Type = source.Rotation.Value is 90 or 270 ? OfdDestinationType.FitH : OfdDestinationType.FitV;
                if (coordinates.Left is double left)
                {
                    var vertical = Point(source, left, source.CropBox.Bounds.Bottom);
                    if (result.Type == OfdDestinationType.FitH) result.Top = vertical.Y; else result.Left = vertical.X;
                }
                return result;
            case ExplicitDestinationType.FitRectangle:
                if (coordinates.Left is not double l || coordinates.Right is not double r ||
                    coordinates.Top is not double t || coordinates.Bottom is not double b)
                    throw new InvalidDataException("PDF rectangle destination is incomplete.");
                var rect = Rectangle(source, l, b, r, t);
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
        if (!(right > left) || !(top > bottom)) throw new InvalidDataException("PDF annotation rectangle is empty or reversed.");
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
        var globalCount = ValidateEmbeddedFileNameTree(source, options, token);
        OfdAttachment Add(string name, StreamToken stream)
        {
            token.ThrowIfCancellationRequested();
            if (++count > options.MaxAttachmentCount) throw new InvalidDataException("PDF attachment count exceeds the configured limit.");
            var remaining = Math.Min(options.MaxAttachmentBytes, options.MaxTotalAttachmentBytes - total);
            if (remaining < 0) throw new InvalidDataException("PDF attachment byte limit exceeded.");
            var bytes = Decode(stream, remaining, options.MaxInputBytes, token);
            total = checked(total + bytes.LongLength);
            var display = Name(name, names);
            var attachment = new OfdAttachment { Id = "800000" + count, Name = display, Data = bytes,
                MediaType = "application/octet-stream", SizeKilobytes = bytes.LongLength / 1024d, Visible = true };
            output.Attachments.Add(attachment);
            return attachment;
        }

        if (source.Advanced.TryGetEmbeddedFiles(out var embedded))
        {
            var files = embedded.ToArray();
            if (files.Length != globalCount) throw new InvalidDataException("PDF embedded-file name tree contains an omitted file.");
            foreach (var file in files)
            {
                token.ThrowIfCancellationRequested();
                Add(file.Name ?? file.FileSpecification ?? "attachment", file.Stream);
            }
        }
        else if (globalCount != 0) throw new InvalidDataException("PDF embedded-file name tree could not be read.");

        for (var index = 0; index < selected.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            var pdfPage = source.GetPage(selected[index] + 1);
            foreach (var annotation in ReadPageAnnotations(source, selected[index], options, token, out _))
            {
                token.ThrowIfCancellationRequested();
                if (annotation.Type != AnnotationType.FileAttachment) continue;
                var spec = Resolve<DictionaryToken>(annotation.AnnotationDictionary.Data.TryGetValue("FS", out var fs) ? fs : null, source);
                if (spec is null) throw new InvalidDataException("PDF FileAttachment has no embedded file specification.");
                var name = Text(spec.Data.TryGetValue("UF", out var unicode) ? unicode : spec.Data.TryGetValue("F", out var fallback) ? fallback : null, source)
                    ?? throw new InvalidDataException("PDF FileAttachment has no file name.");
                var ef = Resolve<DictionaryToken>(spec.Data.TryGetValue("EF", out var efToken) ? efToken : null, source);
                if (ef is null) throw new NotSupportedException("PDF external FileAttachment has no embedded payload.");
                var payload = Resolve<StreamToken>(ef.Data.TryGetValue("UF", out var data) ? data : ef.Data.TryGetValue("F", out var other) ? other : null, source)
                    ?? throw new InvalidDataException("PDF FileAttachment has no embedded stream.");
                var attachment = Add(name, payload);
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

    private static int ValidateEmbeddedFileNameTree(PdfDocument source, PdfToOfdOptions options, CancellationToken token)
    {
        var catalog = source.Structure.Catalog.CatalogDictionary;
        if (!catalog.Data.TryGetValue("Names", out var namesToken)) return 0;
        var names = Resolve<DictionaryToken>(namesToken, source) ?? throw new InvalidDataException("PDF names catalog cannot be resolved.");
        if (!names.Data.TryGetValue("EmbeddedFiles", out var treeToken)) return 0;
        var pending = new Stack<IToken>(); pending.Push(treeToken);
        var nodes = 0; var count = 0;
        while (pending.Count != 0)
        {
            token.ThrowIfCancellationRequested();
            if (++nodes > options.MaxAttachmentCount * 2L) throw new InvalidDataException("PDF embedded-file name tree exceeds the configured limit.");
            var node = Resolve<DictionaryToken>(pending.Pop(), source) ?? throw new InvalidDataException("PDF embedded-file name-tree node cannot be resolved.");
            if (node.Data.TryGetValue("Kids", out var kidsToken))
            {
                if (node.Data.ContainsKey("Names") || Resolve<ArrayToken>(kidsToken, source) is not { } kids || kids.Data.Count == 0)
                    throw new InvalidDataException("PDF embedded-file name tree mixes child and file entries.");
                foreach (var child in kids.Data) pending.Push(child);
            }
            else
            {
                if (!node.Data.TryGetValue("Names", out var filesToken) || Resolve<ArrayToken>(filesToken, source) is not { } files ||
                    files.Data.Count % 2 != 0)
                    throw new InvalidDataException("PDF embedded-file name tree has malformed file pairs.");
                for (var index = 0; index < files.Data.Count; index += 2)
                {
                    token.ThrowIfCancellationRequested();
                    if (++count > options.MaxAttachmentCount || Text(files.Data[index], source) is null ||
                        Resolve<DictionaryToken>(files.Data[index + 1], source) is not { } specification ||
                        !specification.Data.TryGetValue("EF", out var payloadToken) ||
                        Resolve<DictionaryToken>(payloadToken, source) is not { } payload ||
                        !payload.Data.TryGetValue("UF", out var streamToken) && !payload.Data.TryGetValue("F", out streamToken) ||
                        Resolve<StreamToken>(streamToken, source) is null)
                        throw new InvalidDataException("PDF embedded-file name tree has an invalid local payload.");
                }
            }
        }
        return count;
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
                candidate = stem.Substring(0, Math.Min(stem.Length, available)) + suffix + extension;
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
        return DirectObjectFinder.Get<StringToken>(token, source.Structure.TokenScanner)?.Data ??
            DirectObjectFinder.Get<HexToken>(token, source.Structure.TokenScanner)?.Data;
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
