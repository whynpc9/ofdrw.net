using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Signatures.Signing;

/// <summary>Signs OFD packages and writes optional signed seal appearances.</summary>
public sealed partial class OfdSignatureService
{
    private static IReadOnlyList<XElement> BuildStampAppearances(
        IReadOnlyDictionary<string, byte[]> entries, XDocument ofd, XElement documentBody,
        OfdSignatureOptions options, string? sealPath, int signatureId,
        CancellationToken cancellationToken)
    {
        if (options.StampAppearances.Count == 0) return Array.Empty<XElement>();
        if (sealPath is null)
            throw new InvalidDataException("A stamp appearance requires nonempty provider SealData.");
        if (ofd.Descendants().Count(node => node.Name.LocalName == "DocBody") != 1)
            throw new NotSupportedException("Stamp appearances require exactly one DocBody.");

        var pages = ReadStampPages(entries, documentBody, cancellationToken);
        var result = new List<XElement>();
        foreach (var placement in options.StampAppearances)
        {
            cancellationToken.ThrowIfCancellationRequested();
            switch (placement)
            {
                case OfdNormalStampPlacement normal:
                    if (normal.PageNumber < 1 || normal.PageNumber > pages.Count)
                        throw new InvalidDataException("Normal stamp page number is outside the document.");
                    var normalBox = new StampBox(normal.XMillimeters, normal.YMillimeters,
                        normal.WidthMillimeters, normal.HeightMillimeters);
                    EnsureStampBox(normalBox, "Normal stamp Boundary");
                    EnsureContained(pages[normal.PageNumber - 1].Area, normalBox,
                        "Normal stamp Boundary");
                    result.Add(StampXml(signatureId, result.Count + 1,
                        pages[normal.PageNumber - 1].Id, normalBox, null));
                    break;
                case OfdRidingStampPlacement riding:
                    AddRidingStamps(result, signatureId, pages, riding, cancellationToken);
                    break;
                default:
                    throw new NotSupportedException("Unsupported stamp appearance placement.");
            }
            if (result.Count > 1000)
                throw new NotSupportedException("A signature cannot create more than 1000 stamp appearances.");
        }
        return result;
    }

    private static List<StampPage> ReadStampPages(IReadOnlyDictionary<string, byte[]> entries,
        XElement documentBody, CancellationToken token)
    {
        var roots = documentBody.Elements().Where(node => node.Name.LocalName == "DocRoot").ToArray();
        if (roots.Length != 1 || string.IsNullOrWhiteSpace(roots[0].Value))
            throw new InvalidDataException("Stamp appearance requires exactly one DocRoot.");
        var documentPath = NormalizePath(roots[0].Value.Trim());
        if (!entries.TryGetValue(documentPath, out var documentBytes))
            throw new InvalidDataException("Stamp appearance Document.xml is missing.");
        var document = ParseStampXml(documentBytes, documentPath);
        var documentRoot = document.Root;
        if (documentRoot?.Name.LocalName != "Document")
            throw new InvalidDataException("Stamp appearance requires a Document root.");
        var ns = documentRoot.Name.Namespace;
        var defaultArea = ParseStampBox(documentRoot.Element(ns + "CommonData")?
            .Element(ns + "PageArea")?.Element(ns + "PhysicalBox")?.Value,
            "Document PhysicalBox");
        var pageNodes = documentRoot.Element(ns + "Pages")?.Elements(ns + "Page").ToArray()
            ?? throw new InvalidDataException("Stamp appearance requires document pages.");
        if (pageNodes.Length == 0)
            throw new InvalidDataException("Stamp appearance requires at least one page.");
        var pages = new List<StampPage>(pageNodes.Length);
        // Readers match PageRef case-insensitively; reject ambiguous page IDs before signing.
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var node in pageNodes)
        {
            token.ThrowIfCancellationRequested();
            var id = node.Attribute("ID")?.Value;
            var baseLoc = node.Attribute("BaseLoc")?.Value?.Trim();
            if (string.IsNullOrWhiteSpace(id) || !ids.Add(id!) || string.IsNullOrWhiteSpace(baseLoc))
                throw new InvalidDataException("Stamp appearance page has a missing or duplicate ID/BaseLoc.");
            var pagePath = baseLoc!.StartsWith("/", StringComparison.Ordinal)
                ? NormalizePath(baseLoc)
                : NormalizePath(GetDirectory(documentPath) + baseLoc);
            if (!entries.TryGetValue(pagePath, out var pageBytes))
                throw new InvalidDataException("Stamp appearance page entry is missing: /" + pagePath);
            var page = ParseStampXml(pageBytes, pagePath);
            if (page.Root?.Name.LocalName != "Page")
                throw new InvalidDataException("Stamp appearance requires a Page root: /" + pagePath);
            var pageNs = page.Root.Name.Namespace;
            var ownArea = page.Root.Element(pageNs + "Area")?
                .Element(pageNs + "PhysicalBox")?.Value;
            var area = ownArea is null ? defaultArea : ParseStampBox(ownArea, "Page PhysicalBox");
            pages.Add(new StampPage(id!, area));
        }
        return pages;
    }

    private static void AddRidingStamps(List<XElement> destination, int signatureId,
        IReadOnlyList<StampPage> pages, OfdRidingStampPlacement placement, CancellationToken token)
    {
        if (!Enum.IsDefined(typeof(OfdStampEdge), placement.Edge) ||
            placement.FirstPageNumber < 1 || placement.PageCount < 1 ||
            placement.FirstPageNumber > pages.Count ||
            placement.PageCount > pages.Count - placement.FirstPageNumber + 1 ||
            placement.GroupSize < 0 || placement.GroupSize > placement.PageCount)
            throw new InvalidDataException("Invalid riding stamp edge, page range or group size.");
        if (placement.PageCount > 1000 - destination.Count)
            throw new NotSupportedException("A signature cannot create more than 1000 stamp appearances.");
        var whole = new StampBox(0, 0, placement.WidthMillimeters, placement.HeightMillimeters);
        EnsureStampBox(whole, "Riding stamp size");
        if (!Finite(placement.MarginMillimeters) || placement.MarginMillimeters < 0 ||
            placement.OffsetMillimeters.HasValue && !Finite(placement.OffsetMillimeters.Value))
            throw new InvalidDataException("Invalid riding stamp margin or offset.");
        var groupSize = placement.GroupSize == 0 ? placement.PageCount : placement.GroupSize;
        if (placement.SliceWeights.Count != 0 && placement.SliceWeights.Count != groupSize)
            throw new InvalidDataException("Riding stamp weights must match the full group size.");
        if (placement.SliceWeights.Any(weight => !Finite(weight) || weight <= 0))
            throw new InvalidDataException("Riding stamp weights must be finite and positive.");
        var vertical = placement.Edge is OfdStampEdge.Left or OfdStampEdge.Right;
        var span = vertical ? placement.WidthMillimeters : placement.HeightMillimeters;
        for (var groupStart = 0; groupStart < placement.PageCount; groupStart += groupSize)
        {
            var count = Math.Min(groupSize, placement.PageCount - groupStart);
            var weights = Enumerable.Range(0, count).Select(index =>
                placement.SliceWeights.Count == 0 ? 1d : placement.SliceWeights[index]).ToArray();
            var sum = weights.Sum();
            if (!Finite(sum) || sum <= 0)
                throw new InvalidDataException("Riding stamp weight sum is not representable.");
            var cumulative = 0d;
            for (var index = 0; index < count; index++)
            {
                token.ThrowIfCancellationRequested();
                var slice = index == count - 1 ? span - cumulative : span * (weights[index] / sum);
                if (!Finite(slice) || slice <= 0)
                    throw new InvalidDataException("Riding stamp slice is not representable.");
                var reverse = placement.Edge is OfdStampEdge.Left or OfdStampEdge.Top;
                var clipStart = reverse ? span - cumulative - slice : cumulative;
                var clip = vertical
                    ? new StampBox(clipStart, 0, slice, placement.HeightMillimeters)
                    : new StampBox(0, clipStart, placement.WidthMillimeters, slice);
                var page = pages[placement.FirstPageNumber - 1 + groupStart + index];
                var area = page.Area;
                var along = placement.OffsetMillimeters ?? (vertical
                    ? (area.Height - placement.HeightMillimeters) / 2
                    : (area.Width - placement.WidthMillimeters) / 2);
                var x = vertical
                    ? placement.Edge == OfdStampEdge.Left
                        ? area.X + placement.MarginMillimeters - clip.X
                        : area.X + area.Width - placement.MarginMillimeters - clip.X - slice
                    : area.X + along;
                var y = vertical
                    ? area.Y + along
                    : placement.Edge == OfdStampEdge.Top
                        ? area.Y + placement.MarginMillimeters - clip.Y
                        : area.Y + area.Height - placement.MarginMillimeters - clip.Y - slice;
                var boundary = new StampBox(x, y, placement.WidthMillimeters, placement.HeightMillimeters);
                EnsureStampBox(boundary, "Riding stamp Boundary");
                EnsureContained(area, new StampBox(x + clip.X, y + clip.Y, clip.Width, clip.Height),
                    "Riding stamp visible slice");
                destination.Add(StampXml(signatureId, destination.Count + 1, page.Id, boundary, clip));
                cumulative += slice;
            }
        }
    }

    private static XElement StampXml(int signatureId, int ordinal, string pageId,
        StampBox boundary, StampBox? clip)
    {
        XNamespace ns = "http://www.ofdspec.org/2016";
        return new XElement(ns + "StampAnnot",
            new XAttribute("ID", "s" + signatureId.ToString(CultureInfo.InvariantCulture) +
                "_" + ordinal.ToString(CultureInfo.InvariantCulture)),
            new XAttribute("PageRef", pageId),
            new XAttribute("Boundary", boundary.ToString()),
            clip.HasValue ? new XAttribute("Clip", clip.Value.ToString()) : null);
    }

    private static StampBox ParseStampBox(string? text, string field)
    {
        var values = text?.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (values is null || values.Length != 4)
            throw new InvalidDataException(field + " must contain four finite numbers.");
        var numbers = new double[4];
        for (var index = 0; index < 4; index++)
        {
            if (!double.TryParse(values[index], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out numbers[index]))
                throw new InvalidDataException(field + " must contain four finite numbers.");
        }
        var box = new StampBox(numbers[0], numbers[1], numbers[2], numbers[3]);
        EnsureStampBox(box, field);
        return box;
    }

    private static XDocument ParseStampXml(byte[] bytes, string path)
    {
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 16 * 1024 * 1024
            });
            return XDocument.Load(reader);
        }
        catch (Exception error) when (error is XmlException or InvalidOperationException)
        {
            throw new InvalidDataException("Cannot parse stamp appearance page metadata: /" + path, error);
        }
    }

    private static void EnsureStampBox(StampBox box, string field)
    {
        if (!Finite(box.X) || !Finite(box.Y) || !Finite(box.Width) || !Finite(box.Height) ||
            box.Width <= 0 || box.Height <= 0 || !Finite(box.X + box.Width) ||
            !Finite(box.Y + box.Height))
            throw new InvalidDataException(field + " has non-finite coordinates or non-positive dimensions.");
    }

    private static void EnsureContained(StampBox page, StampBox visible, string field)
    {
        EnsureStampBox(visible, field);
        const double epsilon = 1e-7;
        if (visible.X < page.X - epsilon || visible.Y < page.Y - epsilon ||
            visible.X + visible.Width > page.X + page.Width + epsilon ||
            visible.Y + visible.Height > page.Y + page.Height + epsilon)
            throw new InvalidDataException(field + " is outside the referenced page.");
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private readonly struct StampPage
    {
        internal StampPage(string id, StampBox area) { Id = id; Area = area; }
        internal string Id { get; }
        internal StampBox Area { get; }
    }

    private readonly struct StampBox
    {
        internal StampBox(double x, double y, double width, double height)
        { X = x; Y = y; Width = width; Height = height; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        public override string ToString() => string.Join(" ", new[] { X, Y, Width, Height }
            .Select(value => value.ToString("G17", CultureInfo.InvariantCulture)));
    }
}
