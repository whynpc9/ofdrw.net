using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Private, versioned metadata for SDK-created flow areas.</summary>
internal static class OfdNamedAreaIndex
{
    internal const string Key = "ofdrw.net.named-areas.v1";
    internal const int MaxAreas = 10000;
    internal const int MaxXmlLength = 1000000;

    internal sealed class Entry
    {
        internal Entry(string name, OfdPage page, double y, double x, double width, double height)
        { Name = name; Page = page; Y = y; X = x; Width = width; Height = height; }
        internal string Name { get; }
        internal OfdPage Page { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal bool Filled { get; set; }
    }

    internal static void ValidateName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 128 || name.Any(char.IsControl))
            throw new ArgumentException("Named area name must contain 1–128 non-control characters.", nameof(name));
    }

    internal static List<Entry> Load(OfdDocumentPackage package)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (!package.CustomTags.TryGetValue(Key, out var value)) return new List<Entry>();
        if (value.Length == 0 || value.Length > MaxXmlLength) throw new InvalidDataException("Named area index size is invalid.");
        XDocument xml;
        using (var source = new StringReader(value))
        using (var reader = XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
            xml = XDocument.Load(reader);
        var root = xml.Root;
        if (root is null || root.Name != "NamedAreas" || root.Attribute("Version")?.Value != "1" ||
            root.Attributes().Any(a => a.Name != "Version") || root.Nodes().Any(node => node is not XElement))
            throw new InvalidDataException("Unsupported named area index format.");
        var result = new List<Entry>();
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in root.Elements())
        {
            if (node.Name != "Area" || node.HasElements || node.Nodes().Any() ||
                node.Attributes().Any(a => a.Name.LocalName is not ("Name" or "PageIndex" or "PageID" or "Boundary" or "Filled")))
                throw new InvalidDataException("Malformed named area entry.");
            if (result.Count >= MaxAreas) throw new InvalidDataException("Named area limit exceeded.");
            var name = node.Attribute("Name")?.Value;
            try { ValidateName(name); }
            catch (ArgumentException ex) { throw new InvalidDataException("Invalid named area name.", ex); }
            if (!names.Add(name!)) throw new InvalidDataException("Duplicate named area name.");
            if (!int.TryParse(node.Attribute("PageIndex")?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index) ||
                index < 0 || index >= package.Pages.Count) throw new InvalidDataException("Named area page reference is missing.");
            var page = package.Pages[index];
            if (string.IsNullOrEmpty(page.Id) || page.Id != node.Attribute("PageID")?.Value || page.Index != index)
                throw new InvalidDataException("Named area page reference changed.");
            var box = node.Attribute("Boundary")?.Value?.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (box is not { Length: 4 } ||
                !Try(box[0], out var x) || !Try(box[1], out var y) ||
                !Try(box[2], out var width) || !Try(box[3], out var height) ||
                width <= 0 || height <= 0 || x < page.XMillimeters - 0.000001 ||
                y < page.YMillimeters - 0.000001 ||
                x + width > page.XMillimeters + page.WidthMillimeters + 0.000001 ||
                y + height > page.YMillimeters + page.HeightMillimeters + 0.000001)
                throw new InvalidDataException("Named area Boundary is invalid or outside its page.");
            var filled = node.Attribute("Filled")?.Value switch
            {
                "true" => true,
                "false" => false,
                _ => throw new InvalidDataException("Named area fill state is invalid.")
            };
            result.Add(new Entry(name!, page, y, x, width, height) { Filled = filled });
        }
        return result;
    }

    internal static void Save(OfdDocumentPackage package, IReadOnlyList<Entry> entries)
    {
        if (entries.Count > MaxAreas) throw new InvalidOperationException("Named area limit exceeded.");
        var nodes = new List<XElement>(entries.Count);
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var entry in entries)
        {
            ValidateName(entry.Name);
            if (!names.Add(entry.Name)) throw new InvalidOperationException("Duplicate named area name.");
            var index = package.Pages.IndexOf(entry.Page);
            if (index < 0 || string.IsNullOrEmpty(entry.Page.Id)) throw new InvalidOperationException("Named area page is unavailable.");
            nodes.Add(new XElement("Area", new XAttribute("Name", entry.Name),
                new XAttribute("PageIndex", index), new XAttribute("PageID", entry.Page.Id),
                new XAttribute("Boundary", string.Join(" ", new[] { entry.X, entry.Y, entry.Width, entry.Height }
                    .Select(v => v.ToString("R", CultureInfo.InvariantCulture)))),
                new XAttribute("Filled", entry.Filled ? "true" : "false")));
        }
        var value = new XElement("NamedAreas", new XAttribute("Version", "1"), nodes).ToString(SaveOptions.DisableFormatting);
        if (value.Length > MaxXmlLength) throw new InvalidOperationException("Named area index size limit exceeded.");
        package.CustomTags[Key] = value;
    }

    private static bool Try(string value, out double number) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
        !double.IsNaN(number) && !double.IsInfinity(number);
}
