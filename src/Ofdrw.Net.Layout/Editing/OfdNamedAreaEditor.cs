using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Internal.Flow;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A persisted, fixed page rectangle reserved by an AreaHolderBlock.</summary>
public sealed class OfdNamedArea
{
    internal OfdNamedArea(OfdNamedAreaIndex.Entry entry)
    { Name = entry.Name; PageIndex = entry.Page.Index; PageId = entry.Page.Id!;
      XMillimeters = entry.X; YMillimeters = entry.Y;
      WidthMillimeters = entry.Width; HeightMillimeters = entry.Height; IsFilled = entry.Filled; }
    /// <summary>Gets the unique name assigned to the area in this document.</summary>
    public string Name { get; }
    /// <summary>Gets the zero-based page index containing the area.</summary>
    public int PageIndex { get; }
    /// <summary>Gets the persisted identifier of the containing page.</summary>
    public string PageId { get; }
    /// <summary>Gets the left edge of the reserved rectangle in millimeters.</summary>
    public double XMillimeters { get; }
    /// <summary>Gets the top edge of the reserved rectangle in millimeters.</summary>
    public double YMillimeters { get; }
    /// <summary>Gets the width of the reserved rectangle in millimeters.</summary>
    public double WidthMillimeters { get; }
    /// <summary>Gets the height of the reserved rectangle in millimeters.</summary>
    public double HeightMillimeters { get; }
    /// <summary>Gets whether this area has already been filled.</summary>
    public bool IsFilled { get; }
}

/// <summary>Text settings for the explicitly limited legacy Flow measuring path.</summary>
public sealed class NamedAreaTextOptions
{
    /// <summary>Gets or sets the font family used for the fill text.</summary>
    public string FontFamily { get; set; } = "SimSun";
    /// <summary>Gets or sets the font size in millimeters.</summary>
    public double FontSizeMillimeters { get; set; } = 3.704166666666667;
    /// <summary>Gets or sets whether the fill text uses bold weight.</summary>
    public bool Bold { get; set; }
    /// <summary>Gets or sets whether the fill text uses italic style.</summary>
    public bool Italic { get; set; }
    /// <summary>Gets or sets the fill text color.</summary>
    public OfdColor Color { get; set; } = OfdColor.Black;
    /// <summary>Maximum UTF-16 units accepted for one area.</summary>
    public int MaxCharacters { get; set; } = 10000;
}

/// <summary>Queries and fills SDK-created named areas without repaginating other content.</summary>
public sealed class OfdNamedAreaEditor
{
    private sealed class LegacyMetrics : IFlowFontMetrics
    {
        public double AdvanceMillimeters(string grapheme, FlowTextStyle style) =>
            FlowLatinMetrics.AdvanceMillimeters(grapheme, style);
    }

    /// <summary>Reads and validates the persisted index. Documents without an index return an empty list.</summary>
    public IReadOnlyList<OfdNamedArea> GetAreas(OfdDocumentPackage package) =>
        OfdNamedAreaIndex.Load(package).Select(entry => new OfdNamedArea(entry)).ToArray();

    /// <summary>Fills a blank area with wrapped native text using legacy Flow approximate metrics.</summary>
    /// <remarks>The metric is not a font-face or complex-script shaping guarantee. Overflow rejects the fill before mutation.</remarks>
    public void FillText(OfdDocumentPackage package, string name, string text,
        NamedAreaTextOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (text.Length == 0) throw new ArgumentException("Fill text cannot be empty.", nameof(text));
        if (text.IndexOf('\f') >= 0) throw new NotSupportedException("A named area cannot contain a page break.");
        var (areas, area) = Find(package, name);
        options ??= new NamedAreaTextOptions();
        if (string.IsNullOrWhiteSpace(options.FontFamily) || options.FontFamily.Length > 256 ||
            !FinitePositive(options.FontSizeMillimeters) || options.Color is null ||
            options.MaxCharacters < 1 || options.MaxCharacters > 10000 || text.Length > options.MaxCharacters)
            throw new ArgumentException("Invalid named area text options or character limit.", nameof(options));
        var style = new FlowTextStyle { FontFamily = options.FontFamily.Trim(),
            FontSizeMillimeters = options.FontSizeMillimeters, Bold = options.Bold,
            Italic = options.Italic, Source = options.Color };
        var lines = FlowParagraphLayout.Layout(
            new[] { new FlowInline { Text = text, Style = style } }, new FlowParagraphFormat(),
            area.Width, new LegacyMetrics(), style, cancellationToken);
        var created = new List<OfdElement>();
        var top = area.Y;
        foreach (var line in lines)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (line.PageBreak) throw new NotSupportedException("A named area cannot contain a page break.");
            if (top + line.Height > area.Y + area.Height + 0.000001d)
                throw new InvalidOperationException("Named area text exceeds its Boundary.");
            if (line.Glyphs.Count > 0)
            {
                var width = line.Glyphs.Sum(g => g.Width);
                if (width > area.Width + 0.000001d) throw new InvalidOperationException("Named area text exceeds its Boundary.");
                var value = string.Concat(line.Glyphs.Select(g => g.Text));
                var element = new OfdTextElement { LayerType = "Body",
                    XMillimeters = area.X, YMillimeters = top,
                    WidthMillimeters = Math.Min(area.Width, Math.Max(0.1, width)), HeightMillimeters = line.Height,
                    FontName = style.FontFamily, FontSizeMillimeters = style.FontSizeMillimeters,
                    Weight = style.Bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                    Italic = style.Italic, FillColor = options.Color, Text = value };
                element.Runs.Add(new OfdTextRun { Text = value, YMillimeters = line.BaselineMillimeters,
                    DeltaX = line.Glyphs.Count > 1 ? string.Join(" ", line.Glyphs.Take(line.Glyphs.Count - 1)
                        .Select(g => g.Width.ToString("0.######", CultureInfo.InvariantCulture))) : null });
                created.Add(element);
            }
            top += line.Height;
        }
        if (created.Count == 0) throw new ArgumentException("Fill text has no visible content.", nameof(text));
        // BuildFonts resolves each text element through the package font catalog.
        // Declare a matching face once so wrapped lines share one resource.
        if (!package.Fonts.Any(font => string.Equals(font.FontName, style.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            font.Bold == style.Bold && font.Italic == style.Italic))
            package.Fonts.Add(new OfdFontResource { FontName = style.FontFamily, Bold = style.Bold, Italic = style.Italic });
        area.Filled = true;
        OfdNamedAreaIndex.Save(package, areas);
        area.Page.Elements.AddRange(created);
    }

    /// <summary>Fits a PNG or JPEG inside a blank area, retaining the declared source aspect ratio.</summary>
    /// <remarks>The caller supplies trustworthy source pixel dimensions; this method checks the PNG header but does not decode the image.</remarks>
    public void FillImage(OfdDocumentPackage package, string name, byte[] encodedImage, string mediaType,
        double sourceWidthPixels, double sourceHeightPixels, CancellationToken cancellationToken = default)
    {
        var (areas, area) = Find(package, name);
        if (encodedImage is null) throw new ArgumentNullException(nameof(encodedImage));
        if (encodedImage.Length == 0 || encodedImage.Length > 16 * 1024 * 1024)
            throw new InvalidDataException("Named area image size limit exceeded.");
        if (!FinitePositive(sourceWidthPixels) || !FinitePositive(sourceHeightPixels) ||
            sourceWidthPixels > 100000 || sourceHeightPixels > 100000)
            throw new ArgumentOutOfRangeException(nameof(sourceWidthPixels));
        ImageFormatAdmission.RequirePngOrJpeg(encodedImage, encodedImage.Length);
        var png = encodedImage[0] == 137;
        var normalized = mediaType?.Trim().ToLowerInvariant();
        if (png ? normalized != "image/png" : normalized != "image/jpeg" && normalized != "image/jpg")
            throw new InvalidDataException("Named area image media type does not match its signature.");
        if (sourceWidthPixels * sourceHeightPixels > 50000000)
            throw new InvalidDataException("Named area image pixel limit exceeded.");
        if (png)
        {
            if (encodedImage.Length < 33 || encodedImage[12] != (byte)'I' || encodedImage[13] != (byte)'H' ||
                encodedImage[14] != (byte)'D' || encodedImage[15] != (byte)'R')
                throw new InvalidDataException("Named area PNG header is incomplete.");
            var width = ReadBigEndian(encodedImage, 16); var height = ReadBigEndian(encodedImage, 20);
            if (width != sourceWidthPixels || height != sourceHeightPixels)
                throw new InvalidDataException("Declared PNG dimensions do not match the encoded header.");
        }
        else if (TryJpegDimensions(encodedImage, out var width, out var height))
        {
            if (width != sourceWidthPixels || height != sourceHeightPixels)
                throw new InvalidDataException("Declared JPEG dimensions do not match the encoded header.");
        }
        else throw new InvalidDataException("Named area JPEG frame dimensions are unavailable.");
        cancellationToken.ThrowIfCancellationRequested();
        var scale = Math.Min(area.Width / sourceWidthPixels, area.Height / sourceHeightPixels);
        var widthMm = sourceWidthPixels * scale; var heightMm = sourceHeightPixels * scale;
        if (!FinitePositive(widthMm) || !FinitePositive(heightMm) || widthMm < 0.001 || heightMm < 0.001)
            throw new InvalidOperationException("Named area image cannot fit.");
        var element = new OfdImageElement { LayerType = "Body", Data = encodedImage.ToArray(),
            MediaType = png ? "image/png" : "image/jpeg",
            XMillimeters = area.X + (area.Width - widthMm) / 2,
            YMillimeters = area.Y + (area.Height - heightMm) / 2,
            WidthMillimeters = widthMm, HeightMillimeters = heightMm };
        area.Filled = true;
        OfdNamedAreaIndex.Save(package, areas);
        area.Page.Elements.Add(element);
    }

    private static (List<OfdNamedAreaIndex.Entry> Areas, OfdNamedAreaIndex.Entry Area) Find(OfdDocumentPackage package, string name)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        OfdNamedAreaIndex.ValidateName(name);
        if (package.PreservedDocBodyElements.Any(xml => XElement.Parse(xml).Name.LocalName == "Signatures"))
            throw new NotSupportedException("Named area fill cannot rewrite a signed document.");
        var areas = OfdNamedAreaIndex.Load(package);
        var area = areas.SingleOrDefault(item => item.Name == name) ??
            throw new KeyNotFoundException("Named area was not found.");
        if (area.Filled) throw new InvalidOperationException("Named area has already been filled.");
        return (areas, area);
    }

    private static bool FinitePositive(double value) => !double.IsNaN(value) && !double.IsInfinity(value) && value > 0;
    private static uint ReadBigEndian(byte[] data, int offset) =>
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];

    private static bool TryJpegDimensions(byte[] data, out int width, out int height)
    {
        width = height = 0;
        for (var offset = 2; offset + 4 <= data.Length;)
        {
            if (data[offset++] != 0xff) return false;
            while (offset < data.Length && data[offset] == 0xff) offset++;
            if (offset >= data.Length) return false;
            var marker = data[offset++];
            if (marker is 0xd8 or 0xd9 or 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
            if (offset + 2 > data.Length) return false;
            var length = data[offset] * 256 + data[offset + 1];
            if (length < 2 || offset + length > data.Length) return false;
            if (marker is >= 0xc0 and <= 0xcf && marker is not (0xc4 or 0xc8 or 0xcc))
            {
                if (length < 7) return false;
                height = data[offset + 3] * 256 + data[offset + 4];
                width = data[offset + 5] * 256 + data[offset + 6];
                return width > 0 && height > 0;
            }
            if (marker == 0xda) return false;
            offset += length;
        }
        return false;
    }
}
