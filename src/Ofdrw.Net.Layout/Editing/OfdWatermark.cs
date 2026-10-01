using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Placement in millimeters. Existing layer IDs append within that layer; new layers are placed at the front (Background) or back.</summary>
public sealed class OfdWatermarkOptions
{
    /// <summary>Destination layer identifier; null creates a unique layer.</summary>
    public string? LayerId { get; set; }
    /// <summary>OFD layer type: Background, Body or Foreground.</summary>
    public string LayerType { get; set; } = "Foreground";
    /// <summary>Horizontal position in page coordinates, in millimeters.</summary>
    public double XMillimeters { get; set; } = 20;
    /// <summary>Vertical position in page coordinates, in millimeters.</summary>
    public double YMillimeters { get; set; } = 30;
    /// <summary>Object width in millimeters.</summary>
    public double WidthMillimeters { get; set; } = 100;
    /// <summary>Object height in millimeters.</summary>
    public double HeightMillimeters { get; set; } = 20;
    /// <summary>Maximum number of destination pages.</summary>
    public int MaxPageCount { get; set; } = 10_000;
    /// <summary>Maximum text characters summed across generated watermark objects.</summary>
    public long MaxGeneratedTextCharacters { get; set; } = 1_000_000;
    /// <summary>Maximum encoded image bytes.</summary>
    public int MaxImageBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum sum of mutable image payload bytes across generated page objects.</summary>
    public long MaxGeneratedImageBytes { get; set; } = 512L * 1024 * 1024;
    /// <summary>Maximum decoded width times height, validated before copying image bytes.</summary>
    public long MaxDecodedImagePixels { get; set; } = 40_000_000;
}

/// <summary>Adds ordinary text/image objects, without signature declarations or cryptographic validity claims.</summary>
public static class OfdWatermark
{
    /// <summary>Adds one text object per selected zero-based page. All validation precedes mutation.</summary>
    public static void AddText(OfdDocumentPackage package, IReadOnlyList<int> pages, string text,
        OfdWatermarkOptions? options = null, string fontName = "SimSun", double fontSizeMillimeters = 8,
        OfdColor? color = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Watermark text is required.", nameof(text));
        if (string.IsNullOrWhiteSpace(fontName)) throw new ArgumentException("Font name is required.", nameof(fontName));
        if (!Finite(fontSizeMillimeters) || fontSizeMillimeters <= 0) throw new ArgumentOutOfRangeException(nameof(fontSizeMillimeters));
        options ??= new OfdWatermarkOptions();
        if (pages is null) throw new ArgumentNullException(nameof(pages));
        if (options.MaxGeneratedTextCharacters <= 0 || text.Length > options.MaxGeneratedTextCharacters / Math.Max(1, pages.Count))
            throw new ArgumentException("Watermark generated text budget exceeded.", nameof(text));
        Add(package, pages, options, () => new OfdTextElement
        {
            Text = text, FontName = fontName, FontSizeMillimeters = fontSizeMillimeters,
            FillColor = color ?? new OfdColor(160, 80, 80, 160)
        }, cancellationToken);
    }

    /// <summary>Adds one image object per selected zero-based page. Encoded bytes are copied; PDF export applies its own decoded pixel budget.</summary>
    public static void AddImage(OfdDocumentPackage package, IReadOnlyList<int> pages, byte[] data,
        string mediaType, OfdWatermarkOptions? options = null, int alpha = 160, CancellationToken cancellationToken = default)
    {
        options ??= new OfdWatermarkOptions();
        if (data is null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0 || data.Length > options.MaxImageBytes) throw new ArgumentException("Watermark image exceeds its encoded byte budget or is empty.", nameof(data));
        if (mediaType is not ("image/png" or "image/jpeg")) throw new ArgumentException("Watermark image must be PNG or JPEG.", nameof(mediaType));
        if (alpha < 0 || alpha > 255) throw new ArgumentOutOfRangeException(nameof(alpha));
        cancellationToken.ThrowIfCancellationRequested();
        WatermarkImageBudget.Validate(data, mediaType, options.MaxDecodedImagePixels);
        if (pages is null) throw new ArgumentNullException(nameof(pages));
        if (options.MaxGeneratedImageBytes <= 0 || data.Length > options.MaxGeneratedImageBytes / Math.Max(1, pages.Count))
            throw new ArgumentException("Generated watermark image byte budget exceeded.", nameof(data));
        var payload = data.ToArray();
        Add(package, pages, options, () => new OfdImageElement { Data = payload.ToArray(), MediaType = mediaType, Alpha = alpha }, cancellationToken);
    }

    private static void Add(OfdDocumentPackage package, IReadOnlyList<int> pages, OfdWatermarkOptions options,
        Func<OfdElement> create, CancellationToken token)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        OfdDocumentSplitter.ValidateSingleDocument(package);
        OfdDocumentSplitter.ValidatePages(package, pages);
        if (options.MaxPageCount <= 0 || pages.Count > options.MaxPageCount || options.MaxImageBytes <= 0) throw new ArgumentException("Watermark page/image budget exceeded.");
        if (options.LayerType is not ("Background" or "Body" or "Foreground")) throw new ArgumentException("Invalid OFD layer type.");
        if (!Finite(options.XMillimeters) || !Finite(options.YMillimeters) || !Finite(options.WidthMillimeters) || !Finite(options.HeightMillimeters) ||
            options.WidthMillimeters <= 0 || options.HeightMillimeters <= 0) throw new ArgumentException("Watermark geometry must be finite with positive dimensions.");
        var staged = new List<(OfdPage Page, OfdElement Element, int Position)>();
        var layer = options.LayerId ?? "watermark-" + Guid.NewGuid().ToString("N");
        foreach (var index in pages)
        {
            token.ThrowIfCancellationRequested();
            var page = package.Pages[index];
            var matching = page.Elements.Where(element => element.LayerId == layer).ToList();
            if (matching.Any(element => element.LayerType != options.LayerType)) throw new ArgumentException("Existing layer has a different type.");
            var position = matching.Count > 0 ? page.Elements.LastIndexOf(matching.Last()) + 1
                : options.LayerType == "Background" ? 0 : page.Elements.Count;
            var element = create();
            element.LayerId = layer; element.LayerType = options.LayerType;
            element.XMillimeters = options.XMillimeters; element.YMillimeters = options.YMillimeters;
            element.WidthMillimeters = options.WidthMillimeters; element.HeightMillimeters = options.HeightMillimeters;
            staged.Add((page, element, position));
        }
        token.ThrowIfCancellationRequested();
        foreach (var item in staged) item.Page.Elements.Insert(item.Position, item.Element);
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
