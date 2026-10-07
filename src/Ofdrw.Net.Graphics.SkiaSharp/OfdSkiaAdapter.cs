using System;
using System.Collections.Generic;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;

namespace Ofdrw.Net.Graphics.SkiaSharp;

/// <summary>Limits for a producer-supplied event batch, copied at entry. All graphics limits retain 04 semantics.</summary>
public sealed class OfdSkiaAdapterOptions
{
    /// <summary>Millimeters per producer unit. Positive; no implicit pixel/DPI assumption. Default is one millimeter per unit.</summary>
    public double MillimetersPerUnit { get; set; } = 1;
    /// <summary>Maximum events enumerated in one Append batch.</summary>
    public int MaxEvents { get; set; } = 10_000;
    /// <summary>Maximum cumulative snapshot path commands in one batch.</summary>
    public long MaxInputPathCommands { get; set; } = 1_000_000;
    /// <summary>Maximum bytes in one original font payload, before hashing.</summary>
    public int MaxFontBytes { get; set; } = 64_000_000;
    /// <summary>04 output, text, path and geometry limits; original target elements count toward MaxPageElements.</summary>
    public OfdGraphicsOptions Graphics { get; set; } = new();
}

/// <summary>Maps explicit cooperating producer events to native OFD through 04. Does not capture SKCanvas, SKPicture or PDF renderer calls.</summary>
public static class OfdSkiaAdapter
{
    /// <summary>Atomically appends a bounded batch. Any ordinary validation, enumeration, budget or cancellation failure leaves the target and resources unchanged.</summary>
    /// <remarks>No native input objects are retained. Package, page and resources must not be changed concurrently. Fonts must have exactly the original face bytes recorded by Text events.</remarks>
    public static void Append(OfdDocumentPackage package, OfdPage page, IEnumerable<SkiaDrawEvent> events,
        OfdSkiaAdapterOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (page is null) throw new ArgumentNullException(nameof(page));
        if (events is null) throw new ArgumentNullException(nameof(events));
        if (!package.Pages.Contains(page)) throw new ArgumentException("Page must belong to destination package.", nameof(page));
        options ??= new OfdSkiaAdapterOptions();
        var units = options.MillimetersPerUnit; var maxEvents = options.MaxEvents; var maxCommands = options.MaxInputPathCommands; var maxFontBytes = options.MaxFontBytes;
        if (double.IsNaN(units) || double.IsInfinity(units) || units <= 0 || maxEvents <= 0 || maxCommands <= 0 || maxFontBytes <= 0 || options.Graphics is null)
            throw new ArgumentException("Adapter limits and unit scale must be positive and finite.", nameof(options));
        var staged = new OfdDocumentPackage { Options = package.Options };
        var target = new OfdPage { Index = page.Index, XMillimeters = page.XMillimeters, YMillimeters = page.YMillimeters,
            WidthMillimeters = page.WidthMillimeters, HeightMillimeters = page.HeightMillimeters };
        // Count existing elements without reading or modifying their content.
        var originalCount = page.Elements.Count;
        var maxPageElements = options.Graphics.MaxPageElements;
        if (originalCount > maxPageElements) throw new InvalidOperationException("Target already exceeds graphics page element budget.");
        staged.Pages.Add(target); staged.Fonts.AddRange(package.Fonts);
        var limits = new OfdGraphicsOptions { MaxPageElements = maxPageElements - originalCount,
            MaxPathCommands = options.Graphics.MaxPathCommands, MaxTextCharacters = options.Graphics.MaxTextCharacters,
            MaxGeometryCharacters = options.Graphics.MaxGeometryCharacters, MaxSavedStates = options.Graphics.MaxSavedStates, MaxClipRegions = options.Graphics.MaxClipRegions };
        // A full page may still accept an empty batch, but cannot receive another event.
        if (limits.MaxPageElements == 0) limits.MaxPageElements = 1;
        var graphics = new OfdGraphics(staged, target, limits);
        var count = 0; long commands = 0, characters = 0;
        var hashes = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var draw in events)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (draw is null) throw new ArgumentException("Events cannot contain null entries.", nameof(events));
            if (++count > maxEvents || count > maxPageElements - originalCount) throw new InvalidOperationException("Event/page budget exceeded.");
            if (draw.PathCommands > maxCommands - commands || draw.TextCharacters > limits.MaxTextCharacters - characters)
                throw new InvalidOperationException("Cumulative input snapshot budget exceeded.");
            commands += draw.PathCommands; characters += draw.TextCharacters;
            draw.Emit(graphics, staged, units, hashes, maxFontBytes, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        page.Elements.AddRange(target.Elements);
    }
}
