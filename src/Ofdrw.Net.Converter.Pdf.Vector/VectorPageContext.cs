using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Layout.Graphics;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Filters;
using UglyToad.PdfPig.Geometry;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Outline.Destinations;
using UglyToad.PdfPig.Parser;
using UglyToad.PdfPig.Parser.Parts;
using UglyToad.PdfPig.Tokenization.Scanner;
using UglyToad.PdfPig.Tokens;
using SkiaSharp;

namespace Ofdrw.Net.Converter.Pdf.Vector;

// PdfPig injects these public dependencies without decoding content or loading fonts.
internal sealed class VectorContextFactory : IPageFactory<VectorPageContext>
{
    private readonly IPdfTokenScanner _scanner; private readonly IResourceStore _resources;
    private readonly ILookupFilterProvider _filters; private readonly ParsingOptions _parsing;
    public VectorContextFactory(IPdfTokenScanner scanner, IResourceStore resources, ILookupFilterProvider filters, IPageContentParser parser, ParsingOptions parsing)
    { _scanner = scanner; _resources = resources; _filters = filters; _parsing = parsing; }
    public VectorPageContext Create(int number, DictionaryToken dictionary, PageTreeMembers members, NamedDestinations names)
        => new(number, dictionary, _scanner, _resources, _filters, _parsing);
}

internal sealed class VectorPageContext
{
    private readonly int _number; private readonly DictionaryToken _page;
    internal readonly IPdfTokenScanner Scanner; internal readonly IResourceStore Resources;
    internal readonly ILookupFilterProvider Filters; internal readonly ParsingOptions Parsing;
    internal VectorPageContext(int number, DictionaryToken page, IPdfTokenScanner scanner, IResourceStore resources, ILookupFilterProvider filters, ParsingOptions parsing)
    { _number = number; _page = page; Scanner = scanner; Resources = resources; Filters = filters; Parsing = parsing; }
    internal OfdDocumentPackage Convert(PdfVectorToOfdOptions options, CancellationToken token, long remainingImageBytes = long.MaxValue)
    {
        var frame = ReadFrame(options, token);
        var pageFrame = frame.Frame; var resource = frame.Resources;
        var lexical = new CoreTokenScanner(new MemoryInputBytes(frame.Content), false, new StackDepthGuard(options.MaxStackDepth), useLenientParsing: false);
        var pendingOperands = 0;
        while (lexical.MoveNext())
        {
            token.ThrowIfCancellationRequested();
            if (lexical.CurrentToken is CommentToken) continue;
            pendingOperands = lexical.CurrentToken is OperatorToken ? 0 : pendingOperands + 1;
        }
        if (pendingOperands != 0) throw Unsupported("TRAILING_OPERANDS");
        var factory = new VectorOperations(options, token, _number);
        var parser = new PageContentParser(factory, new StackDepthGuard(options.MaxStackDepth), false);
        var operations = parser.Parse(_number, new MemoryInputBytes(frame.Content), Parsing.Logger);
        return ConvertEvents(options, token, pageFrame, resource, parser, operations, remainingImageBytes);
    }
    private (VectorPageFrame Frame, DictionaryToken Resources, byte[] Content, bool SingleStream) ReadFrame(PdfVectorToOfdOptions options, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var ancestry = new List<DictionaryToken>(); var node = _page;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            if (ancestry.Count >= options.MaxStackDepth) throw new InvalidDataException("PDF resource ancestry depth exceeded.");
            ancestry.Add(node);
            if (!node.TryGet(NameToken.Parent, out var parent)) break;
            node = Resolve<DictionaryToken>(parent);
        }
        IToken? Inherited(string key) => ancestry.Select(value => value.Data.TryGetValue(key, out var entry) ? entry : null).FirstOrDefault(value => value is not null);
        var media = Resolve<ArrayToken>(Inherited("MediaBox"));
        var numbers = media.Data.Select(item => Resolve<NumericToken>(item).Double).ToArray();
        if (numbers.Length != 4 || numbers.Any(value => double.IsNaN(value) || double.IsInfinity(value)) ||
            !FinitePositive(numbers[2] - numbers[0]) || !FinitePositive(numbers[3] - numbers[1]))
            throw Unsupported("PAGE_BOX: requires a finite positive MediaBox.");
        var cropToken = Inherited("CropBox");
        var crop = cropToken is null ? numbers : Resolve<ArrayToken>(cropToken).Data.Select(item => Resolve<NumericToken>(item).Double).ToArray();
        if (crop.Length != 4 || crop.Any(value => double.IsNaN(value) || double.IsInfinity(value)) ||
            !FinitePositive(crop[2] - crop[0]) || !FinitePositive(crop[3] - crop[1]) ||
            crop[0] < numbers[0] || crop[1] < numbers[1] || crop[2] > numbers[2] || crop[3] > numbers[3])
            throw Unsupported("CROP_BOX: requires a finite positive box within MediaBox.");
        var rotate = Inherited("Rotate");
        var angle = rotate is null ? 0 : Resolve<NumericToken>(rotate).Double;
        if (double.IsNaN(angle) || double.IsInfinity(angle) || angle != Math.Truncate(angle) || angle % 90 != 0)
            throw Unsupported("PAGE_ROTATION: only integral quarter turns are supported.");
        var rotation = (int)((angle % 360 + 360) % 360);
        var frame = new VectorPageFrame(numbers, crop, rotation);
        if (_page.Data.TryGetValue("UserUnit", out var units) && Resolve<NumericToken>(units).Double != 1) throw Unsupported("USER_UNIT");
        if (_page.ContainsKey(NameToken.Create("Annots")) || _page.ContainsKey(NameToken.Create("Group"))) throw Unsupported("ANNOTATIONS_OR_GROUP");
        var resourceToken = Inherited("Resources");
        var resource = resourceToken is null ? DictionaryToken.Empty : Resolve<DictionaryToken>(resourceToken);
        if (resource.ContainsKey(NameToken.ColorSpace)) throw Unsupported("RESOURCE_COLOR_SPACE: device substitutions not supported.");
        using var content = new MemoryStream(); var singleStream = false;
        if (_page.TryGet(NameToken.Contents, out var contents))
        {
            var directContents = ProfileToken(contents, options.MaxStackDepth, token);
            singleStream = directContents is not ArrayToken;
            var streams = directContents is ArrayToken array ? array.Data : new[] { directContents };
            foreach (var streamToken in streams)
            {
                token.ThrowIfCancellationRequested();
                var bytes = Unfiltered(Resolve<StreamToken>(streamToken), options.MaxContentBytesPerPage - (int)content.Length, "content");
                if (bytes.Length >= options.MaxContentBytesPerPage - content.Length) throw new InvalidDataException("PDF page content byte limit exceeded.");
                content.Write(bytes, 0, bytes.Length); content.WriteByte(10);
            }
        }
        return (frame, resource, content.ToArray(), singleStream);
    }
    private OfdDocumentPackage ConvertEvents(PdfVectorToOfdOptions options, CancellationToken token, VectorPageFrame frame,
        DictionaryToken resource, IPageContentParser parser, IReadOnlyList<IGraphicsStateOperation> operations, long remainingImageBytes)
    {
        // Declared-resource structure/count is a hard input gate, even when no glyph is shown.
        var fontTokens = resource.TryGet(NameToken.Font, out var fontsToken)
            ? FontObject<DictionaryToken>(fontsToken, options, token) : DictionaryToken.Empty;
        if (fontTokens.Data.Count > options.MaxFontsPerPage) throw new InvalidDataException("PDF page font count limit exceeded.");
        var bindings = new Dictionary<string, VectorFont>(StringComparer.Ordinal);
        long attemptedFontBytes = 0; var loadedFrames = 0;
        (IToken Token, DictionaryToken Font) HorizontalFont(string name)
        {
            token.ThrowIfCancellationRequested();
            if (!fontTokens.Data.TryGetValue(name, out var fontToken)) throw new InvalidDataException("PDF referenced font is missing: " + name);
            var font = FontObject<DictionaryToken>(fontToken, options, token);
            if (FontObject<NameToken>(FontEntry(font, "Subtype"), options, token).Data != "Type0") throw Unsupported("FONT_PROFILE: " + name);
            var encoding = ProfileToken(FontEntry(font, "Encoding"), options.MaxStackDepth, token);
            if (encoding is StreamToken) throw Unsupported("FONT_PROFILE: Encoding CMap stream outside Identity-H profile.");
            if (encoding is not NameToken encodingName) throw new InvalidDataException("PDF font Encoding must be a name or CMap stream.");
            if (encodingName.Data != "Identity-H") throw Unsupported("FONT_PROFILE: " + name);
            return (fontToken, font);
        }
        void EnsureFont(string name)
        {
            token.ThrowIfCancellationRequested();
            if (bindings.ContainsKey(name)) return;
            var selected = HorizontalFont(name); var font = selected.Font;
            var descendants = FontObject<ArrayToken>(FontEntry(font, "DescendantFonts"), options, token);
            if (descendants.Length != 1) throw Unsupported("FONT_DESCENDANTS: " + name);
            var cid = FontObject<DictionaryToken>(descendants.Data[0], options, token);
            if (FontObject<NameToken>(FontEntry(cid, "Subtype"), options, token).Data != "CIDFontType2") throw Unsupported("CID_MAPPING: " + name);
            byte[]? cidToGid = null;
            if (cid.Data.TryGetValue("CIDToGIDMap", out var mappingToken))
            {
                var mapping = ProfileToken(mappingToken, options.MaxStackDepth, token);
                if (mapping is StreamToken mapStream)
                {
                    cidToGid = Unfiltered(mapStream, 131072, "CIDToGIDMap");
                    if (cidToGid.Length == 0 || (cidToGid.Length & 1) != 0)
                        throw new InvalidDataException("PDF CIDToGIDMap must contain big-endian 16-bit glyph IDs.");
                }
                else if (mapping is not NameToken mappingName) throw new InvalidDataException("PDF CIDToGIDMap must be a name or stream.");
                else if (mappingName.Data != "Identity") throw Unsupported("CID_MAPPING: " + name);
            }
            _ = Unfiltered(FontObject<StreamToken>(FontEntry(font, "ToUnicode"), options, token), options.MaxContentBytesPerPage, "ToUnicode");
            var descriptor = FontObject<DictionaryToken>(FontEntry(cid, "FontDescriptor"), options, token);
            var allowance = (int)Math.Min(options.MaxFontBytes, options.MaxFontReadBytesPerPage - attemptedFontBytes);
            // Check the original stream before copying or invoking either font loader.
            var bytes = Unfiltered(FontObject<StreamToken>(FontEntry(descriptor, "FontFile2"), options, token), allowance, "font read");
            attemptedFontBytes += bytes.Length; // Attempts are never refunded on rejection or failure.
            token.ThrowIfCancellationRequested();
            bindings.Add(name, new VectorFont("font-" + bindings.Count, bytes, options.MaxFontBytes, cidToGid));
            // PdfPig pushes all four stacks before decoding this single font. Count
            // this frame before Load so the finally also balances a loader failure.
            loadedFrames++;
            Resources.LoadResourceDictionary(new DictionaryToken(new Dictionary<NameToken, IToken> {
                [NameToken.Font] = new DictionaryToken(new Dictionary<NameToken, IToken> { [NameToken.Create(name)] = selected.Token }) }));
        }
        try
        {
            var package = new OfdDocumentPackage();
            // The staged producer fragments and the final package must use the
            // same namespace; otherwise a newly attached clip can be invisible
            // under a TextObject when the caller selects the /2016 default.
            package.Options.Namespace = options.Compatibility.Namespace;
            var page = new OfdPage { WidthMillimeters = frame.Width * 25.4 / 72, HeightMillimeters = frame.Height * 25.4 / 72 };
            package.Pages.Add(page);
            using var processor = new VectorStreamProcessor(this, parser, frame, resource, bindings, options, token, _number,
                remainingImageBytes, EnsureFont, name => { _ = HorizontalFont(name); });
            var events = processor.Process(_number, operations);
            if (processor.PaintCount == 0) throw Unsupported("NO_NATIVE_CONTENT: blank, scan or nonpainted page.");
            // Keep the established exact original-image branch for image-only pages.
            if (events.Count == 0) throw Unsupported("IMAGE_ONLY: use the established original-image or raster page profile.");
            foreach (var binding in bindings.Values)
            {
                if (processor.OutputFonts.Contains(binding)) package.Fonts.Add(binding.Resource);
                else binding.Dispose();
            }
            // Keep parsing, consumed budgets and earlier page profiles ahead of
            // this output-format limit. The existing writer also rejects raster pages of this size.
            token.ThrowIfCancellationRequested();
            if (!OfdNumericFormat.WriterPositive(page.WidthMillimeters) || !OfdNumericFormat.WriterPositive(page.HeightMillimeters))
                throw new UnrepresentablePageException(_number);
            OfdSkiaAdapter.Append(package, page, events, new OfdSkiaAdapterOptions { MillimetersPerUnit = 25.4 / 72,
                MaxEvents = options.MaxEventsPerPage, MaxInputPathCommands = options.MaxPathCommandsPerPage, MaxFontBytes = options.MaxFontBytes,
                Graphics = new OfdGraphicsOptions { MaxPageElements = options.MaxEventsPerPage, MaxPathCommands = options.MaxPathCommandsPerPage,
                    MaxTextCharacters = options.MaxTextCharactersPerPage } }, token);
            processor.ApplyAnnotations(page, package.Options.Namespace);
            return package;
        }
        finally
        {
            try { while (loadedFrames-- > 0) Resources.UnloadResourceDictionary(); }
            finally { foreach (var binding in bindings.Values) binding.Dispose(); }
        }
    }
    private IToken FontEntry(DictionaryToken value, string key) => value.Data.TryGetValue(key, out var entry)
        ? entry : throw new InvalidDataException("PDF font resource entry is missing: " + key);
    private T FontObject<T>(IToken value, PdfVectorToOfdOptions options, CancellationToken token) where T : class, IToken
        => ProfileToken(value, options.MaxStackDepth, token) as T ?? throw new InvalidDataException("PDF font resource requires " + typeof(T).Name + ".");
    internal T Resolve<T>(IToken? value) where T : class, IToken => DirectObjectFinder.Get<T>(value!, Scanner) ?? throw Unsupported("INVALID_RESOURCE: " + typeof(T).Name);

    internal IToken RequireImageXObject(DictionaryToken resource, NameToken name)
    {
        if (!resource.Data.TryGetValue("XObject", out var all)) throw new InvalidDataException("PDF XObject resource dictionary is missing.");
        var objects = Resolve<DictionaryToken>(all);
        if (!objects.Data.TryGetValue(name, out var value)) throw new InvalidDataException("PDF XObject resource is missing: " + name.Data);
        var stream = Resolve<StreamToken>(value);
        if (!stream.StreamDictionary.Data.TryGetValue("Subtype", out var subtype)) throw new InvalidDataException("PDF XObject subtype is missing.");
        if (Resolve<NameToken>(subtype).Data != "Image") throw Unsupported("XOBJECT_PROFILE: only bounded Image XObjects are supported.");
        return value;
    }

    internal OfdImageElement DecodeImageXObject(StreamToken stream, PdfVectorToOfdOptions options, long remainingImageBytes, CancellationToken token)
    {
        var dictionary = stream.StreamDictionary;
        var allowed = new HashSet<string>(new[] { "Type", "Subtype", "Width", "Height", "ColorSpace", "BitsPerComponent", "Length", "Interpolate" });
        var colorSpace = ProfileToken(Entry(dictionary, "ColorSpace"), options.MaxStackDepth, token);
        if (dictionary.Data.Keys.Any(key => !allowed.Contains(key)) || Name(dictionary, "Subtype") != "Image" ||
            dictionary.Data.ContainsKey("Type") && Name(dictionary, "Type") != "XObject" ||
            colorSpace is not NameToken { Data: "DeviceRGB" } || Resolve<NumericToken>(Entry(dictionary, "BitsPerComponent")).Double != 8)
            throw Unsupported("IMAGE_PROFILE: raw opaque DeviceRGB/8 XObject required.");
        int Dimension(string key)
        {
            var value = Resolve<NumericToken>(Entry(dictionary, key)).Double;
            if (!FinitePositive(value) || value > int.MaxValue || value != Math.Truncate(value)) throw Unsupported("IMAGE_DIMENSION");
            return (int)value;
        }
        var width = Dimension("Width"); var height = Dimension("Height");
        var pixels = checked((long)width * height); var rawLength = checked(pixels * 3);
        if (pixels > options.Compatibility.MaxRasterizedPixelsPerPage || rawLength > int.MaxValue || rawLength > remainingImageBytes)
            throw new InvalidDataException("PDF native image pixel/byte limit exceeded.");
        if (stream.Data.Length != rawLength || Resolve<NumericToken>(Entry(dictionary, "Length")).Double != rawLength)
            throw new InvalidDataException("PDF raw image byte length mismatch.");
        var interpolate = false;
        if (dictionary.Data.TryGetValue("Interpolate", out var hint))
        {
            if (ProfileToken(hint, options.MaxStackDepth, token) is not BooleanToken boolean)
                throw Unsupported("IMAGE_PROFILE: Interpolate must be boolean.");
            interpolate = boolean.Data;
        }
        token.ThrowIfCancellationRequested();
        using var bitmap = new Image<Rgb24>(width, height);
        bitmap.ProcessPixelRows(accessor =>
        {
            for (var row = 0; row < height; row++)
            {
                token.ThrowIfCancellationRequested();
                var original = stream.Data.Span.Slice(checked(row * width * 3), checked(width * 3));
                var destination = accessor.GetRowSpan(row);
                for (var column = 0; column < width; column++)
                    destination[column] = new Rgb24(original[column * 3], original[column * 3 + 1], original[column * 3 + 2]);
            }
        });
        using var encoded = new BoundedImageOutput(remainingImageBytes, token);
        bitmap.SaveAsPng(encoded); token.ThrowIfCancellationRequested();
        return new OfdImageElement { Data = encoded.ToArray(), MediaType = "image/png", FileName = "pdf-vector-image.png",
            SourceXml = new XElement(XName.Get("ImageObject", options.Compatibility.Namespace),
                new XAttribute(OfdImageRenderingHints.PdfInterpolateV1, interpolate)).ToString(SaveOptions.DisableFormatting) };
    }
    internal OfdPage? TryOriginalImagePage(PdfVectorToOfdOptions options, long remainingImageBytes, CancellationToken token)
    {
        var encoding = false;
        try
        {
            var allowedPage = new HashSet<string>(new[] { "Type", "Parent", "MediaBox", "CropBox", "Rotate", "UserUnit", "Resources", "Contents" });
            if (_page.Data.Keys.Any(key => !allowedPage.Contains(key))) return null;
            var frame = ReadFrame(options, token);
            if (!frame.Frame.IsLegacyFrame) return null;
            if (!frame.SingleStream) return null;
            // Stock PageContentParser silently drops trailing operands. Consume the entire token stream.
            var tokens = new List<IToken>();
            var scanner = new CoreTokenScanner(new MemoryInputBytes(frame.Content), false, new StackDepthGuard(options.MaxStackDepth), useLenientParsing: false);
            while (scanner.MoveNext())
            {
                token.ThrowIfCancellationRequested();
                if (scanner.CurrentToken is CommentToken) continue;
                if (tokens.Count == 11) return null;
                tokens.Add(scanner.CurrentToken);
            }
            bool Op(int index, string value) => tokens[index] is OperatorToken operation && operation.Data == value;
            if (tokens.Count != 11 || !Op(0, "q") || !Op(7, "cm") || tokens[8] is not NameToken name || !Op(9, "Do") || !Op(10, "Q")) return null;
            var matrix = new double[6];
            for (var index = 0; index < matrix.Length; index++)
            { if (tokens[index + 1] is not NumericToken number) return null; matrix[index] = number.Double; }
            if (!matrix.SequenceEqual(new[] { frame.Frame.Width, 0d, 0d, frame.Frame.Height, 0d, 0d })) return null;
            if (options.MaxOperationsPerPage < 4) throw new InvalidDataException("PDF operation limit exceeded.");
            var xobjects = Resolve<DictionaryToken>(Entry(frame.Resources, "XObject"));
            var stream = Resolve<StreamToken>(Entry(xobjects, name.Data)); var dictionary = stream.StreamDictionary;
            var allowedImage = new HashSet<string>(new[] { "Type", "Subtype", "Width", "Height", "ColorSpace", "BitsPerComponent", "Length", "Interpolate" });
            if (dictionary.Data.Keys.Any(key => !allowedImage.Contains(key)) || Name(dictionary, "Subtype") != "Image" ||
                (dictionary.Data.ContainsKey("Type") && Name(dictionary, "Type") != "XObject") ||
                Resolve<NumericToken>(Entry(dictionary, "BitsPerComponent")).Double != 8) return null;
            var colorSpace = ProfileToken(Entry(dictionary, "ColorSpace"), options.MaxStackDepth, token);
            byte[]? palette = null; var highestIndex = 0;
            if (colorSpace is ArrayToken indexed)
            {
                if (indexed.Data.Count != 4 || ProfileToken(indexed.Data[0], options.MaxStackDepth, token) is not NameToken kind || kind.Data != "Indexed" ||
                    ProfileToken(indexed.Data[1], options.MaxStackDepth, token) is not NameToken basis || basis.Data != "DeviceRGB") return null;
                if (ProfileToken(indexed.Data[2], options.MaxStackDepth, token) is not NumericToken hival || hival.Double < 0 || hival.Double > 255 || Math.Truncate(hival.Double) != hival.Double)
                    throw new InvalidDataException("PDF Indexed hival must be an integer 0..255.");
                highestIndex = (int)hival.Double;
                var lookup = ProfileToken(indexed.Data[3], options.MaxStackDepth, token);
                var paletteBytes = checked(3 * (highestIndex + 1));
                // Bound literal copying before allocating another palette buffer.
                if (lookup is HexToken hex)
                {
                    if (hex.Bytes.Length != paletteBytes) throw new InvalidDataException("PDF Indexed literal palette length mismatch.");
                    palette = hex.Bytes.ToArray();
                }
                else if (lookup is StringToken literal)
                {
                    if (literal.Data.Length > paletteBytes) throw new InvalidDataException("PDF Indexed literal palette length mismatch.");
                    palette = literal.GetBytes();
                }
                if (palette is null) return null; // Stream lookup and other bases remain ordinary fallback.
                if (palette.Length != paletteBytes) throw new InvalidDataException("PDF Indexed literal palette length mismatch.");
            }
            else
            {
                if (colorSpace is not NameToken colorName) throw new InvalidDataException("PDF image ColorSpace must be a name or array.");
                if (colorName.Data != "DeviceRGB") return null;
            }
            int Dimension(string key)
            {
                var number = Resolve<NumericToken>(Entry(dictionary, key)).Double;
                if (!FinitePositive(number) || number > int.MaxValue || Math.Truncate(number) != number) throw Unsupported("IMAGE_DIMENSION");
                return (int)number;
            }
            var width = Dimension("Width"); var height = Dimension("Height");
            var pixels = checked((long)width * height); var rawBytes = checked(pixels * 3);
            var dpi = Math.Max(72, Math.Min(options.Compatibility.Dpi, 300));
            var renderedPixels = Math.Ceiling(frame.Frame.Width * dpi / 72) * Math.Ceiling(frame.Frame.Height * dpi / 72);
            if (pixels > options.Compatibility.MaxRasterizedPixelsPerPage || renderedPixels > options.Compatibility.MaxRasterizedPixelsPerPage || rawBytes > int.MaxValue)
                throw new InvalidDataException("PDF image/page pixel limit exceeded.");
            var sourceBytes = palette is null ? rawBytes : pixels;
            if (stream.Data.Length != sourceBytes || Resolve<NumericToken>(Entry(dictionary, "Length")).Double != sourceBytes) return null;
            // Conservative allocation guard, in addition to bounded encoded PNG writes.
            if (rawBytes > remainingImageBytes) throw new InvalidDataException("Original RGB allocation exceeds remaining image byte budget.");
            var interpolate = false;
            if (dictionary.Data.TryGetValue("Interpolate", out var hint))
            {
                if (ProfileToken(hint, options.MaxStackDepth, token) is not BooleanToken boolean) return null;
                interpolate = boolean.Data;
            }
            token.ThrowIfCancellationRequested();
            encoding = true;
            using var bitmap = new Image<Rgb24>(width, height);
            bitmap.ProcessPixelRows(accessor =>
            {
                for (var row = 0; row < height; row++)
                {
                    token.ThrowIfCancellationRequested(); var destination = accessor.GetRowSpan(row);
                    if (palette is null)
                    {
                        var original = stream.Data.Span.Slice(checked(row * width * 3), checked(width * 3));
                        for (var column = 0; column < width; column++) destination[column] = new Rgb24(original[column * 3], original[column * 3 + 1], original[column * 3 + 2]);
                    }
                    else
                    {
                        var original = stream.Data.Span.Slice(checked(row * width), width);
                        for (var column = 0; column < width; column++)
                        {
                            // Default 8-bit Indexed decode is [0,255]; values clamp to hival.
                            var index = Math.Min(original[column], highestIndex);
                            var offset = index * 3; destination[column] = new Rgb24(palette[offset], palette[offset + 1], palette[offset + 2]);
                        }
                    }
                }
            });
            using var encoded = new BoundedImageOutput(remainingImageBytes, token);
            bitmap.SaveAsPng(encoded); token.ThrowIfCancellationRequested();
            var page = new OfdPage { WidthMillimeters = frame.Frame.Width * 25.4 / 72, HeightMillimeters = frame.Frame.Height * 25.4 / 72 };
            page.Elements.Add(new OfdImageElement { WidthMillimeters = page.WidthMillimeters, HeightMillimeters = page.HeightMillimeters,
                Data = encoded.ToArray(), FileName = "original-page-image.png", MediaType = "image/png",
                SourceXml = new XElement(XName.Get("ImageObject", options.Compatibility.Namespace), new XAttribute(OfdImageRenderingHints.PdfInterpolateV1, interpolate)).ToString(SaveOptions.DisableFormatting) });
            return page;
        }
        catch (NotSupportedException) when (!encoding) { return null; } // Not this profile. Encoding/budget/cancel/I/O failures escape.
    }
    private IToken Entry(DictionaryToken value, string key) => value.Data.TryGetValue(key, out var entry) ? entry : throw Unsupported("MISSING_RESOURCE: " + key);
    private IToken ProfileToken(IToken value, int maxHops, CancellationToken token)
    {
        var visited = new HashSet<IndirectReference>(); var hops = 0;
        token.ThrowIfCancellationRequested();
        while (value is IndirectReferenceToken reference)
        {
            token.ThrowIfCancellationRequested();
            if (hops++ >= maxHops || !visited.Add(reference.Data)) throw new InvalidDataException("PDF profile reference depth/cycle limit exceeded.");
            // TryGet in PdfPig swallows scanner errors. These must escape page fallback.
            var resolved = Scanner.Get(reference.Data);
            if (resolved is null || resolved.Data is NullToken) throw new InvalidDataException("PDF profile reference is missing or null.");
            value = resolved.Data;
        }
        return value;
    }
    private string Name(DictionaryToken value, string key) => Resolve<NameToken>(Entry(value, key)).Data;
    internal byte[] Unfiltered(StreamToken value, int remaining, string kind)
    {
        if (value.StreamDictionary.Data.ContainsKey("Filter") || value.StreamDictionary.Data.ContainsKey("DecodeParms") || value.StreamDictionary.Data.ContainsKey("F") ||
            value.StreamDictionary.Data.ContainsKey("FFilter") || value.StreamDictionary.Data.ContainsKey("FDecodeParms"))
            throw Unsupported("STREAM_FILTER: " + kind + " requires an inline unfiltered stream.");
        if (value.Data.Length > remaining) throw new InvalidDataException("PDF " + kind + " byte limit exceeded.");
        if (!value.StreamDictionary.TryGet(NameToken.Length, out var length) || Resolve<NumericToken>(length).Double != value.Data.Length)
            throw Unsupported("STREAM_LENGTH: " + kind);
        return value.Data.ToArray();
    }
    private static bool FinitePositive(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
    internal NotSupportedException Unsupported(string reason) => new("PDFV_UNSUPPORTED page=" + _number + " " + reason);
}

internal sealed class BoundedImageOutput : Stream
{
    private readonly MemoryStream _buffer = new(); private readonly long _maximum; private readonly CancellationToken _token;
    internal BoundedImageOutput(long maximum, CancellationToken token) { _maximum = maximum; _token = token; }
    internal byte[] ToArray() => _buffer.ToArray();
    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => _buffer.Length;
    public override long Position { get => _buffer.Position; set => throw new NotSupportedException(); }
    public override void Flush() => _buffer.Flush();
    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count)
    {
        _token.ThrowIfCancellationRequested();
        if (count > _maximum - _buffer.Length) throw new InvalidDataException("PNG encoding exceeds remaining image byte budget.");
        _buffer.Write(buffer, offset, count);
    }
    public override void WriteByte(byte value) { var one = new[] { value }; Write(one, 0, 1); }
    protected override void Dispose(bool disposing) { if (disposing) _buffer.Dispose(); base.Dispose(disposing); }
}

internal sealed class VectorOperations : IGraphicsStateOperationFactory
{
    private static readonly HashSet<string> Allowed = new("q Q cm rg RG g G w J j M d m l c v y re h f F f* S s B B* b b* n W W* Do BT ET Tf Tm Td TD T* Tj TJ Tc Tw TL".Split(' '));
    private static int Arity(string operation) => operation switch
    {
        "cm" or "c" or "Tm" => 6,
        "v" or "y" or "re" => 4,
        "rg" or "RG" => 3,
        "m" or "l" or "Tf" or "Td" or "TD" or "d" => 2,
        "g" or "G" or "w" or "J" or "j" or "M" or "Do" or "Tj" or "TJ" or "Tc" or "Tw" or "TL" => 1,
        _ => 0
    };
    private readonly PdfVectorToOfdOptions _limits; private readonly CancellationToken _token; private readonly int _page; private int _count;
    internal VectorOperations(PdfVectorToOfdOptions limits, CancellationToken token, int page) { _limits = limits; _token = token; _page = page; }
    public IGraphicsStateOperation Create(OperatorToken operation, IReadOnlyList<IToken> operands)
    {
        _token.ThrowIfCancellationRequested();
        if (++_count > _limits.MaxOperationsPerPage) throw new InvalidDataException("PDF operation limit exceeded.");
        if (!Allowed.Contains(operation.Data)) throw new NotSupportedException("PDFV_UNSUPPORTED page=" + _page + " operation=" + _count + " operator=" + operation.Data);
        if (operands.Count != Arity(operation.Data)) throw new NotSupportedException("PDFV_OPERATOR_ARITY page=" + _page + " operator=" + operation.Data);
        return ReflectionGraphicsStateOperationFactory.Instance.Create(operation, operands) ?? throw new NotSupportedException("PDFV_UNKNOWN_OPERATOR " + operation.Data);
    }
}

internal sealed class VectorFont : IDisposable
{
    internal readonly SkiaDrawEvent.FixedFontBinding Binding;
    internal readonly OfdFontResource Resource;
    private readonly byte[]? _cidToGid;
    internal VectorFont(string id, byte[] bytes, int maxFontBytes, byte[]? cidToGid = null)
    {
        _cidToGid = cidToGid;
        Binding = SkiaDrawEvent.FixedFontBinding.FromOriginalBytes(bytes, id, maxFontBytes);
        Resource = new OfdFontResource { Id = id, FontName = Binding.FamilyName, FileName = id + ".ttf",
            Data = bytes, Bold = Binding.Bold, Italic = Binding.Italic };
    }
    internal int GlyphForCid(int cid) => _cidToGid is null ? cid : cid >= 0 && cid < _cidToGid.Length / 2
        ? _cidToGid[cid * 2] * 256 + _cidToGid[cid * 2 + 1] : 0;
    public void Dispose() => Binding.Dispose();
}
