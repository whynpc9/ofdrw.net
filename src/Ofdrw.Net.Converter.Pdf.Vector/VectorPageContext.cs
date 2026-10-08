using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
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
    internal OfdDocumentPackage Convert(PdfVectorToOfdOptions options, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var ancestry = new List<DictionaryToken>(); var node = _page;
        while (true)
        {
            if (ancestry.Count >= options.MaxStackDepth) throw new InvalidDataException("PDF resource ancestry depth exceeded.");
            ancestry.Add(node);
            if (!node.TryGet(NameToken.Parent, out var parent)) break;
            node = Resolve<DictionaryToken>(parent);
        }
        IToken? Inherited(string key) => ancestry.Select(value => value.Data.TryGetValue(key, out var entry) ? entry : null).FirstOrDefault(value => value is not null);
        var media = Resolve<ArrayToken>(Inherited("MediaBox"));
        var numbers = media.Data.Select(item => Resolve<NumericToken>(item).Double).ToArray();
        if (numbers.Length != 4 || numbers[0] != 0 || numbers[1] != 0 || !FinitePositive(numbers[2]) || !FinitePositive(numbers[3]))
            throw Unsupported("PAGE_BOX: requires a positive zero-origin MediaBox.");
        var cropToken = Inherited("CropBox");
        if (cropToken is not null && !Resolve<ArrayToken>(cropToken).Data.Select(item => Resolve<NumericToken>(item).Double).SequenceEqual(numbers))
            throw Unsupported("CROP_BOX: requires CropBox=MediaBox.");
        var rotate = Inherited("Rotate");
        if (rotate is not null && Resolve<NumericToken>(rotate).Double != 0) throw Unsupported("PAGE_ROTATION");
        if (_page.Data.TryGetValue("UserUnit", out var units) && Resolve<NumericToken>(units).Double != 1) throw Unsupported("USER_UNIT");
        if (_page.ContainsKey(NameToken.Create("Annots")) || _page.ContainsKey(NameToken.Create("Group"))) throw Unsupported("ANNOTATIONS_OR_GROUP");
        var resourceToken = Inherited("Resources");
        var resource = resourceToken is null ? DictionaryToken.Empty : Resolve<DictionaryToken>(resourceToken);
        if (resource.ContainsKey(NameToken.ColorSpace)) throw Unsupported("RESOURCE_COLOR_SPACE: device substitutions not supported.");
        using var content = new MemoryStream();
        if (_page.TryGet(NameToken.Contents, out var contents))
        {
            var streams = contents is ArrayToken array ? array.Data : DirectObjectFinder.TryGet<ArrayToken>(contents, Scanner, out var resolvedArray) ? resolvedArray.Data : new[] { contents };
            foreach (var streamToken in streams)
            {
                token.ThrowIfCancellationRequested();
                var bytes = Unfiltered(Resolve<StreamToken>(streamToken), options.MaxContentBytesPerPage - (int)content.Length, "content");
                if (bytes.Length >= options.MaxContentBytesPerPage - content.Length) throw new InvalidDataException("PDF page content byte limit exceeded.");
                content.Write(bytes, 0, bytes.Length); content.WriteByte(10);
            }
        }
        var factory = new VectorOperations(options, token, _number);
        var parser = new PageContentParser(factory, new StackDepthGuard(options.MaxStackDepth), false);
        var operations = parser.Parse(_number, new MemoryInputBytes(content.ToArray()), Parsing.Logger);
        var usedNames = operations.OfType<UglyToad.PdfPig.Graphics.Operations.TextState.SetFontAndSize>().Select(value => value.Font.Data).Distinct(StringComparer.Ordinal).ToArray();
        var fontTokens = resource.TryGet(NameToken.Font, out var fontsToken) ? Resolve<DictionaryToken>(fontsToken) : DictionaryToken.Empty;
        if (fontTokens.Data.Count > options.MaxFontsPerPage) throw new InvalidDataException("PDF page font count limit exceeded.");
        var bindings = new Dictionary<string, VectorFont>(StringComparer.Ordinal);
        var filtered = new Dictionary<NameToken, IToken>(); long fontBytes = 0;
        try
        {
            foreach (var name in usedNames)
            {
                token.ThrowIfCancellationRequested();
                if (!fontTokens.Data.TryGetValue(name, out var fontToken)) throw Unsupported("MISSING_FONT: " + name);
                var font = Resolve<DictionaryToken>(fontToken);
                if (Name(font, "Subtype") != "Type0" || Name(font, "Encoding") != "Identity-H") throw Unsupported("FONT_PROFILE: " + name);
                var descendants = Resolve<ArrayToken>(Entry(font, "DescendantFonts"));
                if (descendants.Length != 1) throw Unsupported("FONT_DESCENDANTS: " + name);
                var cid = Resolve<DictionaryToken>(descendants.Data[0]);
                if (Name(cid, "Subtype") != "CIDFontType2" || (cid.Data.ContainsKey("CIDToGIDMap") && Name(cid, "CIDToGIDMap") != "Identity"))
                    throw Unsupported("CID_MAPPING: " + name);
                // Check before the stock font loader can allocate any decoded stream.
                _ = Unfiltered(Resolve<StreamToken>(Entry(font, "ToUnicode")), options.MaxContentBytesPerPage, "ToUnicode");
                var descriptor = Resolve<DictionaryToken>(Entry(cid, "FontDescriptor"));
                var bytes = Unfiltered(Resolve<StreamToken>(Entry(descriptor, "FontFile2")), options.MaxFontBytes, "font");
                if (bytes.Length > options.MaxTotalFontBytes - fontBytes) throw new InvalidDataException("PDF page font byte limit exceeded.");
                fontBytes += bytes.Length;
                bindings.Add(name, new VectorFont("font-" + bindings.Count, bytes)); filtered.Add(NameToken.Create(name), fontToken);
            }
            var package = new OfdDocumentPackage();
            var page = new OfdPage { WidthMillimeters = numbers[2] * 25.4 / 72, HeightMillimeters = numbers[3] * 25.4 / 72 };
            package.Pages.Add(page);
            foreach (var binding in bindings.Values) package.Fonts.Add(binding.Resource);
            // Only the proven effective font resources are loaded. No unused font/XObject decoding.
            Resources.LoadResourceDictionary(new DictionaryToken(new Dictionary<NameToken, IToken> { [NameToken.Font] = new DictionaryToken(filtered) }));
            try
            {
                using var processor = new VectorStreamProcessor(this, parser, numbers[2], numbers[3], bindings, options, token, _number);
                var events = processor.Process(_number, operations);
                if (events.Count == 0) throw Unsupported("NO_NATIVE_CONTENT: blank, scan or nonpainted page.");
                OfdSkiaAdapter.Append(package, page, events, new OfdSkiaAdapterOptions { MillimetersPerUnit = 25.4 / 72,
                    MaxEvents = options.MaxEventsPerPage, MaxInputPathCommands = options.MaxPathCommandsPerPage, MaxFontBytes = options.MaxFontBytes,
                    Graphics = new OfdGraphicsOptions { MaxPageElements = options.MaxEventsPerPage, MaxPathCommands = options.MaxPathCommandsPerPage,
                        MaxTextCharacters = options.MaxTextCharactersPerPage } }, token);
                return package;
            }
            finally { Resources.UnloadResourceDictionary(); }
        }
        finally { foreach (var binding in bindings.Values) binding.Dispose(); }
    }
    internal T Resolve<T>(IToken? value) where T : class, IToken => DirectObjectFinder.Get<T>(value!, Scanner) ?? throw Unsupported("INVALID_RESOURCE: " + typeof(T).Name);
    private IToken Entry(DictionaryToken value, string key) => value.Data.TryGetValue(key, out var entry) ? entry : throw Unsupported("MISSING_RESOURCE: " + key);
    private string Name(DictionaryToken value, string key) => Resolve<NameToken>(Entry(value, key)).Data;
    private byte[] Unfiltered(StreamToken value, int remaining, string kind)
    {
        if (value.StreamDictionary.Data.ContainsKey("Filter") || value.StreamDictionary.Data.ContainsKey("DecodeParms") || value.StreamDictionary.Data.ContainsKey("F"))
            throw Unsupported("STREAM_FILTER: " + kind + " requires an inline unfiltered stream.");
        if (value.Data.Length > remaining) throw new InvalidDataException("PDF " + kind + " byte limit exceeded.");
        return value.Data.ToArray();
    }
    private static bool FinitePositive(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
    internal NotSupportedException Unsupported(string reason) => new("PDFV_UNSUPPORTED page=" + _number + " " + reason);
}

internal sealed class VectorOperations : IGraphicsStateOperationFactory
{
    private static readonly HashSet<string> Allowed = new("q Q cm rg RG g G w J j M m l c v y re h f F f* S s B B* b b* n BT ET Tf Tm Td TD T* Tj TJ Tc Tw TL".Split(' '));
    private readonly PdfVectorToOfdOptions _limits; private readonly CancellationToken _token; private readonly int _page; private int _count;
    internal VectorOperations(PdfVectorToOfdOptions limits, CancellationToken token, int page) { _limits = limits; _token = token; _page = page; }
    public IGraphicsStateOperation Create(OperatorToken operation, IReadOnlyList<IToken> operands)
    {
        _token.ThrowIfCancellationRequested();
        if (++_count > _limits.MaxOperationsPerPage) throw new InvalidDataException("PDF operation limit exceeded.");
        if (!Allowed.Contains(operation.Data)) throw new NotSupportedException("PDFV_UNSUPPORTED page=" + _page + " operation=" + _count + " operator=" + operation.Data);
        return ReflectionGraphicsStateOperationFactory.Instance.Create(operation, operands) ?? throw new NotSupportedException("PDFV_UNKNOWN_OPERATOR " + operation.Data);
    }
}

internal sealed class VectorFont : IDisposable
{
    internal readonly SKTypeface Face; internal readonly OfdFontResource Resource;
    internal VectorFont(string id, byte[] bytes)
    {
        using var data = SKData.CreateCopy(bytes); Face = SKTypeface.FromData(data) ?? throw new NotSupportedException("PDFV_FONT_FACE");
        try
        {
            var selection = Word(Face, 0x4f532f32u, 62); var mac = Word(Face, 0x68656164u, 44);
            var bold = (selection & 0x20) != 0; var italic = (selection & 1) != 0;
            if (((mac & 1) != 0) != bold || ((mac & 2) != 0) != italic) throw new NotSupportedException("PDFV_FONT_STYLE_FLAGS");
            Resource = new OfdFontResource { Id = id, FontName = Face.FamilyName, FileName = id + ".ttf", Data = bytes, Bold = bold, Italic = italic };
        }
        catch { Face.Dispose(); throw; }
    }
    private static ushort Word(SKTypeface face, uint tag, int offset)
    {
        var size = face.GetTableSize(tag);
        if (size < offset + 2 || size > 1024) throw new NotSupportedException("PDFV_FONT_STYLE_TABLE");
        var bytes = face.GetTableData(tag);
        if (bytes is null || bytes.Length < offset + 2) throw new NotSupportedException("PDFV_FONT_STYLE_TABLE");
        return (ushort)((bytes[offset] << 8) | bytes[offset + 1]);
    }
    public void Dispose() => Face.Dispose();
}
