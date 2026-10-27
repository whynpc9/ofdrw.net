using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class DocumentFontContext
{
    private readonly FontParseBudget _parseBudget;
    private readonly CancellationToken _cancellationToken;
    private readonly OfdNameOnlyFontBindings? _nameOnlyBindings;
    private readonly IReadOnlyList<OfdFontResource> _fonts;
    private readonly IFontResolver _fallbackResolver;
    private readonly Dictionary<string, OpenTypeCmap> _fallbackCoverage = new();
    private readonly Dictionary<OfdFontResource, OpenTypeCmap> _coverage = new();
    private readonly Dictionary<OfdFontResource, byte[]> _symbolFonts = new();
    private readonly Dictionary<(OfdFontResource Resource, bool Bold, bool Italic), string> _symbolAliases = new();
    private readonly Dictionary<OfdFontResource, NotSupportedException> _unsupportedCoverage = new();
    private readonly Dictionary<string, PdfSupplementaryFontBinding> _supplementary = new(StringComparer.Ordinal);
    private readonly Dictionary<OfdFontResource, string> _families = new();
    private readonly Dictionary<OfdFontResource, byte[]> _physicalBytes = new();
    private readonly Dictionary<string, byte[]> _snapshots = new();
    private readonly Dictionary<OfdFontResource, string> _nameOnlyFaceIdentities = new();
    private readonly Dictionary<string, byte[]> _selectedVerifiedFaces = new(StringComparer.Ordinal);
    private readonly Dictionary<string, byte[]> _symbolAliasFaces = new(StringComparer.Ordinal);
    private readonly Dictionary<(OfdFontResource Resource, bool Bold, bool Italic), string> _styleAliases = new();
    private readonly Dictionary<(string Family, bool Bold, bool Italic), OfdFontResource> _unbound = new();
    internal Dictionary<string, (SixLabors.Fonts.FontFamily Family, SixLabors.Fonts.FontStyle Style)> OutlineFonts { get; } = new();

    internal DocumentFontContext(IReadOnlyList<OfdFontResource> fonts, IFontResolver? fallbackResolver = null,
        FontParseBudget? parseBudget = null, CancellationToken cancellationToken = default,
        OfdNameOnlyFontBindings? nameOnlyBindings = null)
    {
        _parseBudget = parseBudget ?? new FontParseBudget(cancellationToken: cancellationToken);
        _cancellationToken = cancellationToken;
        _nameOnlyBindings = nameOnlyBindings;
        _parseBudget.Spend(0);
        _fonts = fonts;
        PdfFontRegistry.EnsureInstalled();
        fallbackResolver ??= GlobalFontSettings.FontResolver;
        _fallbackResolver = fallbackResolver;
        foreach (var font in fonts)
        {
            _parseBudget.Spend(0);
            if (font.Data.Length > 0)
            {
                var data = (byte[])OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex).Clone();
                StorePhysical(font, data);
                // Structural corruption must still reject the document at
                // construction. Only a valid glyf face lacking a cmap may
                // wait for the per-element explicit-GID route.
                var parsed = new OpenTypeFace(data);
                try
                {
                    if (parsed.Version == 0x4F54544F &&
                        (parsed.Tables.ContainsKey("CFF ") || parsed.Tables.ContainsKey("CFF2")))
                    {
                        _unsupportedCoverage[font] = new NotSupportedException(
                            "PDF export does not embed CFF font programs in the verified FontFile2 profile.");
                        continue;
                    }
                    var coverage = new OpenTypeCmap(parsed, _parseBudget);
                    _coverage[font] = coverage;
                    if (coverage.IsSymbol) _symbolFonts[font] = data;
                    else _families[font] = PdfFontRegistry.RegisterFontFace(data, font.Bold, font.Italic);
                }
                catch (NotSupportedException exception) { _unsupportedCoverage[font] = exception; }
                catch (System.IO.InvalidDataException exception) when (!parsed.Tables.ContainsKey("cmap"))
                {
                    _ = new TrueTypeGlyphOutlines(data);
                    // A CGTransform source may supply a valid original-GID glyf
                    // face without any Unicode cmap. Ordinary Unicode text must
                    // still fail closed if it selects the same resource.
                    _unsupportedCoverage[font] = new NotSupportedException(
                        "Embedded font cannot supply ordinary Unicode text.", exception);
                }
                continue;
            }

            if (_nameOnlyBindings is not null) continue;

            // Name-only discovery is optional. Unusable host fonts must not make
            // an otherwise renderable document fail before per-element fallback.
            NotSupportedException? discoveredUnsupported = null;
            try
            {
                var local = CjkViewerFontLoader.TryRead(font.FontName) ?? CjkViewerFontLoader.TryRead(font.FamilyName);
                if (local is null && (font.Bold || font.Italic))
                {
                    // Only probe requested styles: plain name-only text can use the
                    // host directly without copying its fonts into our registry.
                    var face = fallbackResolver.ResolveTypeface(font.FontName, font.Bold, font.Italic);
                    if (face is not null) local = fallbackResolver.GetFont(face.FaceName);
                }
                if (IsStandaloneFont(local))
                {
                    local = (byte[])local!.Clone();
                    var coverage = new OpenTypeCmap(new OpenTypeFace(local), _parseBudget);
                    if (coverage.IsSymbol)
                        discoveredUnsupported = new NotSupportedException("Windows symbol cmap character semantics are unmodeled; use an explicit Unicode font.");
                    else
                    {
                        _families[font] = PdfFontRegistry.RegisterFontFace(local!, font.Bold, font.Italic);
                        StorePhysical(font, local!);
                        _coverage[font] = coverage;
                    }
                }
            }
            catch (Exception exception) when (exception is not OutOfMemoryException &&
                                               exception is not OperationCanceledException && !_parseBudget.IsExhausted)
            {
                // Includes malformed font parsing, host I/O, and optional registry
                // budget exhaustion. Embedded OFD fonts above remain strict.
            }
            // Known unsupported semantics must survive the optional-probe catch,
            // rather than allowing the host to draw the same face unverified.
            if (discoveredUnsupported is not null) _unsupportedCoverage[font] = discoveredUnsupported;
        }
    }

    private static bool IsStandaloneFont(byte[]? data) => data is { Length: >= 12 } &&
        ((data[0] == 0 && data[1] == 1 && data[2] == 0 && data[3] == 0) ||
         (data[0] == 'O' && data[1] == 'T' && data[2] == 'T' && data[3] == 'O'));

    internal OpenTypeCmap? Coverage(OfdFontResource? resource, string? family = null) =>
        family is not null && _fallbackCoverage.TryGetValue(family, out var fallback) ? fallback :
        resource is not null && _coverage.TryGetValue(resource, out var value) ? value : null;

    internal IEnumerable<PdfSupplementaryFontBinding> SupplementaryBindings => _supplementary.Values;

    internal IReadOnlyDictionary<string, byte[]> SelectedVerifiedFaces => _selectedVerifiedFaces;

    internal void MarkFontFaceDrawn(OfdFontResource? resource, string? family)
    {
        // Resolve validates every selected face, including text later omitted by
        // the control policy. Only a PDF DrawString needs a FontFile2 descriptor.
        if (resource is not null && _nameOnlyFaceIdentities.TryGetValue(resource, out var identity))
            _selectedVerifiedFaces[identity] = _physicalBytes[resource];
        if (family is not null && _symbolAliasFaces.TryGetValue(family, out var alias))
            _selectedVerifiedFaces[BinaryIdentity.Hash(alias)] = alias;
    }

    internal string PdfText(string family, string logicalText) =>
        _supplementary.TryGetValue(family, out var binding) ? binding.MapText(logicalText) : logicalText;

    internal byte[] OutlineFontBytes(string family, byte[] pdfFontBytes) =>
        _supplementary.TryGetValue(family, out var binding) ? binding.OriginalFont : pdfFontBytes;

    internal string Resolve(OfdTextElement text, out OfdFontResource? resource)
    {
        _parseBudget.Spend(0);
        if (OfdExplicitGlyphRun.Read(text) is { } explicitRun)
        {
            resource = OfdFontSelection.Resolve(_fonts, text);
            PdfTextControlPolicy.Validate(explicitRun.Run.Text);
            var source = OfdExplicitGlyphFont.Read(resource, text, explicitRun, embedFont: true);
            var glyphByUnicode = new Dictionary<int, int>();
            for (var index = 0; index < explicitRun.GlyphIds.Count; index++)
            {
                var code = explicitRun.Run.Text[index];
                var glyph = explicitRun.GlyphIds[index];
                if (glyphByUnicode.TryGetValue(code, out var prior) && prior != glyph)
                    throw new NotSupportedException("One logical character has conflicting original GIDs in a PDF font alias.");
                glyphByUnicode[code] = glyph;
            }
            var alias = new OpenTypeFace(source.FaceBytes);
            alias.Tables["cmap"] = OpenTypeCmap.Build(glyphByUnicode);
            var explicitFamily = PdfFontRegistry.RegisterFontFace(alias.Build(),
                resource!.Bold || text.Weight >= 600, resource.Italic || text.Italic);
            _fallbackCoverage[explicitFamily] = new OpenTypeCmap(alias, _parseBudget);
            return explicitFamily;
        }
        if (EmbeddedFontCoverage.HasExplicitGlyphReferences(text))
            throw new NotSupportedException("PDF export cannot render this CGTransform source.");
        var hasSupplementary = false;
        foreach (var value in text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text))
        {
            PdfTextControlPolicy.Validate(value);
            if (OpenTypeFace.Scalars(value).Any(scalar => scalar > 0xFFFF)) hasSupplementary = true;
        }
        resource = OfdFontSelection.Resolve(_fonts, text);
        if (hasSupplementary && resource?.Data.Length is not > 0)
            throw new NotSupportedException("Supplementary PDF text requires an explicitly embedded static TrueType font with verified Unicode coverage.");
        var hasText = (text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text)).Any(value => !string.IsNullOrEmpty(value));
        if (!hasText) return resource is not null && _families.TryGetValue(resource, out var emptyFamily)
            ? emptyFamily : string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName;
        if (_nameOnlyBindings is not null && resource is null)
            throw new NotSupportedException($"Selected name-only font '{text.FontName}' has no OFD font resource to bind.");
        if (resource is null && hasText)
        {
            var key = (string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName, text.Weight >= 600, text.Italic);
            if (!_unbound.TryGetValue(key, out resource))
                _unbound.Add(key, resource = new OfdFontResource { FontName = key.Item1, Bold = key.Item2, Italic = key.Item3 });
        }
        if (resource is not null && resource.Data.Length == 0 && hasText && _nameOnlyBindings is not null &&
            !_families.ContainsKey(resource))
        {
            OfdNameOnlyFontBindings.ValidateBmpText(text);
            var bound = _nameOnlyBindings.Resolve(resource, _parseBudget);
            _families[resource] = PdfFontRegistry.RegisterFontFace(bound.Data);
            StorePhysical(resource, bound.Data);
            var identity = BinaryIdentity.Hash(bound.Data);
            _nameOnlyFaceIdentities[resource] = identity;
            _coverage[resource] = bound.Coverage;
        }
        if (resource is not null && resource.Data.Length == 0 && hasText &&
            !_families.ContainsKey(resource) && !_unsupportedCoverage.ContainsKey(resource))
        {
            // Regular and unbound names are discovered only when selected. A
            // failed probe may use verified covering defaults, never an unchecked
            // host draw. A known symbol face instead preserves its refusal marker.
            NotSupportedException? unsupportedName = null;
            try
            {
                var selectedFace = _fallbackResolver.ResolveTypeface(resource.FontName, resource.Bold, resource.Italic);
                if (selectedFace is not null)
                {
                    var bytes = (byte[])OpenTypeCollection.SelectNamedFace(_fallbackResolver.GetFont(selectedFace.FaceName), selectedFace.FaceName).Clone();
                    var selectedCoverage = new OpenTypeCmap(new OpenTypeFace(bytes), _parseBudget);
                    if (selectedCoverage.IsSymbol)
                        unsupportedName = new NotSupportedException("Windows symbol cmap character semantics are unmodeled; use an explicit Unicode font.");
                    else
                    {
                        _families[resource] = PdfFontRegistry.RegisterFontFace(bytes, resource.Bold, resource.Italic);
                        StorePhysical(resource, bytes);
                        _coverage[resource] = selectedCoverage;
                    }
                }
            }
            catch (Exception exception) when (exception is not OutOfMemoryException && exception is not OperationCanceledException && !_parseBudget.IsExhausted) { }
            if (unsupportedName is not null) _unsupportedCoverage[resource] = unsupportedName;
            if (!_families.ContainsKey(resource) && !_unsupportedCoverage.ContainsKey(resource))
                return ResolveCoveringFallback(text, resource);
        }
        if (resource is not null && _unsupportedCoverage.TryGetValue(resource, out var unsupported))
            throw new NotSupportedException($"Selected font '{resource.FontName}' has unmodeled font semantics: {unsupported.Message}", unsupported);
        if (resource is not null && _symbolFonts.TryGetValue(resource, out var symbolBytes))
        {
            var symbolCoverage = _coverage[resource];
            SymbolCmapProfile.ValidateText(text, symbolCoverage, _parseBudget, _cancellationToken);
            var key = (resource, resource.Bold || text.Weight >= 600, resource.Italic || text.Italic);
            if (!_symbolAliases.TryGetValue(key, out var symbolFamily))
            {
                var alias = SymbolCmapProfile.CreateUnicodeAlias(symbolBytes, symbolCoverage, _parseBudget);
                symbolFamily = PdfFontRegistry.RegisterFontFace(alias, key.Item2, key.Item3);
                _symbolAliases.Add(key, symbolFamily);
                _symbolAliasFaces[symbolFamily] = alias;
                _fallbackCoverage[symbolFamily] = new OpenTypeCmap(new OpenTypeFace(alias), _parseBudget);
            }
            return symbolFamily;
        }
        if (resource is not null && _coverage.TryGetValue(resource, out var coverage))
        {
            try
            {
                EmbeddedFontCoverage.Validate(text, coverage, resource.FontName, _cancellationToken);
            }
            catch (System.IO.InvalidDataException) when (resource.Data.Length == 0 && _nameOnlyBindings is null && !_parseBudget.IsExhausted)
            {
                return ResolveCoveringFallback(text, resource);
            }
        }
        if (hasSupplementary)
        {
            if (resource is null || !_coverage.TryGetValue(resource, out var sourceCoverage) || !_physicalBytes.TryGetValue(resource, out var originalBytes))
                throw new NotSupportedException("Supplementary PDF text requires a verified embedded Unicode font.");
            var binding = PdfSupplementaryFontBinding.Create(originalBytes, sourceCoverage, text,
                resource.Bold || text.Weight >= 600, resource.Italic || text.Italic, _parseBudget, _cancellationToken);
            if (_supplementary.TryGetValue(binding.Family, out var prior) && !prior.HasSameUnicodeMap(binding))
                throw new NotSupportedException("Supplementary PDF text has conflicting Unicode for one registered font identity.");
            _supplementary[binding.Family] = binding;
            _fallbackCoverage[binding.Family] = sourceCoverage;
            return binding.Family;
        }
        if (resource is not null && _families.TryGetValue(resource, out var family))
        {
            var bold = resource.Bold || text.Weight >= 600;
            var italic = resource.Italic || text.Italic;
            if (bold == resource.Bold && italic == resource.Italic) return family;
            var key = (resource, bold, italic);
            if (!_styleAliases.TryGetValue(key, out var styledFamily))
            {
                styledFamily = PdfFontRegistry.RegisterFontFace(_physicalBytes[resource], bold, italic);
                _styleAliases[key] = styledFamily;
                _fallbackCoverage[styledFamily] = _coverage[resource];
            }
            return styledFamily;
        }
        throw new System.IO.InvalidDataException("No verified font identity was resolved for nonempty text.");
    }
    private void StorePhysical(OfdFontResource resource, byte[] snapshot)
    {
        var identity = BinaryIdentity.Hash(snapshot);
        if (!_snapshots.TryGetValue(identity, out var shared)) _snapshots.Add(identity, shared = snapshot);
        _physicalBytes[resource] = shared;
    }

    private string ResolveCoveringFallback(OfdTextElement text, OfdFontResource resource)
    {
        // Probe the configured default and then the existing Arial fallback.
        // Each candidate must be parsed and cover the text before drawing.
        // Host lookup/parse failures are optional; known missing glyphs are
        // never returned as success or silently painted as boxes.
        Exception? lastFailure = null;
        var candidates = new List<string>();
        try
        {
            var defaultName = _fallbackResolver.DefaultFontName;
            if (!string.IsNullOrWhiteSpace(defaultName)) candidates.Add(defaultName);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException &&
                                           exception is not OperationCanceledException && !_parseBudget.IsExhausted)
        { lastFailure = exception; }
        candidates.Add("Arial");
        foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                var bold = text.Weight >= 600 || resource.Bold;
                var italic = text.Italic || resource.Italic;
                var face = _fallbackResolver.ResolveTypeface(candidate, bold, italic);
                if (face is null) continue;
                var bytes = (byte[])OpenTypeCollection.SelectNamedFace(_fallbackResolver.GetFont(face.FaceName), face.FaceName, allowSingleFace: true).Clone();
                var fallbackCmap = new OpenTypeCmap(new OpenTypeFace(bytes), _parseBudget);
                if (fallbackCmap.IsSymbol) throw new NotSupportedException("A Windows symbol cmap cannot cover a Unicode fallback request.");
                EmbeddedFontCoverage.Validate(text, fallbackCmap, candidate, _cancellationToken);
                var fallbackFamily = PdfFontRegistry.RegisterFontFace(bytes, bold, italic);
                _fallbackCoverage[fallbackFamily] = fallbackCmap;
                return fallbackFamily;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException &&
                                               exception is not OperationCanceledException && !_parseBudget.IsExhausted)
            {
                lastFailure = exception;
            }
        }
        throw new System.IO.InvalidDataException($"No usable covering fallback font was found for '{resource.FontName}'.", lastFailure);
    }

}
