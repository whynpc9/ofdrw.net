using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class DocumentFontContext
{
    private readonly IReadOnlyList<OfdFontResource> _fonts;
    private readonly IFontResolver _fallbackResolver;
    private readonly Dictionary<string, OpenTypeCmap> _fallbackCoverage = new();
    private readonly Dictionary<OfdFontResource, OpenTypeCmap> _coverage = new();
    private readonly Dictionary<OfdFontResource, NotSupportedException> _unsupportedCoverage = new();
    private readonly Dictionary<OfdFontResource, string> _families = new();
    private readonly Dictionary<OfdFontResource, byte[]> _physicalBytes = new();
    private readonly Dictionary<string, byte[]> _snapshots = new();
    private readonly Dictionary<(OfdFontResource Resource, bool Bold, bool Italic), string> _styleAliases = new();
    private readonly Dictionary<(string Family, bool Bold, bool Italic), OfdFontResource> _unbound = new();
    internal Dictionary<string, SixLabors.Fonts.FontFamily> OutlineFonts { get; } = new();

    internal DocumentFontContext(IReadOnlyList<OfdFontResource> fonts, IFontResolver? fallbackResolver = null)
    {
        _fonts = fonts;
        PdfFontRegistry.EnsureInstalled();
        fallbackResolver ??= GlobalFontSettings.FontResolver;
        _fallbackResolver = fallbackResolver;
        foreach (var font in fonts)
        {
            if (font.Data.Length > 0)
            {
                var data = (byte[])OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex).Clone();
                _families[font] = PdfFontRegistry.RegisterFontFace(data, font.Bold, font.Italic);
                StorePhysical(font, data);
                try
                {
                    var coverage = new OpenTypeCmap(new OpenTypeFace(data));
                    if (coverage.IsSymbol) throw new NotSupportedException("Windows symbol cmap character semantics are unmodeled; the original OFD font remains preserved.");
                    _coverage[font] = coverage;
                }
                catch (NotSupportedException exception) { _unsupportedCoverage[font] = exception; }
                continue;
            }

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
                    var coverage = new OpenTypeCmap(new OpenTypeFace(local));
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
                                               exception is not OperationCanceledException)
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

    internal string Resolve(OfdTextElement text, out OfdFontResource? resource)
    {
        if (EmbeddedFontCoverage.HasExplicitGlyphReferences(text))
            throw new NotSupportedException("PDF export does not model CGTransform glyph substitutions; the OFD font and XML remain preserved.");
        foreach (var value in text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text))
        {
            PdfTextControlPolicy.Validate(value);
            if (OpenTypeFace.Scalars(value).Any(scalar => scalar > 0xFFFF))
                throw new NotSupportedException("PDFsharp cannot preserve supplementary Unicode text; use OFD/SVG or a PDF renderer with full scalar support.");
        }
        resource = OfdFontSelection.Resolve(_fonts, text);
        var hasText = (text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text)).Any(value => !string.IsNullOrEmpty(value));
        if (!hasText) return resource is not null && _families.TryGetValue(resource, out var emptyFamily)
            ? emptyFamily : string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName;
        if (resource is null && hasText)
        {
            var key = (string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName, text.Weight >= 600, text.Italic);
            if (!_unbound.TryGetValue(key, out resource))
                _unbound.Add(key, resource = new OfdFontResource { FontName = key.Item1, Bold = key.Item2, Italic = key.Item3 });
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
                    var selectedCoverage = new OpenTypeCmap(new OpenTypeFace(bytes));
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
            catch (Exception exception) when (exception is not OutOfMemoryException && exception is not OperationCanceledException) { }
            if (unsupportedName is not null) _unsupportedCoverage[resource] = unsupportedName;
            if (!_families.ContainsKey(resource) && !_unsupportedCoverage.ContainsKey(resource))
                return ResolveCoveringFallback(text, resource);
        }
        if (resource is not null && _unsupportedCoverage.TryGetValue(resource, out var unsupported))
            throw new NotSupportedException($"Selected font '{resource.FontName}' has unmodeled cmap coverage.", unsupported);
        if (resource is not null && _coverage.TryGetValue(resource, out var coverage))
        {
            try
            {
                foreach (var value in text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text))
                    EmbeddedFontCoverage.Validate(value, coverage, resource.FontName);
            }
            catch (System.IO.InvalidDataException) when (resource.Data.Length == 0)
            {
                return ResolveCoveringFallback(text, resource);
            }
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
                                           exception is not OperationCanceledException)
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
                var fallbackCmap = new OpenTypeCmap(new OpenTypeFace(bytes));
                if (fallbackCmap.IsSymbol) throw new NotSupportedException("A Windows symbol cmap cannot cover a Unicode fallback request.");
                foreach (var value in text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text))
                    EmbeddedFontCoverage.Validate(value, fallbackCmap, candidate);
                var fallbackFamily = PdfFontRegistry.RegisterFontFace(bytes, bold, italic);
                _fallbackCoverage[fallbackFamily] = fallbackCmap;
                return fallbackFamily;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException &&
                                               exception is not OperationCanceledException)
            {
                lastFailure = exception;
            }
        }
        throw new System.IO.InvalidDataException($"No usable covering fallback font was found for '{resource.FontName}'.", lastFailure);
    }

}
