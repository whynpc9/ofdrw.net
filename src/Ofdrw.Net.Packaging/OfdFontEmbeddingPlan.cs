using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
namespace Ofdrw.Net.Packaging;

internal sealed class OfdFontEmbeddingPlan
{
    internal Dictionary<OfdFontResource, OfdFontResource> Resources { get; } = new();
    internal List<OfdFontEmbeddingResult> Results { get; } = new();
    internal List<string> Diagnostics { get; } = new();
    internal bool CanonicalPayloads { get; }

    internal OfdFontEmbeddingPlan(OfdDocumentPackage package, CancellationToken cancellationToken)
    {
        var options = package.Options.FontEmbedding ?? throw new ArgumentException("FontEmbedding is required.");
        if (options.MaximumCollectionBytes <= 0 || options.MaximumFontBytes <= 0 || options.MaximumUsedScalars <= 0 || options.MaximumGlyphClosureOperations <= 0 || !Enum.IsDefined(typeof(OfdFontEmbeddingMode), options.Mode))
            throw new ArgumentOutOfRangeException(nameof(package), "Invalid font embedding options.");
        var unsafeReason = PreservationReason(package);
        CanonicalPayloads = unsafeReason is null;
        // Snapshot selected faces once per content/index, then aggregate aliases
        // by the actual face bytes. Same names with different bytes stay distinct.
        var selected = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var groups = new Dictionary<string, List<OfdFontResource>>(StringComparer.Ordinal);
        var resourceGroups = new Dictionary<OfdFontResource, string>();
        foreach (var font in package.Fonts.Where(font => font.Data.Length > 0))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var collection = font.Data.Length >= 4 && OpenTypeFace.U32(font.Data, 0) == 0x74746366;
            if (font.Data.LongLength > (collection ? options.MaximumCollectionBytes : options.MaximumFontBytes))
                throw new InvalidDataException(collection ? "Font collection exceeds MaximumCollectionBytes." : "Font exceeds MaximumFontBytes.");
            var rawIdentity = BinaryIdentity.Hash(font.Data);
            var key = rawIdentity + ":" + font.CollectionFaceIndex;
            if (!selected.TryGetValue(key, out var bytes))
            {
                if (collection) bytes = OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex, options.MaximumFontBytes);
                else
                {
                    if (font.CollectionFaceIndex != 0) throw new InvalidDataException("Standalone font has only face zero.");
                    bytes = (byte[])font.Data.Clone();
                }
                selected.Add(key, bytes);
            }
            var identity = BinaryIdentity.Hash(bytes);
            if (!groups.TryGetValue(identity, out var aliases)) groups.Add(identity, aliases = new());
            aliases.Add(font);
            resourceGroups[font] = identity;
            Resources[font] = Copy(font, bytes);
        }
        // Resolve every typed binding once, even when it selects no embedded
        // group. Cancellation is checked before every examined element.
        var usage = new Dictionary<string, List<OfdTextElement>>(StringComparer.Ordinal);
        if (options.Mode != OfdFontEmbeddingMode.Full)
            foreach (var page in package.Pages)
                foreach (var element in page.Elements)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (element is not OfdTextElement text) continue;
                    var font = OfdFontSelection.Resolve(package.Fonts, text);
                    if (font is null || !resourceGroups.TryGetValue(font, out var identity)) continue;
                    if (!usage.TryGetValue(identity, out var boundTexts)) usage.Add(identity, boundTexts = new());
                    boundTexts.Add(text);
                }
        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested(); var aliases = group.Value; var source = Resources[aliases[0]].Data;
            var result = new OfdFontEmbeddingResult { SourceIdentity = group.Key, SourceBytes = source.LongLength, ResourceCount = aliases.Count };
            var reason = options.Mode == OfdFontEmbeddingMode.Full ? "Explicit full embedding policy." : unsafeReason;
            var payload = source;
            OpenTypeFace? face = null;
            try { face = new OpenTypeFace(source); }
            catch (InvalidDataException) when (unsafeReason is not null || options.Mode == OfdFontEmbeddingMode.Full)
            { reason += " Unparsed preserved font: typed coverage was not validated."; }
            // License flags apply to the supplied parsed face independently of
            // opaque content, full preservation and the requested subset mode.
            if (face is not null && face.Tables.TryGetValue("OS/2", out var os2))
            {
                var permissions = OpenTypeFace.U16(os2, 8);
                if ((permissions & 0x0002) != 0 || (permissions & 0x0200) != 0)
                    throw new InvalidDataException("Font embedding is restricted by OS/2.fsType; supply a licensed outline font.");
                if ((permissions & 0x0100) != 0)
                    reason = (reason is null ? "" : reason + " ") + "OS/2.fsType prohibits subsetting; full font retained.";
            }
            if (options.Mode != OfdFontEmbeddingMode.Full)
            {
                OpenTypeCmap? cmap = null;
                try { if (face is not null) { cmap = new OpenTypeCmap(face); if (cmap.IsSymbol)
                    {
                        reason ??= "Windows symbol cmap retained in full; Unicode coverage/subsetting is unverified.";
                        cmap = null;
                        Diagnostics.Add("FONT_COVERAGE_UNVERIFIED " + group.Key + ": Windows symbol character semantics are not modeled; original font preserved.");
                    } } }
                catch (NotSupportedException exception) { reason = exception.Message + " Full font retained; coverage was not validated."; }
                var used = new HashSet<int>();
                foreach (var text in usage.TryGetValue(group.Key, out var boundTexts) ? boundTexts : Enumerable.Empty<OfdTextElement>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (EmbeddedFontCoverage.HasExplicitGlyphReferences(text))
                    {
                        Diagnostics.Add("FONT_COVERAGE_UNVERIFIED " + group.Key + ": CGTransform has explicit glyph references; original font preserved.");
                        continue;
                    }
                    var strings = text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text);
                    foreach (var value in strings)
                    {
                        if (cmap is not null) EmbeddedFontCoverage.Validate(value, cmap, aliases[0].FontName);
                        foreach (var normalized in new[] { value, value.Normalize(NormalizationForm.FormC), value.Normalize(NormalizationForm.FormD) })
                            foreach (var scalar in OpenTypeFace.Scalars(normalized))
                            {
                                if (UnicodeFontSubsetProfile.RequiresBidiMirroring(scalar))
                                    reason ??= "RTL/bidi shaping requires a Unicode mirror closure; full font retained.";
                                if (OpenTypeCmap.IsVariationSelector(scalar)) continue;
                                used.Add(scalar);
                                if (used.Count > options.MaximumUsedScalars) throw new InvalidDataException("Font usage exceeds MaximumUsedScalars.");
                            }
                    }
                }
                try
                {
                    if (reason is not null) throw new NotSupportedException(reason);
                    if (face is null || !face.Tables.ContainsKey("glyf")) throw new NotSupportedException("CFF/CFF2 outline subsetting is not implemented; full font retained.");
                    payload = TrueTypeSubset.Create(face, used, cancellationToken, options.MaximumGlyphClosureOperations, out var retained, out var original);
                    payload = OpenTypeFontIdentity.WithUniqueNames(payload, BinaryIdentity.Hash(payload));
                    result.RetainedGlyphs = retained; result.OriginalGlyphs = original; result.IsSubset = retained < original;
                    if (!result.IsSubset) { payload = source; reason = "Glyph closure retains every glyph; full font retained."; }
                }
                catch (NotSupportedException exception) { reason = exception.Message; payload = source; }
            }
            result.PayloadIdentity = BinaryIdentity.Hash(payload); result.PayloadBytes = payload.LongLength; result.PreservationReason = reason;
            foreach (var alias in aliases) Resources[alias] = Copy(alias, payload);
            Results.Add(result);
            if (reason is not null) Diagnostics.Add("FONT_FULL_PRESERVED " + group.Key + ": " + reason);
        }
    }
    private static string? PreservationReason(OfdDocumentPackage package)
    {
        if (package.PreservedEntries.Count > 0 || package.PreservedCommonDataElements.Count > 0 ||
            package.PreservedDocumentElements.Count > 0 || package.PreservedDocBodyElements.Count > 0)
            return "Read/edited package has preserved content; glyph references may be unmodeled.";
        foreach (var page in package.Pages)
        {
            if (page.PreservedPageElements.Count > 0 || page.Templates.Count > 0 || page.AnnotationAppearances.Count > 0)
                return "Preserved page/template/annotation content may contain unmodeled glyph references.";
            foreach (var element in page.Elements)
            {
                if (element is OfdRawElement || element is OfdTextElement { SourceXml: not null } ||
                    element is OfdImageElement { SourceXml: not null } || element is OfdPathElement { SourceXml: not null } ||
                    !string.IsNullOrEmpty(element.ClippingXml))
                    return "SourceXml/Raw/clipping content may contain unmodeled glyph references (including CGTransform).";
            }
        }
        return null;
    }
    private static OfdFontResource Copy(OfdFontResource font, byte[] data) => new()
    {
        Id = font.Id, FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
        Bold = font.Bold, Italic = font.Italic, Data = data,
        FileName = data.Length >= 4 && OpenTypeFace.U32(data, 0) == 0x4F54544F ? ".otf" : ".ttf"
    };
}
