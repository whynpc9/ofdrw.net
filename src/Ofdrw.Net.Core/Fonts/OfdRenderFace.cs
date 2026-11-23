using System;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>
/// The renderer-neutral face policy shared by PDF and SVG/HTML export: which bytes a selected OFD font resource
/// supplies, its parsed Unicode coverage, and whether a given TextObject may be drawn with it. Each renderer keeps
/// only its own adapter: PDF registers faces with PDFsharp (whose process-wide registry lifetime stays PDF-side) and
/// SVG emits <c>@font-face</c> rules. This module owns no resolver, cache or subset.
/// </summary>
internal sealed class OfdRenderFace
{
    private OfdRenderFace(byte[] data, OpenTypeCmap? coverage, bool nameOnlyBinding)
    { Data = data; Coverage = coverage; NameOnlyBinding = nameOnlyBinding; }

    /// <summary>The selected face program: a collection's chosen face, or the explicitly bound name-only face.</summary>
    internal byte[] Data { get; }

    /// <summary>Parsed cmap; null when the face's character semantics are unmodeled.</summary>
    internal OpenTypeCmap? Coverage { get; }

    /// <summary>True when the bytes came from an explicit name-only binding rather than the OFD resource.</summary>
    internal bool NameOnlyBinding { get; }

    internal bool IsSymbol => Coverage?.IsSymbol == true;

    /// <summary>
    /// Loads the face a font resource supplies. Embedded bytes select <see cref="OfdFontResource.CollectionFaceIndex"/>;
    /// a name-only resource requires <paramref name="bindings"/>. An embedded face whose cmap is unsupported loads with
    /// a null <see cref="Coverage"/>; drawing text with it then fails in <see cref="ValidateText"/>.
    /// </summary>
    internal static OfdRenderFace Load(OfdFontResource font, OfdNameOnlyFontBindings? bindings, FontParseBudget budget)
    {
        if (font.Data.Length == 0)
        {
            if (bindings is null) throw new NotSupportedException($"Selected name-only font '{font.FontName}' has no explicit face binding.");
            var bound = bindings.Resolve(font, budget);
            return new OfdRenderFace(bound.Data, bound.Coverage, true);
        }
        var data = OpenTypeCollection.SelectFace(font.Data, font.CollectionFaceIndex);
        var face = new OpenTypeFace(data);
        OpenTypeCmap? coverage = null;
        try { coverage = new OpenTypeCmap(face, budget); }
        catch (NotSupportedException) { /* Unused valid encodings may still be emitted; text validation refuses them. */ }
        return new OfdRenderFace(data, coverage, false);
    }

    /// <summary>
    /// The one coverage policy for drawing a TextObject with this face: name-only bindings are limited to BMP text,
    /// symbol faces follow the symbol cmap profile, and every other face must cover each scalar.
    /// </summary>
    internal void ValidateText(OfdTextElement text, string fontName, FontParseBudget budget, CancellationToken token) =>
        ValidateText(text, Coverage, NameOnlyBinding, fontName, budget, token);

    /// <summary><see cref="ValidateText(OfdTextElement, string, FontParseBudget, CancellationToken)"/> for a coverage a renderer parsed itself.</summary>
    internal static void ValidateText(OfdTextElement text, OpenTypeCmap? coverage, bool nameOnlyBinding, string fontName,
        FontParseBudget budget, CancellationToken token)
    {
        if (nameOnlyBinding) OfdNameOnlyFontBindings.ValidateBmpText(text);
        if (coverage is null)
            throw new NotSupportedException("Selected embedded font has unmodeled cmap character semantics; the original OFD font remains preserved.");
        if (coverage.IsSymbol) SymbolCmapProfile.ValidateText(text, coverage, budget, token);
        else EmbeddedFontCoverage.Validate(text, coverage, fontName, token);
    }
}
