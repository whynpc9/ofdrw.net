using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ofdrw.Net.Layout.Text.Paragraphs;

namespace Ofdrw.Net.Layout;

/// <summary>Opt-in final text layout for <see cref="FlowDocument"/> and <see cref="PlainTextOfdConverter"/>.</summary>
/// <remarks>
/// <para>When <see cref="FlowDocumentOptions.TextEngine"/> is set, paragraph and table-cell text is laid out by
/// <see cref="ParagraphLayoutEngine"/> (Unicode bidi, line breaking and grapheme clusters, run shaping, font
/// fallback, tab stops, synthetic styles) and written as cluster-mapped OFD text with embedded GID-preserving
/// font subsets. Without it, Flow keeps its compatibility metrics and name-only fonts.</para>
/// <para>Font families are explicit: a span's family name selects the matching family (ordinal, case
/// insensitive) first; every other supplied family follows as fallback in the given order. No host
/// fonts are discovered. Paragraphs with inline images keep the compatibility layout.</para>
/// </remarks>
public sealed class FlowTextEngine
{
    /// <summary>Creates an engine configuration with one or more explicit font families.</summary>
    /// <param name="engine">The paragraph layout engine, for example <c>HarfBuzzTextPrelayout.CreateParagraphLayout()</c>.</param>
    /// <param name="families">One to sixteen families; the first is the default for unmatched names.</param>
    /// <param name="syntheticStyles">How bold/italic requests without a declared face are handled.</param>
    /// <param name="language">BCP 47 language passed to shaping.</param>
    public FlowTextEngine(ParagraphLayoutEngine engine, IReadOnlyList<TextFontFamily> families,
        TextSyntheticStylePolicy syntheticStyles = TextSyntheticStylePolicy.Synthesize, string language = "und")
    {
        Engine = engine ?? throw new ArgumentNullException(nameof(engine));
        if (families is null || families.Count is < 1 or > 16 || families.Any(family => family is null))
            throw new ArgumentException("Supply one to sixteen font families.", nameof(families));
        if (!Enum.IsDefined(typeof(TextSyntheticStylePolicy), syntheticStyles))
            throw new ArgumentOutOfRangeException(nameof(syntheticStyles));
        if (string.IsNullOrWhiteSpace(language) || language.Length > 32)
            throw new ArgumentException("A short BCP 47 language tag is required.", nameof(language));
        Families = new ReadOnlyCollection<TextFontFamily>(families.ToArray());
        SyntheticStyles = syntheticStyles;
        Language = language;
    }

    /// <summary>The paragraph layout engine.</summary>
    public ParagraphLayoutEngine Engine { get; }
    /// <summary>Explicit font families in fallback order.</summary>
    public IReadOnlyList<TextFontFamily> Families { get; }
    /// <summary>Synthetic bold/italic policy; synthesis is allowed by default.</summary>
    public TextSyntheticStylePolicy SyntheticStyles { get; }
    /// <summary>Language passed to shaping.</summary>
    public string Language { get; }

    /// <summary>The requested family first (when supplied), then every other family in order.</summary>
    internal IReadOnlyList<TextFontFamily> Chain(string? family)
    {
        var requested = family is null ? null :
            Families.FirstOrDefault(value => string.Equals(value.Name, family.Trim(), StringComparison.OrdinalIgnoreCase));
        return requested is null ? Families : new[] { requested }.Concat(Families.Where(value => !ReferenceEquals(value, requested))).ToArray();
    }
}
