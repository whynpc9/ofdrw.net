using System;

namespace Ofdrw.Net.Medical;

/// <summary>The GB/T 48666-2026 root <c>DocType</c> declaration (clause 7.1(a)).</summary>
public static class MedicalDocTypes
{
    /// <summary>
    /// The ASCII declaration <c>OFD-H</c> (U+002D hyphen-minus). The published PDF typesets it as <c>OFD⁃H</c> (U+2043);
    /// that glyph is a typesetting artefact and is never written or accepted by this package.
    /// </summary>
    public const string Medical = "OFD-H";

    /// <summary>Returns whether a raw, declared <c>DocType</c> is exactly <c>OFD-H</c> (ordinal comparison). Null is not medical.</summary>
    /// <param name="declaredDocType">The raw root attribute value, or null when absent.</param>
    /// <returns>True only for the exact ASCII value.</returns>
    public static bool IsMedical(string? declaredDocType) => string.Equals(declaredDocType, Medical, StringComparison.Ordinal);

    /// <summary>
    /// Returns whether a value is <c>OFD</c>, one non-ASCII dash-like character and <c>H</c>, such as text copied from the
    /// typeset standard (U+2043, U+2010–U+2015, U+2212, U+FE63, U+FF0D).
    /// </summary>
    /// <param name="declaredDocType">The raw root attribute value.</param>
    /// <returns>True for a dash-variant spelling that is not the ASCII declaration.</returns>
    public static bool IsTypesetVariant(string? declaredDocType) =>
        declaredDocType is { Length: 5 } value && value.StartsWith("OFD", StringComparison.Ordinal) && value[4] == 'H' &&
        (value[3] == '⁃' || (value[3] >= '‐' && value[3] <= '―') || value[3] == '−' ||
         value[3] == '﹣' || value[3] == '－');
}
