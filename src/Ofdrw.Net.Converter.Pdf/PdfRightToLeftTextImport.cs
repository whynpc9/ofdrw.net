namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Controls how PDF→OFD semantic text import treats pages that contain
/// right-to-left letters or bidi controls.
/// </summary>
public enum PdfRightToLeftTextImport
{
    /// <summary>
    /// Default. Reject semantic import of any page that contains right-to-left
    /// letters or bidi controls with <c>PDF_TEXT_READING_ORDER_UNVERIFIED</c>.
    /// </summary>
    Reject,

    /// <summary>
    /// Bounded visual-line profile. Each horizontal line is rebuilt from the
    /// positioned PDF glyphs and their decoded character mappings, a logical
    /// order is derived and accepted only when the Unicode 17 UAX #9 forward
    /// reordering of that logical text reproduces the observed glyph order
    /// exactly. Lines outside the profile (for example digits, punctuation,
    /// combining marks, presentation forms, overlapping or rotated glyphs,
    /// tagged <c>ActualText</c>/<c>ReversedChars</c>, or an ambiguous base
    /// direction) are still rejected.
    /// </summary>
    VisualLineProfile,
}

/// <summary>
/// Base (paragraph) direction assumed for each right-to-left line by
/// <see cref="PdfRightToLeftTextImport.VisualLineProfile"/>.
/// </summary>
public enum PdfTextBaseDirection
{
    /// <summary>
    /// Derive the direction of each line independently (UAX #9 P2/P3 on the
    /// reconstructed line). A line whose observed order is consistent with
    /// both directions, or with neither, is rejected.
    /// </summary>
    Auto,

    /// <summary>Every right-to-left line belongs to a left-to-right paragraph.</summary>
    LeftToRight,

    /// <summary>Every right-to-left line belongs to a right-to-left paragraph.</summary>
    RightToLeft,
}
