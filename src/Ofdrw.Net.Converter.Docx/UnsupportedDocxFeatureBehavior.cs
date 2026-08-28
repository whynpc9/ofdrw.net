namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Controls how the in-process renderer handles unsupported DOCX content.
/// </summary>
public enum UnsupportedDocxFeatureBehavior
{
    /// <summary>
    /// Ignore unsupported content when its visible text can still be retained.
    /// </summary>
    BestEffort,

    /// <summary>
    /// Emit a visible placeholder for unsupported content.
    /// </summary>
    Placeholder,

    /// <summary>
    /// Stop conversion when unsupported content is encountered.
    /// </summary>
    Throw
}
