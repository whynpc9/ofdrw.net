namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Native DOCX-to-OFD font payload policy. Neither value grants a licence to embed a font.
/// </summary>
public enum DocxNativeFontSelfContainment
{
    /// <summary>
    /// Default compatibility behavior: system Chinese families (SimSun, SimHei, DengXian, KaiTi,
    /// FangSong, Microsoft YaHei) and Noto Sans CJK without configured bytes are written as
    /// name-only fonts for the viewer to resolve locally; other families are embedded.
    /// </summary>
    Compatibility,

    /// <summary>
    /// Every font resource used by written text must carry the declared family's own font bytes.
    /// Conversion fails, naming the family, when those bytes are unavailable, when only another
    /// family's stand-in exists, or when the face's OS/2.fsType forbids embedding
    /// (restricted-licence or bitmap-only embedding). Applies to Native DOCX-to-OFD output only.
    /// </summary>
    RequireEmbedded
}
