namespace Ofdrw.Net.Converter.Docx;

/// <summary>Native OFD policy for bold text bound to an embedded non-bold face.</summary>
public enum DocxEmbeddedBoldBehavior
{
    /// <summary>Paint a standard OFD text stroke and report simulated emphasis; never alter font data.</summary>
    TextStroke,
    /// <summary>Fail before output when an actual embedded bold face is unavailable.</summary>
    Throw
}
