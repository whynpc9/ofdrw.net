using System;

namespace Ofdrw.Net.Converter.Pdf.Geometry;

/// <summary>An explicitly vetted identity of original embedded font bytes; it supplies no replacement font data.</summary>
/// <remarks>The application is responsible for vetting the exact file, its license, actual style and pinned-backend cost.
/// A digest supplied by an incoming document is not a trust decision. Styled identities do not imply that styled geometry is supported.</remarks>
public sealed class OfdTrustedEmbeddedFontIdentity
{
    /// <summary>Creates a trusted original-byte identity using a SHA256 digest, exact length and actual face flags.</summary>
    public OfdTrustedEmbeddedFontIdentity(string sha256, int byteCount, bool bold = false, bool italic = false)
    {
        if (sha256 is null || sha256.Length != 64) throw new ArgumentException("SHA256 must contain 64 hexadecimal characters.", nameof(sha256));
        foreach (var c in sha256)
            if (!(c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F'))
                throw new ArgumentException("SHA256 must contain 64 hexadecimal characters.", nameof(sha256));
        if (byteCount <= 0 || byteCount > 32 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(byteCount));
        Sha256 = sha256.ToLowerInvariant(); ByteCount = byteCount; Bold = bold; Italic = italic;
    }
    /// <summary>Expected internally computed SHA256 over the unchanged source snapshot.</summary>
    public string Sha256 { get; }
    /// <summary>Exact original byte count, at most 32 MiB.</summary>
    public int ByteCount { get; }
    /// <summary>Vetted actual bold flag, not permission to synthesize bold.</summary>
    public bool Bold { get; }
    /// <summary>Vetted actual italic flag, not permission to synthesize italic.</summary>
    public bool Italic { get; }
}
