using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Crypto.Password;

/// <summary>Explicit selection and bounded processing for the optional self-round-trip password profile.</summary>
public sealed class OfdPasswordOptions
{
    /// <summary>ZIP budgets; no files are extracted to a directory. MaxCompressionRatio must be finite and at least 1 for this uncompressed profile.</summary>
    public OfdPackageLoadOptions LoadOptions { get; set; } = new()
    {
        MaxInputBytes = 64L * 1024 * 1024,
        MaxEntryUncompressedBytes = 16L * 1024 * 1024,
        MaxTotalUncompressedBytes = 64L * 1024 * 1024
    };

    /// <summary>Maximum resulting ZIP bytes. Default 80 MiB; the effective limit is the smaller of this value and LoadOptions.MaxInputBytes (64 MiB by default), so successful output fits the same input budget.</summary>
    public long MaxOutputBytes { get; set; } = 80L * 1024 * 1024;

    /// <summary>Maximum envelope/map XML bytes. Default 4 MiB.</summary>
    public int MaxMetadataBytes { get; set; } = 4 * 1024 * 1024;

    /// <summary>Explicit canonical package paths to encrypt. Empty selection is rejected.</summary>
    public IList<string> EntryNames { get; } = new List<string>();

    /// <summary>Zero-based pages in the sole DocBody. Selects only the Page BaseLoc content entry;
    /// shared fonts, images, templates and annotations require explicit EntryNames.
    /// This is not a promise to remove every representation of a page's content.</summary>
    public IList<int> PageIndices { get; } = new List<int>();

    /// <summary>UserName recorded in the sole UserInfo. Does not impose an access-control policy.</summary>
    public string UserName { get; set; } = "User";
}

/// <summary>Capabilities of this explicitly referenced extension, separate from core SDK capability flags.</summary>
public static class OfdPasswordCapabilities
{
    /// <summary>Supports this package's own bounded password profile and its own outputs only.</summary>
    public static bool SupportsSelfRoundTripPasswordEnvelope => true;

    /// <summary>No vendor interoperability has been established.</summary>
    public static bool VendorInteroperabilityVerified => false;
}
