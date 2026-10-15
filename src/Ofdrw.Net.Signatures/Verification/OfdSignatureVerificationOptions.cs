using System;

namespace Ofdrw.Net.Signatures.Verification;

/// <summary>Finite cumulative limits for one signature verification operation.</summary>
public sealed class OfdSignatureVerificationOptions
{
    /// <summary>Maximum signature records visited across all lists. Defaults to 1,024.</summary>
    public int MaxSignatureRecords { get; set; } = 1024;

    /// <summary>Maximum reference visits, including missing entries and digest cache hits. Defaults to 100,000.</summary>
    public int MaxReferenceVisits { get; set; } = 100000;

    /// <summary>Maximum bytes actually hashed for reference digests. Defaults to 1 GiB.</summary>
    public long MaxDigestBytes { get; set; } = 1024L * 1024 * 1024;

    internal void Validate()
    {
        if (MaxSignatureRecords <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignatureRecords));
        if (MaxReferenceVisits <= 0) throw new ArgumentOutOfRangeException(nameof(MaxReferenceVisits));
        if (MaxDigestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDigestBytes));
    }
}
