using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Signatures.Verification;

public enum OfdCryptographicVerificationStatus
{
    Missing,
    Unsupported,
    Valid,
    Invalid,
    Error
}

public sealed class OfdReferenceVerificationResult
{
    public string FileReference { get; set; } = string.Empty;
    public string ResolvedEntryName { get; set; } = string.Empty;
    public bool EntryExists { get; set; }
    public bool DigestAlgorithmSupported { get; set; }
    public bool DigestMatches { get; set; }
    public string? Error { get; set; }
}

public sealed class OfdSignatureVerificationResult
{
    public string Id { get; set; } = string.Empty;
    public string SignaturePath { get; set; } = string.Empty;
    public string SignatureMethod { get; set; } = string.Empty;
    public string CheckMethod { get; set; } = string.Empty;
    public string? SignedValuePath { get; set; }
    public List<OfdReferenceVerificationResult> References { get; } = new();
    public OfdCryptographicVerificationStatus CryptographicStatus { get; set; }
    public string? CryptographicMessage { get; set; }

    public bool ReferenceIntegrityValid =>
        References.Count > 0 &&
        References.All(reference =>
            reference.EntryExists &&
            reference.DigestAlgorithmSupported &&
            reference.DigestMatches);
}

public sealed class OfdSignatureVerificationReport
{
    public List<OfdSignatureVerificationResult> Signatures { get; } = new();
    public List<string> Issues { get; } = new();
    public bool HasSignatures => Signatures.Count > 0;
    public bool ReferenceIntegrityValid =>
        HasSignatures && Signatures.All(signature => signature.ReferenceIntegrityValid);
    public bool FullyValid =>
        ReferenceIntegrityValid &&
        Signatures.All(signature =>
            signature.CryptographicStatus == OfdCryptographicVerificationStatus.Valid);
}

public static class OfdCryptographicCapabilities
{
    public static bool SupportsSm3ReferenceDigests => true;
    public static bool SupportsPluggableSignedValueVerification => true;
    public static bool SupportsBuiltInSesSm2Verification => false;
    public static bool SupportsGmT0099EncryptionEnvelope => false;
}
