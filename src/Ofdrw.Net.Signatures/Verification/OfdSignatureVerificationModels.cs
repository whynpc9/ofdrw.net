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
    /// <summary>OFD 2.0 base signature ID from the containing list, when declared.</summary>
    public string? RelativeSignatureId { get; set; }
    /// <summary>Whether the currently declared base signature is unique, earlier, and valid; the mutable list relationship is not itself authenticated.</summary>
    public bool RelativeIntegrityValid { get; internal set; } = true;
    /// <summary>Reason a Relative relationship could not be established.</summary>
    public string? RelativeIssue { get; internal set; }
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
    /// <summary>True when OFD.xml declares a signatures list, including a missing, empty or malformed list.</summary>
    public bool HasSignatureDeclarations { get; internal set; }
    /// <summary>True when every declared signatures list has a nonblank location, can be parsed, and contains at least one signature record.</summary>
    public bool SignatureDeclarationsComplete { get; internal set; }
    /// <summary>False when verification stopped because a configured cumulative work limit was exceeded.</summary>
    public bool VerificationWorkComplete { get; internal set; } = true;
    public bool HasSignatures => Signatures.Count > 0;
    public bool ReferenceIntegrityValid =>
        VerificationWorkComplete && SignatureDeclarationsComplete &&
        HasSignatures && Signatures.All(signature => signature.ReferenceIntegrityValid);
    public bool FullyValid =>
        ReferenceIntegrityValid &&
        Signatures.All(signature =>
            signature.CryptographicStatus == OfdCryptographicVerificationStatus.Valid &&
            signature.RelativeIntegrityValid);
}

public static class OfdCryptographicCapabilities
{
    public static bool SupportsSm3ReferenceDigests => true;
    public static bool SupportsPluggableSignedValueVerification => true;
    public static bool SupportsBuiltInSesSm2Verification => false;
    public static bool SupportsGmT0099EncryptionEnvelope => false;
}
