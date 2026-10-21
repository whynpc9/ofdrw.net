using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Crypto.Password;

/// <summary>A named recipient in the explicit Java-compatible SM4 envelope profile.</summary>
public sealed class OfdStandardRecipient
{
    internal string Name { get; }
    internal string? Password { get; }
    internal byte[]? CertificateDer { get; }
    private OfdStandardRecipient(string name, string? password, byte[]? certificateDer)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 256) throw new ArgumentException("Invalid recipient name.", nameof(name));
        Name = name; Password = password; CertificateDer = certificateDer;
    }
    /// <summary>Derive this recipient's wrapping key from the exact UTF-8 password using Java's SM3 KDF.</summary>
    public static OfdStandardRecipient ForPassword(string name, string password) =>
        new(name, password ?? throw new ArgumentNullException(nameof(password)), null);
    /// <summary>Wrap the file key for an SM2 X.509 recipient certificate.</summary>
    public static OfdStandardRecipient ForCertificate(string name, byte[] certificateDer) =>
        new(name, null, (byte[])(certificateDer ?? throw new ArgumentNullException(nameof(certificateDer))).Clone());
}

/// <summary>Signing callback for an independently pinned SM2 certificate. The callback signs the supplied
/// 32-byte SM3 digest with SM3-with-SM2 and returns a DER SM2 signature.</summary>
public sealed class OfdStandardManifestSigner(byte[] certificateDer,
    Func<byte[], CancellationToken, Task<byte[]>> signDigestAsync)
{
    /// <summary>Signer certificate embedded in the signed value; callers must pin it independently when verifying.</summary>
    public byte[] CertificateDer { get; } = (byte[])(certificateDer ?? throw new ArgumentNullException(nameof(certificateDer))).Clone();
    /// <summary>Sign the supplied 32-byte SM3 digest with SM3-with-SM2 and return a DER signature.</summary>
    public Func<byte[], CancellationToken, Task<byte[]>> SignDigestAsync { get; } =
        signDigestAsync ?? throw new ArgumentNullException(nameof(signDigestAsync));
}

/// <summary>Explicit finite profile. All original file entries are encrypted, as in the pinned Java writer.</summary>
public sealed class OfdStandardEncryptionOptions
{
    /// <summary>One to 32 uniquely named recipients, all using the same wrapping scheme.</summary>
    public IList<OfdStandardRecipient> Recipients { get; } = new List<OfdStandardRecipient>();
    /// <summary>Required signing callback and certificate for the package inventory.</summary>
    public OfdStandardManifestSigner? ManifestSigner { get; set; }
    /// <summary>ZIP and metadata limits shared with the optional v2 implementation; KDF settings do not apply.</summary>
    public OfdPasswordOptions Limits { get; set; } = new();
}

/// <summary>One explicit identity for a standard envelope. A pinned certificate is required for certificate identity.</summary>
public sealed class OfdStandardCredential
{
    internal string Name { get; }
    internal string? Password { get; }
    internal byte[]? PrivateKeyPkcs8Der { get; }
    internal byte[]? CertificateDer { get; }
    private OfdStandardCredential(string name, string? password, byte[]? privateKey, byte[]? certificate)
    { Name = name; Password = password; PrivateKeyPkcs8Der = privateKey; CertificateDer = certificate; }
    /// <summary>Select one exact recipient name and password; no other user is tried.</summary>
    public static OfdStandardCredential ForPassword(string name, string password) =>
        new(RequireName(name), password ?? throw new ArgumentNullException(nameof(password)), null, null);
    /// <summary>Select one recipient using its PKCS#8 SM2 private key and independently pinned X.509 certificate.</summary>
    public static OfdStandardCredential ForCertificate(string name, byte[] privateKeyPkcs8Der, byte[] expectedCertificateDer) =>
        new(RequireName(name), null,
            (byte[])(privateKeyPkcs8Der ?? throw new ArgumentNullException(nameof(privateKeyPkcs8Der))).Clone(),
            (byte[])(expectedCertificateDer ?? throw new ArgumentNullException(nameof(expectedCertificateDer))).Clone());
    private static string RequireName(string name) => !string.IsNullOrWhiteSpace(name) && name.Length <= 256 ? name :
        throw new ArgumentException("Invalid recipient name.", nameof(name));
}

/// <summary>Manifest signer must be pinned independently of the OFD package.</summary>
public sealed class OfdStandardDecryptionOptions
{
    /// <summary>Certificate pinned from outside the encrypted package for its manifest signature.</summary>
    public byte[]? ExpectedManifestSignerCertificateDer { get; set; }
    /// <summary>Explicitly allow a Java manifest with only a signed file-name list. Result reports unverified contents.</summary>
    public bool AllowUnverifiedContents { get; set; }
    /// <summary>ZIP and metadata limits shared with the optional v2 implementation; v2 KDF settings do not apply.</summary>
    public OfdPasswordOptions Limits { get; set; } = new();
}

/// <summary>Restored logical OFD plus separately reported authentication properties.</summary>
public sealed record OfdStandardDecryptionResult(byte[] OfdBytes, bool ManifestSignatureVerified,
    bool ExactEntrySetVerified, bool FileContentsVerified);

/// <summary>Integrity status after an atomically committed file recovery.</summary>
public sealed record OfdStandardFileDecryptionStatus(bool ManifestSignatureVerified,
    bool ExactEntrySetVerified, bool FileContentsVerified);
