using System;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>The supported SES container layouts.</summary>
public enum SesVersion
{
    /// <summary>Seven-field TBS and two-field signature container.</summary>
    V1 = 1,
    /// <summary>Five-field TBS and external signer certificate.</summary>
    V4 = 4
}

/// <summary>Capabilities of this optional dev/interop package, without production trust claims.</summary>
public static class SesInteropCapabilities
{
    /// <summary>Common SES V1/V4 fields can be inspected.</summary>
    public static bool SupportsCommonSesFieldInspection => true;
    /// <summary>SM2 test self-signing requires explicit provider/verifier construction.</summary>
    public static bool SupportsExplicitTestSelfSigning => true;
}

/// <summary>Common seal fields. Parsing establishes structure only, never seal authenticity.</summary>
public sealed class SesSealInfo
{
    internal SesSealInfo(string id, string vendor, int version, int type, string name,
        string pictureType, byte[] picture, int width, int height)
    {
        SealId = id; VendorId = vendor; Version = version; SealType = type; Name = name;
        PictureType = pictureType; _picture = picture; WidthMillimeters = width; HeightMillimeters = height;
    }
    private readonly byte[] _picture;
    /// <summary>Seal identifier from the container.</summary>
    public string SealId { get; }
    /// <summary>Vendor identifier from the header; not a trust assertion.</summary>
    public string VendorId { get; }
    /// <summary>Seal header version.</summary>
    public int Version { get; }
    /// <summary>Declared seal type.</summary>
    public int SealType { get; }
    /// <summary>Declared seal name.</summary>
    public string Name { get; }
    /// <summary>Declared picture format, including formats the renderer may not support.</summary>
    public string PictureType { get; }
    /// <summary>A copy of picture bytes. No decoding or rendering is performed.</summary>
    public byte[] PictureData => (byte[])_picture.Clone();
    /// <summary>Declared picture width in millimeters.</summary>
    public int WidthMillimeters { get; }
    /// <summary>Declared picture height in millimeters.</summary>
    public int HeightMillimeters { get; }
}

/// <summary>Read-only common fields of a DER SignedValue.dat; no verification is implied.</summary>
public sealed class SesSignedValue
{
    internal SesSignedValue(SesVersion version, SesSealInfo seal, byte[] time, string timeText,
        byte[] hash, string property, byte[] certificate, string algorithm, byte[] signature,
        byte[] tbs, byte[] sealDer, byte[]? timestamp)
    {
        Version = version; Seal = seal; _time = time; TimeText = timeText; _hash = hash;
        PropertyInformation = property; _certificate = certificate; SignatureAlgorithm = algorithm;
        _signature = signature; _tbs = tbs; _sealDer = sealDer; _timestamp = timestamp;
    }
    private readonly byte[] _time, _hash, _certificate, _signature, _tbs, _sealDer;
    private readonly byte[]? _timestamp;
    /// <summary>Detected V1/V4 layout.</summary>
    public SesVersion Version { get; }
    /// <summary>Common electronic seal fields.</summary>
    public SesSealInfo Seal { get; }
    /// <summary>V1 raw time BIT STRING or V4 DER GeneralizedTime bytes, copied.</summary>
    public byte[] TimeInformation => (byte[])_time.Clone();
    /// <summary>V1 UTF-8 time text without timezone inference, or V4 generalized time text.</summary>
    public string TimeText { get; }
    /// <summary>Copy of the signed document digest.</summary>
    public byte[] DataHash => (byte[])_hash.Clone();
    /// <summary>Signed package property-information path.</summary>
    public string PropertyInformation { get; }
    /// <summary>Copy of signer certificate bytes; not a trusted certificate.</summary>
    public byte[] CertificateDer => (byte[])_certificate.Clone();
    /// <summary>Declared signature algorithm OID.</summary>
    public string SignatureAlgorithm { get; }
    /// <summary>Copy of the signature bytes.</summary>
    public byte[] Signature => (byte[])_signature.Clone();
    /// <summary>Copy of DER-encoded to-be-signed data.</summary>
    public byte[] ToBeSignedDer => (byte[])_tbs.Clone();
    /// <summary>Copy of the complete embedded seal.</summary>
    public byte[] SealDer => (byte[])_sealDer.Clone();
    /// <summary>Optional V4 timestamp token bytes, copied; never validated as a TSA token.</summary>
    public byte[]? Timestamp => _timestamp is null ? null : (byte[])_timestamp.Clone();
}
