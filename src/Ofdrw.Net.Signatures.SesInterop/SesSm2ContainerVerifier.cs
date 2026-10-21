using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit format-aware dispatcher for SES and GB/T 35275 containers sharing the SM2/SM3 OID.</summary>
/// <remarks>Only the supplied pinned verifiers are used; no certificate or format is trusted from the container.</remarks>
public sealed class SesSm2ContainerVerifier : IOfdSignedValueVerifier
{
    private readonly SesV5PinnedVerifier? _v5;
    private readonly Gbt35275PinnedVerifier? _signedData;
    private readonly SesTestSignedValueVerifier? _test;

    /// <summary>Register any required, independently configured profile verifiers under one algorithm OID.</summary>
    public SesSm2ContainerVerifier(SesV5PinnedVerifier? v5 = null,
        Gbt35275PinnedVerifier? signedData = null, SesTestSignedValueVerifier? test = null)
    {
        if (v5 is null && signedData is null && test is null)
            throw new ArgumentException("At least one explicit verifier is required.");
        _v5 = v5; _signedData = signedData; _test = test;
    }

    /// <summary>Shared SM2-with-SM3 OID.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;

    /// <summary>Dispatch by bounded container structure and exact SES version, never by trial verification.</summary>
    public async Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (signatureXml is null) throw new ArgumentNullException(nameof(signatureXml));
        if (signedValue is null) throw new ArgumentNullException(nameof(signedValue));
        if (propertyInformation is null) throw new ArgumentNullException(nameof(propertyInformation));
        try
        {
            var outer = SesDer.Sequence(SesDer.Read(signedValue), 2, 4, 5);
            if (outer[0].ToAsn1Object() is DerObjectIdentifier)
                return _signedData is not null && await _signedData.VerifyAsync(
                    signatureXml, signedValue, propertyInformation, cancellationToken).ConfigureAwait(false);
            if (outer[0].ToAsn1Object() is not Asn1Sequence) return false;
            var version = SesSignedValueReader.Parse(signedValue).Version;
            if (version == SesVersion.V5)
                return _v5 is not null && await _v5.VerifyAsync(
                    signatureXml, signedValue, propertyInformation, cancellationToken).ConfigureAwait(false);
            return _test is not null && await _test.VerifyAsync(
                signatureXml, signedValue, propertyInformation, cancellationToken).ConfigureAwait(false);
        }
        catch (InvalidDataException) { return false; }
    }
}
