using System;
using System.Collections.Generic;
using System.Threading;

namespace Ofdrw.Net.Signatures.Verification;

internal sealed class OfdSignatureVerificationBudget
{
    private int _signatureRecordsRemaining;
    private int _referenceVisitsRemaining;
    private long _digestBytesRemaining;

    // Archive lookups normalize slash/leading slash and compare entry names ignoring case.
    // The archive rejects duplicate normalized names, so this identifies the actual entry.
    internal Dictionary<string, Dictionary<string, byte[]>> Digests { get; } =
        new(StringComparer.OrdinalIgnoreCase);

    internal OfdSignatureVerificationBudget(OfdSignatureVerificationOptions options)
    {
        _signatureRecordsRemaining = options.MaxSignatureRecords;
        _referenceVisitsRemaining = options.MaxReferenceVisits;
        _digestBytesRemaining = options.MaxDigestBytes;
    }

    internal void VisitSignature(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_signatureRecordsRemaining == 0) throw new OfdSignatureVerificationBudgetExceededException("signature record");
        _signatureRecordsRemaining--;
    }

    internal void VisitReference(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_referenceVisitsRemaining == 0) throw new OfdSignatureVerificationBudgetExceededException("reference visit");
        _referenceVisitsRemaining--;
    }

    internal void HashBytes(int length, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (length > _digestBytesRemaining) throw new OfdSignatureVerificationBudgetExceededException("digest byte");
        _digestBytesRemaining -= length;
    }

    // Keep this aligned with the supported identifiers in OfdDigestAlgorithms.
    internal static string NormalizeAlgorithm(string algorithm) => algorithm.Trim().ToLowerInvariant() switch
    {
        "1.2.156.10197.1.401" or "sm3" => "sm3",
        "1.3.14.3.2.26" or "sha-1" or "sha1" => "sha1",
        "2.16.840.1.101.3.4.2.1" or "sha-256" or "sha256" => "sha256",
        var value => value
    };
}

internal sealed class OfdSignatureVerificationBudgetExceededException : Exception
{
    internal OfdSignatureVerificationBudgetExceededException(string limit)
        : base($"Signature verification is incomplete: the {limit} budget was exceeded.") { }
}
