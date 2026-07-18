using System;
using System.Security.Cryptography;

namespace Ofdrw.Net.Signatures.Crypto;

/// <summary>
/// OFD digest algorithm identifiers and implementations used by reference checks.
/// </summary>
public static class OfdDigestAlgorithms
{
    public const string Sm3Oid = "1.2.156.10197.1.401";
    public const string Sha1Oid = "1.3.14.3.2.26";
    public const string Sha256Oid = "2.16.840.1.101.3.4.2.1";

    public static bool IsSupported(string? algorithm)
    {
        return Normalize(algorithm) is "sm3" or "sha1" or "sha256";
    }

    public static byte[] Compute(string algorithm, byte[] data)
    {
        if (data is null)
        {
            throw new ArgumentNullException(nameof(data));
        }

        return Normalize(algorithm) switch
        {
            "sm3" => Sm3Digest.ComputeHash(data),
            "sha1" => ComputeSha1(data),
            "sha256" => ComputeSha256(data),
            _ => throw new NotSupportedException(
                $"Unsupported OFD digest algorithm: {algorithm}")
        };
    }

    private static string Normalize(string? algorithm)
    {
        var value = algorithm?.Trim().ToLowerInvariant();
        return value switch
        {
            Sm3Oid or "sm3" => "sm3",
            Sha1Oid or "sha-1" or "sha1" => "sha1",
            Sha256Oid or "sha-256" or "sha256" => "sha256",
            _ => value ?? string.Empty
        };
    }

    private static byte[] ComputeSha1(byte[] data)
    {
        using var algorithm = SHA1.Create();
        return algorithm.ComputeHash(data);
    }

    private static byte[] ComputeSha256(byte[] data)
    {
        using var algorithm = SHA256.Create();
        return algorithm.ComputeHash(data);
    }
}
