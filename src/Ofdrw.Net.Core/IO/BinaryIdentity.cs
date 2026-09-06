using System;
using System.Security.Cryptography;

namespace Ofdrw.Net.Core.IO;

internal static class BinaryIdentity
{
    internal static string Hash(byte[] data)
    {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", string.Empty).ToLowerInvariant();
    }
}
