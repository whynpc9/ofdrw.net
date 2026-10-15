using System.Security.Cryptography;
using System.Buffers.Binary;
using System.Text;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Digests;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Modes;
using Org.BouncyCastle.Crypto.Paddings;
using Org.BouncyCastle.Crypto.Parameters;

namespace Ofdrw.Net.Crypto.Password.Internal;

internal static class PasswordPrimitives
{
    internal static byte[] Sm3(byte[] input)
    {
        var digest = new SM3Digest();
        digest.BlockUpdate(input, 0, input.Length);
        var output = new byte[32];
        digest.DoFinal(output, 0);
        return output;
    }

    // Fixed RFC 7914 scrypt policy for private profile v2. No untrusted cost
    // parameter reaches the backend. The gate is per loaded assembly instance,
    // not a bound on all loaded archives or arbitrary assembly load contexts.
    internal const long KdfMemoryBudget = 160L * 1024 * 1024;
    internal const long KdfWorkUnits = 131072L * 8;
    internal static readonly SemaphoreSlim KdfGate = new(1, 1);

    internal static byte[] PasswordBytes(string password)
    {
        if (string.IsNullOrEmpty(password) || password.Length > 1024)
            throw new ArgumentException("Password must contain 1 to 1024 UTF-16 code units.", nameof(password));
        // Reject unpaired surrogates; never normalize or trim a password.
        return new UTF8Encoding(false, true).GetBytes(password);
    }

    internal static byte[] PasswordKey(byte[] password, byte[] salt, CancellationToken token)
    {
        if (salt.Length != 32) throw new InvalidDataException("Profile v2 requires a 32-byte salt.");
        token.ThrowIfCancellationRequested();
        var result = SCrypt.Generate(password, salt, 131072, 8, 1, 48);
        try { token.ThrowIfCancellationRequested(); return result; }
        catch { CryptographicOperations.ZeroMemory(result); throw; }
    }

    internal static byte[] Authenticate(Dictionary<string, byte[]> entries, string authPath, byte[] key, CancellationToken token)
    {
        using var mac = IncrementalHash.CreateHMAC(HashAlgorithmName.SHA256, key);
        mac.AppendData("Ofdrw.Net.Password.Profile.v2\0"u8);
        var files = entries.Where(pair => !pair.Key.Equals(authPath, StringComparison.Ordinal)).OrderBy(pair => pair.Key, StringComparer.Ordinal).ToArray();
        AppendLength(files.Length);
        foreach (var entry in files)
        {
            token.ThrowIfCancellationRequested();
            var name = new UTF8Encoding(false, true).GetBytes(entry.Key);
            AppendLength(name.Length); mac.AppendData(name); AppendLength(entry.Value.LongLength);
            for (int offset = 0; offset < entry.Value.Length; offset += 16384)
            {
                token.ThrowIfCancellationRequested();
                mac.AppendData(entry.Value, offset, Math.Min(16384, entry.Value.Length - offset));
            }
        }
        return mac.GetHashAndReset();

        void AppendLength(long value)
        {
            Span<byte> bytes = stackalloc byte[8];
            BinaryPrimitives.WriteInt64BigEndian(bytes, value); mac.AppendData(bytes);
        }
    }

    internal static byte[] Cbc(bool encrypt, byte[] input, byte[] key, byte[] iv, CancellationToken token)
    {
        if (key.Length != 16 || iv.Length != 16) throw new InvalidDataException("SM4 key and IV must be 16 bytes.");
        if (!encrypt && (input.Length == 0 || input.Length % 16 != 0))
            throw new InvalidDataException("SM4 ciphertext must contain complete blocks.");
        var cipher = new PaddedBufferedBlockCipher(new CbcBlockCipher(new SM4Engine()), new Pkcs7Padding());
        cipher.Init(encrypt, new ParametersWithIV(new KeyParameter(key), iv));
        var buffer = new byte[cipher.GetOutputSize(input.Length)];
        try
        {
            int count = 0;
            for (int offset = 0; offset < input.Length; offset += 16384)
            {
                token.ThrowIfCancellationRequested();
                count += cipher.ProcessBytes(input, offset, Math.Min(16384, input.Length - offset), buffer, count);
            }
            token.ThrowIfCancellationRequested();
            count += cipher.DoFinal(buffer, count);
            return buffer.AsSpan(0, count).ToArray();
        }
        catch (InvalidCipherTextException ex)
        {
            throw new InvalidDataException("Password or encrypted data is invalid.", ex);
        }
        finally { CryptographicOperations.ZeroMemory(buffer); cipher.Reset(); }
    }
}
