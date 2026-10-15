using System.Security.Cryptography;
using System.Text;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Digests;
using Org.BouncyCastle.Crypto.Engines;
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

    // GB/T 32918.3 5.4.3, counter starts at 1 and is encoded big endian.
    internal static byte[] Kdf(byte[] input, int length)
    {
        if (length <= 0 || length > 4096) throw new ArgumentOutOfRangeException(nameof(length));
        var result = new byte[length];
        for (int offset = 0, counter = 1; offset < length; offset += 32, counter++)
        {
            var digest = new SM3Digest();
            digest.BlockUpdate(input, 0, input.Length);
            digest.Update((byte)(counter >> 24)); digest.Update((byte)(counter >> 16));
            digest.Update((byte)(counter >> 8)); digest.Update((byte)counter);
            var block = new byte[32];
            digest.DoFinal(block, 0);
            Buffer.BlockCopy(block, 0, result, offset, Math.Min(32, length - offset));
            CryptographicOperations.ZeroMemory(block);
        }
        return result;
    }

    internal static byte[] PasswordKey(string password)
    {
        if (string.IsNullOrEmpty(password) || password.Length > 1024)
            throw new ArgumentException("Password must contain 1 to 1024 UTF-16 code units.", nameof(password));
        // Reject unpaired surrogates; never normalize or trim a password.
        var bytes = new UTF8Encoding(false, true).GetBytes(password);
        try { return Kdf(bytes, 16); }
        finally { CryptographicOperations.ZeroMemory(bytes); }
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
