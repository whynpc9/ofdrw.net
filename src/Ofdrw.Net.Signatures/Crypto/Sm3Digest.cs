using System;

namespace Ofdrw.Net.Signatures.Crypto;

internal static class Sm3Digest
{
    private static readonly uint[] InitialVector =
    {
        0x7380166f, 0x4914b2b9, 0x172442d7, 0xda8a0600,
        0xa96f30bc, 0x163138aa, 0xe38dee4d, 0xb0fb0e4e
    };

    public static byte[] ComputeHash(byte[] input)
    {
        var padded = Pad(input);
        var state = (uint[])InitialVector.Clone();
        var w = new uint[68];
        var wPrime = new uint[64];

        for (var offset = 0; offset < padded.Length; offset += 64)
        {
            for (var index = 0; index < 16; index++)
            {
                var position = offset + index * 4;
                w[index] =
                    ((uint)padded[position] << 24) |
                    ((uint)padded[position + 1] << 16) |
                    ((uint)padded[position + 2] << 8) |
                    padded[position + 3];
            }

            for (var index = 16; index < 68; index++)
            {
                w[index] = P1(
                    w[index - 16] ^
                    w[index - 9] ^
                    RotateLeft(w[index - 3], 15)) ^
                    RotateLeft(w[index - 13], 7) ^
                    w[index - 6];
            }

            for (var index = 0; index < 64; index++)
            {
                wPrime[index] = w[index] ^ w[index + 4];
            }

            Compress(state, w, wPrime);
        }

        var output = new byte[32];
        for (var index = 0; index < state.Length; index++)
        {
            output[index * 4] = (byte)(state[index] >> 24);
            output[index * 4 + 1] = (byte)(state[index] >> 16);
            output[index * 4 + 2] = (byte)(state[index] >> 8);
            output[index * 4 + 3] = (byte)state[index];
        }

        return output;
    }

    private static void Compress(uint[] state, uint[] w, uint[] wPrime)
    {
        var a = state[0];
        var b = state[1];
        var c = state[2];
        var d = state[3];
        var e = state[4];
        var f = state[5];
        var g = state[6];
        var h = state[7];

        for (var index = 0; index < 64; index++)
        {
            var constant = index < 16 ? 0x79cc4519u : 0x7a879d8au;
            var a12 = RotateLeft(a, 12);
            var ss1 = RotateLeft(
                unchecked(a12 + e + RotateLeft(constant, index)),
                7);
            var ss2 = ss1 ^ a12;
            var tt1 = unchecked(
                Ff(a, b, c, index) + d + ss2 + wPrime[index]);
            var tt2 = unchecked(
                Gg(e, f, g, index) + h + ss1 + w[index]);
            d = c;
            c = RotateLeft(b, 9);
            b = a;
            a = tt1;
            h = g;
            g = RotateLeft(f, 19);
            f = e;
            e = P0(tt2);
        }

        state[0] ^= a;
        state[1] ^= b;
        state[2] ^= c;
        state[3] ^= d;
        state[4] ^= e;
        state[5] ^= f;
        state[6] ^= g;
        state[7] ^= h;
    }

    private static byte[] Pad(byte[] input)
    {
        var bitLength = checked((ulong)input.Length * 8);
        var paddedLength = ((input.Length + 9 + 63) / 64) * 64;
        var padded = new byte[paddedLength];
        Buffer.BlockCopy(input, 0, padded, 0, input.Length);
        padded[input.Length] = 0x80;
        for (var index = 0; index < 8; index++)
        {
            padded[padded.Length - 1 - index] =
                (byte)(bitLength >> (index * 8));
        }

        return padded;
    }

    private static uint RotateLeft(uint value, int bits)
    {
        bits &= 31;
        return bits == 0 ? value : (value << bits) | (value >> (32 - bits));
    }

    private static uint P0(uint value)
    {
        return value ^ RotateLeft(value, 9) ^ RotateLeft(value, 17);
    }

    private static uint P1(uint value)
    {
        return value ^ RotateLeft(value, 15) ^ RotateLeft(value, 23);
    }

    private static uint Ff(uint x, uint y, uint z, int round)
    {
        return round < 16 ? x ^ y ^ z : (x & y) | (x & z) | (y & z);
    }

    private static uint Gg(uint x, uint y, uint z, int round)
    {
        return round < 16 ? x ^ y ^ z : (x & y) | (~x & z);
    }
}
