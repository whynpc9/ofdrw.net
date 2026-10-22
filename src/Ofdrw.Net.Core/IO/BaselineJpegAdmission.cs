using System;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.IO;

// A deliberately narrow baseline JPEG syntax verifier: single interleaved scan,
// 8-bit Huffman DCT, <=4 components, bounded dimensions. Entropy is walked (not
// merely SOI/EOI checked); unsupported coding modes are rejected explicitly.
internal static class BaselineJpegAdmission
{
    internal static void Validate(byte[] data, CancellationToken token, Action<long>? reservePixels = null)
    {
        if (data.Length < 4 || data[0] != 255 || data[1] != 216) Bad();
        var dc = new Table?[4]; var ac = new Table?[4]; var quant = new bool[4];
        var components = new Component[4]; var count = 0; var width = 0; var height = 0; var maxH = 0; var maxV = 0; var restart = 0;
        var p = 2;
        while (p < data.Length)
        {
            token.ThrowIfCancellationRequested(); if (data[p++] != 255) Bad();
            while (p < data.Length && data[p] == 255) p++;
            if (p >= data.Length) Bad(); var marker = data[p++];
            if (data.Length - p < 2) Bad(); var length = data[p] * 256 + data[p + 1];
            if (length < 2 || length > data.Length - p) Bad(); var end = p + length; p += 2;
            switch (marker)
            {
                case 0xdb:
                    while (p < end)
                    {
                        var id = data[p++]; if (id > 3 || end - p < 64) Bad();
                        for (var i = 0; i < 64; i++) if (data[p++] == 0) Bad(); quant[id] = true;
                    }
                    break;
                case 0xc4:
                    while (p < end)
                    {
                        var id = data[p++]; if ((id & 15) > 3 || id >> 4 > 1 || end - p < 16) Bad();
                        var counts = new int[17]; var total = 0;
                        for (var bits = 1; bits <= 16; bits++) { counts[bits] = data[p++]; total += counts[bits]; }
                        if (total == 0 || total > 256 || total > end - p) Bad();
                        var values = new byte[total]; Buffer.BlockCopy(data, p, values, 0, total); p += total;
                        var table = new Table(counts, values);
                        if (id >> 4 == 0) dc[id & 15] = table; else ac[id & 15] = table;
                    }
                    break;
                case 0xc0:
                    if (count != 0 || end - p < 6 || data[p++] != 8) Bad();
                    height = data[p] * 256 + data[p + 1]; width = data[p + 2] * 256 + data[p + 3]; p += 4; count = data[p++];
                    if (count < 1 || count > 4 || end - p != 3 * count || width == 0 || height == 0 || (long)width * height > 100_000_000) Bad();
                    reservePixels?.Invoke(checked((long)width * height)); token.ThrowIfCancellationRequested();
                    for (var i = 0; i < count; i++)
                    {
                        var id = data[p++]; var sampling = data[p++]; var q = data[p++]; var h = sampling >> 4; var v = sampling & 15;
                        if (h < 1 || h > 4 || v < 1 || v > 4 || q > 3) Bad();
                        for (var j = 0; j < i; j++) if (components[j].Id == id) Bad();
                        components[i] = new Component { Id = id, H = h, V = v, Quant = q };
                        maxH = Math.Max(maxH, h); maxV = Math.Max(maxV, v);
                    }
                    break;
                case 0xdd:
                    if (end - p != 2) Bad(); restart = data[p] * 256 + data[p + 1]; p += 2; break;
                case 0xda:
                    if (count == 0 || end - p != 1 + 2 * count + 3 || data[p++] != count) Unsupported();
                    var order = new Component[count]; var seen = new bool[count];
                    for (var i = 0; i < count; i++)
                    {
                        var id = data[p++]; var tables = data[p++]; var found = -1;
                        for (var j = 0; j < count; j++) if (components[j].Id == id) found = j;
                        if (found < 0 || seen[found] || tables >> 4 > 3 || (tables & 15) > 3) Bad();
                        seen[found] = true; var component = components[found]; component.Dc = dc[tables >> 4]; component.Ac = ac[tables & 15];
                        if (component.Dc is null || component.Ac is null || !quant[component.Quant]) Bad(); order[i] = component;
                    }
                    if (data[p++] != 0 || data[p++] != 63 || data[p++] != 0) Unsupported();
                    var entropy = new Bits(data, p); var across = (width + 8 * maxH - 1) / (8 * maxH); var down = (height + 8 * maxV - 1) / (8 * maxV);
                    // A single component is non-interleaved: one block per MCU.
                    if (count == 1) { across = (width + 7) / 8; down = (height + 7) / 8; order[0].H = order[0].V = 1; }
                    var mcus = checked(across * down); var nextRestart = 0;
                    for (var mcu = 0; mcu < mcus; mcu++)
                    {
                        token.ThrowIfCancellationRequested();
                        if (mcu > 0 && restart != 0 && mcu % restart == 0) entropy.Marker(0xd0 + nextRestart++ % 8);
                        foreach (var component in order)
                            for (var block = 0; block < component.H * component.V; block++)
                            {
                                var category = component.Dc!.Decode(entropy); if (category > 11) Bad(); entropy.Skip(category);
                                for (var coefficient = 1; coefficient < 64;)
                                {
                                    var symbol = component.Ac!.Decode(entropy); var size = symbol & 15; var zeros = symbol >> 4;
                                    if (size == 0) { if (zeros == 0) break; if (zeros != 15) Bad(); coefficient += 16; if (coefficient > 64) Bad(); }
                                    else { if (size > 10) Bad(); coefficient += zeros; if (coefficient >= 64) Bad(); entropy.Skip(size); coefficient++; }
                                }
                            }
                    }
                    entropy.Marker(0xd9); if (entropy.Position != data.Length) Bad(); return;
                case 0xe0: case 0xe1: case 0xe2: case 0xee: p = end; break;
                default: Unsupported(); break;
            }
            if (p != end) Bad();
        }
        Bad();
    }
    private sealed class Component
    { internal int Id, H, V, Quant; internal Table? Dc, Ac; }
    private sealed class Table
    {
        private readonly int[] _minimum = new int[17], _maximum = new int[17], _offset = new int[17]; private readonly byte[] _values;
        internal Table(int[] counts, byte[] values)
        {
            _values = values; var code = 0; var offset = 0;
            for (var length = 1; length <= 16; length++)
            { _minimum[length] = code; _maximum[length] = code + counts[length] - 1; _offset[length] = offset; code += counts[length]; if (code >= 1 << length) Bad(); code <<= 1; offset += counts[length]; }
        }
        internal int Decode(Bits bits)
        { var code = 0; for (var length = 1; length <= 16; length++) { code = code * 2 + bits.Read(); if (code >= _minimum[length] && code <= _maximum[length]) return _values[_offset[length] + code - _minimum[length]]; } Bad(); return 0; }
    }
    private sealed class Bits(byte[] data, int position)
    {
        private int _position = position, _value, _remaining;
        internal int Position => _position;
        internal int Read()
        {
            if (_remaining == 0)
            {
                if (_position >= data.Length) Bad(); _value = data[_position++];
                if (_value == 255 && (_position >= data.Length || data[_position++] != 0)) Bad(); _remaining = 8;
            }
            return (_value >> --_remaining) & 1;
        }
        internal void Skip(int count) { for (var i = 0; i < count; i++) Read(); }
        internal void Marker(int marker)
        {
            if (_remaining > 0 && (_value & ((1 << _remaining) - 1)) != (1 << _remaining) - 1) Bad(); _remaining = 0;
            if (_position >= data.Length || data[_position++] != 255) Bad(); while (_position < data.Length && data[_position] == 255) _position++;
            if (_position >= data.Length || data[_position++] != marker) Bad();
        }
    }
    private static void Unsupported() => throw new NotSupportedException("HTML preview supports only a single baseline Huffman JPEG scan.");
    private static void Bad() => throw new InvalidDataException("Invalid baseline JPEG payload.");
}
