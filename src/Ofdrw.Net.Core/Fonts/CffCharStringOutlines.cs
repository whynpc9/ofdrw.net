using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>An original-GID outline source for one selected face, independent of its outline format.</summary>
internal interface IGlyphOutlineSource
{
    int GlyphCount { get; }
    int UnitsPerEm { get; }
    TrueTypeGlyphOutline Read(int glyphId, CancellationToken cancellationToken = default);
}

internal static class GlyphOutlineSources
{
    private const int MaximumFaceBytes = 32 * 1024 * 1024;

    /// <summary>Opens the TrueType (glyf) or OpenType/CFF1 outline reader matching the selected face.</summary>
    internal static IGlyphOutlineSource Open(byte[] sourceBytes, int faceIndex = 0)
    {
        if (sourceBytes is null) throw new ArgumentNullException(nameof(sourceBytes));
        if (sourceBytes.Length < 12) throw new NotSupportedException("GID outlines require a bounded font source.");
        var selected = OpenTypeCollection.SelectFace((byte[])sourceBytes.Clone(), faceIndex, MaximumFaceBytes);
        return OpenTypeFace.U32(selected, 0) == 0x4F54544F
            ? new CffCharStringOutlines(sourceBytes, faceIndex)
            : new TrueTypeGlyphOutlines(sourceBytes, faceIndex);
    }

    internal static bool IsCff(byte[] faceBytes) => faceBytes.Length >= 4 && OpenTypeFace.U32(faceBytes, 0) == 0x4F54544F;
}

/// <summary>A parsed standalone CFF1 font program (plain or CID-keyed) with Type 2 CharStrings.</summary>
internal sealed class CffFontProgram
{
    private readonly byte[] _cff;
    private readonly byte[]? _fdSelect; // FD index per GID for CID-keyed fonts
    private readonly CffIndex[] _localSubrs;

    private CffFontProgram(byte[] cff, CffIndex charStrings, CffIndex globalSubrs, CffIndex[] localSubrs,
        byte[]? fdSelect, double scale, bool cid)
    {
        _cff = cff; CharStrings = charStrings; GlobalSubrs = globalSubrs; _localSubrs = localSubrs;
        _fdSelect = fdSelect; Scale = scale; IsCidKeyed = cid;
    }

    internal CffIndex CharStrings { get; }
    internal CffIndex GlobalSubrs { get; }
    internal int GlyphCount => CharStrings.Count;
    /// <summary>Factor from charstring units to font design units (FontMatrix × unitsPerEm).</summary>
    internal double Scale { get; }
    internal bool IsCidKeyed { get; }
    internal byte[] Bytes => _cff;
    internal CffIndex LocalSubrs(int glyphId) => _localSubrs[FdIndex(glyphId)];
    internal int FdIndex(int glyphId) => _fdSelect is null ? 0 : _fdSelect[glyphId];
    internal int FdCount => _localSubrs.Length;
    internal CffIndex LocalSubrsOfFd(int fd) => _localSubrs[fd];

    internal static CffFontProgram Parse(OpenTypeFace face)
    {
        var cff = face.Table("CFF ", 4);
        if (cff[0] != 1 || cff[2] < 4 || cff[2] > cff.Length || cff[3] is < 1 or > 4)
            throw new NotSupportedException("The CFF table is not a CFF1 font program.");
        var names = CffIndex.Read(cff, cff[2]);
        var top = CffIndex.Read(cff, names.End);
        if (names.Count != 1 || top.Count != 1)
            throw new NotSupportedException("The CFF table must contain exactly one font.");
        var strings = CffIndex.Read(cff, top.End);
        var globalSubrs = CffIndex.Read(cff, strings.End);
        var topDict = CffDict.Parse(cff, top.ItemStart(0), top.ItemEnd(0));
        if (topDict.TryGetValue(CffDict.CharstringType, out var type) && (type.Length != 1 || type[0] != 2))
            throw new NotSupportedException("Only Type 2 CharStrings are supported.");
        if (topDict.ContainsKey(CffDict.SyntheticBase))
            throw new NotSupportedException("Synthetic CFF fonts are not supported.");
        var unitsPerEm = OpenTypeFace.U16(face.Table("head", 54), 18);
        var matrix = topDict.TryGetValue(CffDict.FontMatrix, out var values) ? values : new[] { 0.001, 0, 0, 0.001, 0, 0 };
        if (matrix.Length != 6 || matrix[1] != 0 || matrix[2] != 0 || matrix[4] != 0 || matrix[5] != 0 ||
            matrix[0] <= 0 || Math.Abs(matrix[0] - matrix[3]) > 1e-12)
            throw new NotSupportedException("Only a uniform CFF FontMatrix is supported.");
        var scale = matrix[0] * unitsPerEm;
        if (scale <= 0 || scale > 64) throw new NotSupportedException("The CFF FontMatrix does not match unitsPerEm.");
        var charStrings = CffIndex.Read(cff, Offset(topDict, CffDict.CharStrings, cff));
        if (charStrings.Count == 0) throw new InvalidDataException("The CFF font has no CharStrings.");
        CffIndex[] localSubrs;
        byte[]? fdSelect = null;
        var cid = topDict.ContainsKey(CffDict.Ros);
        if (cid)
        {
            var fdArray = CffIndex.Read(cff, Offset(topDict, CffDict.FdArray, cff));
            if (fdArray.Count is < 1 or > 256) throw new InvalidDataException("Invalid CFF FDArray.");
            localSubrs = new CffIndex[fdArray.Count];
            for (var index = 0; index < fdArray.Count; index++)
            {
                var fd = CffDict.Parse(cff, fdArray.ItemStart(index), fdArray.ItemEnd(index));
                if (fd.TryGetValue(CffDict.FontMatrix, out var fdMatrix) &&
                    !(fdMatrix.Length == 6 && fdMatrix[0] == fdMatrix[3] && fdMatrix[1] == 0 && fdMatrix[2] == 0 &&
                      fdMatrix[4] == 0 && fdMatrix[5] == 0 && Math.Abs(fdMatrix[0] * 1000 - 1) < 1e-9))
                    throw new NotSupportedException("A non-default CID FD FontMatrix is not supported.");
                localSubrs[index] = PrivateSubrs(cff, fd);
            }
            fdSelect = ReadFdSelect(cff, Offset(topDict, CffDict.FdSelect, cff), charStrings.Count, fdArray.Count);
        }
        else localSubrs = new[] { PrivateSubrs(cff, topDict) };
        return new CffFontProgram(cff, charStrings, globalSubrs, localSubrs, fdSelect, scale, cid);
    }

    private static int Offset(Dictionary<int, double[]> dict, int key, byte[] cff)
    {
        if (!dict.TryGetValue(key, out var values) || values.Length != 1 || values[0] < 0 || values[0] >= cff.Length ||
            values[0] != Math.Floor(values[0]))
            throw new InvalidDataException("A required CFF DICT offset is missing or invalid.");
        return (int)values[0];
    }

    private static CffIndex PrivateSubrs(byte[] cff, Dictionary<int, double[]> dict)
    {
        if (!dict.TryGetValue(CffDict.Private, out var range)) return CffIndex.Empty;
        if (range.Length != 2 || range[0] < 0 || range[1] < 0 || range[1] + range[0] > cff.Length ||
            range[0] != Math.Floor(range[0]) || range[1] != Math.Floor(range[1]))
            throw new InvalidDataException("Invalid CFF Private DICT range.");
        var start = (int)range[1];
        var privateDict = CffDict.Parse(cff, start, start + (int)range[0]);
        if (!privateDict.TryGetValue(CffDict.Subrs, out var subrs)) return CffIndex.Empty;
        if (subrs.Length != 1 || subrs[0] <= 0 || start + subrs[0] >= cff.Length)
            throw new InvalidDataException("Invalid CFF local Subrs offset.");
        return CffIndex.Read(cff, start + (int)subrs[0]);
    }

    private static byte[] ReadFdSelect(byte[] cff, int offset, int glyphCount, int fdCount)
    {
        OpenTypeFace.Require(cff, offset, 1);
        var result = new byte[glyphCount];
        switch (cff[offset])
        {
            case 0:
                OpenTypeFace.Require(cff, offset + 1, glyphCount);
                Array.Copy(cff, offset + 1, result, 0, glyphCount);
                break;
            case 3:
                var ranges = OpenTypeFace.U16(cff, offset + 1);
                OpenTypeFace.Require(cff, offset + 3, ranges * 3 + 2);
                if (ranges == 0) throw new InvalidDataException("Empty CFF FDSelect.");
                for (var index = 0; index < ranges; index++)
                {
                    var first = OpenTypeFace.U16(cff, offset + 3 + index * 3);
                    var fd = cff[offset + 5 + index * 3];
                    var next = OpenTypeFace.U16(cff, offset + 6 + index * 3);
                    if (index == 0 && first != 0 || next <= first || next > glyphCount)
                        throw new InvalidDataException("Invalid CFF FDSelect range.");
                    for (var gid = first; gid < next; gid++) result[gid] = fd;
                }
                if (OpenTypeFace.U16(cff, offset + 3 + ranges * 3) != glyphCount)
                    throw new InvalidDataException("CFF FDSelect does not cover every glyph.");
                break;
            default:
                throw new NotSupportedException("Unsupported CFF FDSelect format.");
        }
        if (result.Any(fd => fd >= fdCount)) throw new InvalidDataException("CFF FDSelect selects a missing FD.");
        return result;
    }
}

/// <summary>A CFF INDEX: count, offsets and the absolute byte range of each item.</summary>
internal readonly struct CffIndex
{
    private readonly int[]? _offsets;
    private readonly int _dataStart;

    private CffIndex(int start, int dataStart, int end, int[] offsets)
    { Start = start; _dataStart = dataStart; End = end; _offsets = offsets; }

    internal static CffIndex Empty => new(0, 0, 0, Array.Empty<int>());
    internal int Start { get; }
    internal int Count => _offsets is null || _offsets.Length == 0 ? 0 : _offsets.Length - 1;
    internal int End { get; }
    internal int ItemStart(int index) => checked(_dataStart + _offsets![index] - 1);
    internal int ItemEnd(int index) => checked(_dataStart + _offsets![index + 1] - 1);
    internal int ItemLength(int index) => ItemEnd(index) - ItemStart(index);

    internal static CffIndex Read(byte[] cff, int start)
    {
        OpenTypeFace.Require(cff, start, 2);
        var count = OpenTypeFace.U16(cff, start);
        if (count == 0) return new CffIndex(start, start + 2, start + 2, Array.Empty<int>());
        OpenTypeFace.Require(cff, start + 2, 1);
        var offSize = cff[start + 2];
        if (offSize is < 1 or > 4) throw new InvalidDataException("Invalid CFF INDEX offset size.");
        var offsetsStart = checked(start + 3);
        OpenTypeFace.Require(cff, offsetsStart, checked((count + 1) * offSize));
        var dataStart = checked(offsetsStart + (count + 1) * offSize);
        var offsets = new int[count + 1];
        for (var i = 0; i <= count; i++)
        {
            uint value = 0;
            for (var j = 0; j < offSize; j++) value = (value << 8) | cff[offsetsStart + i * offSize + j];
            if (value == 0 || value > int.MaxValue || i == 0 && value != 1 || i > 0 && value < offsets[i - 1])
                throw new InvalidDataException("Invalid CFF INDEX offsets.");
            offsets[i] = (int)value;
        }
        var end = checked(dataStart + offsets[count] - 1);
        OpenTypeFace.Require(cff, dataStart, end - dataStart);
        return new CffIndex(start, dataStart, end, offsets);
    }
}

/// <summary>CFF DICT parsing into operator → operands; escaped operators are 1200 + the second byte.</summary>
internal static class CffDict
{
    internal const int CharStrings = 17, Private = 18, Subrs = 19, CharstringType = 1206, FontMatrix = 1207,
        SyntheticBase = 1220, Ros = 1230, FdArray = 1236, FdSelect = 1237, Charset = 15, Encoding = 16;

    internal static Dictionary<int, double[]> Parse(byte[] data, int start, int end)
    {
        if (start < 0 || end > data.Length || end < start) throw new InvalidDataException("Invalid CFF DICT range.");
        var result = new Dictionary<int, double[]>();
        var operands = new List<double>();
        for (var cursor = start; cursor < end;)
        {
            var value = data[cursor++];
            if (value <= 21)
            {
                var op = (int)value;
                if (value == 12)
                {
                    if (cursor >= end) throw new InvalidDataException("Truncated CFF escaped DICT operator.");
                    op = 1200 + data[cursor++];
                }
                result[op] = operands.ToArray();
                operands.Clear();
                continue;
            }
            if (operands.Count >= 48) throw new InvalidDataException("CFF DICT operand stack exceeded.");
            if (value == 28 || value == 29)
            {
                var bytes = value == 28 ? 2 : 4;
                if (cursor + bytes > end) throw new InvalidDataException("Truncated CFF DICT operand.");
                operands.Add(value == 28 ? unchecked((short)OpenTypeFace.U16(data, cursor)) : unchecked((int)OpenTypeFace.U32(data, cursor)));
                cursor += bytes;
            }
            else if (value == 30) operands.Add(ReadReal(data, ref cursor, end));
            else if (value >= 32 && value <= 246) operands.Add(value - 139);
            else if (value >= 247 && value <= 254)
            {
                if (cursor >= end) throw new InvalidDataException("Truncated CFF DICT operand.");
                var tail = data[cursor++];
                operands.Add(value <= 250 ? (value - 247) * 256 + tail + 108 : -(value - 251) * 256 - tail - 108);
            }
            else throw new InvalidDataException("Invalid CFF DICT byte.");
        }
        return result;
    }

    private static double ReadReal(byte[] data, ref int cursor, int end)
    {
        var text = new System.Text.StringBuilder();
        for (var bytes = 0; bytes < 32; bytes++)
        {
            if (cursor >= end) throw new InvalidDataException("Unterminated CFF real operand.");
            var pair = data[cursor++];
            foreach (var nibble in new[] { pair >> 4, pair & 15 })
            {
                switch (nibble)
                {
                    case <= 9: text.Append((char)('0' + nibble)); break;
                    case 0xa: text.Append('.'); break;
                    case 0xb: text.Append('E'); break;
                    case 0xc: text.Append("E-"); break;
                    case 0xe: text.Append('-'); break;
                    case 0xf:
                        return double.TryParse(text.ToString(), System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var number) &&
                            !double.IsNaN(number) && !double.IsInfinity(number) ? number :
                            throw new InvalidDataException("Invalid CFF real operand.");
                    default: throw new InvalidDataException("Reserved CFF real nibble.");
                }
            }
        }
        throw new InvalidDataException("Unterminated CFF real operand.");
    }
}

/// <summary>Managed Type 2 CharString interpreter producing original-GID cubic outlines in design units.</summary>
/// <remarks>Hints are parsed and ignored. seac-style accented endchar, CFF2 blending and the
/// random operator are rejected. Budgets bound subroutine depth, operators and points.</remarks>
internal sealed class CffCharStringOutlines : IGlyphOutlineSource
{
    private const int MaximumFaceBytes = 32 * 1024 * 1024;
    private const int MaximumSourceBytes = 64 * 1024 * 1024;
    private const int MaximumDepth = 10;
    private const int MaximumOperations = 200_000;
    private const int MaximumPoints = 32768;
    private static readonly string[] UnsupportedTables =
    [
        "CFF2", "glyf", "loca", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "sbix"
    ];
    private readonly CffFontProgram _program;
    private readonly byte[] _hmtx;
    private readonly int _metricCount;

    internal CffCharStringOutlines(byte[] sourceBytes, int faceIndex = 0)
    {
        if (sourceBytes is null) throw new ArgumentNullException(nameof(sourceBytes));
        if (sourceBytes.Length is < 12 or > MaximumSourceBytes)
            throw new NotSupportedException("CFF GID outlines require a bounded font source.");
        var ownedSource = (byte[])sourceBytes.Clone();
        var selected = OpenTypeCollection.SelectFace(ownedSource, faceIndex, MaximumFaceBytes);
        var face = new OpenTypeFace(selected);
        if (face.Version != 0x4F54544F || !face.Tables.ContainsKey("CFF ") || UnsupportedTables.Any(face.Tables.ContainsKey))
            throw new NotSupportedException("CFF GID outlines require a static monochrome OpenType/CFF1 face.");
        UnitsPerEm = OpenTypeFace.U16(face.Table("head", 54), 18);
        GlyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        if (GlyphCount == 0 || UnitsPerEm is < 16 or > 16384)
            throw new InvalidDataException("Invalid CFF face metrics.");
        _program = CffFontProgram.Parse(face);
        if (_program.GlyphCount != GlyphCount)
            throw new InvalidDataException("CFF CharStrings count differs from maxp.");
        _metricCount = OpenTypeFace.U16(face.Table("hhea", 36), 34);
        if (_metricCount == 0 || _metricCount > GlyphCount)
            throw new InvalidDataException("Invalid CFF horizontal metric count.");
        _hmtx = face.Table("hmtx", checked(_metricCount * 4 + (GlyphCount - _metricCount) * 2));
        SourceSha256 = BinaryIdentity.Hash(ownedSource);
        SelectedFaceSha256 = BinaryIdentity.Hash(selected);
        FaceIndex = faceIndex;
    }

    internal string SourceSha256 { get; }
    internal string SelectedFaceSha256 { get; }
    internal int FaceIndex { get; }
    public int GlyphCount { get; }
    public int UnitsPerEm { get; }

    internal CffFontProgram Program => _program;

    /// <summary>Interprets one glyph (including .notdef) and records the global and local subroutine
    /// indexes it calls; used by GID-preserving subsetting.</summary>
    internal void CollectSubroutines(int glyphId, HashSet<int> global, HashSet<int> local, CancellationToken cancellationToken)
    {
        if (glyphId < 0 || glyphId >= GlyphCount) throw new InvalidDataException("GID outside the CFF font.");
        var interpreter = new Interpreter(_program, _program.LocalSubrs(glyphId), cancellationToken, global, local);
        interpreter.Run(_program.CharStrings.ItemStart(glyphId), _program.CharStrings.ItemEnd(glyphId));
        _ = interpreter.Finish(1);
    }

    public TrueTypeGlyphOutline Read(int glyphId, CancellationToken cancellationToken = default)
    {
        if (glyphId <= 0 || glyphId >= GlyphCount)
            throw new NotSupportedException("The requested original GID is outside the admitted font.");
        var interpreter = new Interpreter(_program, _program.LocalSubrs(glyphId), cancellationToken);
        var charStrings = _program.CharStrings;
        interpreter.Run(charStrings.ItemStart(glyphId), charStrings.ItemEnd(glyphId));
        var contours = interpreter.Finish(_program.Scale);
        var advance = OpenTypeFace.U16(_hmtx, Math.Min(glyphId, _metricCount - 1) * 4);
        return new TrueTypeGlyphOutline(SourceSha256, SelectedFaceSha256, FaceIndex, glyphId, UnitsPerEm, advance,
            contours, cubic: true);
    }

    private sealed class Interpreter
    {
        private readonly byte[] _data;
        private readonly CffIndex _global, _local;
        private readonly int _globalBias, _localBias;
        private readonly CancellationToken _token;
        private readonly double[] _stack = new double[48];
        private readonly double[] _transient = new double[32];
        private readonly List<List<TrueTypeOutlinePoint>> _contours = new();
        private List<TrueTypeOutlinePoint>? _current;
        private int _sp, _stems, _operations, _points;
        private bool _widthSeen, _ended;
        private double _x, _y;
        private readonly HashSet<int>? _globalUsed, _localUsed;

        internal Interpreter(CffFontProgram program, CffIndex local, CancellationToken token,
            HashSet<int>? globalUsed = null, HashSet<int>? localUsed = null)
        {
            _data = program.Bytes; _global = program.GlobalSubrs; _local = local; _token = token;
            _globalBias = Bias(_global.Count); _localBias = Bias(_local.Count);
            _globalUsed = globalUsed; _localUsed = localUsed;
        }

        private static int Bias(int count) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768;

        internal IReadOnlyList<TrueTypeGlyphContour> Finish(double scale)
        {
            if (!_ended) throw new InvalidDataException("A CFF CharString ended without endchar.");
            return new ReadOnlyCollection<TrueTypeGlyphContour>(_contours.Select(points =>
                new TrueTypeGlyphContour(new ReadOnlyCollection<TrueTypeOutlinePoint>(scale == 1 ? points.ToArray() :
                    points.Select(point => new TrueTypeOutlinePoint(point.X * scale, point.Y * scale, point.OnCurve)).ToArray())))
                .ToArray());
        }

        internal void Run(int start, int end) => Execute(start, end, 0);

        private void Push(double value)
        {
            if (_sp >= _stack.Length) throw new InvalidDataException("CFF CharString operand stack overflow.");
            _stack[_sp++] = value;
        }

        private double Pop() => _sp > 0 ? _stack[--_sp] : throw new InvalidDataException("CFF CharString operand stack underflow.");

        private void Require(int count)
        {
            if (_sp < count) throw new InvalidDataException("CFF CharString operator lacks operands.");
        }

        // The first stack-clearing operator may carry the advance width as an extra leading operand.
        private int Width(bool odd)
        {
            var first = 0;
            if (!_widthSeen && odd) first = 1;
            _widthSeen = true;
            return first;
        }

        private void Execute(int start, int end, int depth)
        {
            if (depth > MaximumDepth) throw new InvalidDataException("CFF subroutine nesting exceeded its budget.");
            var cursor = start;
            while (cursor < end)
            {
                if (++_operations > MaximumOperations) throw new InvalidDataException("CFF CharString exceeded its operation budget.");
                if ((_operations & 1023) == 0) _token.ThrowIfCancellationRequested();
                var b0 = _data[cursor++];
                if (b0 >= 32)
                {
                    if (b0 <= 246) Push(b0 - 139);
                    else if (b0 <= 250) { Need(cursor, end, 1); Push((b0 - 247) * 256 + _data[cursor++] + 108); }
                    else if (b0 <= 254) { Need(cursor, end, 1); Push(-(b0 - 251) * 256 - _data[cursor++] - 108); }
                    else
                    {
                        Need(cursor, end, 4);
                        Push(unchecked((int)OpenTypeFace.U32(_data, cursor)) / 65536.0);
                        cursor += 4;
                    }
                    continue;
                }
                if (b0 == 28)
                {
                    Need(cursor, end, 2);
                    Push(unchecked((short)OpenTypeFace.U16(_data, cursor)));
                    cursor += 2;
                    continue;
                }
                switch (b0)
                {
                    case 1: case 3: case 18: case 23: // hstem vstem hstemhm vstemhm
                        _stems += (_sp - Width(_sp % 2 == 1)) / 2;
                        _sp = 0;
                        break;
                    case 19: case 20: // hintmask cntrmask (operands are an implied vstem list)
                        _stems += (_sp - Width(_sp % 2 == 1)) / 2;
                        _sp = 0;
                        var maskBytes = (_stems + 7) / 8;
                        Need(cursor, end, maskBytes);
                        cursor += maskBytes;
                        break;
                    case 21: // rmoveto
                    {
                        var first = Width(_sp > 2);
                        Require(first + 2);
                        MoveTo(_x + _stack[first], _y + _stack[first + 1]);
                        _sp = 0;
                        break;
                    }
                    case 22: // hmoveto
                    {
                        var first = Width(_sp > 1);
                        Require(first + 1);
                        MoveTo(_x + _stack[first], _y);
                        _sp = 0;
                        break;
                    }
                    case 4: // vmoveto
                    {
                        var first = Width(_sp > 1);
                        Require(first + 1);
                        MoveTo(_x, _y + _stack[first]);
                        _sp = 0;
                        break;
                    }
                    case 5: // rlineto
                        for (var i = 0; i + 1 < _sp; i += 2) LineTo(_x + _stack[i], _y + _stack[i + 1]);
                        _sp = 0;
                        break;
                    case 6: case 7: // hlineto vlineto
                    {
                        var horizontal = b0 == 6;
                        for (var i = 0; i < _sp; i++, horizontal = !horizontal)
                            if (horizontal) LineTo(_x + _stack[i], _y); else LineTo(_x, _y + _stack[i]);
                        _sp = 0;
                        break;
                    }
                    case 8: // rrcurveto
                        for (var i = 0; i + 5 < _sp; i += 6)
                            CurveRelative(_stack[i], _stack[i + 1], _stack[i + 2], _stack[i + 3], _stack[i + 4], _stack[i + 5]);
                        _sp = 0;
                        break;
                    case 27: // hhcurveto
                    {
                        var i = 0;
                        double dy1 = 0;
                        if (_sp % 2 == 1) dy1 = _stack[i++];
                        for (; i + 3 < _sp; i += 4)
                        {
                            CurveRelative(_stack[i], dy1, _stack[i + 1], _stack[i + 2], _stack[i + 3], 0);
                            dy1 = 0;
                        }
                        _sp = 0;
                        break;
                    }
                    case 26: // vvcurveto
                    {
                        var i = 0;
                        double dx1 = 0;
                        if (_sp % 2 == 1) dx1 = _stack[i++];
                        for (; i + 3 < _sp; i += 4)
                        {
                            CurveRelative(dx1, _stack[i], _stack[i + 1], _stack[i + 2], 0, _stack[i + 3]);
                            dx1 = 0;
                        }
                        _sp = 0;
                        break;
                    }
                    case 30: case 31: // vhcurveto hvcurveto
                    {
                        var horizontal = b0 == 31;
                        for (var i = 0; i + 3 < _sp; i += 4, horizontal = !horizontal)
                        {
                            var last = _sp - i == 5 ? _stack[i + 4] : 0;
                            if (horizontal) CurveRelative(_stack[i], 0, _stack[i + 1], _stack[i + 2], last, _stack[i + 3]);
                            else CurveRelative(0, _stack[i], _stack[i + 1], _stack[i + 2], _stack[i + 3], last);
                        }
                        _sp = 0;
                        break;
                    }
                    case 24: // rcurveline
                    {
                        var i = 0;
                        for (; i + 7 < _sp; i += 6)
                            CurveRelative(_stack[i], _stack[i + 1], _stack[i + 2], _stack[i + 3], _stack[i + 4], _stack[i + 5]);
                        if (i + 1 < _sp) LineTo(_x + _stack[i], _y + _stack[i + 1]);
                        _sp = 0;
                        break;
                    }
                    case 25: // rlinecurve
                    {
                        var i = 0;
                        for (; i + 7 < _sp; i += 2) LineTo(_x + _stack[i], _y + _stack[i + 1]);
                        if (i + 5 < _sp)
                            CurveRelative(_stack[i], _stack[i + 1], _stack[i + 2], _stack[i + 3], _stack[i + 4], _stack[i + 5]);
                        _sp = 0;
                        break;
                    }
                    case 10: case 29: // callsubr callgsubr
                    {
                        var subrs = b0 == 10 ? _local : _global;
                        var index = (int)Pop() + (b0 == 10 ? _localBias : _globalBias);
                        if (index < 0 || index >= subrs.Count) throw new InvalidDataException("CFF subroutine index out of range.");
                        (b0 == 10 ? _localUsed : _globalUsed)?.Add(index);
                        Execute(subrs.ItemStart(index), subrs.ItemEnd(index), depth + 1);
                        if (_ended) return;
                        break;
                    }
                    case 11: // return
                        return;
                    case 14: // endchar
                        if (_sp - Width(_sp == 1 || _sp == 5) >= 4)
                            throw new NotSupportedException("Accented (seac) CFF endchar is outside the outline profile.");
                        CloseContour();
                        _sp = 0;
                        _ended = true;
                        return;
                    case 12:
                        Need(cursor, end, 1);
                        Escape(_data[cursor++]);
                        break;
                    default:
                        throw new InvalidDataException("Reserved CFF CharString operator.");
                }
            }
        }

        private void Escape(byte op)
        {
            switch (op)
            {
                case 35: // flex
                    Require(13);
                    CurveRelative(_stack[0], _stack[1], _stack[2], _stack[3], _stack[4], _stack[5]);
                    CurveRelative(_stack[6], _stack[7], _stack[8], _stack[9], _stack[10], _stack[11]);
                    _sp = 0;
                    break;
                case 34: // hflex
                    Require(7);
                    CurveRelative(_stack[0], 0, _stack[1], _stack[2], _stack[3], 0);
                    CurveRelative(_stack[4], 0, _stack[5], -_stack[2], _stack[6], 0);
                    _sp = 0;
                    break;
                case 36: // hflex1
                    Require(9);
                    CurveRelative(_stack[0], _stack[1], _stack[2], _stack[3], _stack[4], 0);
                    CurveRelative(_stack[5], 0, _stack[6], _stack[7], _stack[8], -(_stack[1] + _stack[3] + _stack[7]));
                    _sp = 0;
                    break;
                case 37: // flex1
                {
                    Require(11);
                    double dx = 0, dy = 0;
                    for (var i = 0; i < 10; i += 2) { dx += _stack[i]; dy += _stack[i + 1]; }
                    CurveRelative(_stack[0], _stack[1], _stack[2], _stack[3], _stack[4], _stack[5]);
                    if (Math.Abs(dx) > Math.Abs(dy))
                        CurveRelative(_stack[6], _stack[7], _stack[8], _stack[9], _stack[10], -dy);
                    else CurveRelative(_stack[6], _stack[7], _stack[8], _stack[9], -dx, _stack[10]);
                    _sp = 0;
                    break;
                }
                case 3: { Require(2); var b = Pop(); var a = Pop(); Push(a != 0 && b != 0 ? 1 : 0); break; }
                case 4: { Require(2); var b = Pop(); var a = Pop(); Push(a != 0 || b != 0 ? 1 : 0); break; }
                case 5: Push(Pop() == 0 ? 1 : 0); break;
                case 9: Push(Math.Abs(Pop())); break;
                case 10: { Require(2); var b = Pop(); var a = Pop(); Push(a + b); break; }
                case 11: { Require(2); var b = Pop(); var a = Pop(); Push(a - b); break; }
                case 12: { Require(2); var b = Pop(); var a = Pop(); Push(b == 0 ? 0 : a / b); break; }
                case 14: Push(-Pop()); break;
                case 15: { Require(2); var b = Pop(); var a = Pop(); Push(a == b ? 1 : 0); break; }
                case 18: Pop(); break;
                case 20: { Require(2); var index = (int)Pop(); var value = Pop(); Transient(index) = value; break; }
                case 21: Push(Transient((int)Pop())); break;
                case 22: { Require(4); var v2 = Pop(); var v1 = Pop(); var s2 = Pop(); var s1 = Pop(); Push(v1 <= v2 ? s1 : s2); break; }
                case 24: { Require(2); var b = Pop(); var a = Pop(); Push(a * b); break; }
                case 26: { var a = Pop(); Push(a < 0 ? 0 : Math.Sqrt(a)); break; }
                case 27: { var a = Pop(); Push(a); Push(a); break; }
                case 28: { Require(2); var b = Pop(); var a = Pop(); Push(b); Push(a); break; }
                case 29:
                {
                    var index = (int)Pop();
                    if (index < 0) index = 0;
                    if (index >= _sp) throw new InvalidDataException("CFF index operator out of range.");
                    Push(_stack[_sp - 1 - index]);
                    break;
                }
                case 30:
                {
                    Require(2);
                    var shift = (int)Pop(); var count = (int)Pop();
                    if (count <= 0 || count > _sp) throw new InvalidDataException("CFF roll operator out of range.");
                    var values = new double[count];
                    Array.Copy(_stack, _sp - count, values, 0, count);
                    for (var i = 0; i < count; i++)
                        _stack[_sp - count + (((i + shift) % count) + count) % count] = values[i];
                    break;
                }
                default:
                    throw new NotSupportedException("A CFF CharString operator is outside the outline profile.");
            }
        }

        private ref double Transient(int index)
        {
            if (index < 0 || index >= _transient.Length) throw new InvalidDataException("CFF transient array index out of range.");
            return ref _transient[index];
        }

        private static void Need(int cursor, int end, int bytes)
        {
            if (cursor + bytes > end) throw new InvalidDataException("Truncated CFF CharString.");
        }

        private void MoveTo(double x, double y)
        {
            CloseContour();
            _x = x; _y = y;
            _current = new List<TrueTypeOutlinePoint> { Point(x, y, true) };
        }

        private void LineTo(double x, double y)
        {
            EnsureContour();
            _x = x; _y = y;
            _current!.Add(Point(x, y, true));
        }

        private void CurveRelative(double dx1, double dy1, double dx2, double dy2, double dx3, double dy3)
        {
            EnsureContour();
            var x1 = _x + dx1; var y1 = _y + dy1;
            var x2 = x1 + dx2; var y2 = y1 + dy2;
            _x = x2 + dx3; _y = y2 + dy3;
            _current!.Add(Point(x1, y1, false));
            _current.Add(Point(x2, y2, false));
            _current.Add(Point(_x, _y, true));
        }

        private void EnsureContour()
        {
            if (_current is null) throw new InvalidDataException("A CFF path operator precedes its moveto.");
        }

        private TrueTypeOutlinePoint Point(double x, double y, bool onCurve)
        {
            if (++_points > MaximumPoints) throw new InvalidDataException("CFF outline point budget exceeded.");
            if (Math.Abs(x) > 1_000_000 || Math.Abs(y) > 1_000_000)
                throw new InvalidDataException("CFF point exceeds bounded design coordinates.");
            return new TrueTypeOutlinePoint(x, y, onCurve);
        }

        private void CloseContour()
        {
            if (_current is null) return;
            // A closing lineto back to the start is implied by the contour.
            if (_current.Count > 1 && _current[_current.Count - 1].OnCurve &&
                _current[_current.Count - 1].X == _current[0].X && _current[_current.Count - 1].Y == _current[0].Y)
                _current.RemoveAt(_current.Count - 1);
            if (_current.Count > 1) _contours.Add(_current);
            _current = null;
        }
    }
}
