using System.Globalization;
using System.Linq;
using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Graphics.SkiaSharp;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp.Tests;

public sealed class CffGlyphOutlinesTests
{
    private static byte[] Fixture() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
        "cff", "SourceCodePro-Regular.otf"));

    // Independent fontTools 4.60.1 RecordingPen of the pinned Adobe CFF1
    // charstrings. Each digest covers every design-unit path command in the
    // adjacent .path fixture. Skia's native backends may quantize coordinates
    // and make a closing line explicit, so compare the complete geometry.
    [Theory]
    [InlineData(2, 16, "246cf95385f4ca91e6048151b21fce9f907b5afd2fc3d9af0cd23212e7df41a3")]
    [InlineData(34, 36, "69931ee0a24a958361e221c911c3da6768046f405bf69da791267754c7e97a1a")]
    [InlineData(298, 21, "3a69a6e9deda4131fd49479b389f30942c3ddad14bbf8a56cc20fc768cdd30ea")]
    [InlineData(998, 25, "b8c5eb3e26740cc366973e0e86120ac2db1f1f759e150d7aa43d437ebe666292")]
    public void OriginalGidMatchesIndependentCffOracle(int gid, int count, string expectedHash)
    {
        var input = Fixture();
        using var reader = new CffGlyphOutlines(input);
        Array.Clear(input);
        var outline = reader.Read(gid);
        Assert.Equal("9f9664e2edf6f045c11e774f9bd0be6993971f2544a39061a5ce478b96b051f8", outline.SourceSha256);
        Assert.Equal(gid, outline.GlyphId);
        Assert.Equal(1000, outline.UnitsPerEm);
        Assert.Equal(600, outline.AdvanceWidth);
        var oraclePath = Path.Combine(AppContext.BaseDirectory, "cff", $"gid-{gid}.path");
        var oracleText = File.ReadAllText(oraclePath);
        Assert.Equal(expectedHash, Hash(oracleText));
        var expected = oracleText.Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(ParseOracleCommand).ToArray();
        Assert.Equal(count, expected.Length);
        AssertMatchesOracle(gid, outline.Commands, expected);
    }

    [Fact]
    public void UnsupportedFaceAndInvalidGidFailClosed()
    {
        var bytes = Fixture();
        Assert.Throws<NotSupportedException>(() => new CffGlyphOutlines(Array.Empty<byte>()));
        using var reader = new CffGlyphOutlines(bytes);
        Assert.Throws<NotSupportedException>(() => reader.Read(0));
        Assert.Throws<NotSupportedException>(() => reader.Read(reader.GlyphCount));
        var cancelled = new CancellationToken(true);
        Assert.Throws<OperationCanceledException>(() => reader.Read(2, cancelled));
        // Native null is ambiguous even for this valid space GID; fail closed.
        Assert.Throws<NotSupportedException>(() => reader.Read(1));
        reader.Dispose();
        Assert.Throws<ObjectDisposedException>(() => reader.Read(2));
    }

    [Fact]
    public void SfntCountCannotInventAnAbsentCffCharstring()
    {
        // The original hmtx has two zero padding bytes. Extend its directory
        // length and maxp count by one without adding a CFF CharString.
        var bytes = Fixture();
        var maxp = Record(bytes, "maxp");
        var hmtx = Record(bytes, "hmtx");
        var head = Record(bytes, "head");
        var count = Read16(bytes, Read32(bytes, maxp + 8) + 4);
        Write16(bytes, Read32(bytes, maxp + 8) + 4, count + 1);
        var hmtxLength = Read32(bytes, hmtx + 12);
        Assert.Equal(2, hmtxLength % 4);
        Write32(bytes, hmtx + 12, hmtxLength + 2);
        Write32(bytes, Read32(bytes, head + 8) + 8, 0);
        foreach (var tag in new[] { "maxp", "hmtx", "head" })
        {
            var record = Record(bytes, tag);
            Write32(bytes, record + 4, Checksum(bytes, Read32(bytes, record + 8), Read32(bytes, record + 12)));
        }
        Write32(bytes, Read32(bytes, head + 8) + 8, unchecked(0xB1B0AFBAu - Checksum(bytes, 0, bytes.Length)));
        using var data = SKData.CreateCopy(bytes);
        using var native = SKTypeface.FromData(data);
        Assert.NotNull(native);
        Assert.Equal(count, native.GlyphCount); // Skia sees 1568 CFF charstrings, not mutated maxp 1569.
        Assert.Throws<InvalidDataException>(() => new CffGlyphOutlines(bytes));
    }

    [Fact]
    public void OriginalGidPathsMakeAnInspectablePage()
    {
        using var reader = new CffGlyphOutlines(Fixture());
        var glyphs = new[] { 2, 34, 298, 998 }.Select(gid => reader.Read(gid)).ToArray();
        Assert.All(glyphs, value => Assert.Contains(value.Commands, command => command.Verb == CffPathVerb.Cubic));
        var svg = new StringBuilder("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"600\" height=\"220\" viewBox=\"0 0 600 220\"><rect width=\"600\" height=\"220\" fill=\"white\"/>");
        for (var i = 0; i < glyphs.Length; i++)
        {
            svg.Append("<path d=\"").Append(SvgPath(glyphs[i]))
                .Append("\" transform=\"translate(").Append(30 + i * 145)
                .Append(" 175) scale(0.18 -0.18)\" fill=\"#173c78\"/>");
        }
        svg.Append("</svg>");
        Assert.Equal(4, svg.ToString().Split("<path ").Length - 1);
        var artifactDirectory = Environment.GetEnvironmentVariable("OFDRW_R08_CFF_ARTIFACT_DIR");
        if (!string.IsNullOrWhiteSpace(artifactDirectory))
        {
            Directory.CreateDirectory(artifactDirectory);
            File.WriteAllText(Path.Combine(artifactDirectory, "cff-skia-gids.svg"), svg.ToString());
        }
    }

    private static string SvgPath(CffGlyphOutline outline)
    {
        var result = new StringBuilder();
        foreach (var command in outline.Commands)
        {
            switch (command.Verb)
            {
                case CffPathVerb.Move: result.Append("M ").Append(SvgPoint(command.X, command.Y)); break;
                case CffPathVerb.Line: result.Append(" L ").Append(SvgPoint(command.X, command.Y)); break;
                case CffPathVerb.Cubic:
                    result.Append(" C ").Append(SvgPoint(command.ControlX1, command.ControlY1))
                        .Append(' ').Append(SvgPoint(command.ControlX2, command.ControlY2))
                        .Append(' ').Append(SvgPoint(command.X, command.Y)); break;
                case CffPathVerb.Close: result.Append(" Z"); break;
                default: throw new InvalidOperationException("Unexpected CFF command.");
            }
        }
        return result.ToString();
    }

    private static string SvgPoint(double x, double y) =>
        x.ToString("0.###", CultureInfo.InvariantCulture) + " " + y.ToString("0.###", CultureInfo.InvariantCulture);

    private const double CoordinateTolerance = 0.125;

    private static (CffPathVerb Verb, double[] Values) ParseOracleCommand(string line)
    {
        var parts = line.Split(':');
        var verb = parts[0] switch
        {
            "M" => CffPathVerb.Move,
            "L" => CffPathVerb.Line,
            "C" => CffPathVerb.Cubic,
            "Z" => CffPathVerb.Close,
            _ => throw new InvalidDataException("Unexpected independent CFF oracle verb.")
        };
        var values = parts.Skip(1).SelectMany(point => point.Split(','))
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(verb == CffPathVerb.Cubic ? 6 : verb == CffPathVerb.Close ? 0 : 2, values.Length);
        return (verb, values);
    }

    private static void AssertMatchesOracle(int gid, IReadOnlyList<CffPathCommand> actual,
        IReadOnlyList<(CffPathVerb Verb, double[] Values)> expected)
    {
        var actualIndex = 0;
        var contourX = 0d;
        var contourY = 0d;
        for (var expectedIndex = 0; expectedIndex < expected.Count; expectedIndex++)
        {
            var (verb, values) = expected[expectedIndex];
            // FreeType can materialize the implicit closing edge before Z.
            // Only that exact edge is skipped; every oracle segment still
            // has to match its native counterpart in order and coordinates.
            if (verb == CffPathVerb.Close && actualIndex + 1 < actual.Count &&
                actual[actualIndex].Verb == CffPathVerb.Line &&
                actual[actualIndex + 1].Verb == CffPathVerb.Close)
            {
                Compare(gid, expectedIndex, "closing X", contourX, actual[actualIndex].X);
                Compare(gid, expectedIndex, "closing Y", contourY, actual[actualIndex].Y);
                actualIndex++;
            }
            Assert.True(actualIndex < actual.Count, $"CFF GID {gid} ended before oracle command {expectedIndex}.");
            var command = actual[actualIndex++];
            Assert.Equal(verb, command.Verb);
            if (verb == CffPathVerb.Close) continue;
            var native = verb == CffPathVerb.Cubic
                ? new[] { command.ControlX1, command.ControlY1, command.ControlX2, command.ControlY2, command.X, command.Y }
                : new[] { command.X, command.Y };
            for (var i = 0; i < values.Length; i++)
                Compare(gid, expectedIndex, $"coordinate {i}", values[i], native[i]);
            if (verb == CffPathVerb.Move) (contourX, contourY) = (command.X, command.Y);
        }
        Assert.Equal(actual.Count, actualIndex);
    }

    private static void Compare(int gid, int index, string coordinate, double expected, double actual) =>
        Assert.True(Math.Abs(expected - actual) <= CoordinateTolerance,
            $"CFF GID {gid} command {index} {coordinate}: expected {expected}, actual {actual}, tolerance {CoordinateTolerance}.");

    private static string Hash(string text) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

    private static int Record(byte[] bytes, string tag)
    {
        var count = Read16(bytes, 4);
        for (var i = 0; i < count; i++)
        {
            var offset = 12 + i * 16;
            if (Encoding.ASCII.GetString(bytes, offset, 4) == tag) return offset;
        }
        throw new InvalidOperationException("Pinned font lacks " + tag);
    }

    private static int Read16(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));
    private static int Read32(byte[] bytes, int offset) => checked((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4)));
    private static void Write16(byte[] bytes, int offset, int value) =>
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(offset, 2), checked((ushort)value));
    private static void Write32(byte[] bytes, int offset, int value) =>
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(offset, 4), checked((uint)value));
    private static void Write32(byte[] bytes, int offset, uint value) =>
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(offset, 4), value);
    private static uint Checksum(byte[] bytes, int offset, int length)
    {
        uint sum = 0;
        for (var i = 0; i < length; i += 4)
        {
            uint word = 0;
            for (var j = 0; j < 4; j++)
                word = (word << 8) | (i + j < length ? bytes[offset + i + j] : 0u);
            sum = unchecked(sum + word);
        }
        return sum;
    }
}
