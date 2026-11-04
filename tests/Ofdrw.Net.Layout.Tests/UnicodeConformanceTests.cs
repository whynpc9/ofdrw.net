using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Layout.Text.Unicode;
using Xunit.Abstractions;

namespace Ofdrw.Net.Layout.Tests;

// Full Unicode 17.0.0 conformance files for the algorithms used by paragraph layout.
// Every data line is executed; the counts are written to the test output.
public sealed class UnicodeConformanceTests(ITestOutputHelper output)
{
    [Fact]
    public void GraphemeBreakTestPassesEveryCase()
    {
        int cases = 0, failures = 0;
        var messages = new List<string>();
        foreach (var line in Lines("auxiliary/GraphemeBreakTest.txt"))
        {
            var (codePoints, expected) = ParseBoundaryLine(line);
            if (codePoints.Length == 0) continue;
            cases++;
            var actual = GraphemeBreaker.Boundaries(codePoints);
            if (!actual.SequenceEqual(expected)) Fail(ref failures, messages, line, string.Join(",", actual));
        }
        Report("GraphemeBreakTest", cases, failures, messages);
    }

    [Fact]
    public void LineBreakTestPassesEveryCase()
    {
        int cases = 0, failures = 0;
        var messages = new List<string>();
        foreach (var line in Lines("auxiliary/LineBreakTest.txt"))
        {
            var (codePoints, expected) = ParseBoundaryLine(line);
            if (codePoints.Length == 0) continue;
            cases++;
            var kinds = LineBreaker.Analyze(codePoints);
            var actual = Enumerable.Range(1, codePoints.Length).Where(index => kinds[index] != LineBreakKind.Prohibited);
            if (!actual.SequenceEqual(expected.Where(index => index > 0)))
                Fail(ref failures, messages, line, string.Join(",", actual));
        }
        Report("LineBreakTest", cases, failures, messages);
    }

    [Fact]
    public void BidiTestPassesEveryCase()
    {
        int cases = 0, failures = 0;
        var messages = new List<string>();
        string[] levels = [];
        int[] reorder = [];
        var names = Enum.GetNames<BidiClass>();
        foreach (var line in Lines("BidiTest.txt"))
        {
            if (line.StartsWith("@Levels:", StringComparison.Ordinal))
            { levels = Split(line.Substring(8)); continue; }
            if (line.StartsWith("@Reorder:", StringComparison.Ordinal))
            { reorder = Split(line.Substring(9)).Select(int.Parse).ToArray(); continue; }
            if (line.StartsWith('@')) continue;
            var parts = line.Split(';');
            var types = Split(parts[0]).Select(name => (BidiClass)Array.IndexOf(names, name)).ToArray();
            var bitset = int.Parse(parts[1].Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            foreach (var (bit, paragraph) in new[] { (1, -1), (2, 0), (4, 1) })
            {
                if ((bitset & bit) == 0) continue;
                cases++;
                var bidi = BidiAlgorithm.Resolve(types, null, paragraph);
                var lineLevels = bidi.LineLevels(0, types.Length, keepRemovedMarker: true);
                var actualLevels = lineLevels.Select(level => level == BidiAlgorithm.Removed ? "x" : level.ToString(CultureInfo.InvariantCulture));
                var actualOrder = BidiAlgorithm.VisualOrder(lineLevels);
                if (!actualLevels.SequenceEqual(levels) || !actualOrder.SequenceEqual(reorder))
                    Fail(ref failures, messages, line + " P=" + paragraph,
                        string.Join(" ", actualLevels) + " / " + string.Join(" ", actualOrder));
            }
        }
        Report("BidiTest", cases, failures, messages);
    }

    [Fact]
    public void BidiCharacterTestPassesEveryCase()
    {
        int cases = 0, failures = 0;
        var messages = new List<string>();
        foreach (var line in Lines("BidiCharacterTest.txt"))
        {
            var fields = line.Split(';');
            var codePoints = Split(fields[0]).Select(value => int.Parse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToArray();
            var direction = int.Parse(fields[1], CultureInfo.InvariantCulture);
            var paragraphLevel = int.Parse(fields[2], CultureInfo.InvariantCulture);
            var expectedLevels = Split(fields[3]);
            var expectedOrder = Split(fields[4]).Select(int.Parse).ToArray();
            cases++;
            var types = codePoints.Select(UnicodeProperties.Bidi).ToArray();
            var bidi = BidiAlgorithm.Resolve(types, codePoints, direction == 2 ? -1 : direction);
            var lineLevels = bidi.LineLevels(0, types.Length, keepRemovedMarker: true);
            var actualLevels = lineLevels.Select(level => level == BidiAlgorithm.Removed ? "x" : level.ToString(CultureInfo.InvariantCulture));
            var actualOrder = BidiAlgorithm.VisualOrder(lineLevels);
            if (bidi.ParagraphLevel != paragraphLevel || !actualLevels.SequenceEqual(expectedLevels) ||
                !actualOrder.SequenceEqual(expectedOrder))
                Fail(ref failures, messages, line,
                    bidi.ParagraphLevel + ";" + string.Join(" ", actualLevels) + ";" + string.Join(" ", actualOrder));
        }
        Report("BidiCharacterTest", cases, failures, messages);
    }

    private static (int[] CodePoints, int[] Boundaries) ParseBoundaryLine(string line)
    {
        var content = line.Split('#')[0];
        var codePoints = new List<int>();
        var boundaries = new List<int>();
        foreach (var token in Split(content))
        {
            if (token == "÷") boundaries.Add(codePoints.Count);
            else if (token == "×") { }
            else codePoints.Add(int.Parse(token, NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        }
        return (codePoints.ToArray(), boundaries.ToArray());
    }

    private void Report(string name, int cases, int failures, List<string> messages)
    {
        output.WriteLine($"{name}: Unicode {UnicodeData.Version}, {cases} cases, {failures} failures.");
        foreach (var message in messages) output.WriteLine(message);
        Assert.True(cases > 0, name + " executed no cases.");
        Assert.True(failures == 0, $"{name}: {failures} of {cases} failed.\n" + string.Join("\n", messages));
    }

    private static void Fail(ref int failures, List<string> messages, string line, string actual)
    {
        failures++;
        if (messages.Count < 25) messages.Add(line.Trim() + "  => " + actual);
    }

    private static string[] Split(string value) =>
        value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

    private static IEnumerable<string> Lines(string manifestKey)
    {
        var root = RepositoryRoot();
        var path = Path.Combine(root, "tests", "Ofdrw.Net.Layout.Tests", "testdata", "unicode-17.0.0",
            Path.GetFileName(manifestKey) + ".gz");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "scripts", "unicode", "ucd-17.0.0-manifest.json")));
        var expected = manifest.RootElement.GetProperty("conformanceTests").GetProperty(manifestKey).GetString();
        byte[] bytes;
        using (var stream = new GZipStream(File.OpenRead(path), CompressionMode.Decompress))
        using (var buffer = new MemoryStream())
        {
            stream.CopyTo(buffer);
            bytes = buffer.ToArray();
        }
        Assert.Equal(expected, Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        using var reader = new StringReader(Encoding.UTF8.GetString(bytes));
        while (reader.ReadLine() is { } line)
        {
            var trimmed = line.Trim();
            if (trimmed.Length == 0 || trimmed.StartsWith('#')) continue;
            yield return trimmed;
        }
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("Repository root not found.");
    }
}
