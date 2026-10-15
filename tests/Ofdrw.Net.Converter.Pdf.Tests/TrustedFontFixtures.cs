using System.Security.Cryptography;
using System.Text.Json;

namespace Ofdrw.Net.Converter.Pdf.Tests;

// Test-only fixture paths; never an OS font resolver or production fallback.
internal static class TrustedFontFixtures
{
    internal static bool CanRun => Environment.GetEnvironmentVariable("OFDRW_REQUIRE_TRUSTED_FONTS") == "1" ||
        !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFDRW_TEST_TRUSTED_CJK_FONT")) ||
        !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFDRW_TEST_TRUSTED_LATIN_FONT")) ||
        (DefaultDirectory() is { } directory && File.Exists(Path.Combine(directory, "fonts", "NotoSans-Regular.ttf")) &&
            File.Exists(Path.Combine(directory, "fonts", "Ofdrw-CI-NotoSansCJKsc-Regular.ttf")));

    internal static byte[] Read(string face)
    {
        var cjk = face switch { "cjk" => true, "latin" => false, _ => throw new ArgumentOutOfRangeException(nameof(face)) };
        var path = PathFor(face);
        var expectedBytes = cjk ? 21659544 : 569208;
        using var source = File.OpenRead(path);
        if (source.Length != expectedBytes) throw new InvalidDataException("Trusted fixture size mismatch: " + face);
        var bytes = new byte[expectedBytes];
        source.ReadExactly(bytes);
        if (source.ReadByte() != -1) throw new InvalidDataException("Trusted fixture grew during capture: " + face);
        var expected = cjk ? "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa" :
            "b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5";
        if (Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() != expected)
            throw new InvalidDataException("Trusted fixture SHA mismatch: " + face);
        return bytes;
    }

    internal static string PathFor(string face)
    {
        var cjk = face switch { "cjk" => true, "latin" => false, _ => throw new ArgumentOutOfRangeException(nameof(face)) };
        var path = Environment.GetEnvironmentVariable(cjk ? "OFDRW_TEST_TRUSTED_CJK_FONT" : "OFDRW_TEST_TRUSTED_LATIN_FONT");
        if (string.IsNullOrEmpty(path))
            path = Path.Combine(DefaultDirectory() ?? throw new FileNotFoundException("Prepare repository-local trusted fonts or set both explicit fixture paths."),
                "fonts", cjk ? "Ofdrw-CI-NotoSansCJKsc-Regular.ttf" : "NotoSans-Regular.ttf");
        return path;
    }

    private static string? DefaultDirectory()
    {
        foreach (var start in new[] { AppContext.BaseDirectory, Directory.GetCurrentDirectory() })
            for (var current = new DirectoryInfo(start); current is not null; current = current.Parent)
            {
                var manifest = Path.Combine(current.FullName, "tests", "fixtures", "trusted-fonts", "bundle.json");
                if (!File.Exists(manifest)) continue;
                using var stream = File.OpenRead(manifest);
                var metadata = new byte[65537]; var count = 0;
                while (count < metadata.Length)
                {
                    var read = stream.Read(metadata, count, metadata.Length - count);
                    if (read == 0) break;
                    count += read;
                }
                if (count > 65536) throw new InvalidDataException("Oversized trusted fixture manifest.");
                using var json = JsonDocument.Parse(metadata.AsMemory(0, count));
                var digest = json.RootElement.GetProperty("archive").GetProperty("sha256").GetString()!;
                if (digest.Length != 64 || digest.Any(c => c is not (>= '0' and <= '9') and not (>= 'a' and <= 'f')))
                    throw new InvalidDataException("Invalid fixture bundle identity.");
                return Path.Combine(current.FullName, "artifacts", "trusted-fonts", "sha256-" + digest);
            }
        return null;
    }
}
