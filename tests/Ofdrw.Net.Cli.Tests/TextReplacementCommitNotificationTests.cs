using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Converter.Pdf.Geometry;
using Xunit.Abstractions;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Cli.Tests;

[CollectionDefinition("Owned CLI console", DisableParallelization = true)]
public sealed class OwnedCliConsoleCollection { }

[Collection("Owned CLI console")]
public sealed class TextReplacementCommitNotificationTests
{
    private readonly ITestOutputHelper testOutput;
    public TextReplacementCommitNotificationTests(ITestOutputHelper output) => testOutput = output;
    [Fact]
    public async Task ThrowingSuccessWriterCannotFailActualNoopRawZipCommit()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-committed-console-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var font = await File.ReadAllBytesAsync(FontFixture());
            var input = Path.Combine(directory, "input.ofd"); var output = Path.Combine(directory, "output.ofd");
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new() { Id = "10", FontName = "Ofdrw Test Face", Data = font });
            var page = new OfdPage { Id = "1", WidthMillimeters = 200, HeightMillimeters = 200 };
            var text = new OfdTextElement { ObjectId = "30", LayerId = "20", FontResourceId = "10", FontName = "Ofdrw Test Face",
                XMillimeters = 20, YMillimeters = 30, WidthMillimeters = 100, HeightMillimeters = 50, FontSizeMillimeters = 4, Text = "AAA" };
            text.Runs.Add(new() { Text = "AAA", XMillimeters = 2, YMillimeters = 8, DeltaX = "2 2" });
            page.Elements.Add(text); package.Pages.Add(page);
            using (var file = File.Create(input)) await new OfdPackageWriter().WriteAsync(package, file);
            var original = await File.ReadAllBytesAsync(input); await File.WriteAllTextAsync(output, "existing target sentinel");
            var throwing = new ThrowingWriter(); var saved = Console.Out; int code;
            try
            {
                Console.SetOut(throwing);
                code = await global::Cli.RunAsync(["replace-text", input, output, "--find", "AAA", "--replace", "AAA",
                    "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            }
            finally { Console.SetOut(saved); }
            Assert.Equal(0, code); Assert.Equal(1, throwing.Attempts);
            Assert.Equal(original, await File.ReadAllBytesAsync(input)); Assert.Equal(original, await File.ReadAllBytesAsync(output));
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task FailedInputStillReportsFailureAndPreservesTargetWithThrowingStdout()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-failed-console-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        try
        {
            var input = Path.Combine(directory, "input.ofd"); var output = Path.Combine(directory, "output.ofd");
            await File.WriteAllTextAsync(input, "invalid OFD original sentinel"); await File.WriteAllTextAsync(output, "existing target sentinel");
            var throwing = new ThrowingWriter(); var saved = Console.Out; int code;
            try
            {
                Console.SetOut(throwing);
                code = await global::Cli.RunAsync(["replace-text", input, output, "--find", "A", "--replace", "A",
                    "--trusted-font", new string('0', 64) + ":100"]);
            }
            finally { Console.SetOut(saved); }
            Assert.Equal(1, code); Assert.Equal(0, throwing.Attempts);
            Assert.Equal("invalid OFD original sentinel", await File.ReadAllTextAsync(input));
            Assert.Equal("existing target sentinel", await File.ReadAllTextAsync(output));
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }
    [Theory]
    [InlineData("AAA")]
    [InlineData("A")]
    public async Task CaseVariantOwnedPathsFollowActualPlatformPolicyWithoutChangingInput(string replacement)
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-case-path-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var font = await File.ReadAllBytesAsync(FontFixture());
            var input = Path.Combine(directory, "Input.ofd"); var target = Path.Combine(directory, "input.ofd");
            await WriteCaseFixtureAsync(input, font);
            var original = await File.ReadAllBytesAsync(input);
            var namesAliasOnVolume = File.Exists(target);
            var insensitivePolicy = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS();
            testOutput.WriteLine($"Actual OS: {System.Runtime.InteropServices.RuntimeInformation.OSDescription}; case names already alias on this volume: {namesAliasOnVolume}; insensitive policy: {insensitivePolicy}");
            // On a case-insensitive volume target already names the input. Never
            // write a supposed independent target sentinel through that alias.
            var targetBefore = original;
            if (!namesAliasOnVolume)
            {
                targetBefore = Encoding.UTF8.GetBytes("existing case-distinct target sentinel");
                using var separate = new FileStream(target, FileMode.CreateNew, FileAccess.Write);
                await separate.WriteAsync(targetBefore);
            }
            if (!insensitivePolicy) Assert.False(namesAliasOnVolume); // successful branch requires real independent files, no OS simulation/skip
            var code = await global::Cli.RunAsync(["replace-text", input, target, "--find", "AAA", "--replace", replacement,
                "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            Assert.Equal(original, await File.ReadAllBytesAsync(input));
            if (insensitivePolicy)
            {
                Assert.Equal(1, code); Assert.Equal(targetBefore, await File.ReadAllBytesAsync(target));
                testOutput.WriteLine("Executed conservative Windows/macOS rejection; Linux independent-file publication is not exercised on this host.");
            }
            else
            {
                Assert.Equal(0, code);
                if (replacement == "AAA") Assert.Equal(original, await File.ReadAllBytesAsync(target));
                using var actual = File.OpenRead(target); var readback = await new OfdReader().ReadAsync(actual);
                var texts = readback.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
                Assert.Equal(replacement, Assert.Single(texts.Where(t => t.ObjectId == "30")).Text);
                Assert.Equal("A", Assert.Single(texts.Where(t => t.ObjectId == "31")).Text);
                var profile = new OfdTrustedEmbeddedFontIdentity(Convert.ToHexString(SHA256.HashData(font)), font.Length);
                var locations = new OfdKeywordLocator(new OfdTrustedEmbeddedFontGeometryProvider([profile])).Find(readback, "A");
                Assert.True(locations.IsComplete);
                var neighbor = Assert.Single(locations.Hits.Where(hit => hit.Fragments[0].BaselineMillimeters.Y == 93));
                Assert.Equal(22, Assert.Single(neighbor.Fragments).BaselineMillimeters.X);
                testOutput.WriteLine("Executed real case-distinct file publication and fresh public Reader/locator replacement and unchanged-neighbor readback.");
            }
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ValidInputCanonicalAndActualSymlinkAliasesRejectWithoutChangingFiles(bool symbolicLink)
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-actual-alias-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var font = await File.ReadAllBytesAsync(FontFixture()); var input = Path.Combine(directory, "input.ofd");
            await WriteCaseFixtureAsync(input, font);
            var original = await File.ReadAllBytesAsync(input);
            using (var actual = File.OpenRead(input))
            {
                var valid = await new OfdReader().ReadAsync(actual);
                Assert.Equal("AAA", Assert.Single(valid.Pages[0].Elements.OfType<OfdTextElement>().Where(t => t.ObjectId == "30")).Text);
            }
            var target = input;
            if (symbolicLink)
            {
                target = Path.Combine(directory, "output-link.ofd");
                try { File.CreateSymbolicLink(target, input); }
                catch (Exception error) when (OperatingSystem.IsWindows() &&
                    (error is UnauthorizedAccessException or PlatformNotSupportedException ||
                     error is IOException io && (io.HResult & 0xffff) is 1314 or 50))
                {
                    testOutput.WriteLine($"SYMLINK BRANCH NOT EXECUTED: actual Windows link API unavailable: {error.Message}. Linux/macOS real link evidence is required separately.");
                    Assert.Equal(original, await File.ReadAllBytesAsync(input));
                    Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
                    return; // no fabricated link, OS simulation, permission escalation or claimed rejection
                }
                Assert.Equal(input, new FileInfo(target).LinkTarget);
            }
            var targetBefore = await File.ReadAllBytesAsync(target);
            var code = await global::Cli.RunAsync(["replace-text", input, target, "--find", "AAA", "--replace", "AAA",
                "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            Assert.Equal(1, code); Assert.Equal(original, await File.ReadAllBytesAsync(input));
            Assert.Equal(targetBefore, await File.ReadAllBytesAsync(target));
            if (symbolicLink) Assert.Equal(input, new FileInfo(target).LinkTarget);
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
            testOutput.WriteLine(symbolicLink ? "Executed actual valid-OFD output symlink-to-input rejection." : "Executed actual valid-OFD same-canonical-path rejection.");
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("AAA")]
    [InlineData("A")]
    public async Task MissingNestedOutputParentIsCreatedAfterActualVerifiedSave(string replacement)
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-missing-parent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var font = await File.ReadAllBytesAsync(FontFixture()); var input = Path.Combine(directory, "input.ofd");
            await WriteCaseFixtureAsync(input, font); var original = await File.ReadAllBytesAsync(input);
            var parent = Path.Combine(directory, "new", "sub"); var target = Path.Combine(parent, "output.ofd");
            Assert.False(Directory.Exists(Path.Combine(directory, "new")));
            var code = await global::Cli.RunAsync(["replace-text", input, target, "--find", "AAA", "--replace", replacement,
                "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            Assert.Equal(0, code); Assert.Equal(original, await File.ReadAllBytesAsync(input));
            Assert.True(Directory.Exists(parent));
            if (replacement == "AAA") Assert.Equal(original, await File.ReadAllBytesAsync(target));
            using var actual = File.OpenRead(target); var readback = await new OfdReader().ReadAsync(actual);
            var texts = readback.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
            Assert.Equal(replacement, Assert.Single(texts.Where(t => t.ObjectId == "30")).Text);
            Assert.Equal("A", Assert.Single(texts.Where(t => t.ObjectId == "31")).Text);
            var profile = new OfdTrustedEmbeddedFontIdentity(Convert.ToHexString(SHA256.HashData(font)), font.Length);
            var locations = new OfdKeywordLocator(new OfdTrustedEmbeddedFontGeometryProvider([profile])).Find(readback, "A");
            Assert.True(locations.IsComplete);
            var neighbor = Assert.Single(locations.Hits.Where(hit => hit.Fragments[0].BaselineMillimeters.Y == 93));
            Assert.Equal(22, Assert.Single(neighbor.Fragments).BaselineMillimeters.X);
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp", SearchOption.AllDirectories));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileAndActualSymlinkOutputAncestorsRejectWithoutChangingSentinels(bool symbolicLink)
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-parent-blocker-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var font = await File.ReadAllBytesAsync(FontFixture()); var input = Path.Combine(directory, "input.ofd");
            await WriteCaseFixtureAsync(input, font); var original = await File.ReadAllBytesAsync(input);
            var ancestor = Path.Combine(directory, "blocked");
            var target = Path.Combine(ancestor, "output.ofd");
            var sentinel = Encoding.UTF8.GetBytes("existing ancestor/target sentinel");
            string? realDirectory = null;
            if (symbolicLink)
            {
                realDirectory = Path.Combine(directory, "real-target"); Directory.CreateDirectory(realDirectory);
                await File.WriteAllBytesAsync(Path.Combine(realDirectory, "output.ofd"), sentinel);
                try { Directory.CreateSymbolicLink(ancestor, realDirectory); }
                catch (Exception error) when (OperatingSystem.IsWindows() &&
                    (error is UnauthorizedAccessException or PlatformNotSupportedException ||
                     error is IOException io && (io.HResult & 0xffff) is 1314 or 50))
                {
                    testOutput.WriteLine($"SYMLINK ANCESTOR BRANCH NOT EXECUTED: actual Windows link API unavailable: {error.Message}. Linux/macOS real link evidence is required separately.");
                    Assert.Equal(original, await File.ReadAllBytesAsync(input));
                    Assert.Equal(sentinel, await File.ReadAllBytesAsync(Path.Combine(realDirectory, "output.ofd")));
                    Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp", SearchOption.AllDirectories));
                    return;
                }
                Assert.Equal(realDirectory, new DirectoryInfo(ancestor).LinkTarget);
            }
            else await File.WriteAllBytesAsync(ancestor, sentinel);
            var code = await global::Cli.RunAsync(["replace-text", input, target, "--find", "AAA", "--replace", "AAA",
                "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            Assert.Equal(1, code); Assert.Equal(original, await File.ReadAllBytesAsync(input));
            if (symbolicLink)
            {
                Assert.Equal(realDirectory, new DirectoryInfo(ancestor).LinkTarget);
                Assert.Equal(sentinel, await File.ReadAllBytesAsync(Path.Combine(realDirectory!, "output.ofd")));
            }
            else
            {
                Assert.Equal(sentinel, await File.ReadAllBytesAsync(ancestor)); Assert.False(File.Exists(target));
            }
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp", SearchOption.AllDirectories));
        }
        finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public async Task InvalidInputDoesNotCreateMissingOutputParent()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "ofdrw-invalid-missing-parent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var input = Path.Combine(directory, "input.ofd"); var parent = Path.Combine(directory, "new", "sub");
            var target = Path.Combine(parent, "output.ofd"); var original = Encoding.UTF8.GetBytes("invalid OFD original sentinel");
            await File.WriteAllBytesAsync(input, original);
            var font = await File.ReadAllBytesAsync(FontFixture());
            var code = await global::Cli.RunAsync(["replace-text", input, target, "--find", "AAA", "--replace", "AAA",
                "--trusted-font", Convert.ToHexString(SHA256.HashData(font)) + ":" + font.Length]);
            Assert.Equal(1, code); Assert.Equal(original, await File.ReadAllBytesAsync(input));
            Assert.False(Directory.Exists(Path.Combine(directory, "new"))); Assert.False(File.Exists(target));
            Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp", SearchOption.AllDirectories));
        }
        finally { Directory.Delete(directory, true); }
    }

    private static async Task WriteCaseFixtureAsync(string input, byte[] font)
    {
        var package = new OfdDocumentPackage(); package.Fonts.Add(new() { Id = "10", FontName = "Ofdrw Test Face", Data = font });
        var page = new OfdPage { Id = "1", WidthMillimeters = 200, HeightMillimeters = 200 };
        foreach (var (id, value, y) in new[] { ("30", "AAA", 30d), ("31", "A", 85d) })
        {
            var text = new OfdTextElement { ObjectId = id, LayerId = "20", FontResourceId = "10", FontName = "Ofdrw Test Face",
                XMillimeters = 20, YMillimeters = y, WidthMillimeters = 100, HeightMillimeters = 50, FontSizeMillimeters = 4, Text = value };
            text.Runs.Add(new() { Text = value, XMillimeters = 2, YMillimeters = 8, DeltaX = value.Length == 3 ? "2 2" : null });
            page.Elements.Add(text);
    }
        package.Pages.Add(page);
        using (var file = File.Create(input)) await new OfdPackageWriter().WriteAsync(package, file);
    }

    private static string FontFixture()
    {
        for (var current = new DirectoryInfo(AppContext.BaseDirectory); current is not null; current = current.Parent)
        {
            var path = Path.Combine(current.FullName, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "testdata", "fonts", "narrow.ttf");
            if (File.Exists(path)) return path;
        }
        throw new FileNotFoundException("Deterministic narrow.ttf test fixture is required.");
    }
    private sealed class ThrowingWriter : TextWriter
    {
        internal int Attempts { get; private set; }
        public override Encoding Encoding => Encoding.UTF8;
        public override void WriteLine(string? value) { Attempts++; throw new IOException("closed success output"); }
    }
}
