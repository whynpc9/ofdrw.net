using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Cli.Tests;

[CollectionDefinition("Owned CLI console", DisableParallelization = true)]
public sealed class OwnedCliConsoleCollection { }

[Collection("Owned CLI console")]
public sealed class TextReplacementCommitNotificationTests
{
    [Fact]
    public async Task ThrowingSuccessWriterCannotFailActualNoopRawZipCommit()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-committed-console-" + Guid.NewGuid().ToString("N"));
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
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-failed-console-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
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
