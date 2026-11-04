using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class PreflightCliTests : IDisposable
{
    private static readonly string[] Outcomes = ["Pass", "Fail", "NotApplicable", "Unknown"];
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-preflight-" + Guid.NewGuid().ToString("N"));

    public PreflightCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task Preflight_Human_PassExitsZero()
    {
        var (code, stdout, _) = await Run(["preflight", await Valid()]);

        Assert.Equal(0, code);
        Assert.StartsWith("Preflight ", stdout);
        Assert.Contains("Coverage gaps:", stdout);
    }

    [Fact]
    public async Task Preflight_Json_IsParseableAndStable()
    {
        var input = await Valid();
        var (code, stdout, _) = await Run(["preflight", input, "--json"]);

        Assert.Equal(0, code);
        using var json = JsonDocument.Parse(stdout);
        var root = json.RootElement;
        Assert.Equal("ofdrw-cli-preflight/1", root.GetProperty("schema").GetString());
        Assert.Equal(Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(input))).ToLowerInvariant(),
            root.GetProperty("input").GetProperty("sha256").GetString());
        Assert.Equal("valid.ofd", root.GetProperty("input").GetProperty("fileName").GetString());
        Assert.Equal(JsonValueKind.Array, root.GetProperty("report").GetProperty("ruleResults").ValueKind);
        Assert.NotEmpty(root.GetProperty("ruleOutcomes").EnumerateObject());
        Assert.All(root.GetProperty("ruleOutcomes").EnumerateObject(), outcome => Assert.Contains(outcome.Value.GetString(), Outcomes));
        Assert.DoesNotContain(_directory, stdout);
    }

    [Fact]
    public async Task Preflight_Fail_ExitsTwo()
    {
        var (code, stdout, _) = await Run(["preflight", await Failing(), "--json"]);

        Assert.Equal(2, code);
        using var json = JsonDocument.Parse(stdout);
        Assert.Contains(json.RootElement.GetProperty("ruleOutcomes").EnumerateObject(), outcome => outcome.Value.GetString() == "Fail");
    }

    [Fact]
    public async Task Preflight_Output_WritesAtomically()
    {
        var report = Path.Combine(_directory, "r.json");
        var (code, stdout, _) = await Run(["preflight", await Valid(), "--json", "--output", report]);

        Assert.Equal(0, code);
        Assert.Equal(string.Empty, stdout);
        using (var json = JsonDocument.Parse(await File.ReadAllTextAsync(report)))
            Assert.Equal("ofdrw-cli-preflight/1", json.RootElement.GetProperty("schema").GetString());
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Preflight_OutputWithoutJson_IsError()
    {
        var (code, _, stderr) = await Run(["preflight", await Valid(), "--output", Path.Combine(_directory, "r.json")]);

        Assert.Equal(1, code);
        Assert.Contains("--output requires --json", stderr);
    }

    [Fact]
    public async Task Preflight_MissingInput_IsError() =>
        Assert.Equal(1, (await Run(["preflight", Path.Combine(_directory, "missing.ofd")])).Code);

    [Fact]
    public async Task Preflight_UnknownOption_IsError() =>
        Assert.Equal(1, (await Run(["preflight", await Valid(), "--frob"])).Code);

    [Fact]
    public async Task Preflight_Garbage_IsIncomplete()
    {
        var input = Path.Combine(_directory, "garbage.ofd");
        await File.WriteAllTextAsync(input, "this is not a zip archive");

        Assert.Equal(3, (await Run(["preflight", input])).Code);
    }

    [Fact]
    public async Task Preflight_OversizedInput_SkipsHashAndIsIncomplete()
    {
        var input = Path.Combine(_directory, "big.ofd");
        await File.WriteAllBytesAsync(input, new byte[4096]);
        var options = new ArchivalPreflightOptions();
        options.LoadOptions.MaxInputBytes = 1024;
        var saved = Console.Out;
        using var output = new StringWriter();
        int code;
        try { Console.SetOut(output); code = await global::Cli.PreflightAsync([input, "--json"], options, CancellationToken.None); }
        finally { Console.SetOut(saved); }

        Assert.Equal(3, code);
        using var json = JsonDocument.Parse(output.ToString());
        var described = json.RootElement.GetProperty("input");
        Assert.Equal(JsonValueKind.Null, described.GetProperty("sha256").ValueKind);
        Assert.Equal(4096, described.GetProperty("bytes").GetInt64());
    }

    [Fact]
    public async Task Preflight_Cancelled_Returns130AndWritesNothing()
    {
        var report = Path.Combine(_directory, "r.json");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        Assert.Equal(130, await global::Cli.RunAsync(["preflight", await Valid(), "--json", "--output", report], cancelled.Token));

        Assert.False(File.Exists(report));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    private async Task<string> Valid()
    {
        var path = Path.Combine(_directory, "valid.ofd");
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        await using (var output = File.Create(path)) await new OfdPackageWriter().WriteAsync(package, output);
        return path;
    }

    // A second DocBody for the same document violates APF-001 (single document).
    private async Task<string> Failing()
    {
        var path = await Valid();
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry("OFD.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open())) xml = await reader.ReadToEndAsync();
            var body = xml.Substring(xml.IndexOf("<ofd:DocBody", StringComparison.Ordinal),
                xml.IndexOf("</ofd:DocBody>", StringComparison.Ordinal) + "</ofd:DocBody>".Length - xml.IndexOf("<ofd:DocBody", StringComparison.Ordinal));
            xml = xml.Replace("</ofd:OFD>", body + "</ofd:OFD>");
            entry.Delete();
            await using var stream = zip.CreateEntry("OFD.xml").Open();
            await stream.WriteAsync(Encoding.UTF8.GetBytes(xml));
        }
        return path;
    }

    private static async Task<(int Code, string Stdout, string Stderr)> Run(string[] args)
    {
        var savedOut = Console.Out; var savedError = Console.Error;
        using var stdout = new StringWriter(); using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout); Console.SetError(stderr);
            return (await global::Cli.RunAsync(args), stdout.ToString(), stderr.ToString());
        }
        finally { Console.SetOut(savedOut); Console.SetError(savedError); }
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
