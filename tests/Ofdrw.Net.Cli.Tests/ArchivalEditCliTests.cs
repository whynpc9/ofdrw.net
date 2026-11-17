using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

/// <summary>A2: write verbs refuse OFD-A input with exit code 5; read-only verbs and derive-ofd keep working.</summary>
[Collection("Owned CLI console")]
public sealed class ArchivalEditCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-archival-" + Guid.NewGuid().ToString("N"));
    public ArchivalEditCliTests() => Directory.CreateDirectory(_directory);
    public void Dispose() { try { Directory.Delete(_directory, true); } catch (IOException) { } }
    private string PathFor(string name) => Path.Combine(_directory, name);

    // {in} = OFD-A input, {ord} = ordinary input, {out} = new output path.
    public static TheoryData<string> WriteVerbs() => new()
    {
        "replace-text {in} {out} --find FIRST --replace LAST --trusted-font " + new string('0', 64) + ":1",
        "reorder {in} {out} --pages 2,1",
        "delete-pages {in} {out} --pages 2",
        "crop {in} {out} --box 5,5,90,90 --pages 1",
        "attachments remove {in} {out} --id {attachment}",
        "areas fill {in} {out} --name name --text Alice",
        "watermark {in} {out} --pages 1 --text DRAFT",
        "split {in} {out} --pages 2",
        "mix {out} {ord} 1 {in} 1",
        "clean-signatures {in} {out}",
        "merge {out} {ord} {in}",
        "merge {out} {in} {ord}"
    };

    [Theory]
    [MemberData(nameof(WriteVerbs))]
    public async Task WriteVerb_RefusesArchivalInputWithExitCode5(string command)
    {
        var archival = await Fixture("archival.ofd", OfdDocTypes.Archival);
        var ordinary = await Fixture("ordinary.ofd", OfdDocTypes.Ordinary);
        var output = PathFor("out.ofd");
        var hash = Sha(archival);
        var args = Expand(command, archival, ordinary, output, await AttachmentId(archival));

        var (code, _, stderr) = await Run(args);

        Assert.True(code == 5, $"exit {code}: {stderr}");
        Assert.Contains("DocType=\"OFD-A\"", stderr);
        Assert.Contains("ofdrw derive-ofd", stderr);
        Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
        Assert.Equal(hash, Sha(archival));
    }

    [Theory]
    [InlineData("reorder {ord} {out} --pages 2,1")]
    [InlineData("delete-pages {ord} {out} --pages 2")]
    [InlineData("crop {ord} {out} --box 5,5,90,90 --pages 1")]
    [InlineData("attachments remove {ord} {out} --id {attachment}")]
    [InlineData("areas fill {ord} {out} --name name --text Alice")]
    [InlineData("watermark {ord} {out} --pages 1 --text DRAFT")]
    [InlineData("split {ord} {out} --pages 2")]
    [InlineData("mix {out} {ord} 1 {ord} 2")]
    [InlineData("clean-signatures {ord} {out}")]
    [InlineData("merge {out} {ord} {ord}")]
    public async Task WriteVerb_StillWritesOrdinaryOfd(string command)
    {
        var ordinary = await Fixture("ordinary.ofd", OfdDocTypes.Ordinary);
        var output = PathFor("out.ofd");
        var (code, _, stderr) = await Run(Expand(command, ordinary, ordinary, output, await AttachmentId(ordinary)));
        Assert.True(code == 0, $"exit {code}: {stderr}");
        Assert.Equal("OFD", RootDocType(output));
    }

    [Theory]
    [InlineData("extract-text {in}")]
    [InlineData("ofd-to-pdf -i {in} -o {out}.pdf")]
    [InlineData("ofd-to-svg -i {in} -o {out}.svg --pages 1")]
    [InlineData("attachments list {in}")]
    [InlineData("attachments extract {in} --id {attachment} --output {out}.bin")]
    [InlineData("areas list {in}")]
    [InlineData("verify-signatures -i {in}")]
    public async Task ReadOnlyVerb_AcceptsArchivalInput(string command)
    {
        var archival = await Fixture("archival.ofd", OfdDocTypes.Archival);
        var hash = Sha(archival);
        var (code, _, stderr) = await Run(Expand(command, archival, archival, PathFor("out"), await AttachmentId(archival)));
        Assert.True(code == 0, $"exit {code}: {stderr}");
        Assert.Equal(hash, Sha(archival));
    }

    [Fact]
    public async Task Preflight_AcceptsArchivalInput()
    {
        var archival = await Fixture("archival.ofd", OfdDocTypes.Archival);
        var (code, _, stderr) = await Run(["preflight", archival]);
        Assert.True(code is 0 or 2 or 3, $"exit {code}: {stderr}");
        Assert.DoesNotContain("derive-ofd", stderr);
    }

    [Fact]
    public async Task DeriveOfd_WritesRecordAndDerivativeCanBeEdited()
    {
        var archival = await Fixture("archival.ofd", OfdDocTypes.Archival);
        var hash = Sha(archival);
        var derived = PathFor("derived.ofd");

        var (code, stdout, stderr) = await Run(["derive-ofd", archival, derived, "--json"]);

        Assert.True(code == 0, $"exit {code}: {stderr}");
        using var json = JsonDocument.Parse(stdout);
        var root = json.RootElement;
        Assert.Equal("ofdrw-cli-derive-ofd/1", root.GetProperty("schema").GetString());
        Assert.Equal(OfdOrdinaryDerivation.OperationName, root.GetProperty("operation").GetString());
        Assert.Equal(hash, root.GetProperty("source").GetProperty("sha256").GetString());
        Assert.Equal("OFD-A", root.GetProperty("source").GetProperty("docType").GetString());
        Assert.Equal(Sha(derived), root.GetProperty("output").GetProperty("sha256").GetString());
        Assert.Equal(new FileInfo(derived).Length, root.GetProperty("output").GetProperty("bytes").GetInt64());
        Assert.Equal("OFD", root.GetProperty("output").GetProperty("docType").GetString());
        Assert.Equal("OFD.xml", root.GetProperty("changedEntry").GetString());
        Assert.False(root.GetProperty("inheritsArchivalConclusion").GetBoolean());
        Assert.Equal(hash, Sha(archival));
        Assert.Equal("OFD", RootDocType(derived));

        var edited = PathFor("edited.ofd");
        Assert.Equal(0, (await Run(["delete-pages", derived, edited, "--pages", "2"])).Code);
        Assert.Equal("OFD", RootDocType(edited));

        // Text mode prints the same record for humans.
        var (textCode, text, _) = await Run(["derive-ofd", archival, PathFor("derived-text.ofd")]);
        Assert.Equal(0, textCode);
        Assert.Contains(hash, text);
        Assert.Contains("does not inherit any archival conclusion", text);
    }

    [Theory]
    [InlineData("ordinary")]
    [InlineData("existing")]
    [InlineData("same")]
    public async Task DeriveOfd_FailsWithExitCode1WithoutOutput(string scenario)
    {
        var input = await Fixture("input.ofd", scenario == "ordinary" ? OfdDocTypes.Ordinary : OfdDocTypes.Archival);
        var output = scenario == "same" ? input : PathFor("derived.ofd");
        if (scenario == "existing") await File.WriteAllTextAsync(output, "keep");
        var hash = Sha(input);
        var files = Directory.GetFiles(_directory).OrderBy(x => x, StringComparer.Ordinal).ToArray();

        var (code, _, stderr) = await Run(["derive-ofd", input, output]);

        Assert.Equal(1, code);
        if (scenario == "ordinary") Assert.Contains("needs no derivation", stderr);
        Assert.Equal(hash, Sha(input));
        Assert.Equal(files, Directory.GetFiles(_directory).OrderBy(x => x, StringComparer.Ordinal).ToArray());
    }

    private async Task<string> Fixture(string name, string docType)
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new Paragraph("FIRST"));
        flow.Blocks.Add(new AreaHolderBlock("name", 18));
        flow.Blocks.Add(new Paragraph("SECOND") { PageBreakBefore = true });
        var package = flow.Render();
        package.Attachments.Add(new OfdAttachment { Name = "note.txt", MediaType = "text/plain", Data = [1, 2, 3] });
        package.Options.DocType = docType;
        var path = PathFor(name);
        await using (var stream = File.Create(path)) await new OfdPackageWriter().WriteAsync(package, stream);
        return path;
    }

    private static async Task<string> AttachmentId(string path)
    {
        await using var stream = File.OpenRead(path);
        return (await new OfdReader().ReadAsync(stream)).Attachments.Single().Id;
    }

    private static string[] Expand(string command, string archival, string ordinary, string output, string attachment) =>
        command.Split(' ').Select(part => part.Replace("{in}", archival).Replace("{ord}", ordinary)
            .Replace("{out}", output).Replace("{attachment}", attachment)).ToArray();

    private static string? RootDocType(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        using var stream = zip.GetEntry("OFD.xml")!.Open();
        return XDocument.Load(stream).Root!.Attribute("DocType")?.Value;
    }

    private static string Sha(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

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
}
