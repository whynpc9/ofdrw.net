using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class AreasCliTests : IDisposable
{
    private const string FillText = "Alice 中文";
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-areas-" + Guid.NewGuid().ToString("N"));

    public AreasCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task List_PrintsOneRowPerArea()
    {
        var template = await Template();
        var previous = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
        try
        {
            var (code, stdout, stderr) = await Run("areas", "list", template);

            Assert.True(code == 0, stderr);
            var rows = stdout.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
                .Select(row => row.Split('\t')).ToArray();
            Assert.Collection(rows,
                row => Assert.Equal(new[] { "name", "1", "25.4", "30.525", "159.2", "18", "blank" }, row),
                row => Assert.Equal(new[] { "logo", "1", "25.4", "48.525", "159.2", "15", "blank" }, row));
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    [Fact]
    public async Task List_Json()
    {
        var (code, stdout, stderr) = await Run("areas", "list", await Template(), "--json");

        Assert.True(code == 0, stderr);
        Assert.Empty(stderr);
        using var json = JsonDocument.Parse(stdout);
        Assert.Equal("ofdrw-cli-areas/1", json.RootElement.GetProperty("schema").GetString());
        var areas = json.RootElement.GetProperty("areas").EnumerateArray().ToArray();
        Assert.Equal(2, areas.Length);
        var area = Assert.Single(areas, item => item.GetProperty("name").GetString() == "name");
        Assert.Equal(1, area.GetProperty("page").GetInt32());
        Assert.Equal(25.4, area.GetProperty("xMillimeters").GetDouble());
        Assert.Equal(30.525, area.GetProperty("yMillimeters").GetDouble(), 6);
        Assert.Equal(159.2, area.GetProperty("widthMillimeters").GetDouble(), 6);
        Assert.Equal(18, area.GetProperty("heightMillimeters").GetDouble());
        Assert.False(area.GetProperty("filled").GetBoolean());
    }

    [Fact]
    public async Task List_DocumentWithoutAreas_PrintsNothing()
    {
        var input = PathFor("plain.txt");
        var document = PathFor("plain.ofd");
        await File.WriteAllTextAsync(input, "no areas");
        Assert.Equal(0, (await Run("text-to-ofd", input, document)).Code);

        var (code, stdout, stderr) = await Run("areas", "list", document);
        Assert.True(code == 0, stderr);
        Assert.Empty(stdout);
        var machine = await Run("areas", "list", document, "--json");
        Assert.True(machine.Code == 0, machine.Stderr);
        using var json = JsonDocument.Parse(machine.Stdout);
        Assert.Empty(json.RootElement.GetProperty("areas").EnumerateArray());
    }

    [Fact]
    public async Task FillText_WritesTextInsideTheArea()
    {
        var template = await Template();
        var output = PathFor("out.ofd");
        var (code, stdout, stderr) = await Run("areas", "fill", template, output, "--name", "name", "--text", FillText);

        Assert.True(code == 0, stderr);
        Assert.Equal($"Filled area name -> {output}{Environment.NewLine}", stdout);
        var source = await Read(template);
        var result = await Read(output);
        var editor = new OfdNamedAreaEditor();
        var areas = editor.GetAreas(result);
        var area = Assert.Single(areas, item => item.Name == "name");
        Assert.True(area.IsFilled);
        Assert.False(Assert.Single(areas, item => item.Name == "logo").IsFilled);
        var after = Assert.Single(result.Pages).Elements.OfType<OfdTextElement>().ToArray();
        AssertInside(Assert.Single(after, text => text.Text == FillText), area);
        var before = Assert.Single(source.Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "HEADER", "FOOTER" }, before.Select(text => text.Text));
        Assert.Equal(before.Length + 1, after.Length);
        foreach (var original in before)
        {
            var current = Assert.Single(after, text => text.Text == original.Text);
            Assert.Equal((original.XMillimeters, original.YMillimeters),
                (current.XMillimeters, current.YMillimeters));
        }

        var artifacts = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (!string.IsNullOrEmpty(artifacts))
        {
            Directory.CreateDirectory(artifacts);
            File.Copy(template, Path.Combine(artifacts, "template.ofd"), overwrite: true);
            File.Copy(output, Path.Combine(artifacts, "out.ofd"), overwrite: true);
        }
    }

    [Fact]
    public async Task FillText_InPlace()
    {
        var template = await Template();

        var (code, _, stderr) = await Run("areas", "fill", template, template,
            "--name", "name", "--text", FillText, "--font", "Arial", "--font-size", "4.25");

        Assert.True(code == 0, stderr);
        var package = await Read(template);
        var text = Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>(), item => item.Text == FillText);
        Assert.Equal("Arial", text.FontName);
        Assert.Equal(4.25, text.FontSizeMillimeters);
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task FillImage_FitsTheImage()
    {
        var input = PathFor("dot.PNG");
        // Valid 1 × 1 grayscale PNG, including the IDAT CRC checked by ImageSharp.
        var payload = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGNQAgAAJAAjRKTbngAAAABJRU5ErkJggg==");
        await File.WriteAllBytesAsync(input, payload);
        var template = await Template();
        var output = PathFor("out.ofd");

        var (code, _, stderr) = await Run("areas", "fill", template, output, "--name", "logo", "--image", input);

        Assert.True(code == 0, stderr);
        var package = await Read(output);
        var area = Assert.Single(new OfdNamedAreaEditor().GetAreas(package), item => item.Name == "logo");
        var image = Assert.Single(package.Pages[area.PageIndex].Elements.OfType<OfdImageElement>());
        AssertInside(image, area);
        Assert.Equal(area.HeightMillimeters, image.HeightMillimeters);
        Assert.Equal(image.HeightMillimeters, image.WidthMillimeters);
        Assert.Equal(payload, image.Data);
        Assert.True(area.IsFilled);

        var alias = await Run("areas", "fill", template, input, "--name", "logo", "--image", input);
        Assert.Equal(1, alias.Code);
        Assert.Contains("different paths", alias.Stderr);
        Assert.Equal(payload, await File.ReadAllBytesAsync(input));
    }

    [Fact]
    public async Task FillImage_RejectsOversizedAndMislabeledFilesBeforeWriting()
    {
        var template = await Template();
        var image = PathFor("picture.png");
        var output = PathFor("existing.ofd");
        await File.WriteAllTextAsync(output, "unchanged");
        using (var stream = File.Create(image))
            stream.SetLength(16 * 1024 * 1024 + 1L);

        var oversized = await Run("areas", "fill", template, output, "--name", "logo", "--image", image);
        Assert.Equal(1, oversized.Code);
        Assert.Contains("Image encoded byte budget exceeded.", oversized.Stderr);

        // A BigTIFF header with a PNG extension must never reach image identification.
        await File.WriteAllBytesAsync(image, [0x49, 0x49, 0x2b, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
        var mislabeled = await Run("areas", "fill", template, output, "--name", "logo", "--image", image);
        Assert.Equal(1, mislabeled.Code);
        Assert.Contains("Image must be PNG/JPEG.", mislabeled.Stderr);
        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Fill_Twice_Fails()
    {
        var first = PathFor("first.ofd");
        var second = PathFor("second.ofd");
        var initial = await Run("areas", "fill", await Template(), first, "--name", "name", "--text", FillText);
        Assert.True(initial.Code == 0, initial.Stderr);

        var (code, _, stderr) = await Run("areas", "fill", first, second, "--name", "name", "--text", "again");
        Assert.Equal(1, code);
        Assert.Contains("already been filled", stderr);
        Assert.False(File.Exists(second));
    }

    [Fact]
    public async Task Fill_UnknownName_Fails()
    {
        var output = PathFor("out.ofd");

        var (code, _, stderr) = await Run("areas", "fill", await Template(), output, "--name", "nope", "--text", FillText);

        Assert.Equal(1, code);
        Assert.Contains("not found", stderr);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Fill_TooMuchText_FailsWithoutOutput()
    {
        var template = await Template();
        var output = PathFor("out.ofd");
        var original = await File.ReadAllBytesAsync(template);

        var (code, _, stderr) = await Run("areas", "fill", template, output, "--name", "name", "--text", new string('W', 5000));

        Assert.Equal(1, code);
        Assert.Contains("exceeds its Boundary", stderr);
        Assert.False(File.Exists(output));
        var inPlace = await Run("areas", "fill", template, template, "--name", "name", "--text", new string('W', 5000));
        Assert.Equal(1, inPlace.Code);
        Assert.Equal(original, await File.ReadAllBytesAsync(template));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("requires exactly one", "--name", "name")]
    [InlineData("requires exactly one", "--name", "name", "--text", "x", "--image", "dot.png")]
    [InlineData("apply only to --text", "--name", "logo", "--image", "dot.png", "--font", "X")]
    [InlineData("--name is required", "--text", "x")]
    [InlineData("Invalid --font-size", "--name", "name", "--text", "x", "--font-size", "0")]
    [InlineData("Invalid --font-size", "--name", "name", "--text", "x", "--font-size", "Infinity")]
    [InlineData("Invalid --font-size", "--name", "name", "--text", "x", "--font-size", "4,25")]
    [InlineData("Image must be PNG/JPEG", "--name", "logo", "--image", "file.gif")]
    [InlineData("Unknown areas fill option: --bold", "--name", "name", "--text", "x", "--bold")]
    [InlineData("Duplicate option: --name", "--name", "name", "--text", "x", "--name", "name")]
    public async Task Fill_ArgumentErrors(string diagnostic, params string[] options)
    {
        var output = PathFor("out.ofd");

        var (code, _, stderr) = await Run(["areas", "fill", await Template(), output, .. options]);

        Assert.Equal(1, code);
        Assert.Contains(diagnostic, stderr);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task UnknownSubcommand_Fails()
    {
        var (code, _, stderr) = await Run("areas", "frob", await Template());

        Assert.Equal(1, code);
        Assert.Contains("Unknown areas subcommand: frob", stderr);
    }

    [Fact]
    public async Task List_UsesTerminalEscapingAndKeepsJsonNamesExact()
    {
        const string name = "中\u202e文";
        var path = await Template(name);

        var list = await Run("areas", "list", path);
        Assert.True(list.Code == 0, list.Stderr);
        Assert.StartsWith(@"中\u202E文" + "\t", list.Stdout);
        var machine = await Run("areas", "list", path, "--json");
        Assert.True(machine.Code == 0, machine.Stderr);
        using var json = JsonDocument.Parse(machine.Stdout);
        Assert.Equal(name, json.RootElement.GetProperty("areas")[0].GetProperty("name").GetString());
        var output = PathFor("out-\u202e.ofd");
        var fill = await Run("areas", "fill", path, output, "--name", name, "--text", FillText);
        Assert.True(fill.Code == 0, fill.Stderr);
        Assert.Equal($@"Filled area 中\u202E文 -> {PathFor(@"out-\u202E.ofd")}" + Environment.NewLine, fill.Stdout);
        var filled = await Run("areas", "list", output);
        Assert.True(filled.Code == 0, filled.Stderr);
        Assert.EndsWith("\tfilled", filled.Stdout.Split(Environment.NewLine)[0]);
    }

    [Fact]
    public async Task MultiDocument_ListsSelectedDocumentAndRejectsFill()
    {
        var entries = await MultiDocumentFixture.CreateEntriesAsync();
        var root = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"]));
        var bodies = root.Root!.Elements().ToArray();
        var custom = CreateTemplate("second-area");
        custom.Options.DocumentId = "Named";
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(custom, stream);
        var namedEntries = MultiDocumentFixture.Unzip(stream);
        var namedRoot = XDocument.Parse(Encoding.UTF8.GetString(namedEntries["OFD.xml"]));
        bodies[1].ReplaceWith(namedRoot.Root!.Elements().Single());
        foreach (var (path, data) in namedEntries.Where(entry => entry.Key != "OFD.xml"))
            entries.Add(path, data);
        entries["OFD.xml"] = Encoding.UTF8.GetBytes(root.ToString());
        var input = PathFor("multi.ofd");
        using (var zip = MultiDocumentFixture.Zip(entries))
            await File.WriteAllBytesAsync(input, zip.ToArray());

        var first = await Run("areas", "list", input);
        Assert.True(first.Code == 0, first.Stderr);
        Assert.Empty(first.Stdout);
        var second = await Run("areas", "list", input, "--document", "2");
        Assert.True(second.Code == 0, second.Stderr);
        Assert.StartsWith("second-area\t1\t", second.Stdout);
        var output = PathFor("out.ofd");
        var fill = await Run("areas", "fill", input, output, "--document", "2", "--name", "second-area", "--text", FillText);
        Assert.Equal(1, fill.Code);
        Assert.Contains("single DocBody", fill.Stderr);
        Assert.False(File.Exists(output));
    }

    private string PathFor(string name) => Path.Combine(_directory, name);

    private async Task<string> Template(string name = "name")
    {
        var path = PathFor("template.ofd");
        await using var stream = File.Create(path);
        await new OfdPackageWriter().WriteAsync(CreateTemplate(name), stream);
        return path;
    }

    private static OfdDocumentPackage CreateTemplate(string name)
    {
        var flow = new FlowDocument();
        flow.Blocks.Add(new Paragraph("HEADER") { MinimumLineHeightMillimeters = 5.125 });
        flow.Blocks.Add(new AreaHolderBlock(name, 18));
        flow.Blocks.Add(new AreaHolderBlock("logo", 15));
        flow.Blocks.Add(new Paragraph("FOOTER"));
        return flow.Render();
    }

    private static async Task<OfdDocumentPackage> Read(string path)
    {
        await using var stream = File.OpenRead(path);
        return await new OfdReader().ReadAsync(stream);
    }

    private static void AssertInside(OfdElement element, OfdNamedArea area)
    {
        const double tolerance = 0.001; // OFD coordinates are serialized to three decimal places.
        Assert.InRange(element.XMillimeters, area.XMillimeters - tolerance, area.XMillimeters + area.WidthMillimeters);
        Assert.InRange(element.YMillimeters, area.YMillimeters - tolerance, area.YMillimeters + area.HeightMillimeters);
        Assert.InRange(element.XMillimeters + element.WidthMillimeters, area.XMillimeters, area.XMillimeters + area.WidthMillimeters + tolerance);
        Assert.InRange(element.YMillimeters + element.HeightMillimeters, area.YMillimeters, area.YMillimeters + area.HeightMillimeters + tolerance);
    }

    private static async Task<(int Code, string Stdout, string Stderr)> Run(params string[] args)
    {
        var savedOut = Console.Out;
        var savedError = Console.Error;
        using var stdout = new StringWriter();
        using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);
            return (await global::Cli.RunAsync(args), stdout.ToString(), stderr.ToString());
        }
        finally
        {
            Console.SetOut(savedOut);
            Console.SetError(savedError);
        }
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
