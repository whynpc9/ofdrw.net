using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class PageEditCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-page-edit-" + Guid.NewGuid().ToString("N"));

    public PageEditCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task DeletePages_RemovesSelected()
    {
        var (input, output) = (await Fixture(), PathFor("out.ofd"));

        Assert.Equal(0, await Run("delete-pages", input, output, "--pages", "2,4"));

        Assert.Equal(new[] { "P1", "P3" }, await Texts(output));
    }

    [Fact]
    public async Task DeletePages_InPlace()
    {
        var input = await Fixture();

        Assert.Equal(0, await Run("delete-pages", input, input, "--pages", "1,1"));

        Assert.Equal(new[] { "P2", "P3", "P4" }, await Texts(input));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task DeletePages_All_Fails()
    {
        var (input, output) = (await Fixture(), PathFor("out.ofd"));
        await File.WriteAllTextAsync(output, "unchanged");

        Assert.Equal(1, await Run("delete-pages", input, output, "--pages", "1-4"));

        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task DeletePages_OutOfRange_Fails() =>
        Assert.Equal(1, await Run("delete-pages", await Fixture(), PathFor("out.ofd"), "--pages", "9"));

    [Fact]
    public async Task DeletePages_RequiresPages() =>
        Assert.Equal(1, await Run("delete-pages", await Fixture(), PathFor("out.ofd")));

    [Fact]
    public async Task Crop_SetsBox()
    {
        var (input, output) = (await Fixture(), PathFor("out.ofd"));

        Assert.Equal(0, await Run("crop", input, output, "--box", "10,20,80,60", "--pages", "1"));

        var pages = await Pages(output);
        Assert.Equal(new[] { 10d, 20, 80, 60 }, new[] { pages[0].XMillimeters, pages[0].YMillimeters, pages[0].WidthMillimeters, pages[0].HeightMillimeters });
        Assert.All(pages.Skip(1), page => Assert.Equal((0d, 0d, 210d, 297d), (page.XMillimeters, page.YMillimeters, page.WidthMillimeters, page.HeightMillimeters)));
    }

    [Fact]
    public async Task Crop_AllPagesByDefault()
    {
        var (input, output) = (await Fixture(), PathFor("out.ofd"));

        Assert.Equal(0, await Run("crop", input, output, "--box", "5,5,100,100"));

        Assert.All(await Pages(output), page => Assert.Equal((5d, 5d, 100d, 100d), (page.XMillimeters, page.YMillimeters, page.WidthMillimeters, page.HeightMillimeters)));
    }

    [Theory]
    [InlineData("1,2,3")]
    [InlineData("1,2,0,5")]
    [InlineData("1,2,NaN,5")]
    [InlineData("1,2,Infinity,5")]
    public async Task Crop_InvalidBox(string box) =>
        Assert.Equal(1, await Run("crop", await Fixture(), PathFor("out.ofd"), "--box", box));

    [Fact]
    public async Task Box_IsRejectedForOtherCommands() =>
        Assert.Equal(1, await Run("reorder", await Fixture(), PathFor("out.ofd"), "--pages", "1,2,3,4", "--box", "1,1,1,1"));

    private string PathFor(string name) => Path.Combine(_directory, name);

    private async Task<string> Fixture()
    {
        var path = PathFor("input.ofd");
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 4; index++)
            package.Pages.Add(new OfdPage
            {
                Index = index, WidthMillimeters = 210, HeightMillimeters = 297,
                Elements = { new OfdTextElement { Text = $"P{index + 1}", XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 20, HeightMillimeters = 8 } }
            });
        await using (var output = File.Create(path)) await new OfdPackageWriter().WriteAsync(package, output);
        return path;
    }

    private static async Task<IReadOnlyList<OfdPage>> Pages(string path)
    {
        await using var stream = File.OpenRead(path);
        return (await new OfdReader().ReadAsync(stream)).Pages.OrderBy(page => page.Index).ToList();
    }

    private static async Task<string[]> Texts(string path) =>
        (await Pages(path)).Select(page => page.Elements.OfType<OfdTextElement>().Single().Text).ToArray();

    private static async Task<int> Run(params string[] args)
    {
        var saved = Console.Error; var savedOut = Console.Out;
        using var errors = new StringWriter(); using var output = new StringWriter();
        try { Console.SetError(errors); Console.SetOut(output); return await global::Cli.RunAsync(args); }
        finally { Console.SetError(saved); Console.SetOut(savedOut); }
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
