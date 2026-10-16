using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit023RegressionTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "ofdrw-audit023-" + Guid.NewGuid().ToString("N"));
    private readonly List<DocxConversionDiagnostic> _diagnostics = new();

    public Audit023RegressionTests() => Directory.CreateDirectory(_root);

    [Fact]
    public void Audit023_EnumerationCancellation_StopsSingleDirectoryBeforeBufferingAndSorting()
    {
        for (var index = 0; index < 256; index++) File.WriteAllText(Path.Combine(_root, $"{index:D3}.ttf"), "fixture");
        using var cancellation = new CancellationTokenSource();
        var visited = 0;
        var disposed = false;
        IEnumerable<string> Entries(string directory)
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(directory))
                {
                    if (++visited == 8) cancellation.Cancel();
                    yield return file;
                }
            }
            finally { disposed = true; }
        }
        var exception = Assert.Throws<OperationCanceledException>(() =>
            DocxFontCatalog.DiscoverFontFiles(_root, cancellation.Token, _diagnostics, Entries));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(8, visited);
        Assert.True(disposed);
        Assert.Empty(_diagnostics);
    }

    [Fact]
    public void Audit023_EnumerationCancellation_StopsBetweenDirectoryEntries()
    {
        using var cancellation = new CancellationTokenSource();
        var visited = 0;
        IEnumerable<string> Children(string directory)
        {
            for (var index = 0; index < 16; index++)
            {
                var child = Directory.CreateDirectory(Path.Combine(_root, $"child-{index}"));
                if (++visited == 3) cancellation.Cancel();
                yield return child.FullName;
            }
        }
        Assert.Throws<OperationCanceledException>(() => DocxFontCatalog.DiscoverFontFiles(
            _root, cancellation.Token, _diagnostics, enumerateDirectories: Children));
        Assert.Equal(3, visited);
        Assert.Empty(_diagnostics);
    }

    [Fact]
    public void Audit023_DeterministicOrder_FiltersNonFontsAndKeepsOrdinalPreference()
    {
        var child = Directory.CreateDirectory(Path.Combine(_root, "child")).FullName;
        foreach (var name in new[] { "z.ttf", "A.otf", "a.TTF", "skip.txt", "skip.ttf.bak" })
            File.WriteAllText(Path.Combine(_root, name), "fixture");
        File.WriteAllText(Path.Combine(child, "collection.TTC"), "fixture");
        Assert.Equal(new[] { "A.otf", "a.TTF", Path.Combine("child", "collection.TTC"), "z.ttf" }
            .Select(name => Path.Combine(_root, name)),
            DocxFontCatalog.DiscoverFontFiles(_root, CancellationToken.None, _diagnostics));
        Assert.Empty(_diagnostics);
    }

    [Fact]
    public void Audit023_LinkedChild_SkipsCyclesAndRepeatedTargetsWithDiagnostics()
    {
        var child = Directory.CreateDirectory(Path.Combine(_root, "real")).FullName;
        var font = Path.Combine(child, "face.ttf");
        File.WriteAllText(font, "fixture");
        Directory.CreateSymbolicLink(Path.Combine(child, "cycle"), _root);
        Directory.CreateSymbolicLink(Path.Combine(_root, "repeat"), child);
        Assert.Equal(new[] { font }, DocxFontCatalog.DiscoverFontFiles(_root, CancellationToken.None, _diagnostics));
        Assert.Equal(2, _diagnostics.Count(diagnostic => diagnostic.Code == "DOCX_FONT_DIRECTORY_LINK_SKIPPED"));
    }

    [Fact]
    public void Audit023_RootLink_TraversesExplicitlyConfiguredRoot()
    {
        var target = Directory.CreateDirectory(Path.Combine(_root, "target")).FullName;
        File.WriteAllText(Path.Combine(target, "face.ttf"), "fixture");
        var linkedRoot = Path.Combine(_root, "configured");
        Directory.CreateSymbolicLink(linkedRoot, target);
        Assert.Equal(new[] { Path.Combine(linkedRoot, "face.ttf") },
            DocxFontCatalog.DiscoverFontFiles(linkedRoot, CancellationToken.None, _diagnostics));
        Assert.Empty(_diagnostics);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Audit023_UnreadableDirectory_HandlesCreationAndLazyIterationFailures(bool lazy)
    {
        IEnumerable<string> Broken(string directory)
        {
            if (lazy) return FailDuringIteration();
            throw new UnauthorizedAccessException("controlled creation failure");
        }
        IEnumerable<string> FailDuringIteration()
        {
            yield return Path.Combine(_root, "visible.ttf");
            throw new IOException("controlled MoveNext failure");
        }
        var files = DocxFontCatalog.DiscoverFontFiles(_root, CancellationToken.None, _diagnostics, Broken);
        Assert.Equal(lazy ? new[] { Path.Combine(_root, "visible.ttf") } : Array.Empty<string>(), files);
        Assert.Equal("DOCX_FONT_DIRECTORY_READ_FAILED", Assert.Single(_diagnostics).Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Audit023_UnreadableDirectory_DoesNotSwallowCallerCancellation(bool lazy)
    {
        using var cancellation = new CancellationTokenSource();
        IEnumerable<string> Broken(string directory)
        {
            if (lazy) return FailDuringIteration();
            cancellation.Cancel();
            throw new UnauthorizedAccessException("controlled creation failure");
        }
        IEnumerable<string> FailDuringIteration()
        {
            yield return Path.Combine(_root, "visible.ttf");
            cancellation.Cancel();
            throw new IOException("controlled MoveNext failure");
        }
        var exception = Assert.Throws<OperationCanceledException>(() =>
            DocxFontCatalog.DiscoverFontFiles(_root, cancellation.Token, _diagnostics, Broken));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Empty(_diagnostics);
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);
}
