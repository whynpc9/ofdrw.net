using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Tests;

public sealed class PrivateTemporaryDirectoryTests
{
    [Fact]
    public void Create_ReturnsAnOwnerOnlyDirectory()
    {
        var directory = PrivateTemporaryDirectory.Create("ofdrw-test-private-");
        try
        {
            Assert.True(Directory.Exists(directory));
            Assert.StartsWith(Path.Combine(Path.GetTempPath(), "ofdrw-test-private-"), directory);
            if (!OperatingSystem.IsWindows())
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute, File.GetUnixFileMode(directory));
        }
        finally { PrivateTemporaryDirectory.TryDelete(directory); }
    }

    [Fact]
    public void Create_ReturnsDistinctPaths()
    {
        var first = PrivateTemporaryDirectory.Create("ofdrw-test-private-");
        var second = PrivateTemporaryDirectory.Create("ofdrw-test-private-");
        try { Assert.NotEqual(first, second); }
        finally { PrivateTemporaryDirectory.TryDelete(first); PrivateTemporaryDirectory.TryDelete(second); }
    }

    [Fact]
    public void TryDelete_RemovesContentsAndToleratesMissingPaths()
    {
        var directory = PrivateTemporaryDirectory.Create("ofdrw-test-private-");
        File.WriteAllText(Path.Combine(directory, "copy.tmp"), "document");

        PrivateTemporaryDirectory.TryDelete(directory);

        Assert.False(Directory.Exists(directory));
        PrivateTemporaryDirectory.TryDelete(directory);
        PrivateTemporaryDirectory.TryDelete(null);
    }

    [Theory]
    [InlineData("nested/prefix-")]
    [InlineData("nested\\prefix-")]
    [InlineData("")]
    public void Create_RejectsPrefixesThatAreNotPlainNames(string prefix) =>
        Assert.Throws<ArgumentException>(() => PrivateTemporaryDirectory.Create(prefix));
}
