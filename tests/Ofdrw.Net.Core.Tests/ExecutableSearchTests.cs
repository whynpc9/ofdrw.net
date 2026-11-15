using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class ExecutableSearchTests
{
    private static readonly string BaseDirectory = Path.Combine(Path.GetTempPath(), "app");
    private static readonly string FirstDirectory = Path.Combine(Path.GetTempPath(), "first");
    private static readonly string SecondDirectory = Path.Combine(Path.GetTempPath(), "second");

    [Fact]
    public void BareName_IsFoundInPathDirectory()
    {
        var expected = Path.Combine(SecondDirectory, "tool");
        Assert.Equal(expected, Find("tool", PathList(FirstDirectory, SecondDirectory), false, expected));
    }

    [Fact]
    public void BareName_InWorkingDirectoryOnly_IsNotFound()
    {
        Assert.Null(Find("tool", null, false, Path.Combine(Directory.GetCurrentDirectory(), "tool")));
    }

    [Fact]
    public void EmptyAndRelativePathEntries_AreSkipped()
    {
        var expected = Path.Combine(SecondDirectory, "tool");
        var relative = Path.Combine("relative", "dir");
        Assert.Equal(expected, Find("tool", PathList("", " ", ".", relative, SecondDirectory), false,
            Path.Combine(Directory.GetCurrentDirectory(), "tool"), Path.GetFullPath(Path.Combine(relative, "tool")), expected));
    }

    [Fact]
    public void BaseDirectory_IsSearchedFirst()
    {
        var expected = Path.Combine(BaseDirectory, "tool");
        Assert.Equal(expected, Find("tool", FirstDirectory, false, expected, Path.Combine(FirstDirectory, "tool")));
    }

    [Fact]
    public void ExplicitRelativePath_PreservesPlatformBehavior()
    {
        var name = Path.Combine("sub", "tool");
        Assert.Equal(Path.GetFullPath(name), Find(name, null, false));
        Assert.Equal(name, Find(name, null, true));
    }

    [Fact]
    public void RootedName_IsReturnedUnchanged()
    {
        var name = Path.Combine(FirstDirectory, "..", "tool");
        Assert.Equal(name, Find(name, null, false));
    }

    [Fact]
    public void WindowsMode_ProbesExeThenCom()
    {
        var exe = Path.Combine(FirstDirectory, "pdftoppm.exe");
        var com = Path.Combine(FirstDirectory, "pdftoppm.com");
        Assert.Equal(com, Find("pdftoppm", FirstDirectory, true, com));
        Assert.Equal(exe, Find("pdftoppm", FirstDirectory, true, exe, com));
    }

    [Fact]
    public void WindowsMode_NameWithExtension_IsNotExtended()
    {
        Assert.Null(Find("soffice.com", FirstDirectory, true, Path.Combine(FirstDirectory, "soffice.com.exe")));
    }

    [Fact]
    public void WindowsMode_UnquotesPathEntry()
    {
        var expected = Path.Combine(FirstDirectory, "soffice.com");
        Assert.Equal(expected, Find("soffice.com", $"\"{FirstDirectory}\"", true, expected));
    }

    [Fact]
    public void MalformedPathEntry_IsSkipped()
    {
        var expected = Path.Combine(SecondDirectory, "tool");
        Assert.Equal(expected, Find("tool", PathList(FirstDirectory + "\0", SecondDirectory), false, expected));
    }

    private static string PathList(params string[] entries) => string.Join(Path.PathSeparator, entries);

    private static string? Find(string name, string? path, bool windows, params string[] files) =>
        ExecutableSearch.Find(name, BaseDirectory, path, windows, new HashSet<string>(files).Contains);
}
