using Ofdrw.Net.Converter.Docx.Internal;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class LibreOfficeFontStagingTests
{
    [Fact]
    public void PortableProfile_ShouldResolveDataSettingsWithoutIsolation()
    {
        var root = Path.Combine(Path.GetTempPath(), "ofdrw-lo-portable-" + Guid.NewGuid().ToString("N"));
        var program = Path.Combine(root, "LibreOfficePortable", "App", "libreoffice", "program");
        Directory.CreateDirectory(program);
        var executable = Path.Combine(program, "soffice.exe");
        File.WriteAllBytes(executable, Array.Empty<byte>());
        try
        {
            Assert.False(LibreOfficeExecutableResolver.ShouldIsolateUserProfile(executable));
            Assert.Equal(
                Path.Combine(root, "LibreOfficePortable", "Data", "settings"),
                LibreOfficeExecutableResolver.TryGetPortableUserProfile(executable));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Stage_ShouldCopyEveryFontFromALargeConfiguredDirectory()
    {
        var root = Path.Combine(Path.GetTempPath(), "ofdrw-lo-fonts-" + Guid.NewGuid().ToString("N"));
        var fonts = Path.Combine(root, "Fonts");
        var profile = Path.Combine(root, "settings");
        Directory.CreateDirectory(fonts);
        try
        {
            File.WriteAllBytes(Path.Combine(fonts, "simhei.ttf"), new byte[] { 1, 2, 3 });
            File.WriteAllBytes(Path.Combine(fonts, "corporate.ttf"), new byte[] { 4, 5, 6 });
            for (var index = 0; index < 40; index++)
                File.WriteAllBytes(Path.Combine(fonts, $"filler-{index}.ttf"), new byte[] { 7 });

            var options = new DocxConversionOptions();
            options.FontDirectories.Add(fonts);
            LibreOfficeFontStager.Stage(profile, options);

            var staged = Path.Combine(profile, "user", "fonts");
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(staged, "simhei.ttf")));
            Assert.Equal(new byte[] { 4, 5, 6 }, File.ReadAllBytes(Path.Combine(staged, "corporate.ttf")));
            Assert.True(File.Exists(Path.Combine(staged, "filler-0.ttf")));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Stage_ShouldReplaceASameLengthFontWhenContentChanges()
    {
        var root = Path.Combine(Path.GetTempPath(), "ofdrw-lo-stale-" + Guid.NewGuid().ToString("N"));
        var fonts = Path.Combine(root, "Fonts");
        var profile = Path.Combine(root, "settings");
        Directory.CreateDirectory(fonts);
        try
        {
            var source = Path.Combine(fonts, "corporate.ttf");
            File.WriteAllBytes(source, new byte[] { 1, 2, 3 });
            var options = new DocxConversionOptions();
            options.FontDirectories.Add(fonts);
            LibreOfficeFontStager.Stage(profile, options);
            File.WriteAllBytes(source, new byte[] { 4, 5, 6 });
            LibreOfficeFontStager.Stage(profile, options);

            Assert.Equal(new byte[] { 4, 5, 6 }, File.ReadAllBytes(Path.Combine(profile, "user", "fonts", "corporate.ttf")));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ShouldRestrictPreferredCjkOnlyForLargePlatformFontDirectories()
    {
        var custom = Path.Combine(Path.GetTempPath(), "ofdrw-corporate-fonts");
        Assert.False(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(custom, 80));
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        if (string.IsNullOrEmpty(windows))
            return;
        var fonts = Path.Combine(windows, "Fonts");
        if (!Directory.Exists(fonts))
            return;
        Assert.False(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(fonts, 32));
        Assert.True(LibreOfficeFontStager.ShouldRestrictToPreferredCjk(fonts, 33));
    }
}
