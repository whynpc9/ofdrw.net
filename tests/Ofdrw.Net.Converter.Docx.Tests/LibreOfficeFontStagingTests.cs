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
    public void Stage_ShouldCopyCjkFontsFromALargeDirectoryIntoThePortableProfile()
    {
        var root = Path.Combine(Path.GetTempPath(), "ofdrw-lo-fonts-" + Guid.NewGuid().ToString("N"));
        var fonts = Path.Combine(root, "Fonts");
        var profile = Path.Combine(root, "settings");
        Directory.CreateDirectory(fonts);
        try
        {
            File.WriteAllBytes(Path.Combine(fonts, "simhei.ttf"), new byte[] { 1, 2, 3 });
            File.WriteAllBytes(Path.Combine(fonts, "ignored-latin.ttf"), new byte[] { 4, 5, 6 });
            for (var index = 0; index < 40; index++)
                File.WriteAllBytes(Path.Combine(fonts, $"filler-{index}.ttf"), new byte[] { 7 });

            var options = new DocxConversionOptions();
            options.FontDirectories.Add(fonts);
            LibreOfficeFontStager.Stage(profile, options);

            var staged = Path.Combine(profile, "user", "fonts", "simhei.ttf");
            Assert.True(File.Exists(staged));
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(staged));
            Assert.False(File.Exists(Path.Combine(profile, "user", "fonts", "ignored-latin.ttf")));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
