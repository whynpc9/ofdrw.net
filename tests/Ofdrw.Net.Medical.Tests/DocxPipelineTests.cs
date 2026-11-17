using System.IO.Compression;
using System.Security.Cryptography;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.Reader.Readers;
using static Ofdrw.Net.Medical.Tests.MedicalTestSupport;

namespace Ofdrw.Net.Medical.Tests;

/// <summary>
/// The documented route for DOCX sources: DOCX → native OFD (/2016 namespace) → read → OFD-H profile → write. Only package-level
/// declarations may change; page content and text must survive unchanged. Set OFDRW_MEDICAL_ARTIFACTS to keep the
/// outputs for page rendering and visual review.
/// </summary>
public class DocxPipelineTests
{
    private const string Fixture = "tests/fixtures/trusted-fonts/sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip";

    internal static string? SkipReason
    {
        get
        {
            var path = Path.Combine(RepositoryRoot(), Fixture);
            if (!File.Exists(path)) return "Pinned CJK font fixture missing: " + Fixture;
            using var stream = File.OpenRead(path);
            return stream.ReadByte() == 'P' && stream.ReadByte() == 'K' ? null : "Pinned CJK font fixture is not materialized (git lfs pull): " + Fixture;
        }
    }

    private static string FontDirectory()
    {
        var path = Path.Combine(RepositoryRoot(), Fixture);
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-medical-cjk-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        using var archive = ZipFile.OpenRead(path);
        archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.ExtractToFile(Path.Combine(directory, "noto-cjk.ttf"));
        return directory;
    }

    private static string Text(OfdDocumentPackage package) =>
        string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(t => t.Text));

    [PinnedFontTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task NativeDocxOutputBecomesOfdHWithoutChangingPages(bool explicitNative)
    {
        var fonts = FontDirectory();
        try
        {
            // A read package keeps its source namespace (the writer refuses to remap it), so convert with /2016 up front.
            var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false, OfdNamespace = OfdConstants.StandardNamespace };
            options.FontDirectories.Add(fonts);
            if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
            using var input = File.OpenRead(Path.Combine(RepositoryRoot(), "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"));
            using var ordinary = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, ordinary);

            var package = await new OfdReader().ReadAsync(new MemoryStream(ordinary.ToArray()));
            var sourceText = Text(package);
            var pages = package.Pages.Count;
            MedicalDocumentProfile.Apply(package.Options, Metadata());
            using var medical = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, medical);

            var reread = await new OfdReader().ReadAsync(new MemoryStream(medical.ToArray()));
            Assert.Equal(pages, reread.Pages.Count);
            Assert.Equal(sourceText, Text(reread));
            Assert.False(string.IsNullOrWhiteSpace(sourceText));

            var report = await Assess(medical.ToArray());
            foreach (var clause in new[] { "7.1(a)", "6.1.2(a)", "6.1.3(a)", "6.1.3(b)", "6.2(c)", "7.3(a)", "7.3(b)" })
                Assert.True(report.GetClause(clause).Outcome == PreflightOutcome.Pass,
                    clause + " is " + report.GetClause(clause).Outcome + ": " + string.Join("; ", report.GetClause(clause).Observations.Select(o => o.ReasonCode + " " + o.Evidence.Observed)));

            var keep = Environment.GetEnvironmentVariable("OFDRW_MEDICAL_ARTIFACTS");
            if (!string.IsNullOrEmpty(keep))
            {
                var mode = explicitNative ? "native" : "default";
                Directory.CreateDirectory(keep);
                File.WriteAllBytes(Path.Combine(keep, $"generated-layout-{mode}.ofd"), ordinary.ToArray());
                File.WriteAllBytes(Path.Combine(keep, $"generated-layout-{mode}-ofd-h.ofd"), medical.ToArray());
                File.WriteAllText(Path.Combine(keep, $"generated-layout-{mode}-ofd-h.sha256"),
                    Convert.ToHexString(SHA256.HashData(medical.ToArray())).ToLowerInvariant() + "\n");
            }
        }
        finally { Directory.Delete(fonts, true); }
    }
}

/// <summary>A theory that is reported as skipped, not passed, when the pinned font fixture is unavailable.</summary>
internal sealed class PinnedFontTheoryAttribute : TheoryAttribute
{
    public PinnedFontTheoryAttribute() { if (DocxPipelineTests.SkipReason is { } reason) Skip = reason; }
}
