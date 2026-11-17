using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical.Tests;

internal static class MedicalTestSupport
{
    internal static MedicalDocumentMetadata Metadata() => new()
    {
        DocumentIdentifier = "EMR-2026-0001", PatientId = "P-000123", ResponsibleParty = "Dr. Example",
        CreationTime = new DateTimeOffset(2026, 11, 17, 9, 30, 0, TimeSpan.FromHours(8)), Keywords = "出院小结 discharge summary",
        FileType = "出院小结", SecretLevel = "Internal", Title = "出院小结"
    };

    internal static string RepositoryRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return root.FullName;
    }

    internal static byte[] NotoSans() =>
        File.ReadAllBytes(Path.Combine(RepositoryRoot(), "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));

    /// <summary>Two-page Layout output with an embedded open font; <paramref name="medical"/> applies the OFD-H profile.</summary>
    internal static async Task<byte[]> LayoutOfd(bool medical, bool embed = true)
    {
        var document = new FlowDocument();
        document.Options.DefaultFontFamily = "Noto Sans";
        document.Blocks.Add(new Paragraph("Discharge summary: medical record package."));
        // A large space before pushes the second paragraph onto page 2.
        document.Blocks.Add(new Paragraph("Second page.") { SpaceBeforeMillimeters = 250 });
        var package = document.Render();
        if (embed) foreach (var font in package.Fonts) font.Data = NotoSans();
        if (medical) MedicalDocumentProfile.Apply(package.Options, Metadata());
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    internal static async Task<MedicalAssessmentReport> Assess(byte[] package, MedicalAssessmentOptions? options = null)
    {
        using var input = new MemoryStream(package, false);
        var report = await new MedicalAssessmentService().AssessAsync(input, options);
        Assert.Equal(0, input.Position);
        Assert.Equal(package, input.ToArray());
        Assert.Equal(MedicalClauseMatrix.Entries.Select(x => x.Clause), report.Clauses.Select(x => x.Clause));
        Assert.NotEqual(PreflightOutcome.Pass, report.Document.StandardOutcome);
        Assert.Equal("Incomplete", report.Document.CoverageStatus);
        return report;
    }

    internal static Dictionary<string, byte[]> Unzip(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, false), ZipArchiveMode.Read);
        var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries.Where(x => x.Name.Length > 0))
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            result[entry.FullName] = copy.ToArray();
        }
        return result;
    }

    internal static List<string> Order(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, false), ZipArchiveMode.Read);
        return zip.Entries.Select(x => x.FullName).ToList();
    }

    /// <summary>Rewrites a package in the given entry order; <paramref name="stored"/> entries use no compression.</summary>
    internal static byte[] Zip(IEnumerable<KeyValuePair<string, byte[]>> entries, Func<string, bool>? stored = null)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries)
                using (var stream = zip.CreateEntry(entry.Key, stored?.Invoke(entry.Key) == true ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open())
                    stream.Write(entry.Value);
        return output.ToArray();
    }

    /// <summary>Applies text edits to selected entries and keeps the original physical order.</summary>
    internal static byte[] Mutate(byte[] package, Action<Dictionary<string, byte[]>> change, Func<string, bool>? stored = null)
    {
        var order = Order(package);
        var entries = Unzip(package);
        change(entries);
        var ordered = order.Where(entries.ContainsKey).Select(x => new KeyValuePair<string, byte[]>(x, entries[x]))
            .Concat(entries.Where(x => !order.Contains(x.Key)));
        return Zip(ordered, stored);
    }

    internal static string Text(this Dictionary<string, byte[]> entries, string name) => Encoding.UTF8.GetString(entries[name]);
    internal static void Set(this Dictionary<string, byte[]> entries, string name, string text) => entries[name] = Encoding.UTF8.GetBytes(text);

    internal static MedicalRuleObservation Single(this MedicalAssessmentReport report, string clause, string rule) =>
        Assert.Single(report.GetClause(clause).Observations, x => x.RuleId == rule);
}
