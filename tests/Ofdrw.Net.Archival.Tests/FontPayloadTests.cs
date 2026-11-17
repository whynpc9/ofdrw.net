using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Tests;

// A1c: AC-017..AC-020 judge the embedded bytes against the characters and glyph IDs actually used.
// Variants are derived at test time from the open NotoSans fixture; no font binary is added.
public sealed class FontPayloadTests
{
    private const string Ns = "http://www.ofdspec.org";
    private const string Res = "Doc_0/PublicRes.xml";
    private const string Page = "Doc_0/Pages/P0/Content.xml";
    private const string FontPath = "Doc_0/Res/font.ttf";
    private const string B1 = "6.2.6(b)(1)", B2 = "6.2.6(b)(2)", C1 = "6.2.6(c)(1)", C2 = "6.2.6(c)(2)";

    private static string Text(string font, string text, string id = "10", string extra = "") =>
        $"<TextObject ID='{id}' Boundary='0 0 100 10' Font='{font}' Size='3'{extra}><TextCode X='0' Y='3'>{text}</TextCode></TextObject>";

    private static string ClipText(string font, string text) =>
        $"<PathObject ID='20' Boundary='0 0 10 10'><Clips><Clip><Area><Text Font='{font}' Size='3' Boundary='0 0 10 10'>" +
        $"<TextCode X='0' Y='3'>{text}</TextCode></Text></Area></Clip></Clips><AbbreviatedData>M 0 0 L 1 1</AbbreviatedData></PathObject>";

    private static string Font(string id, string? file = "font.ttf") =>
        $"<Font ID='{id}' FontName='Noto Sans' FamilyName='Noto Sans'>{(file == null ? "" : $"<FontFile>{file}</FontFile>")}</Font>";

    private static Dictionary<string, byte[]> Package(string fonts, string objects, byte[]? payload = null, string baseLoc = "Res") =>
        new(StringComparer.Ordinal)
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><CommonData><PublicRes>PublicRes.xml</PublicRes></CommonData>" +
                "<Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages></Document>"),
            [Res] = Utf8($"<Res xmlns='{Ns}' BaseLoc='{baseLoc}'><Fonts>{fonts}</Fonts></Res>"),
            [FontPath] = payload ?? NotoSans(),
            [Page] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'>{objects}</Layer></Content></Page>")
        };

    [Fact]
    public async Task CompleteFontResolvedThroughBaseLocPassesEveryFontClause()
    {
        var report = await Assess(Package(Font("1"), Text("1", "Ofdrw é") + ClipText("1", "Clip")));
        Assert.Equal(ArchivalExecutionState.Complete, report.ExecutionState);
        foreach (var clause in new[] { B1, B2, C1, C2 })
            Assert.Equal(PreflightOutcome.Pass, report.GetClause(clause).Outcome);
        var used = Assert.Single(report.GetClause(B1).Observations);
        Assert.Equal(("AC-017", "UsedGlyphsCovered", Res), (used.RuleId, used.ReasonCode, used.Evidence.Entry));
        Assert.Contains("Font ID=1", used.Evidence.Observed);
        Assert.Contains("FontFile=" + FontPath, used.Evidence.Observed);
        Assert.Contains("licensing not assessed", used.Evidence.Observed);
        Assert.Equal("AC-018", Assert.Single(report.GetClause(B2).Observations).RuleId);
        Assert.Equal(("AC-019", "OutlineTablesRetained"), Pair(report, C1));
        Assert.Equal(("AC-020", "RequiredTablesRetained"), Pair(report, C2));
        Assert.DoesNotContain("FontPayloadNotAssessed", report.Document.Gaps);
    }

    [Fact]
    public async Task FontFileResolvesAgainstTheResourceBaseNotTheZipRoot()
    {
        var entries = Package(Font("1"), Text("1", "Ofdrw"));
        entries["font.ttf"] = entries[FontPath];
        entries.Remove(FontPath);
        var report = await Assess(entries);
        var failure = Assert.Single(report.GetClause(B1).Observations);
        Assert.Equal((PreflightOutcome.Fail, "FontFileMissing"), (failure.Outcome, failure.ReasonCode));
        Assert.Contains("FontFile=" + FontPath, failure.Evidence.Observed);
        // An absolute FontFile is package-root relative.
        report = await Assess(Package(Font("1", "/Doc_0/Res/font.ttf"), Text("1", "Ofdrw"), baseLoc: "Elsewhere"));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(B1).Outcome);
    }

    [Fact]
    public async Task WriterSubsetIsAValidSubsetAndAMissingUsedGlyphFails()
    {
        var subset = await Written("Ofdrw");
        var entries = Unzip(subset);
        var file = entries.Keys.Single(x => x.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) || x.Contains("/Font_", StringComparison.Ordinal));
        Assert.True(entries[file].Length < NotoSans().Length); // a subset written by OfdPackageWriter
        var report = await Assess(entries);
        foreach (var clause in new[] { B1, C1, C2 }) Assert.Equal(PreflightOutcome.Pass, report.GetClause(clause).Outcome);

        // The subset parses, but 'Z' was never kept: the check goes beyond the presence of a FontFile.
        var content = entries.Keys.Single(x => x.EndsWith("Content.xml", StringComparison.Ordinal));
        entries[content] = Utf8(Encoding.UTF8.GetString(entries[content]).Replace(">Ofdrw<", ">OfdrwZ<", StringComparison.Ordinal));
        report = await Assess(entries);
        var failure = Assert.Single(report.GetClause(B1).Observations);
        Assert.Equal((PreflightOutcome.Fail, "UsedGlyphMissing"), (failure.Outcome, failure.ReasonCode));
        Assert.Contains("U+005A", failure.Evidence.Observed);
        Assert.Contains("FontFile=" + file, failure.Evidence.Observed);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(C1).Outcome); // tables are intact
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
    }

    [Theory]
    [InlineData(ArchivalFontExemptionPolicy.ExemptionBasisUnverified, PreflightOutcome.Unknown, "FontExemptionBasisUnverified")]
    [InlineData(ArchivalFontExemptionPolicy.NoExemptionsAccepted, PreflightOutcome.Fail, "FontPayloadNotEmbedded")]
    public async Task MissingPayloadFollowsTheExemptionPolicy(ArchivalFontExemptionPolicy policy, PreflightOutcome outcome, string reason)
    {
        var entries = Package(Font("1", null) + Font("2", null), Text("1", "Ofdrw") + ClipText("1", "Clip"));
        var report = await Assess(entries, new ArchivalAssessmentOptions { FontExemptionPolicy = policy });
        foreach (var (clause, rule) in new[] { (B1, "AC-017"), (B2, "AC-018") })
        {
            var observation = Assert.Single(report.GetClause(clause).Observations);
            Assert.Equal((rule, outcome, reason, Res), (observation.RuleId, observation.Outcome, observation.ReasonCode, observation.Evidence.Entry));
            Assert.Contains("Font ID=1", observation.Evidence.Observed); // unused font 2 is not reported
            Assert.EndsWith("{http://www.ofdspec.org}Font[1]", observation.Evidence.Selector, StringComparison.Ordinal);
            Assert.Equal(outcome, report.GetClause(clause).Outcome);
        }
        Assert.Equal(PreflightOutcome.NotApplicable, report.GetClause(C1).Outcome); // nothing embedded
        Assert.Equal(policy == ArchivalFontExemptionPolicy.NoExemptionsAccepted, !report.Document.Gaps.Contains("FontPayloadNotAssessed"));
    }

    [Theory]
    [InlineData("truncated", PreflightOutcome.Fail, "FontPayloadUnparseable")]
    [InlineData("garbage", PreflightOutcome.Unknown, "FontFormatUnsupported")]
    public async Task CorruptedOrUnknownBytesAreJudgedByContent(string variant, PreflightOutcome outcome, string reason)
    {
        var bytes = variant == "truncated" ? NotoSans().Take(4096).ToArray() : Enumerable.Range(0, 4096).Select(x => (byte)(x * 31)).ToArray();
        var report = await Assess(Package(Font("1"), Text("1", "Ofdrw"), bytes));
        foreach (var clause in new[] { B1, C1, C2 })
        {
            var observation = Assert.Single(report.GetClause(clause).Observations);
            Assert.Equal((outcome, reason, Res), (observation.Outcome, observation.ReasonCode, observation.Evidence.Entry));
        }
    }

    [Fact]
    public async Task ClippingTextUsesItsOwnClause()
    {
        var report = await Assess(Package(Font("1"), Text("1", "Ofdrw") + ClipText("1", "中")));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(B1).Outcome);
        var failure = Assert.Single(report.GetClause(B2).Observations);
        Assert.Equal(("AC-018", PreflightOutcome.Fail, "UsedGlyphMissing"), (failure.RuleId, failure.Outcome, failure.ReasonCode));
        Assert.Contains("U+4E2D", failure.Evidence.Observed);
        Assert.Contains(Page, failure.Evidence.Observed); // first use is located on the page
        // No clipping text at all: the conditional clause is not applicable.
        report = await Assess(Package(Font("1"), Text("1", "Ofdrw")));
        Assert.Equal(PreflightOutcome.NotApplicable, report.GetClause(B2).Outcome);
    }

    [Theory]
    [InlineData("wrong-id", PreflightOutcome.Fail, "UnresolvedFontReference")]
    [InlineData("gid-out-of-font", PreflightOutcome.Fail, "UsedGlyphMissing")]
    [InlineData("bad-mapping", PreflightOutcome.Unknown, "UnmodeledGlyphMapping")]
    public async Task WrongFontIdOrGlyphMappingIsLocated(string variant, PreflightOutcome outcome, string reason)
    {
        var text = variant switch
        {
            "wrong-id" => Text("99", "Ofdrw"),
            "gid-out-of-font" => "<TextObject ID='10' Boundary='0 0 100 10' Font='1' Size='3'><CGTransform CodePosition='0' CodeCount='1' GlyphCount='1'>" +
                "<Glyphs>60000</Glyphs></CGTransform><TextCode X='0' Y='3'>O</TextCode></TextObject>",
            _ => "<TextObject ID='10' Boundary='0 0 100 10' Font='1' Size='3'><CGTransform CodePosition='0' CodeCount='1' GlyphCount='2'>" +
                "<Glyphs>36</Glyphs></CGTransform><TextCode X='0' Y='3'>O</TextCode></TextObject>"
        };
        var report = await Assess(Package(Font("1"), text));
        var observation = Assert.Single(report.GetClause(B1).Observations, x => x.ReasonCode == reason);
        Assert.Equal(outcome, observation.Outcome);
        if (variant == "gid-out-of-font") Assert.Contains("GID 60000", observation.Evidence.Observed);
        else
        {
            Assert.Equal(Page, observation.Evidence.Entry);
            Assert.EndsWith("TextObject[1]", observation.Evidence.Selector, StringComparison.Ordinal);
            Assert.Contains("Object ID=10", observation.Evidence.Observed);
        }
    }

    [Fact]
    public async Task ExplicitGlyphIdsAreCheckedInsteadOfTheCmap()
    {
        // GID 36 is 'A' in NotoSans; the CGTransform covers the only character, so the cmap is not consulted for it.
        var text = "<TextObject ID='10' Boundary='0 0 100 10' Font='1' Size='3'><CGTransform CodePosition='0' CodeCount='1' GlyphCount='1'>" +
            "<Glyphs>36</Glyphs></CGTransform><TextCode X='0' Y='3'>中</TextCode></TextObject>";
        var report = await Assess(Package(Font("1"), text));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(B1).Outcome);
        Assert.Contains("glyphIds=1", Assert.Single(report.GetClause(B1).Observations).Evidence.Observed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ThreeFontResourcesSharingOneFileAreJudgedByBytesEach(bool corrupt)
    {
        var fonts = Font("1") + Font("2") + Font("3");
        var objects = Text("1", "Bold", "10", " Weight='700'") + Text("2", "Italic", "11", " Italic='true'") + Text("3", "Regular", "12");
        var report = await Assess(Package(fonts, objects, corrupt ? NotoSans().Take(4096).ToArray() : null));
        var observations = report.GetClause(B1).Observations;
        Assert.Equal(3, observations.Count);
        Assert.Equal(new[] { "Font ID=1", "Font ID=2", "Font ID=3" }, observations.Select(x => x.Evidence.Observed.Split(';')[0]));
        Assert.All(observations, x => Assert.Equal(corrupt ? PreflightOutcome.Fail : PreflightOutcome.Pass, x.Outcome));
        Assert.All(observations, x => Assert.Contains("FontFile=" + FontPath, x.Evidence.Observed));
        Assert.Equal(3, report.GetClause(C1).Observations.Count);
    }

    [Fact]
    public async Task UnmodeledTextScopeKeepsTheGap()
    {
        var entries = Package(Font("1"), Text("1", "Ofdrw") + "<CompositeObject ID='30' Boundary='0 0 1 1' ResourceID='5'/>");
        var report = await Assess(entries);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause(B1).Outcome);
        Assert.Contains(report.GetClause(B1).Observations, x => x.ReasonCode == "TextScopeIncomplete:CompositeGraphicsPresent");
        Assert.Contains("FontPayloadNotAssessed", report.Document.Gaps);
    }

    [Fact]
    public async Task FontByteBudgetMakesTheRunIncomplete()
    {
        var report = await Assess(Package(Font("1"), Text("1", "Ofdrw")), new ArchivalAssessmentOptions { MaxTotalFontPayloadBytes = 1024 });
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Contains(report.GetClause(B1).Observations, x => x.ReasonCode == "InspectionBudgetExceeded");
        Assert.Contains("FontPayloadNotAssessed", report.Document.Gaps);
        report = await Assess(Package(Font("1"), Text("1", "Ofdrw")), new ArchivalAssessmentOptions { MaxFontPayloadBytes = 1024 });
        Assert.Equal("FontPayloadBudgetExceeded", Assert.Single(report.GetClause(B1).Observations).ReasonCode);
    }

    // Review finding 1: {pageDirectory}/PageRes.xml is a fixed page resource even without a PageRes link.
    [Fact]
    public async Task SiblingPageResDeclaresPageFonts()
    {
        var entries = Package("", Text("1", "Ofdrw"));
        entries["Doc_0/Pages/P0/PageRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='/Doc_0/Res'><Fonts>{Font("1")}</Fonts></Res>");
        var report = await Assess(entries);
        var used = Assert.Single(report.GetClause(B1).Observations);
        Assert.Equal((PreflightOutcome.Pass, "Doc_0/Pages/P0/PageRes.xml"), (used.Outcome, used.Evidence.Entry));
        Assert.Contains("FontFile=" + FontPath, used.Evidence.Observed);

        // An unusable sibling table leaves the reference Unknown, never a false UnresolvedFontReference Fail.
        entries["Doc_0/Pages/P0/PageRes.xml"] = Utf8("<Unrelated/>");
        report = await Assess(entries);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause(B1).Outcome);
        Assert.DoesNotContain(report.GetClause(B1).Observations, x => x.Outcome == PreflightOutcome.Fail);
    }

    // Review finding 2: a valid composite feature the outline reader does not interpret is Unknown, not a Fail.
    [Fact]
    public async Task UninterpretedCompositeGlyphIsUnknown()
    {
        var (bytes, gid) = TestFont.PointAttachedComposite(NotoSans());
        var report = await Assess(Package(Font("1"), Glyphs("1", gid), bytes));
        foreach (var clause in new[] { B1, C1 })
        {
            var observation = Assert.Single(report.GetClause(clause).Observations);
            Assert.Equal((PreflightOutcome.Unknown, "GlyphOutlineUninterpreted"), (observation.Outcome, observation.ReasonCode));
            Assert.Contains("GID " + gid, observation.Evidence.Observed);
        }
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(C2).Outcome);
        Assert.Contains("FontPayloadNotAssessed", report.Document.Gaps);
    }

    // Review finding 3: an unreadable glyph used only by clipping text fails AC-018, not AC-017.
    [Fact]
    public async Task UnreadableGlyphFailsOnlyTheCategoryThatUsesIt()
    {
        var bytes = TestFont.TruncateGlyph(NotoSans(), 37);
        var clip = $"<PathObject ID='20' Boundary='0 0 10 10'><Clips><Clip><Area><Text Font='1' Size='3' Boundary='0 0 10 10'>" +
            "<CGTransform CodePosition='0' CodeCount='1' GlyphCount='1'><Glyphs>37</Glyphs></CGTransform><TextCode X='0' Y='3'>B</TextCode>" +
            "</Text></Area></Clip></Clips><AbbreviatedData>M 0 0 L 1 1</AbbreviatedData></PathObject>";
        var report = await Assess(Package(Font("1"), Glyphs("1", 36) + clip, bytes));
        Assert.Equal("UsedGlyphsCovered", Assert.Single(report.GetClause(B1).Observations).ReasonCode);
        var failure = Assert.Single(report.GetClause(B2).Observations);
        Assert.Equal((PreflightOutcome.Fail, "UsedGlyphMissing"), (failure.Outcome, failure.ReasonCode));
        Assert.Contains("unreadable GID 37", failure.Evidence.Observed);
        Assert.Equal("GlyphOutlineUnreadable", Assert.Single(report.GetClause(C1).Observations).ReasonCode);
    }

    // Review finding 4: the gap closes only when AC-017..AC-020 are all decided.
    [Fact]
    public async Task UndecidedClippingFontKeepsTheFontGap()
    {
        var entries = Package(Font("1") + Font("2", null), Text("1", "Ofdrw") + ClipText("2", "Clip"));
        var report = await Assess(entries);
        Assert.Equal(ArchivalExecutionState.Complete, report.ExecutionState);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(B1).Outcome);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(C1).Outcome);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(C2).Outcome);
        Assert.Equal("FontExemptionBasisUnverified", Assert.Single(report.GetClause(B2).Observations).ReasonCode);
        Assert.Contains("FontPayloadNotAssessed", report.Document.Gaps);
    }

    [Fact]
    public async Task PackageOwnedObservationsRespectTheResultBudget()
    {
        // One unresolved font reference per text object would otherwise add an observation per object without bound.
        var objects = string.Concat(Enumerable.Range(0, 200).Select(i => Text("99", "O", (100 + i).ToString())));
        var options = new ArchivalAssessmentOptions();
        options.Preflight.MaxResults = 40;
        var report = await Assess(Package(Font("1"), objects), options);
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Contains(report.GetClause(B1).Observations, x => x.ReasonCode == "InspectionBudgetExceeded");
        Assert.DoesNotContain(report.GetClause(B1).Observations, x => x.ReasonCode == "UnresolvedFontReference");
        Assert.NotEqual(PreflightOutcome.Pass, report.Document.StandardOutcome);
        // The default budget still decides the same package.
        report = await Assess(Package(Font("1"), objects));
        Assert.Equal(200, report.GetClause(B1).Observations.Count(x => x.ReasonCode == "UnresolvedFontReference"));
    }

    private static string Glyphs(string font, int gid) =>
        $"<TextObject ID='10' Boundary='0 0 100 10' Font='{font}' Size='3'><CGTransform CodePosition='0' CodeCount='1' GlyphCount='1'>" +
        $"<Glyphs>{gid}</Glyphs></CGTransform><TextCode X='0' Y='3'>A</TextCode></TextObject>";

    private static (string, string) Pair(ArchivalAssessmentReport report, string clause)
    {
        var observation = Assert.Single(report.GetClause(clause).Observations);
        return (observation.RuleId, observation.ReasonCode);
    }

    private static async Task<byte[]> Written(string text)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "1", FontName = "Noto Sans", FamilyName = "Noto Sans", Data = NotoSans() });
        var page = new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 };
        page.Elements.Add(new OfdTextElement { Text = text, FontName = "Noto Sans", FontResourceId = "1", XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 50, HeightMillimeters = 5 });
        package.Pages.Add(page);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }

    private static async Task<ArchivalAssessmentReport> Assess(Dictionary<string, byte[]> entries, ArchivalAssessmentOptions? options = null)
        => await Assess(Zip(entries), options);

    private static async Task<ArchivalAssessmentReport> Assess(byte[] package, ArchivalAssessmentOptions? options = null)
    {
        using var input = new MemoryStream(package, false);
        var report = await new ArchivalAssessmentService().AssessAsync(input, options ?? new ArchivalAssessmentOptions());
        Assert.Equal(package, input.ToArray());
        return report;
    }

    internal static byte[] NotoSans()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
    }

    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);

    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
                using (var stream = zip.CreateEntry(entry.Key).Open()) stream.Write(entry.Value);
        return output.ToArray();
    }

    private static Dictionary<string, byte[]> Unzip(byte[] package)
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
}

// In-place sfnt edits for test-time font variants (no table checksums are validated by the reader).
internal static class TestFont
{
    private static int U16(byte[] f, int o) => f[o] << 8 | f[o + 1];
    private static int U32(byte[] f, int o) => f[o] << 24 | f[o + 1] << 16 | f[o + 2] << 8 | f[o + 3];

    private static (int Glyf, Func<int, int> Offset, int Count) Layout(byte[] f)
    {
        var tables = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < U16(f, 4); i++)
            tables[Encoding.ASCII.GetString(f, 12 + 16 * i, 4)] = U32(f, 12 + 16 * i + 8);
        var head = tables["head"]; var loca = tables["loca"];
        var longLoca = U16(f, head + 50) != 0;
        return (tables["glyf"], gid => longLoca ? U32(f, loca + 4 * gid) : U16(f, loca + 2 * gid) * 2, U16(f, tables["maxp"] + 4));
    }

    // numberOfContours = 32767 makes the simple glyph record truncated: malformed data.
    internal static byte[] TruncateGlyph(byte[] bytes, int gid)
    {
        var f = (byte[])bytes.Clone();
        var (glyf, offset, _) = Layout(f);
        Assert.True(U16(f, glyf + offset(gid)) < 0x8000);
        f[glyf + offset(gid)] = 0x7F; f[glyf + offset(gid) + 1] = 0xFF;
        return f;
    }

    // Clears ARGS_ARE_XY_VALUES on the first component of an XY-placed composite: valid, but outside the reader's profile.
    internal static (byte[] Bytes, int GlyphId) PointAttachedComposite(byte[] bytes)
    {
        var f = (byte[])bytes.Clone();
        var (glyf, offset, count) = Layout(f);
        for (var gid = 1; gid < count; gid++)
        {
            var start = glyf + offset(gid);
            if (offset(gid + 1) - offset(gid) < 14 || U16(f, start) != 0xFFFF) continue;
            var flags = U16(f, start + 10);
            if ((flags & 0x0002) == 0 || (flags & 0x00C8) != 0) continue;
            flags &= ~0x0002;
            f[start + 10] = (byte)(flags >> 8); f[start + 11] = (byte)flags;
            return (f, gid);
        }
        throw new InvalidOperationException("No XY-placed composite glyph in the fixture.");
    }
}
