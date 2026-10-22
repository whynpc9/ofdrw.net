using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Packaging.Tests;

public class ArchivalStructuralCoverageTests
{
    private const string Ns = "http://www.ofdspec.org";
    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);
    private static Dictionary<string, byte[]> Seed(string documentExtra = "", string pageExtra = "", string objects = "<PathObject/>") =>
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages>{documentExtra}</Document>"),
            ["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'>{pageExtra}<Content><Layer>{objects}</Layer></Content></Page>")
        };
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
                using (var stream = zip.CreateEntry(entry.Key).Open()) stream.Write(entry.Value);
        return output.ToArray();
    }
    private static async Task<ArchivalCoverageReport> Check(Dictionary<string, byte[]> entries)
    {
        using var input = new MemoryStream(Zip(entries), false);
        var original = SHA256.HashData(input.ToArray());
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.Equal(0, input.Position);
        Assert.Equal(original, SHA256.HashData(input.ToArray()));
        Assert.Equal(12, report.RuleResults.Count);
        return report;
    }
    private static ArchivalCoverageRuleResult Rule(ArchivalCoverageReport report, string id) =>
        Assert.Single(report.RuleResults, x => x.Rule.RuleId == id);

    [Theory]
    [InlineData("<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>", PreflightOutcome.Fail)]
    [InlineData("<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>", PreflightOutcome.Pass)]
    [InlineData("<x:Actions xmlns:x='urn:vendor'><x:Action/></x:Actions>", PreflightOutcome.Unknown)]
    public async Task PageActionTypeIsCheckedWithoutChangingPage(string actions, PreflightOutcome expected)
    {
        var report = await Check(Seed(pageExtra: actions));
        Assert.Equal(expected, Rule(report, "AC-004").Outcome);
        Assert.Equal("6.2.3(c)", Rule(report, "AC-004").Rule.Clause);
        Assert.Equal("Incomplete", report.CoverageStatus);
    }

    [Fact]
    public async Task KnownPageActionFailureDominatesUnresolvedSiblingPage()
    {
        var entries = Seed(pageExtra: "<Actions><Action><Sound ResourceID='9'/></Action></Actions>");
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>", "<Page ID='2' BaseLoc='missing.xml'/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Fail, report.StandardOutcome);
    }

    [Theory]
    [InlineData("<URI URI='https://example.invalid'/>", PreflightOutcome.Fail)]
    [InlineData("<Goto><Dest Type='Fit' PageID='1'/></Goto>", PreflightOutcome.Pass)]
    [InlineData("<v:Future xmlns:v='urn:vendor'/>", PreflightOutcome.Unknown)]
    public async Task NestedOutlineActionTypesAreChecked(string kind, PreflightOutcome expected)
    {
        var outline = $"<Outlines><OutlineElem Title='Parent'><OutlineElem Title='Child'><Actions><Action>{kind}</Action></Actions></OutlineElem></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline));
        Assert.Equal(expected, Rule(report, "AC-005").Outcome);
        Assert.Equal("6.2.5(a)", Rule(report, "AC-005").Rule.Clause);
    }

    [Theory]
    [InlineData("<Goto><Dest Type='Fit' PageID='1'/></Goto>", "", PreflightOutcome.Pass)]
    [InlineData("<Goto><Dest Type='Fit' PageID='99'/></Goto>", "", PreflightOutcome.Fail)]
    [InlineData("<Goto><Bookmark Name='missing'/></Goto>", "", PreflightOutcome.Fail)]
    [InlineData("<Goto><Bookmark Name='kept'/></Goto>", "<Bookmarks><Bookmark Name='kept'><Dest Type='Fit' PageID='1'/></Bookmark></Bookmarks>", PreflightOutcome.Pass)]
    [InlineData("<Goto><Bookmark Name='kept'/></Goto>", "<Bookmarks><Bookmark Name='kept'><Dest Type='Fit' PageID='99'/></Bookmark></Bookmarks>", PreflightOutcome.Fail)]
    public async Task OutlineGotoTargetMustExist(string gotoXml, string bookmarks, PreflightOutcome expected)
    {
        var outline = $"<Outlines><OutlineElem Title='A'><Actions><Action>{gotoXml}</Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline + bookmarks));
        Assert.Equal(expected, Rule(report, "AC-006").Outcome);
        Assert.Equal("6.2.5(b)", Rule(report, "AC-006").Rule.Clause);
    }

    [Theory]
    [InlineData("<ImageObject Interpolate='true'/>", PreflightOutcome.Fail)]
    [InlineData("<ImageObject Interpolate='false'/>", PreflightOutcome.Pass)]
    [InlineData("<ImageObject/>", PreflightOutcome.NotApplicable)]
    [InlineData("<ImageObject xmlns:v='urn:vendor' v:Interpolate='true'/>", PreflightOutcome.Unknown)]
    public async Task InterpolateIsConditionalOnAttributeUse(string image, PreflightOutcome expected)
    {
        var report = await Check(Seed(objects: image));
        Assert.Equal(expected, Rule(report, "AC-007").Outcome);
        Assert.Equal("6.5(b)", Rule(report, "AC-007").Rule.Clause);
    }

    [Theory]
    [InlineData("<TextObject Size='3'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Pass)]
    [InlineData("<TextObject><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
    [InlineData("<TextObject Size='0'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
    public async Task TextObjectSizeIsExplicitAndPositive(string text, PreflightOutcome expected)
    {
        var report = await Check(Seed(objects: text));
        Assert.Equal(expected, Rule(report, "AC-008").Outcome);
        Assert.Equal("6.6(c)", Rule(report, "AC-008").Rule.Clause);
    }

    [Fact]
    public async Task KnownGraphicFailuresDominateAnUnresolvedSiblingPage()
    {
        var entries = Seed(objects: "<ImageObject Interpolate='true'/><TextObject><TextCode>missing size</TextCode></TextObject>");
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>", "<Page ID='2' BaseLoc='missing.xml'/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("ReadOnly='true' NoZoom='true' NoRotate='true'", PreflightOutcome.Pass)]
    [InlineData("ReadOnly='false' NoZoom='true' NoRotate='true'", PreflightOutcome.Fail)]
    [InlineData("ReadOnly='true' NoZoom='true'", PreflightOutcome.Fail)]
    public async Task RetainedAnnotationRequiresAllThreeFlags(string flags, PreflightOutcome expected)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' {flags}><Appearance><PathObject/></Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-009").Outcome);
        Assert.Equal("6.10.1(a)(3)", Rule(report, "AC-009").Rule.Clause);
    }

    [Fact]
    public async Task UnknownAnnotationPageDoesNotBecomeNotApplicable()
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/missing.xml</Annotations>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-009").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
    }

    [Fact]
    public async Task RetainedObjectActionIsKnownFailureDespiteMissingMigrationProvenance()
    {
        var report = await Check(Seed(objects: "<PathObject><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></PathObject>"));
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-010").Outcome);
        Assert.Equal("6.3.3(b)", Rule(report, "AC-010").Rule.Clause);
        Assert.Equal(PreflightOutcome.Fail, report.StandardOutcome);
        var clean = await Check(Seed());
        Assert.Equal(PreflightOutcome.Unknown, Rule(clean, "AC-010").Outcome);
    }

    [Fact]
    public async Task UnsupportedPageGraphicsPreventAnAbsenceClaim()
    {
        var report = await Check(Seed(objects: "<v:ImageObject xmlns:v='urn:vendor' Interpolate='true'/>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("<Layer Name='body'/><Layer Name='body'/>", PreflightOutcome.Fail)]
    [InlineData("<Layer Name='body'/><Layer Name='background'/>", PreflightOutcome.Pass)]
    [InlineData("<Layer xmlns:v='urn:vendor' v:Name='opaque'/>", PreflightOutcome.Unknown)]
    public async Task LayerNamesAreUniqueWithinRecognizedPage(string layers, PreflightOutcome expected)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content>{layers}</Content></Page>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-011").Outcome);
        Assert.Equal("6.2.3(d)", Rule(report, "AC-011").Rule.Clause);
    }

    [Theory]
    [InlineData("<PageBlock ID='3'><PageBlock ID='4'><PathObject/></PageBlock></PageBlock>", PreflightOutcome.Fail)]
    [InlineData("<PageBlock ID='3'><PathObject/></PageBlock>", PreflightOutcome.Pass)]
    [InlineData("<v:PageBlock xmlns:v='urn:vendor'/>", PreflightOutcome.Unknown)]
    public async Task RetainedAnnotationAppearanceHasNoNestedPageBlocks(string appearance,
        PreflightOutcome expected)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance>{appearance}</Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-012").Outcome);
        Assert.Equal("6.10.1(a)(2)", Rule(report, "AC-012").Rule.Clause);
    }

    [Fact]
    public async Task OpaqueWrapperBlocksLayerClaimAndSelectorRedactsItsNamespace()
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><v:Wrapper xmlns:v='https://user:secret@example.invalid/scope'><Layer Name='hidden'><TextObject><TextCode>missing size</TextCode></TextObject></Layer></v:Wrapper></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-011").Outcome);
        var text = Rule(report, "AC-008");
        Assert.Equal(PreflightOutcome.Fail, text.Outcome);
        Assert.DoesNotContain("secret", text.Evidence.Selector, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unsupported-namespace", text.Evidence.Selector, StringComparison.Ordinal);
    }
}
