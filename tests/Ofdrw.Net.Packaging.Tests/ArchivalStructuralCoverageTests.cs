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
    private static async Task<ArchivalCoverageReport> Check(Dictionary<string, byte[]> entries,
        ArchivalPreflightOptions? options = null)
    {
        using var input = new MemoryStream(Zip(entries), false);
        var original = SHA256.HashData(input.ToArray());
        var report = await new ArchivalCoverageService().CheckAsync(input, options ?? new());
        Assert.Equal(0, input.Position);
        Assert.Equal(original, SHA256.HashData(input.ToArray()));
        Assert.Equal(12, report.RuleResults.Count);
        return report;
    }
    private static ArchivalCoverageRuleResult Rule(ArchivalCoverageReport report, string id) =>
        Assert.Single(report.RuleResults, x => x.Rule.RuleId == id);

    [Fact]
    public async Task OpaqueDocumentChildTaintsStructuralScopeButKnownFailureWins()
    {
        var entries = Seed(documentExtra: "<Future><Outlines><OutlineElem Title='hidden'/></Outlines></Future>",
            pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>",
            objects: "<ImageObject Interpolate='false'/><TextObject Size='3'/>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-005", "AC-006", "AC-007", "AC-008", "AC-009", "AC-011", "AC-012" })
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, id).Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Actions><Action><URI URI='https://example.invalid'/></Action></Actions><Content><Layer><ImageObject Interpolate='true'/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-007").Outcome);
    }

    [Theory]
    [InlineData("<!--opaque-->")]
    [InlineData("<?vendor opaque?>")]
    public async Task OpaqueDocRootLocatorPreventsStructuralPass(string payload)
    {
        var entries = Seed(pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>");
        var root = Encoding.UTF8.GetString(entries["OFD.xml"]);
        entries["OFD.xml"] = Utf8(root.Replace("<DocRoot>Doc_0/Document.xml</DocRoot>",
            $"<DocRoot>{payload}Doc_0/Document.xml</DocRoot>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        Assert.Equal("DocumentScopeUnavailable", Rule(report, "AC-004").ReasonCode);
    }

    [Theory]
    [InlineData("</DocBody>", "<FutureDocBody/>")]
    [InlineData("</DocRoot>", "<FutureRoot/>")]
    public async Task OpaqueOfdEnvelopePreventsStructuralPass(string anchor, string payload)
    {
        var entries = Seed(pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>");
        var root = Encoding.UTF8.GetString(entries["OFD.xml"]);
        entries["OFD.xml"] = Utf8(root.Replace(anchor, anchor + payload, StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("<DocInfo><Outlines><OutlineElem Title='hidden'/></Outlines></DocInfo>")]
    [InlineData("<Versions><Outlines><OutlineElem Title='hidden'/></Outlines></Versions>")]
    [InlineData("<Signatures><Outlines><OutlineElem Title='hidden'/></Outlines></Signatures>")]
    public async Task OpaqueDocBodyChildCannotHideOutlineScope(string child)
    {
        var entries = Seed();
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody>{child}<DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task RecognizedDocBodyMetadataDoesNotTaintOutlineAbsence()
    {
        var entries = Seed();
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocInfo><DocID>123</DocID><CustomDatas><CustomData Name='owner'>local</CustomData></CustomDatas></DocInfo><DocRoot>Doc_0/Document.xml</DocRoot><Versions><Version ID='v1' Index='1' BaseLoc='Doc_0/Version.xml'/></Versions><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.NotApplicable, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.NotApplicable, Rule(report, "AC-006").Outcome);
    }

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

    [Fact]
    public async Task TemplatePageActionIsPartOfPageActionScope()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='8' BaseLoc='Templates/Content.xml'/></CommonData>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Actions><Action><URI URI='https://example.invalid'/></Action></Actions><Content><Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
        entries.Remove("Doc_0/Templates/Content.xml");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("<v:CommonData xmlns:v='urn:vendor'><v:TemplatePage BaseLoc='hidden.xml'/></v:CommonData>")]
    [InlineData("<CommonData><v:TemplatePage xmlns:v='urn:vendor' BaseLoc='hidden.xml'/></CommonData>")]
    [InlineData("<CommonData><Wrapper><TemplatePage BaseLoc='hidden.xml'/></Wrapper></CommonData>")]
    public async Task UnreadTemplateDeclarationsPreventPageActionPass(string commonData)
    {
        var report = await Check(Seed(documentExtra: commonData,
            pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        report = await Check(Seed(documentExtra: commonData,
            pageExtra: "<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>"));
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("<CommonData vendor='x'><MaxUnitID>2</MaxUnitID></CommonData>")]
    [InlineData("<CommonData><FutureTemplate BaseLoc='hidden.xml'/></CommonData>")]
    public async Task OpaqueCommonDataPreventsPageActionPass(string commonData)
    {
        var actions = "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>";
        var report = await Check(Seed(documentExtra: commonData, pageExtra: actions));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("<MaxUnitID><v:TemplatePage xmlns:v='urn:vendor' BaseLoc='hidden.xml'/></MaxUnitID>")]
    [InlineData("<PageArea><Future><TemplatePage ID='9' BaseLoc='hidden.xml'/></Future></PageArea>")]
    public async Task OpaqueCommonDataChildTaintsPageContentRules(string child)
    {
        var entries = Seed(documentExtra: $"<CommonData>{child}</CommonData>",
            pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>",
            objects: "<ImageObject Interpolate='false'/><TextObject Size='3'/>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, id).Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body'><ImageObject Interpolate='true'/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-007").Outcome);
    }

    [Fact]
    public async Task TemplateIdsUsePositiveNumericIdentity()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='01' BaseLoc='Templates/Content.xml'/></CommonData>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer/></Content></Page>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='1'/><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions><Content><Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-004").Outcome);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</CommonData>",
            "<TemplatePage ID='1' BaseLoc='Templates/Content.xml'/></CommonData>", StringComparison.Ordinal));
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Fact]
    public async Task StandardTemplateNameKeepsPageChecksDecidable()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='8' Name='cover' BaseLoc='Templates/Content.xml'/></CommonData>",
            pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>",
            objects: "<ImageObject Interpolate='false'/><TextObject Size='3'/>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer/></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Pass, Rule(report, id).Outcome);
    }

    [Fact]
    public async Task MissingNestedTemplateTaintsTemplatePageContentClaims()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='8' BaseLoc='Templates/Content.xml'/></CommonData>",
            pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>",
            objects: "<ImageObject Interpolate='false'/><TextObject Size='3'/>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='999'/><Content><Layer Name='template'><ImageObject Interpolate='false'/><TextObject Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, id).Outcome);
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='999'/><Content><Layer><ImageObject Interpolate='true'/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-007").Outcome);
    }

    [Theory]
    [InlineData("<Permissions><Outlines><OutlineElem Title='hidden'/></Outlines></Permissions>")]
    [InlineData("<VPreferences><Outlines><OutlineElem Title='hidden'/></Outlines></VPreferences>")]
    [InlineData("<Attachments><Outlines><OutlineElem Title='hidden'/></Outlines></Attachments>")]
    [InlineData("<Actions><Action><Goto><Outlines><OutlineElem Title='hidden'/></Outlines></Goto></Action></Actions>")]
    [InlineData("<Bookmarks><Bookmark Name='x'><Dest Type='Fit' PageID='1'><Outlines><OutlineElem Title='hidden'/></Outlines></Dest></Bookmark></Bookmarks>")]
    public async Task OpaqueDocumentChildCannotHideOutlineScope(string child)
    {
        var report = await Check(Seed(documentExtra: child));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task RecognizedDocumentPolicyFieldsDoNotTaintOutlineAbsence()
    {
        var report = await Check(Seed(documentExtra:
            "<Permissions><Edit>false</Edit><Print Printable='true' Copies='2'/></Permissions>" +
            "<VPreferences><PageMode>UseOutlines</PageMode></VPreferences>"));
        Assert.Equal(PreflightOutcome.NotApplicable, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.NotApplicable, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task SharedPageTreeIsInspectedOnceAcrossPageAndTemplateDeclarations()
    {
        var entries = Seed();
        var pageDeclarations = string.Concat(Enumerable.Range(1, 100).Select(id =>
            $"<Page ID='{id}' BaseLoc='Pages/P0/Content.xml'/>"));
        var templateDeclarations = string.Concat(Enumerable.Range(200, 100).Select(id =>
            $"<TemplatePage ID='{id}' BaseLoc='Pages/P0/Content.xml'/>"));
        entries["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><Pages>{pageDeclarations}</Pages><CommonData>{templateDeclarations}</CommonData></Document>");
        var images = string.Concat(Enumerable.Repeat("<ImageObject Interpolate='false'/>", 10));
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer>{images}</Layer></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-007").Outcome);
        Assert.StartsWith("Recognized objects=10;", Rule(report, "AC-007").Evidence.Observed, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MixedPageNamespaceKeepsKnownFailureButPreventsPass()
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8("<Page xmlns='http://www.ofdspec.org/2016'><Actions><Action><URI URI='https://example.invalid'/></Action></Actions><Content><Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    [InlineData("", "<?vendor opaque?>")]
    [InlineData("", "<FutureActions><Action><URI URI='https://example.invalid'/></Action></FutureActions>")]
    public async Task OpaquePageRootPreventsFalseActionClaim(string attributes, string payload)
    {
        var entries = Seed(pageExtra: "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}' {attributes}>{payload}<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions><Content><Layer><PathObject/></Layer></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}' {attributes}>{payload}<Content><Layer><PathObject/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}' {attributes}>{payload}<Actions><Action><URI URI='https://example.invalid'/></Action></Actions><Content><Layer><PathObject/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("<Template><Actions><Action><URI URI='https://example.invalid'/></Action></Actions></Template>")]
    [InlineData("<Area><FutureActions><Action><URI URI='https://example.invalid'/></Action></FutureActions></Area>")]
    [InlineData("<PageRes><Actions><Action><URI URI='https://example.invalid'/></Action></Actions></PageRes>")]
    public async Task OpaqueAllowedPageChildPreventsFalseActionClaim(string child)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'>{child}<Content><Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Fact]
    public async Task RecognizedPageAreaAndTemplateReferenceRemainCheckable()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='8' BaseLoc='Templates/Content.xml'/></CommonData>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer/></Content></Page>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Area><PhysicalBox>0 0 100 100</PhysicalBox></Area><Template TemplateID='8' ZOrder='Background'/><PageRes>Res.xml</PageRes><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions><Content><Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-004").Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='missing'/><Content><Layer/></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Fact]
    public async Task MissingTemplateTaintsAllPageContentClaims()
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='missing'/><Content><Layer Name='body'><ImageObject Interpolate='false'/><TextObject Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, id).Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='missing'/><Content><Layer Name='body'><ImageObject Interpolate='true'/><TextObject Size='0'/></Layer><Layer Name='body'/></Content></Page>");
        report = await Check(entries);
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Fail, Rule(report, id).Outcome);
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
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    [InlineData("", "<?vendor opaque?>")]
    [InlineData("Expanded='maybe'", "")]
    public async Task OpaqueOutlineNodePreventsPass(string attributes, string payload)
    {
        var outline = $"<Outlines><OutlineElem Title='A' {attributes}>{payload}<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<Goto><Dest Type='Fit' PageID='1'/></Goto>", "", PreflightOutcome.Pass)]
    [InlineData("<Goto><Dest Type='Fit' PageID='01'/></Goto>", "", PreflightOutcome.Pass)]
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
    [InlineData("<Goto vendor='x'><Dest Type='Fit' PageID='1'/></Goto>")]
    [InlineData("<Goto>opaque<Dest Type='Fit' PageID='1'/></Goto>")]
    [InlineData("opaque<Goto><Dest Type='Fit' PageID='1'/></Goto>")]
    public async Task OpaqueActionOrGotoCannotPassPageOrOutlineChecks(string actionBody)
    {
        var action = $"<Actions><Action>{actionBody}</Action></Actions>";
        var report = await Check(Seed(documentExtra: $"<Outlines><OutlineElem Title='A'>{action}</OutlineElem></Outlines>", pageExtra: action));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<Dest PageID='1'/>")]
    [InlineData("<Dest Type='Future' PageID='1'/>")]
    [InlineData("<Dest Type='Fit' PageID='1' Zoom='not-a-number'/>")]
    public async Task MalformedGotoDestinationPreventsPass(string destination)
    {
        var action = $"<Actions><Action><Goto>{destination}</Goto></Action></Actions>";
        var report = await Check(Seed(documentExtra: $"<Outlines><OutlineElem Title='A'>{action}</OutlineElem></Outlines>", pageExtra: action));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<Region><URI URI='https://example.invalid'/></Region>", PreflightOutcome.Unknown)]
    [InlineData("<Region vendor='x'><Area Start='0 0'><Move Point1='0 0'/></Area></Region>", PreflightOutcome.Unknown)]
    [InlineData("<Region><Area Start='0 0'><Move Point1='0 0'/><Line Point1='1 1'/><Close/></Area></Region>", PreflightOutcome.Pass)]
    public async Task ActionRegionMustHaveSupportedShape(string region, PreflightOutcome expected)
    {
        var action = $"<Actions><Action>{region}<Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>";
        var report = await Check(Seed(documentExtra: $"<Outlines><OutlineElem Title='A'>{action}</OutlineElem></Outlines>", pageExtra: action));
        Assert.Equal(expected, Rule(report, "AC-004").Outcome);
        Assert.Equal(expected, Rule(report, "AC-005").Outcome);
        Assert.Equal(expected, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<FutureAction/>", PreflightOutcome.Unknown)]
    [InlineData("<FutureAction/><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action>", PreflightOutcome.Unknown)]
    [InlineData("<FutureAction/><Action><URI URI='https://example.invalid'/></Action>", PreflightOutcome.Fail)]
    public async Task OpaqueActionsChildPreventsFalseTypeOrTargetClaim(string children, PreflightOutcome expectedType)
    {
        var actions = $"<Actions>{children}</Actions>";
        var outline = $"<Outlines><OutlineElem Title='A'>{actions}</OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline, pageExtra: actions));
        Assert.Equal(expectedType, Rule(report, "AC-004").Outcome);
        Assert.Equal(expectedType, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task OpaquePageInventoryPreventsFalseMissingTargetFailure()
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='99'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>", "<FuturePage/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    [InlineData("", "<?vendor opaque?>")]
    public async Task OpaquePageTreePayloadPreventsFalseMissingTargetFailure(string attributes, string payload)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='99'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("<Pages>", $"<Pages {attributes}>{payload}", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<Page ID='99' BaseLoc='hidden.xml'/>", "", PreflightOutcome.Unknown)]
    [InlineData("", "vendor='x'", PreflightOutcome.Fail)]
    [InlineData("opaque", "", PreflightOutcome.Fail)]
    public async Task OpaquePageDeclarationScopesMissingTargetUncertainty(string payload, string attributes, PreflightOutcome expected)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='99'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("<Page ID='1' BaseLoc='Pages/P0/Content.xml'/>",
            $"<Page ID='1' BaseLoc='Pages/P0/Content.xml' {attributes}>{payload}</Page>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-006").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Actions><Action><URI URI='https://example.invalid'/></Action></Actions><Content><Layer/></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-004").Outcome);
    }

    [Theory]
    [InlineData("1", "vendor='x'", PreflightOutcome.Pass)]
    [InlineData("2", "vendor='x'", PreflightOutcome.Unknown)]
    [InlineData("99", "vendor='x'", PreflightOutcome.Fail)]
    [InlineData("1", "xmlns:v='urn:vendor' v:ID='99'", PreflightOutcome.Pass)]
    [InlineData("2", "xmlns:v='urn:vendor' v:ID='99'", PreflightOutcome.Unknown)]
    [InlineData("99", "xmlns:v='urn:vendor' v:ID='99'", PreflightOutcome.Fail)]
    public async Task OpaquePageAttributesTaintOnlyTheDeclaredId(string target, string attributes, PreflightOutcome expected)
    {
        var outline = $"<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='{target}'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>",
            $"<Page ID='2' BaseLoc='Pages/P0/Content.xml' {attributes}/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("ID='bad'")]
    [InlineData("ID='0'")]
    [InlineData("")]
    public async Task DefiniteInvalidPageIdDoesNotHideAnotherTarget(string idAttribute)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='99'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>",
            $"<Page {idAttribute} BaseLoc='Pages/P0/Content.xml'/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-006").Outcome);
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("PageID='99'", "PageID='1'", StringComparison.Ordinal));
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    [InlineData("", "<?vendor opaque?>")]
    public async Task OpaqueOutlineContainerPreventsPass(string attributes, string payload)
    {
        var actions = "<OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem>";
        var report = await Check(Seed(documentExtra: $"<Outlines {attributes}>{payload}{actions}</Outlines>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task DeepOutlinesUseBoundedIterativeTraversal()
    {
        const int depth = 1500;
        var opens = string.Concat(Enumerable.Repeat("<OutlineElem Title='A'>", depth));
        var closes = string.Concat(Enumerable.Repeat("</OutlineElem>", depth));
        var outline = "<Outlines>" + opens +
            "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>" + closes + "</Outlines>";
        var report = await Check(Seed(documentExtra: outline), new ArchivalPreflightOptions { MaxXmlDepth = 4096 });
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-005").Outcome);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task OpaqueBookmarkInventoryPreventsFalseMissingTargetFailure()
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='future'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline + "<Bookmarks><FutureBookmark/></Bookmarks>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    public async Task OpaqueBookmarkContainerPreventsFalseMissingTargetFailure(string attributes, string payload)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='future'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline + $"<Bookmarks {attributes}>{payload}</Bookmarks>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task OpaqueBookmarkDeclarationPreventsFalseMissingTargetFailure()
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='hidden'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline + "<Bookmarks><Bookmark Name='known'><Bookmark Name='hidden'/></Bookmark></Bookmarks>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
        report = await Check(Seed(documentExtra: outline + "<Bookmarks><Bookmark><Bookmark Name='hidden'/></Bookmark></Bookmarks>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Name=''")]
    public async Task DefiniteEmptyBookmarkNameDoesNotHideAnotherTarget(string attributes)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='missing'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var bookmarks = $"<Bookmarks><Bookmark {attributes}/><Bookmark Name='good'><Dest Type='Fit' PageID='1'/></Bookmark></Bookmarks>";
        var report = await Check(Seed(documentExtra: outline + bookmarks));
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-006").Outcome);
        report = await Check(Seed(documentExtra: outline.Replace("missing", "good", StringComparison.Ordinal) + bookmarks));
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicateBookmarkNameCannotSelectAnArbitraryDestination(bool reverse)
    {
        var declarations = new[]
        {
            "<Bookmark Name='same'><Dest Type='Fit' PageID='99'/></Bookmark>",
            "<Bookmark Name='same'><Dest Type='Fit' PageID='1'/></Bookmark>"
        };
        if (reverse) Array.Reverse(declarations);
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='same'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var report = await Check(Seed(documentExtra: outline + $"<Bookmarks>{string.Concat(declarations)}</Bookmarks>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task DirectPageTargetIgnoresUnrelatedDuplicateBookmarks()
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var bookmarks = "<Bookmarks><Bookmark Name='other'><Dest Type='Fit' PageID='1'/></Bookmark><Bookmark Name='other'><Dest Type='Fit' PageID='99'/></Bookmark></Bookmarks>";
        var report = await Check(Seed(documentExtra: outline + bookmarks));
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'")]
    [InlineData("xmlns:v='urn:vendor' v:Name='vendor'")]
    public async Task BookmarkTargetIgnoresOpaqueDeclarationWithDifferentKnownName(string attributes)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Bookmark Name='good'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var bookmarks = $"<Bookmarks><Bookmark Name='other' {attributes}><Dest Type='Fit' PageID='1'/></Bookmark><Bookmark Name='good'><Dest Type='Fit' PageID='1'/></Bookmark></Bookmarks>";
        var report = await Check(Seed(documentExtra: outline + bookmarks));
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
        report = await Check(Seed(documentExtra: outline.Replace("good", "other", StringComparison.Ordinal) + bookmarks));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
        report = await Check(Seed(documentExtra: outline.Replace("good", "missing", StringComparison.Ordinal) + bookmarks));
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("<Page ID='2' BaseLoc='Pages/P0/Content.xml'/><Page ID='02' BaseLoc='Pages/P0/Content.xml'/>")]
    [InlineData("<Page ID='2' BaseLoc='Pages/P0/missing.xml'/>")]
    public async Task DirectPageTargetIgnoresUnrelatedPageAmbiguity(string extraPages)
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>", extraPages + "</Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-006").Outcome);
    }

    [Theory]
    [InlineData("")]
    [InlineData("<Outlines><OutlineElem Title='A'/></Outlines>")]
    public async Task AmbiguousTargetInventoryWithoutGotoIsNotApplicable(string outlines)
    {
        var entries = Seed(documentExtra: outlines + "<Bookmarks><Bookmark Name='same'/><Bookmark Name='same'/></Bookmarks>");
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>",
            "<Page ID='01' BaseLoc='Pages/P0/Content.xml'/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.NotApplicable, Rule(report, "AC-006").Outcome);
    }

    [Fact]
    public async Task LexicallyDifferentDuplicatePageIdsMakeTargetAmbiguous()
    {
        var outline = "<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem></Outlines>";
        var entries = Seed(documentExtra: outline);
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(document.Replace("</Pages>", "<Page ID='01' BaseLoc='Pages/P0/Content.xml'/></Pages>", StringComparison.Ordinal));
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-006").Outcome);
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

    [Fact]
    public async Task NamespaceDeclarationsDoNotImpersonateForeignAttributes()
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Actions><Action><Goto><Dest xmlns:PageID='urn:x' Type='Fit' PageID='1'/></Goto></Action></Actions><Content><Layer xmlns:Name='urn:x' Name='body'><ImageObject xmlns:Interpolate='urn:x' Interpolate='false'/><TextObject xmlns:Size='urn:x' Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-004", "AC-007", "AC-008", "AC-011" })
            Assert.Equal(PreflightOutcome.Pass, Rule(report, id).Outcome);
    }

    [Theory]
    [InlineData("<TextObject Size='3'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Pass)]
    [InlineData("<TextObject Size='100000000000000000000000000000'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Pass)]
    [InlineData("<TextObject><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
    [InlineData("<TextObject Size='0'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
    [InlineData("<TextObject Size='1e1000'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
    [InlineData("<TextObject Size='not-a-number'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Fail)]
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
    public async Task SharedAnnotationTreeIsInspectedOnce()
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        var pages = string.Concat(Enumerable.Repeat("<Page PageID='1'><FileLoc>Page.xml</FileLoc></Page>", 100));
        var annots = string.Concat(Enumerable.Range(2, 10).Select(id =>
            $"<Annot ID='{id}' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PathObject/></Appearance></Annot>"));
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'>{pages}</Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'>{annots}</PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-009").Outcome);
        Assert.StartsWith("Recognized objects=10;", Rule(report, "AC-009").Evidence.Observed, StringComparison.Ordinal);
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

    [Theory]
    [InlineData("<FutureFileLoc>hidden.xml</FutureFileLoc>")]
    [InlineData("opaque")]
    public async Task OpaqueAnnotationIndexPagePreventsFalsePass(string extra)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc>{extra}</Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PathObject/></Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-009").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-012").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
    }

    [Fact]
    public async Task NestedAppearanceCannotBeExcludedFromAnnotationScope()
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><FutureAppearance><Appearance><PageBlock ID='3'/></Appearance></FutureAppearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-012").Outcome);
    }

    [Theory]
    [InlineData("<FutureAppearance><PageBlock ID='3'/></FutureAppearance>", PreflightOutcome.Unknown)]
    [InlineData("<Remark>reviewed</Remark><Parameters><Parameter Name='owner'>local</Parameter></Parameters>", PreflightOutcome.Pass)]
    public async Task AnnotationWrapperRecognizesOnlySupportedChildren(string extra, PreflightOutcome expected)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PageBlock ID='4'><PathObject/></PageBlock></Appearance>{extra}</Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-009").Outcome);
        Assert.Equal(expected, Rule(report, "AC-012").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "<!--opaque-->")]
    [InlineData("", "<?vendor opaque?>")]
    public async Task OpaqueAnnotationLocatorPreventsFalsePass(string attributes, string payload)
    {
        var entries = Seed(documentExtra: $"<Annotations {attributes}>{payload}Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PathObject/></Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-009").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-012").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    [InlineData("", "<?vendor opaque?>")]
    public async Task OpaquePageAnnotContainerPreventsFalsePass(string attributes, string payload)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}' {attributes}>{payload}<Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PageBlock ID='3'><PathObject/></PageBlock></Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-009").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-012").Outcome);
    }

    [Fact]
    public async Task XmlNodeBudgetRejectsDensePageBeforeItCanBeClaimedClean()
    {
        var entries = Seed(objects: string.Concat(Enumerable.Repeat("<PathObject/>", 200)));
        var report = await Check(entries, new ArchivalPreflightOptions { MaxXmlNodes = 30 });
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
    }

    [Fact]
    public async Task XmlNodeBudgetChargesAttributesBeforeMaterializingPage()
    {
        var attributes = string.Concat(Enumerable.Range(0, 200).Select(i => $" a{i}='{i}'"));
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'{attributes}><Content><Layer/></Content></Page>");
        var report = await Check(entries, new ArchivalPreflightOptions { MaxXmlNodes = 100 });
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        Assert.Equal("PageContentUnavailable", Rule(report, "AC-004").ReasonCode);
    }

    [Fact]
    public async Task XmlDepthBudgetRejectsDeepPageBeforeItCanBeClaimedClean()
    {
        var nested = string.Concat(Enumerable.Repeat("<Wrapper>", 20)) + "<PathObject/>" +
            string.Concat(Enumerable.Repeat("</Wrapper>", 20));
        var report = await Check(Seed(objects: nested), new ArchivalPreflightOptions { MaxXmlDepth = 8 });
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
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
    public async Task AnnotationObjectActionMatchesPublishedRuleScope()
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance><PathObject><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></PathObject></Appearance></Annot></PageAnnot>");
        var result = Rule(await Check(entries), "AC-010");
        Assert.Equal(PreflightOutcome.Fail, result.Outcome);
        Assert.Contains("annotation appearances", result.Rule.Condition, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<Actions/>")]
    [InlineData("<Actions><FutureAction/></Actions>")]
    public async Task ObjectWithoutRecognizedActionDoesNotBecomeKnownFailure(string payload)
    {
        var report = await Check(Seed(objects: $"<PathObject>{payload}</PathObject>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-010").Outcome);
    }

    [Theory]
    [InlineData("<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>", PreflightOutcome.Pass)]
    [InlineData("<Actions><FutureAction/></Actions>", PreflightOutcome.Unknown)]
    public async Task ObjectActionsKeepIndependentGraphicChecksDecidableWhenSupported(string actions,
        PreflightOutcome expectedGraphics)
    {
        var objects = $"<PathObject>{actions}</PathObject><ImageObject Interpolate='false'/><TextObject Size='3'/>";
        var report = await Check(Seed(objects: objects));
        Assert.Equal(expectedGraphics, Rule(report, "AC-007").Outcome);
        Assert.Equal(expectedGraphics, Rule(report, "AC-008").Outcome);
    }

    [Fact]
    public async Task UnsupportedPageGraphicsPreventAnAbsenceClaim()
    {
        var report = await Check(Seed(objects: "<v:ImageObject xmlns:v='urn:vendor' Interpolate='true'/>"));
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("<FutureObject/>", PreflightOutcome.Unknown)]
    [InlineData("<FutureObject/><ImageObject Interpolate='false'/><TextObject Size='3'/>", PreflightOutcome.Unknown)]
    [InlineData("<FutureObject/><ImageObject Interpolate='true'/><TextObject Size='0'/>", PreflightOutcome.Fail)]
    [InlineData("<FutureObject><ImageObject Interpolate='true'/><TextObject Size='0'/></FutureObject>", PreflightOutcome.Unknown)]
    [InlineData("<FutureObject><ImageObject Interpolate='true'/><TextObject Size='0'/></FutureObject><ImageObject Interpolate='true'/><TextObject Size='0'/>", PreflightOutcome.Fail)]
    public async Task SameNamespaceUnknownGraphicKeepsFailDominance(string objects, PreflightOutcome expected)
    {
        var report = await Check(Seed(objects: objects));
        Assert.Equal(expected, Rule(report, "AC-007").Outcome);
        Assert.Equal(expected, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("<PathObject><FutureObject/></PathObject><ImageObject Interpolate='false'/><TextObject Size='3'/>", PreflightOutcome.Unknown)]
    [InlineData("<PathObject><AbbreviatedData><FutureObject/></AbbreviatedData></PathObject><ImageObject Interpolate='false'/><TextObject Size='3'/>", PreflightOutcome.Unknown)]
    [InlineData("<PathObject><FutureObject/></PathObject><ImageObject Interpolate='true'/><TextObject Size='0'/>", PreflightOutcome.Fail)]
    [InlineData("<PathObject><AbbreviatedData>M 0 0 L 1 1</AbbreviatedData></PathObject><ImageObject Interpolate='false'/><TextObject Size='3'><TextCode>OK</TextCode></TextObject>", PreflightOutcome.Pass)]
    public async Task GraphicObjectChildrenAreShapeChecked(string objects, PreflightOutcome expected)
    {
        var report = await Check(Seed(objects: objects));
        Assert.Equal(expected, Rule(report, "AC-007").Outcome);
        Assert.Equal(expected, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("<ImageObject vendor='x' Interpolate='false'/><TextObject Size='3'/>", PreflightOutcome.Unknown)]
    [InlineData("<ImageObject Interpolate='false'/><TextObject vendor='x' Size='3'/>", PreflightOutcome.Unknown)]
    [InlineData("<ImageObject Interpolate='false'/><TextObject Size='3'><TextCode vendor='x'>OK</TextCode></TextObject>", PreflightOutcome.Unknown)]
    [InlineData("<ImageObject ResourceID='1' Interpolate='false'/><TextObject Font='1' Size='3'><TextCode X='0' Y='0'>OK</TextCode></TextObject>", PreflightOutcome.Pass)]
    public async Task GraphicAttributesAreShapeChecked(string objects, PreflightOutcome expected)
    {
        var report = await Check(Seed(objects: objects));
        Assert.Equal(expected, Rule(report, "AC-007").Outcome);
        Assert.Equal(expected, Rule(report, "AC-008").Outcome);
    }

    [Theory]
    [InlineData("DrawParam='9'", PreflightOutcome.Pass)]
    [InlineData("DrawParam='0'", PreflightOutcome.Unknown)]
    [InlineData("DrawParam='invalid'", PreflightOutcome.Unknown)]
    [InlineData("xmlns:v='urn:vendor' v:DrawParam='9'", PreflightOutcome.Unknown)]
    public async Task GraphicDrawParamHasNumericReferenceShape(string attributes, PreflightOutcome expected)
    {
        var objects = $"<PathObject {attributes}/><ImageObject {attributes} Interpolate='false'/><TextObject {attributes} Size='3'/>";
        var report = await Check(Seed(objects: objects));
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(expected, Rule(report, id).Outcome);
    }

    [Theory]
    [InlineData("ImageMask='123'", PreflightOutcome.Pass)]
    [InlineData("ImageMask='0'", PreflightOutcome.Unknown)]
    [InlineData("ImageMask='invalid'", PreflightOutcome.Unknown)]
    [InlineData("xmlns:v='urn:vendor' v:ImageMask='123'", PreflightOutcome.Unknown)]
    public async Task ImageMaskHasNumericReferenceShape(string attributes, PreflightOutcome expected)
    {
        var objects = $"<ImageObject {attributes} Interpolate='false'/><TextObject Size='3'/>";
        var report = await Check(Seed(objects: objects));
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(expected, Rule(report, id).Outcome);
    }

    [Theory]
    [InlineData("ColorSpace='9'", PreflightOutcome.Pass)]
    [InlineData("ColorSpace='0'", PreflightOutcome.Unknown)]
    public async Task ColorSpaceHasNumericReferenceShape(string attributes, PreflightOutcome expected)
    {
        var objects = $"<PathObject><FillColor {attributes} Value='0 0 0'/></PathObject><ImageObject Interpolate='false'/><TextObject Size='3'/>";
        var report = await Check(Seed(objects: objects));
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(expected, Rule(report, id).Outcome);
    }

    [Fact]
    public async Task DenseGraphicSiblingEvidenceKeepsTheCorrectOrdinal()
    {
        const int count = 3000;
        var objects = string.Concat(Enumerable.Repeat("<ImageObject Interpolate='false'/>", count)) +
            "<ImageObject Interpolate='true'/>";
        var report = await Check(Seed(objects: objects));
        var result = Rule(report, "AC-007");
        Assert.Equal(PreflightOutcome.Fail, result.Outcome);
        Assert.EndsWith($"ImageObject[{count + 1}]", result.Evidence.Selector, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("vendor='x'", "")]
    [InlineData("", "opaque")]
    public async Task OpaqueContentPayloadPreventsGraphicAndLayerPass(string attributes, string payload)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content {attributes}>{payload}<Layer Name='body'><ImageObject Interpolate='false'/><TextObject Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-011").Outcome);
    }

    [Theory]
    [InlineData("", "<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>", PreflightOutcome.Unknown)]
    [InlineData("<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>", "<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>", PreflightOutcome.Unknown)]
    [InlineData("<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>", "<Future/>", PreflightOutcome.Unknown)]
    [InlineData("<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>", "<Layer><Actions><Action><URI URI='https://example.invalid'/></Action></Actions></Layer>", PreflightOutcome.Unknown)]
    [InlineData("", "<Layer><PageBlock><Actions><Action><URI URI='https://example.invalid'/></Action></Actions></PageBlock></Layer>", PreflightOutcome.Unknown)]
    [InlineData("<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>", "<Actions><Action><URI URI='https://example.invalid'/></Action></Actions>", PreflightOutcome.Fail)]
    public async Task OpaqueContentPayloadTaintsPageActionInventory(string pageActions, string payload, PreflightOutcome expected)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'>{pageActions}<Content>{payload}<Layer/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-004").Outcome);
    }

    [Fact]
    public async Task RecognizedObjectActionsKeepPageActionInventoryIndependent()
    {
        var pageActions = "<Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions>";
        var objects = "<PathObject><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></PathObject>";
        var report = await Check(Seed(pageExtra: pageActions, objects: objects));
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-004").Outcome);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-010").Outcome);
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
    [InlineData("<Future><Layer Name='body'/></Future>")]
    [InlineData("<PageBlock><Future><Layer Name='body'/></Future></PageBlock>")]
    [InlineData("<PathObject><Future><Layer Name='body'/></Future></PathObject>")]
    [InlineData("<!--opaque-->")]
    public async Task OpaqueLayerPayloadPreventsFalseUniqueNameClaim(string payload)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body'>{payload}</Layer></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-011").Outcome);
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body'>{payload}</Layer><Layer Name='body'/></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-011").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", PreflightOutcome.Unknown)]
    [InlineData("ID='2' Type='Body' Name='body'", PreflightOutcome.Pass)]
    public async Task LayerAttributesControlGraphicAndNameClaims(string attributes, PreflightOutcome expected)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer {attributes}><ImageObject Interpolate='false'/><TextObject Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(expected, Rule(report, id).Outcome);
    }

    [Theory]
    [InlineData("DrawParam='9'", PreflightOutcome.Pass)]
    [InlineData("DrawParam='0'", PreflightOutcome.Unknown)]
    [InlineData("DrawParam='invalid'", PreflightOutcome.Unknown)]
    [InlineData("xmlns:v='urn:vendor' v:DrawParam='9'", PreflightOutcome.Unknown)]
    public async Task StandardLayerDrawParamHasNumericReferenceShape(string attributes, PreflightOutcome expected)
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body' {attributes}><ImageObject Interpolate='false'/><TextObject Size='3'/></Layer></Content></Page>");
        var report = await Check(entries);
        foreach (var id in new[] { "AC-007", "AC-008", "AC-011" })
            Assert.Equal(expected, Rule(report, id).Outcome);
    }

    [Fact]
    public async Task AppliedTemplateLayerNamesShareTheEffectivePageScope()
    {
        var entries = Seed(documentExtra: "<CommonData><TemplatePage ID='8' BaseLoc='Templates/Content.xml'/></CommonData>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='8'/><Content><Layer Name='body'><PathObject/></Layer></Content></Page>");
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body'><PathObject/></Layer></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-011").Outcome);
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='template'><PathObject/></Layer></Content></Page>");
        report = await Check(entries);
        Assert.Equal(PreflightOutcome.Pass, Rule(report, "AC-011").Outcome);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicateTemplateIdCannotSelectAnArbitraryEffectiveLayerName(bool reverse)
    {
        var declarations = new[]
        {
            "<TemplatePage ID='8' BaseLoc='Templates/Colliding.xml'/>",
            "<TemplatePage ID='08' BaseLoc='Templates/Distinct.xml'/>"
        };
        if (reverse) Array.Reverse(declarations);
        var entries = Seed(documentExtra: $"<CommonData>{string.Concat(declarations)}</CommonData>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Template TemplateID='8'/><Content><Layer Name='body'/></Content></Page>");
        entries["Doc_0/Templates/Colliding.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='body'/></Content></Page>");
        entries["Doc_0/Templates/Distinct.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer Name='other'/></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-011").Outcome);
    }

    [Theory]
    [InlineData("<PageBlock ID='3'><PageBlock ID='4'><PathObject/></PageBlock></PageBlock>", PreflightOutcome.Fail)]
    [InlineData("<PageBlock ID='3'><PathObject/></PageBlock>", PreflightOutcome.Pass)]
    [InlineData("<PageBlock ID='3' vendor='x'><PathObject/></PageBlock>", PreflightOutcome.Unknown)]
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
    public async Task DeepNestedAnnotationBlocksUseBoundedIterativeTraversal()
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        var opens = string.Concat(Enumerable.Repeat("<PageBlock>", 1500));
        var closes = string.Concat(Enumerable.Repeat("</PageBlock>", 1500));
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance>{opens}{closes}</Appearance></Annot></PageAnnot>");
        var report = await Check(entries, new ArchivalPreflightOptions { MaxXmlDepth = 4096 });
        Assert.Equal(PreflightOutcome.Fail, Rule(report, "AC-012").Outcome);
    }

    [Theory]
    [InlineData("vendor='x'", "<PageBlock ID='3'><PathObject/></PageBlock>", PreflightOutcome.Unknown)]
    [InlineData("", "opaque<PageBlock ID='3'><PathObject/></PageBlock>", PreflightOutcome.Unknown)]
    [InlineData("", "<FutureWrapper><PageBlock ID='3'><PathObject/></PageBlock></FutureWrapper>", PreflightOutcome.Unknown)]
    [InlineData("ID='7' Boundary='0 0 10 10' CTM='1 0 0 1 0 0'", "<PageBlock ID='3'><PathObject/></PageBlock>", PreflightOutcome.Pass)]
    [InlineData("vendor='x'", "<PageBlock ID='3'><PageBlock ID='4'><PathObject/></PageBlock></PageBlock>", PreflightOutcome.Fail)]
    public async Task AnnotationAppearanceShapeControlsNestedBlockClaim(string attributes, string content,
        PreflightOutcome expected)
    {
        var entries = Seed(documentExtra: "<Annotations>Annots/Annotations.xml</Annotations>");
        entries["Doc_0/Annots/Annotations.xml"] = Utf8($"<Annotations xmlns='{Ns}'><Page PageID='1'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        entries["Doc_0/Annots/Page.xml"] = Utf8($"<PageAnnot xmlns='{Ns}'><Annot ID='2' ReadOnly='true' NoZoom='true' NoRotate='true'><Appearance {attributes}>{content}</Appearance></Annot></PageAnnot>");
        var report = await Check(entries);
        Assert.Equal(expected, Rule(report, "AC-012").Outcome);
        if (expected == PreflightOutcome.Unknown)
        {
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-007").Outcome);
            Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-008").Outcome);
        }
    }

    [Fact]
    public async Task OpaqueWrapperBlocksLayerAndTextClaimsWithoutLeakingItsNamespace()
    {
        var entries = Seed();
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><v:Wrapper xmlns:v='https://user:secret@example.invalid/scope'><Layer Name='hidden'><TextObject><TextCode>missing size</TextCode></TextObject></Layer></v:Wrapper></Content></Page>");
        var report = await Check(entries);
        Assert.Equal(PreflightOutcome.Unknown, Rule(report, "AC-011").Outcome);
        var text = Rule(report, "AC-008");
        Assert.Equal(PreflightOutcome.Unknown, text.Outcome);
        Assert.DoesNotContain("secret", text.Evidence.Selector, StringComparison.OrdinalIgnoreCase);
    }
}
