using Ofdrw.Net.Medical.Internal;
using System.Text;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.TestSupport;
using static Ofdrw.Net.Medical.Tests.MedicalTestSupport;

namespace Ofdrw.Net.Medical.Tests;

public class MedicalAssessmentServiceTests
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private static string Open(string name, string attributes = "") => $"<ofd:{name} xmlns:ofd='{Ns}'{attributes}>";

    private static byte[] AddToDocument(byte[] package, string afterPages, Action<Dictionary<string, byte[]>>? more = null) =>
        Mutate(package, entries =>
        {
            entries.Set("Doc_0/Document.xml", entries.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "</ofd:Pages>" + afterPages));
            more?.Invoke(entries);
        });

    [Fact]
    public async Task ProfileOutputPassesThePackageDecidableRules()
    {
        var report = await Assess(await LayoutOfd(true));

        Assert.Equal(MedicalExecutionState.Complete, report.ExecutionState);
        Assert.Equal(MedicalDeclarationStatus.Medical, report.Declaration.Status);
        Assert.Equal("OFD-H", report.Declaration.RawDocType);
        foreach (var clause in new[] { "6.1.1", "6.1.2(a)", "6.1.3(a)", "6.1.3(b)", "6.2(a)", "6.2(b)", "6.2(c)", "6.2(d)",
                     "7.1(a)", "7.1(b)", "7.1(c)", "7.2(b)", "7.3(a)", "7.3(b)", "7.3(g)" })
            Assert.True(report.GetClause(clause).Outcome == PreflightOutcome.Pass, clause + " is " + report.GetClause(clause).Outcome);
        // The two clauses disagree about the font table; neither reading is chosen silently.
        Assert.Equal("ResourceTableConflict", report.Single("7.2(a)", "MH-007").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("7.2(a)").Outcome);
        // 7.3(d) is a recommendation: an unsigned document is advised, not failed.
        Assert.Equal("NoSignature", Assert.Single(report.GetClause("7.3(d)").Advisories).ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("7.3(d)").Outcome);
        Assert.Equal(PreflightOutcome.Unknown, report.Document.StandardOutcome);
        Assert.Contains("NoFullStandardAssessment", report.Document.Gaps);
        Assert.Contains("5.3.3", report.Document.UncheckedClauses);

        // GB/T 42133 runs on the same bytes; its OFD-A declaration failure is superseded, not inherited.
        Assert.Equal(Ofdrw.Net.Archival.ArchivalExecutionState.Complete, report.Archival!.ExecutionState);
        var docType = Assert.Single(report.InheritedArchivalClauses, x => x.ArchivalClause == "6.2.1(a)");
        Assert.Equal(PreflightOutcome.Fail, docType.ArchivalOutcome);
        Assert.Equal(MedicalArchivalDisposition.Superseded, docType.Inheritance.Disposition);
        Assert.False(docType.ContributesFail);
        Assert.DoesNotContain(report.InheritedArchivalClauses, x => x.ContributesFail);
        Assert.Equal(report.Input!.Sha256, report.Archival.Input!.Sha256);
    }

    [Fact]
    public async Task OrdinaryOutputFailsTheMedicalDeclarations()
    {
        var report = await Assess(await LayoutOfd(false));

        Assert.Equal(MedicalDeclarationStatus.NotMedical, report.Declaration.Status);
        Assert.Equal("NonMedicalDocType", report.Single("7.1(a)", "MH-001").ReasonCode);
        Assert.Equal("NamespaceNotStandard", report.Single("6.2(c)", "MH-003").ReasonCode);
        Assert.Equal("MetadataMissing", report.Single("7.3(b)", "MH-011").ReasonCode);
        Assert.Contains("DocID", report.Single("7.3(b)", "MH-011").Evidence.Observed);
        // Linearization is a capability: an ordinal entry order is advised, not failed.
        Assert.Equal("EntryOrderNotLinearized", Assert.Single(report.GetClause("6.1.3(a)").Advisories).ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("6.1.3(a)").Outcome);
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
    }

    [Fact]
    public async Task TypesetDashFromThePublishedStandardIsNotTheDeclaration()
    {
        var package = Mutate(await LayoutOfd(true), e => e.Set("OFD.xml", e.Text("OFD.xml").Replace("DocType=\"OFD-H\"", "DocType=\"OFD⁃H\"")));
        var report = await Assess(package);
        Assert.Equal(MedicalDeclarationStatus.TypesetVariant, report.Declaration.Status);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("7.1(a)").Outcome);
        Assert.Contains("U+2043", report.Single("7.1(a)", "MH-001").Evidence.Observed);
    }

    [Fact]
    public async Task StoredXmlAndNonLinearOrderAreReported()
    {
        var source = await LayoutOfd(true);
        var stored = Mutate(source, _ => { }, name => name == "Doc_0/Document.xml");
        Assert.Equal("CompressibleEntryStored", (await Assess(stored)).Single("6.1.3(b)", "MH-004").ReasonCode);

        var entries = Unzip(source);
        var reversed = Zip(entries.OrderByDescending(x => x.Key, StringComparer.Ordinal));
        var report = await Assess(reversed);
        Assert.Equal("EntryOrderNotLinearized", report.Single("6.1.3(a)", "MH-005").ReasonCode);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("6.1.3(b)").Outcome);
    }

    [Fact]
    public async Task NonCanonicalPageFolderFailsTheFigure3Structure()
    {
        var package = Mutate(await LayoutOfd(true), e =>
        {
            e["Doc_0/Pages/P1/Content.xml"] = e["Doc_0/Pages/Page_1/Content.xml"];
            e.Remove("Doc_0/Pages/Page_1/Content.xml");
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("Pages/Page_1/Content.xml", "Pages/P1/Content.xml"));
        });
        var observation = (await Assess(package)).Single("6.2(a)", "MH-006");
        Assert.Equal("NonCanonicalLocation", observation.ReasonCode);
        Assert.Contains("Doc_0/Pages/P1/Content.xml", observation.Evidence.Observed);
    }

    [Fact]
    public async Task FontsInPageResFailBothResourceTableClauses()
    {
        var package = Mutate(await LayoutOfd(true), e =>
        {
            e.Set("Doc_0/Pages/Page_1/Content.xml", e.Text("Doc_0/Pages/Page_1/Content.xml").Replace("<ofd:Content>", "<ofd:PageRes>PageRes.xml</ofd:PageRes><ofd:Content>"));
            e.Set("Doc_0/Pages/Page_1/PageRes.xml", Open("Res") + "<ofd:Fonts><ofd:Font ID='99' FontName='SimSun'/></ofd:Fonts></ofd:Res>");
        });
        var report = await Assess(package);
        Assert.Equal("DefinedInPageRes", report.Single("6.2(b)", "MH-007").ReasonCode);
        Assert.Equal("FontInPageRes", report.Single("7.2(a)", "MH-007").ReasonCode);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("7.2(a)").Outcome);
    }

    [Fact]
    public async Task TemplatePageResourcesAreAssessedAndRootBaseLocIsAccepted()
    {
        var gif = new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 1, 0, 1, 0, 0, 0, 0 };
        var package = Mutate(await LayoutOfd(true), e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:PublicRes>",
                "</ofd:PublicRes><ofd:TemplatePage ID='80' BaseLoc='Tpls/Tpl_0/Content.xml'/>"));
            e.Set("Doc_0/Tpls/Tpl_0/Content.xml", Open("Page") + "<ofd:PageRes>Res.xml</ofd:PageRes><ofd:Content/></ofd:Page>");
            // BaseLoc ".." from Doc_0/Tpls/Tpl_0 is Doc_0/Tpls; "/" would be the package root.
            e.Set("Doc_0/Tpls/Tpl_0/Res.xml", Open("Res", " BaseLoc='/'") +
                "<ofd:MultiMedias><ofd:MultiMedia ID='81' Type='Image'><ofd:MediaFile>logo.gif</ofd:MediaFile></ofd:MultiMedia></ofd:MultiMedias></ofd:Res>");
            e["logo.gif"] = gif;
        });
        var report = await Assess(package);
        Assert.Equal(MedicalExecutionState.Complete, report.ExecutionState);
        Assert.Equal("ImageFormatNotPermitted", report.Single("7.2(d)", "MH-009").ReasonCode);
    }

    [Fact]
    public async Task StructuralEntriesMustUseTheOfdNamespaceAndTheOfdPrefix()
    {
        var source = await LayoutOfd(true);
        // A structural entry in an unrelated namespace is not exempt like a custom tag payload.
        var wrong = await Assess(Mutate(source, e => e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace(MedicalPackage.Namespace2016, "urn:wrong"))));
        Assert.Equal("NamespaceNotStandard", wrong.Single("6.2(c)", "MH-003").ReasonCode);
        Assert.Contains("Doc_0/Document.xml", wrong.Single("6.2(c)", "MH-003").Evidence.Observed);

        // Declaring xmlns:ofd but writing every element with another prefix still violates 6.2(d).
        var other = await Assess(Mutate(source, e =>
        {
            var text = e.Text("Doc_0/Document.xml").Replace("<ofd:", "<o:").Replace("</ofd:", "</o:");
            e.Set("Doc_0/Document.xml", text.Replace("xmlns:ofd=", "xmlns:o=\"" + MedicalPackage.Namespace2016 + "\" xmlns:ofd="));
        }));
        Assert.Equal("ElementsNotOfdPrefixed", other.Single("6.2(d)", "MH-003").ReasonCode);
    }

    [Fact]
    public async Task StructuralEntriesAreInspectedWhateverTheirExtension()
    {
        var package = Mutate(await LayoutOfd(true), e =>
        {
            var page = e.Text("Doc_0/Pages/Page_1/Content.xml").Replace("<ofd:", "<o:").Replace("</ofd:", "</o:");
            e.Set("Doc_0/Pages/Page_1/Content", page.Replace("xmlns:ofd=", "xmlns:o=\"" + MedicalPackage.Namespace2016 + "\" xmlns:ofd="));
            e.Remove("Doc_0/Pages/Page_1/Content.xml");
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("Pages/Page_1/Content.xml", "Pages/Page_1/Content"));
        });
        var observation = (await Assess(package)).Single("6.2(d)", "MH-003");
        Assert.Equal("ElementsNotOfdPrefixed", observation.ReasonCode);
        Assert.Contains("Doc_0/Pages/Page_1/Content", observation.Evidence.Observed);
    }

    [Fact]
    public async Task UnreadableEntriesKeepLocationAndSizeClausesUnknown()
    {
        var source = await LayoutOfd(true);
        // A page resource table that cannot be read may declare further locations and payloads.
        var report = await Assess(Mutate(source, e =>
            e.Set("Doc_0/Pages/Page_1/Content.xml", e.Text("Doc_0/Pages/Page_1/Content.xml").Replace("<ofd:Content>", "<ofd:PageRes>PageRes.xml</ofd:PageRes><ofd:Content>"))));
        Assert.Equal("StructureIncomplete", report.Single("6.2(a)", "MH-006").ReasonCode);
        Assert.Equal("ResourceScopeIncomplete", report.Single("7.1(b)", "MH-006").ReasonCode);

        // An unparsable Document.xml may declare tags, attachments and annotations: unknown, not "not declared".
        foreach (var broken in new Action<Dictionary<string, byte[]>>[]
        {
            e => e.Set("Doc_0/Document.xml", "<not-closed"),
            e => e.Set("OFD.xml", e.Text("OFD.xml").Replace(">Doc_0/Document.xml<", ">../../escape.xml<")),
        })
        {
            report = await Assess(Mutate(source, broken));
            Assert.Equal(MedicalExecutionState.Complete, report.ExecutionState);
            foreach (var clause in new[] { "7.1(d)", "7.1(e)", "7.1(f)", "7.4(c)", "7.5(a)", "7.5(b)", "7.2(c)", "7.2(e)", "7.2(f)" })
                Assert.True(report.GetClause(clause).Outcome == PreflightOutcome.Unknown, clause + " is " + report.GetClause(clause).Outcome);
        }
        var signedBroken = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
        {
            e.Set("Doc_0/Document.xml", "<not-closed");
            e.Set("Doc_0/Annots/Page_0/Annotation.xml", Open("PageAnnot") + "</ofd:PageAnnot>");
        });
        Assert.Equal("AnnotationIndexUnreadable", (await Assess(signedBroken)).Single("8(c)", "MH-013").ReasonCode);

        // An unreadable attachment index leaves attachment sizes unknown.
        report = await Assess(AddToDocument(source, "<ofd:Attachments>Attachs/Attachments.xml</ofd:Attachments>", e => e.Set("Doc_0/Attachs/Attachments.xml", "<not-closed")));
        Assert.Equal("AttachmentUnavailable", report.Single("7.4(c)", "MH-010").ReasonCode);
    }

    [Fact]
    public async Task ZipLoadingBudgetsBoundMedicalInspection()
    {
        var package = await LayoutOfd(true);
        var options = new MedicalAssessmentOptions();
        options.Archival.Preflight.LoadOptions.MaxEntryUncompressedBytes = 1024;
        var report = await Assess(package, options);
        Assert.Equal(MedicalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.All(report.Clauses, x => Assert.Equal(PreflightOutcome.Unknown, x.Outcome));
    }

    [Fact]
    public async Task RereadingXmlIsChargedOnceAgainstTheBudget()
    {
        var package = await LayoutOfd(true);
        var xml = Unzip(package).Where(x => x.Key.EndsWith(".xml", StringComparison.Ordinal)).Sum(x => (long)x.Value.Length);
        var options = new MedicalAssessmentOptions();
        // Enough for every distinct XML entry once, but not twice (MH-003 re-reads each one lexically).
        options.Archival.Preflight.MaxXmlBytes = xml + xml / 2;
        options.Archival.Preflight.MaxSingleXmlBytes = (int)xml;
        Assert.Equal(MedicalExecutionState.Complete, (await Assess(package, options)).ExecutionState);

        // Entry identity is case-insensitive: a DocRoot spelled in another case is the same entry, charged once.
        var recased = Mutate(package, e => e.Set("OFD.xml", e.Text("OFD.xml").Replace("Doc_0/Document.xml", "dOC_0/document.xml")));
        Assert.Equal(MedicalExecutionState.Complete, (await Assess(recased, options)).ExecutionState);
    }

    private static Task<byte[]> WithMedia(byte[] package, string type, string file, byte[] payload) =>
        Task.FromResult(Mutate(package, e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:PublicRes>", "</ofd:PublicRes><ofd:DocumentRes>DocumentRes.xml</ofd:DocumentRes>"));
            e.Set("Doc_0/DocumentRes.xml", Open("Res", " BaseLoc='Res'") + $"<ofd:MultiMedias><ofd:MultiMedia ID='90' Type='{type}'><ofd:MediaFile>{file}</ofd:MediaFile></ofd:MultiMedia></ofd:MultiMedias></ofd:Res>");
            if (payload.Length > 0) e["Doc_0/Res/" + file] = payload;
        }));

    [Fact]
    public async Task ImagesAreCheckedByPayloadSignatureNotFileName()
    {
        var source = await LayoutOfd(true);
        var png = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0 };
        var gif = Encoding.ASCII.GetBytes("GIF89a\0\0\0\0\0\0");

        var report = await Assess(await WithMedia(source, "Image", "scan.png", png));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("7.2(c)").Outcome);
        Assert.Equal("PermittedImageFormats", report.Single("7.2(d)", "MH-009").ReasonCode);

        report = await Assess(await WithMedia(source, "Image", "scan.png", gif));
        Assert.Equal("ImageFormatNotPermitted", report.Single("7.2(d)", "MH-009").ReasonCode);

        report = await Assess(await WithMedia(source, "Image", "missing.png", Array.Empty<byte>()));
        Assert.Equal("ImageNotInFile", report.Single("7.2(c)", "MH-009").ReasonCode);
        Assert.Equal("ResourceNotInContainer", report.Single("7.1(b)", "MH-006").ReasonCode);

        // A readable PNG does not vouch for a page resource table that could not be read.
        var partial = Mutate(await WithMedia(source, "Image", "scan.png", png), e =>
            e.Set("Doc_0/Pages/Page_1/Content.xml", e.Text("Doc_0/Pages/Page_1/Content.xml").Replace("<ofd:Content>", "<ofd:PageRes>PageRes.xml</ofd:PageRes><ofd:Content>")));
        report = await Assess(partial);
        foreach (var clause in new[] { "7.2(c)", "7.2(d)" })
            Assert.Equal("ResourceScopeIncomplete", report.Single(clause, "MH-009").ReasonCode);

        // A readable PNG beside an image whose payload is missing does not establish the permitted formats.
        var oneMissing = Mutate(await WithMedia(source, "Image", "scan.png", png), e =>
            e.Set("Doc_0/DocumentRes.xml", e.Text("Doc_0/DocumentRes.xml").Replace("</ofd:MultiMedias>",
                "<ofd:MultiMedia ID='91' Type='Image'><ofd:MediaFile>gone.png</ofd:MediaFile></ofd:MultiMedia></ofd:MultiMedias>")));
        Assert.Equal("ImageUnavailable", (await Assess(oneMissing)).Single("7.2(d)", "MH-009").ReasonCode);

        // An unresolved page leaves the page sequence, and so linearization, unassessed.
        Assert.Equal("PagesUnresolved", (await Assess(Mutate(source, e =>
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "<ofd:Page ID='98' BaseLoc='../../../escape.xml'/></ofd:Pages>")))))
            .Single("6.1.3(a)", "MH-005").ReasonCode);

        // A page whose BaseLoc does not resolve may declare further images in its PageRes.
        var unresolvedPage = Mutate(await WithMedia(source, "Image", "scan.png", png), e =>
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "<ofd:Page ID='99' BaseLoc='../../../escape.xml'/></ofd:Pages>")));
        Assert.Equal("ResourceScopeIncomplete", (await Assess(unresolvedPage)).Single("7.2(d)", "MH-009").ReasonCode);
    }

    [Fact]
    public async Task AudioFormatsAndTheEmbeddedObjectLimit()
    {
        var source = await LayoutOfd(true);
        var mp3 = Encoding.ASCII.GetBytes("ID3\u0004\0\0\0\0\0\0");
        var report = await Assess(await WithMedia(source, "Audio", "voice.mp3", mp3));
        Assert.Equal("MP3Payloads", report.Single("7.2(f)", "MH-010").ReasonCode);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("7.2(e)").Outcome);

        // Declared audio whose payload is missing, or audio that may sit in an unreadable table, is not "no audio".
        report = await Assess(await WithMedia(source, "Audio", "gone.mp3", Array.Empty<byte>()));
        Assert.Equal("MediaUnavailable", report.Single("7.2(f)", "MH-010").ReasonCode);
        Assert.Equal("MediaUnavailable", report.Single("7.2(e)", "MH-010").ReasonCode);
        report = await Assess(Mutate(await WithMedia(source, "Audio", "voice.mp3", mp3), e =>
            e.Set("Doc_0/Pages/Page_1/Content.xml", e.Text("Doc_0/Pages/Page_1/Content.xml").Replace("<ofd:Content>", "<ofd:PageRes>PageRes.xml</ofd:PageRes><ofd:Content>"))));
        foreach (var clause in new[] { "7.2(f)", "7.2(g)", "7.2(e)" })
            Assert.Equal("ResourceScopeIncomplete", report.Single(clause, "MH-010").ReasonCode);

        // "MP3 等" admits other common formats: an unrecognized payload stays Unknown.
        report = await Assess(await WithMedia(source, "Audio", "voice.ogg", Encoding.ASCII.GetBytes("OggS\0\0\0\0\0\0")));
        Assert.Equal("OtherAudioFormat", report.Single("7.2(f)", "MH-010").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("7.2(f)").Outcome);

        report = await Assess(await WithMedia(source, "Audio", "voice.mp3", mp3), new MedicalAssessmentOptions { MaxEmbeddedObjectBytes = 4 });
        Assert.Equal("EmbeddedObjectOverLimit", Assert.Single(report.GetClause("7.2(e)").Advisories).ReasonCode);
        Assert.NotEqual(PreflightOutcome.Fail, report.GetClause("7.2(e)").Outcome);
    }

    [Fact]
    public async Task MetadataMappingIsCallerOwned()
    {
        var package = await LayoutOfd(true);
        var report = await Assess(package, new MedicalAssessmentOptions { Metadata = null });
        Assert.Equal("FieldMappingNotProvided", report.Single("7.3(a)", "MH-011").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("7.3(b)").Outcome);

        report = await Assess(package, new MedicalAssessmentOptions { Metadata = new MedicalMetadataMapping { PatientField = "HIS.Patient", FileTypeField = "EMR.FileType", SecretLevelField = "EMR.SecretLevel" } });
        Assert.Contains("HIS.Patient", report.Single("7.3(a)", "MH-011").Evidence.Observed);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause("7.3(a)").Outcome);
    }

    [Fact]
    public async Task PermissionsMustNotProhibitAnnotationOrSigning()
    {
        var package = AddToDocument(await LayoutOfd(true), "<ofd:Permissions><ofd:Edit>false</ofd:Edit><ofd:Annot>false</ofd:Annot></ofd:Permissions>");
        var observation = (await Assess(package)).Single("7.3(g)", "MH-015");
        Assert.Equal("OperationProhibited", observation.ReasonCode);
        Assert.Contains("Annot", observation.Evidence.Observed);
        Assert.DoesNotContain("Edit", observation.Evidence.Observed);
    }

    private static byte[] Signed(byte[] package, string signatureMethod, string checkMethod, Action<Dictionary<string, byte[]>>? afterSigning = null) =>
        Mutate(package, e =>
        {
            e.Set("OFD.xml", e.Text("OFD.xml").Replace("</ofd:DocRoot>", "</ofd:DocRoot><ofd:Signatures>Doc_0/Signs/Signatures.xml</ofd:Signatures>"));
            var references = string.Concat(e.Keys.OrderBy(x => x, StringComparer.Ordinal).Select(x => $"<ofd:Reference FileRef='/{x}'><ofd:CheckValue>AA==</ofd:CheckValue></ofd:Reference>"));
            e.Set("Doc_0/Signs/Signatures.xml", Open("Signatures") + "<ofd:MaxSignId>1</ofd:MaxSignId><ofd:Signature ID='1' BaseLoc='Sign_0/Signature.xml'/></ofd:Signatures>");
            e.Set("Doc_0/Signs/Sign_0/Signature.xml", Open("Signature") + $"<ofd:SignedInfo><ofd:Provider ProviderName='test'/><ofd:SignatureMethod>{signatureMethod}</ofd:SignatureMethod>" +
                $"<ofd:References CheckMethod='{checkMethod}'>{references}</ofd:References></ofd:SignedInfo><ofd:SignedValue>/Doc_0/Signs/Sign_0/SignedValue.dat</ofd:SignedValue></ofd:Signature>");
            e["Doc_0/Signs/Sign_0/SignedValue.dat"] = new byte[] { 1, 2, 3 };
            afterSigning?.Invoke(e);
        });

    [Fact]
    public async Task SignatureCoverageExcludesOnlyAnnotationAndSignatureFiles()
    {
        var source = await LayoutOfd(true);
        var report = await Assess(Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid));
        foreach (var clause in new[] { "6.3(b)", "6.3(e)", "6.3(g)", "7.3(d)", "8(c)", "8(d)" })
            Assert.True(report.GetClause(clause).Outcome == PreflightOutcome.Pass, clause + " is " + report.GetClause(clause).Outcome);
        Assert.Contains("digests not verified", report.Single("8(c)", "MH-013").Evidence.Observed);
        Assert.Contains("CryptographicVerificationNotPerformed", report.Document.Gaps);

        // Annotations added after signing stay outside the protected scope (6.3(f), 8(c)).
        var annotated = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "</ofd:Pages><ofd:Annotations>Annots/Annotations.xml</ofd:Annotations>"));
            e.Set("Doc_0/Annots/Annotations.xml", Open("Annotations") + "<ofd:Page PageID='1'><ofd:FileLoc>Page_0/Annotation.xml</ofd:FileLoc></ofd:Page></ofd:Annotations>");
            e.Set("Doc_0/Annots/Page_0/Annotation.xml", Open("PageAnnot") + "</ofd:PageAnnot>");
        });
        // Structural coverage only: the changed Document.xml is still referenced; digests are not verified here.
        var annotatedReport = await Assess(annotated);
        Assert.Equal(PreflightOutcome.Pass, annotatedReport.GetClause("8(c)").Outcome);
        Assert.Equal(PreflightOutcome.Pass, annotatedReport.GetClause("7.1(f)").Outcome);

        // A reference that cannot be resolved may name a page, so an unlisted page is Unknown, not Fail.
        var unresolved = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
            e.Set("Doc_0/Signs/Sign_0/Signature.xml", e.Text("Doc_0/Signs/Sign_0/Signature.xml")
                .Replace("FileRef='/Doc_0/Pages/Page_0/Content.xml'", "FileRef='../../Pages/Page_0/Content.xml'")));
        var unresolvedReport = await Assess(unresolved);
        Assert.Equal("UnresolvableReference", unresolvedReport.Single("6.3(e)", "MH-013").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, unresolvedReport.GetClause("6.3(e)").Outcome);

        // An unresolvable last signature is not replaced by the earlier, readable one.
        foreach (var baseLoc in new[] { "Sign_9/Signature.xml", "../../../../escape.xml" })
        {
            var trailing = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
                e.Set("Doc_0/Signs/Signatures.xml", e.Text("Doc_0/Signs/Signatures.xml").Replace("</ofd:Signatures>", $"<ofd:Signature ID='2' BaseLoc='{baseLoc}'/></ofd:Signatures>")));
            var trailingReport = await Assess(trailing);
            foreach (var clause in new[] { "6.3(e)", "6.3(g)", "8(c)", "8(d)", "6.3(b)" })
                Assert.True(trailingReport.GetClause(clause).Outcome == PreflightOutcome.Unknown, baseLoc + " " + clause + " is " + trailingReport.GetClause(clause).Outcome);
        }

        // A second annotation index and its page files are excluded from the signed scope like the first.
        var twoAnnotationIndexes = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>",
                "</ofd:Pages><ofd:Annotations>Annots/Annotations.xml</ofd:Annotations><ofd:Annotations>Annots/Second.xml</ofd:Annotations>"));
            e.Set("Doc_0/Annots/Annotations.xml", Open("Annotations") + "<ofd:Page PageID='1'><ofd:FileLoc>Page_0/Annotation.xml</ofd:FileLoc></ofd:Page></ofd:Annotations>");
            e.Set("Doc_0/Annots/Page_0/Annotation.xml", Open("PageAnnot") + "</ofd:PageAnnot>");
            e.Set("Doc_0/Annots/Second.xml", Open("Annotations") + "<ofd:Page PageID='2'><ofd:FileLoc>Page_1/Annotation.xml</ofd:FileLoc></ofd:Page></ofd:Annotations>");
            e.Set("Doc_0/Annots/Page_1/Annotation.xml", Open("PageAnnot") + "</ofd:PageAnnot>");
        });
        Assert.Equal(PreflightOutcome.Pass, (await Assess(twoAnnotationIndexes)).GetClause("8(c)").Outcome);

        // An unreadable annotation index leaves its page files' status unknown, not unsigned content.
        var unreadableAnnotationIndex = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "</ofd:Pages><ofd:Annotations>Annots/Annotations.xml</ofd:Annotations>"));
            e.Set("Doc_0/Annots/Annotations.xml", "<not-closed");
            e.Set("Doc_0/Annots/Page_0/Annotation.xml", Open("PageAnnot") + "</ofd:PageAnnot>");
        });
        var unreadableAnnotationReport = await Assess(unreadableAnnotationIndex);
        foreach (var clause in new[] { "8(c)", "6.3(g)" })
            Assert.Equal("AnnotationIndexUnreadable", unreadableAnnotationReport.Single(clause, "MH-013").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, unreadableAnnotationReport.GetClause("7.1(f)").Outcome);

        // Only signature files and their declared payloads are exempt, not everything beside them.
        var besideSignature = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e => e.Set("Doc_0/Signs/Sign_0/record.json", "{}"));
        var besideReport = await Assess(besideSignature);
        Assert.Equal("ContentNotProtected", besideReport.Single("8(c)", "MH-013").ReasonCode);
        Assert.Contains("Doc_0/Signs/Sign_0/record.json", besideReport.Single("8(c)", "MH-013").Evidence.Observed);

        // A seal declared the way OfdSignatureService writes it (Seal/@BaseLoc) stays exempt.
        var sealed_ = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
        {
            e.Set("Doc_0/Signs/Sign_0/Signature.xml", e.Text("Doc_0/Signs/Sign_0/Signature.xml").Replace("<ofd:References", "<ofd:Seal BaseLoc='/Doc_0/Signs/Sign_0/Seal.esl'/><ofd:References"));
            e["Doc_0/Signs/Sign_0/Seal.esl"] = new byte[] { 0x30, 0x00 };
        });
        Assert.Equal(PreflightOutcome.Pass, (await Assess(sealed_)).GetClause("8(c)").Outcome);

        // A page declared at an unresolvable location is not shown to be covered.
        var unresolvedPage = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:Pages>", "<ofd:Page ID='99' BaseLoc='../../../escape.xml'/></ofd:Pages>")));
        Assert.Equal("PagesUnresolved", (await Assess(unresolvedPage)).Single("6.3(e)", "MH-013").ReasonCode);

        // A Signatures index that does not resolve is not "no signature declared".
        var unresolvedIndex = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e =>
            e.Set("OFD.xml", e.Text("OFD.xml").Replace("<ofd:Signatures>Doc_0/Signs/Signatures.xml", "<ofd:Signatures>Doc_0\\Signs\\Signatures.xml")));
        var unresolvedIndexReport = await Assess(unresolvedIndex);
        foreach (var clause in new[] { "6.3(e)", "8(c)", "8(d)" })
            Assert.True(unresolvedIndexReport.GetClause(clause).Outcome == PreflightOutcome.Unknown, clause + " is " + unresolvedIndexReport.GetClause(clause).Outcome);
        Assert.Equal("SignatureUnreadable", unresolvedIndexReport.Single("7.3(d)", "MH-013").ReasonCode);

        // A new attachment after signing needs a new signature (6.3(g)).
        var extended = Signed(source, MedicalRules.Sm2Oid, MedicalRules.Sm3Oid, e => e.Set("Doc_0/Attachs/later.xml", "<later/>"));
        Assert.Equal("UnsignedModification", (await Assess(extended)).Single("6.3(g)", "MH-013").ReasonCode);
    }

    [Fact]
    public async Task SignatureAlgorithmsMustBeSmWithGbt33560Identifiers()
    {
        var source = await LayoutOfd(true);
        var named = await Assess(Signed(source, "SM2", "sm3"));
        Assert.Equal(PreflightOutcome.Pass, named.GetClause("8(d)").Outcome);
        Assert.Equal("IdentifierNotOid", named.Single("6.3(b)", "MH-014").ReasonCode);

        // SM identifiers in swapped or duplicated roles are not SM2 signing with an SM3 digest.
        foreach (var (method, check) in new[] { (MedicalRules.Sm3Oid, MedicalRules.Sm2Oid), (MedicalRules.Sm3Oid, MedicalRules.Sm3Oid) })
        {
            var swapped = await Assess(Signed(source, method, check));
            Assert.Equal("AlgorithmRoleMismatch", swapped.Single("8(d)", "MH-014").ReasonCode);
            Assert.Equal(PreflightOutcome.Fail, swapped.GetClause("6.3(b)").Outcome);
        }

        var rsa = await Assess(Signed(source, "1.2.840.113549.1.1.11", "2.16.840.1.101.3.4.2.1"));
        Assert.Equal("NonSmAlgorithm", rsa.Single("8(d)", "MH-014").ReasonCode);
        Assert.Equal(PreflightOutcome.Fail, rsa.GetClause("8(d)").Outcome);
    }

    [Fact]
    public async Task MachineReadableTagsMustBeUtf8()
    {
        byte[] WithTag(byte[] payload) => AddToDocument(LayoutOfd(true).GetAwaiter().GetResult(), "<ofd:CustomTags>Tags/CustomTags.xml</ofd:CustomTags>", e =>
        {
            e.Set("Doc_0/Tags/CustomTags.xml", Open("CustomTags") + "<ofd:CustomTag NameSpace='urn:example:emr'><ofd:FileLoc>EMR.xml</ofd:FileLoc></ofd:CustomTag></ofd:CustomTags>");
            e["Doc_0/Tags/EMR.xml"] = payload;
        });
        var report = await Assess(WithTag(Encoding.UTF8.GetBytes("<?xml version='1.0' encoding='UTF-8'?><emr><patient>张三</patient></emr>")));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("7.5(a)").Outcome);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("7.5(b)").Outcome);
        Assert.Equal(PreflightOutcome.Pass, report.GetClause("7.1(d)").Outcome);

        report = await Assess(WithTag(Encoding.Latin1.GetBytes("<?xml version='1.0' encoding='ISO-8859-1'?><emr>café</emr>")));
        Assert.Equal("NotUtf8", report.Single("7.5(b)", "MH-012").ReasonCode);

        // A second CustomTags index is assessed too, not only the first.
        var twoIndexes = Mutate(WithTag(Encoding.UTF8.GetBytes("<emr>ok</emr>")), e =>
        {
            e.Set("Doc_0/Document.xml", e.Text("Doc_0/Document.xml").Replace("</ofd:CustomTags>", "</ofd:CustomTags><ofd:CustomTags>Tags/Second.xml</ofd:CustomTags>"));
            e.Set("Doc_0/Tags/Second.xml", Open("CustomTags") + "<ofd:CustomTag NameSpace='urn:example:emr2'><ofd:FileLoc>EMR2.xml</ofd:FileLoc></ofd:CustomTag></ofd:CustomTags>");
            e["Doc_0/Tags/EMR2.xml"] = Encoding.Latin1.GetBytes("<?xml version='1.0' encoding='ISO-8859-1'?><emr>café</emr>");
        });
        report = await Assess(twoIndexes);
        Assert.Equal("NotUtf8", report.Single("7.5(b)", "MH-012").ReasonCode);
        Assert.Contains("EMR2.xml", report.Single("7.5(b)", "MH-012").Evidence.Observed);

        // A declared payload that is missing, or an unreadable index, is unknown rather than absent.
        var missing = Mutate(WithTag(Encoding.UTF8.GetBytes("<emr>ok</emr>")), e =>
            e.Set("Doc_0/Tags/CustomTags.xml", e.Text("Doc_0/Tags/CustomTags.xml").Replace("</ofd:CustomTags>",
                "<ofd:CustomTag NameSpace='urn:example:gone'><ofd:FileLoc>Gone.xml</ofd:FileLoc></ofd:CustomTag></ofd:CustomTags>")));
        report = await Assess(missing);
        foreach (var clause in new[] { "7.5(a)", "7.5(b)" })
            Assert.Equal("PayloadUnavailable", report.Single(clause, "MH-012").ReasonCode);
        var unreadableIndex = Mutate(WithTag(Encoding.UTF8.GetBytes("<emr>ok</emr>")), e => e.Set("Doc_0/Tags/CustomTags.xml", "<not-closed"));
        report = await Assess(unreadableIndex);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("7.5(b)").Outcome);

        // Attachments may give FileLoc as an attribute; the payload is still inspected.
        var attributeForm = AddToDocument(LayoutOfd(true).GetAwaiter().GetResult(), "<ofd:Attachments>Attachs/Attachments.xml</ofd:Attachments>", e =>
        {
            e.Set("Doc_0/Attachs/Attachments.xml", Open("Attachments") + "<ofd:Attachment ID='1' Name='emr' FileLoc='emr.xml'/></ofd:Attachments>");
            e["Doc_0/Attachs/emr.xml"] = Encoding.Latin1.GetBytes("<?xml version='1.0' encoding='ISO-8859-1'?><emr>café</emr>");
        });
        Assert.Equal("NotUtf8", (await Assess(attributeForm)).Single("7.5(b)", "MH-012").ReasonCode);

        // Bytes that are not UTF-8 leave the structure unassessed rather than passing 7.5(a).
        var invalidJson = AddToDocument(LayoutOfd(true).GetAwaiter().GetResult(), "<ofd:Attachments>Attachs/Attachments.xml</ofd:Attachments>", e =>
        {
            e.Set("Doc_0/Attachs/Attachments.xml", Open("Attachments") + "<ofd:Attachment ID='1' Name='record'><ofd:FileLoc>record.json</ofd:FileLoc></ofd:Attachment></ofd:Attachments>");
            e["Doc_0/Attachs/record.json"] = new byte[] { 0xFF };
        });
        report = await Assess(invalidJson);
        Assert.Equal("StructureNotAssessed", report.Single("7.5(a)", "MH-012").ReasonCode);
        Assert.Equal("NotUtf8", report.Single("7.5(b)", "MH-012").ReasonCode);
    }

    [Fact]
    public async Task EncryptionIsOnlyTemporary()
    {
        var package = Mutate(await LayoutOfd(true), e => e.Set("Encryptions.xml", Open("Encryptions") + "</ofd:Encryptions>"));
        var report = await Assess(package);
        Assert.Equal("TemporaryEncryptionDeclared", report.Single("8(b)", "MH-016").ReasonCode);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("8(b)").Outcome);
    }

    [Fact]
    public async Task BudgetsAndMalformedInputsAreExecutionStatesNotVerdicts()
    {
        var package = await LayoutOfd(true);
        var options = new MedicalAssessmentOptions();
        options.Archival.Preflight.LoadOptions.MaxInputBytes = 100;
        var report = await Assess(package, options);
        Assert.Equal(MedicalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Null(report.Input);
        Assert.All(report.Clauses, x => Assert.Equal(PreflightOutcome.Unknown, x.Outcome));

        report = await Assess(Encoding.ASCII.GetBytes("not a zip archive at all"));
        Assert.Equal(MedicalExecutionState.Error, report.ExecutionState);
        Assert.Equal(MedicalDeclarationStatus.Unavailable, report.Declaration.Status);

        await Assert.ThrowsAsync<ArgumentException>(() => new MedicalAssessmentService().AssessAsync(new NonSeekableInput(package)));
        await Assert.ThrowsAsync<ArgumentException>(() => new MedicalAssessmentService().AssessAsync(new MemoryStream(package), new MedicalAssessmentOptions { MaxEmbeddedObjectBytes = 0 }));
    }

    [Fact]
    public async Task ImplicitSiblingPageResIsAPageResourceLikeTheReaderAndArchivalChecksTreatIt()
    {
        // Readers and archival preflight treat {page}/PageRes.xml as the page's resource table without a PageRes link;
        // the shared scope walk now does too, so a font defined there is in PageRes for 6.2(b).
        var package = Mutate(await LayoutOfd(true), e => e.Set("Doc_0/Pages/Page_0/PageRes.xml",
            "<ofd:Res xmlns:ofd=\"http://www.ofdspec.org/2016\"><ofd:Fonts><ofd:Font ID=\"900\" FontName=\"X\"/></ofd:Fonts></ofd:Res>"));
        var observation = (await Assess(package)).Single("6.2(b)", "MH-007");
        Assert.Equal(PreflightOutcome.Fail, observation.Outcome);
        Assert.Equal("DefinedInPageRes", observation.ReasonCode);
    }

    [Fact]
    public async Task PercentEscapedLocationIsUnresolvableUnderTheSharedReferencePolicy()
    {
        // Medical once accepted %-escapes that Archival and Preflight reject; one policy now decides for all of them.
        var package = Mutate(await LayoutOfd(true), e => e.Set("Doc_0/Document.xml",
            e.Text("Doc_0/Document.xml").Replace("Pages/Page_1/Content.xml", "Pages/Page%5F1/Content.xml")));
        var observation = (await Assess(package)).Single("6.2(a)", "MH-006");
        Assert.Equal(PreflightOutcome.Unknown, observation.Outcome);
        Assert.Equal("UnresolvableLocation", observation.ReasonCode);
    }
}
