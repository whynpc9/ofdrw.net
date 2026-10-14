using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Keywords.Internal;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Packaging.Tests;

// Existing fake metrics qualify only this private component. These tests do not
// claim CLI sealed-provider eligibility, fresh measurement, receipt or publish.
public sealed class TextGeometryInputEquivalenceTests
{
    [Fact]
    public async Task ActualWriterReaderInputsReplayCompiledSlotsWithoutAnotherProvider()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        var metrics = new TextReplacementTests.Metrics();
        var report = await session.PrepareAndWriteOwnedAsync(metrics, "B", "--");
        await session.ReadOwnedOutputAsync();
        var counts = session.VerifyOwnedGeometryInputs();
        Assert.True(report.Changed); Assert.Equal(1, counts.Pages); Assert.Equal(1, counts.Roots);
        Assert.Equal(1, counts.Slots); Assert.Equal(4, counts.Scalars); Assert.Equal(0, counts.ProviderOpens); Assert.Equal(0, counts.ProviderMeasures);
        Assert.Equal(1, metrics.Opens); Assert.Equal(2, metrics.Measures); Assert.Equal(1, metrics.Disposals);
        Assert.Equal(1, counts.FreshFontHashOperations); Assert.Equal(3, counts.FreshFontHashBytes);
        Assert.Throws<InvalidOperationException>(() => session.VerifyOwnedGeometryInputs());
        Assert.Throws<InvalidOperationException>(() => session.ReserveOwnedAtomicMetadata(fixture.Input + ".stage", default));
    }

    [Theory]
    [InlineData("--", 6)]
    [InlineData("", 4)]
    public async Task EveryAffectedNeighborAndZeroSlotSurvivesTheActualCompiledCapsule(string replacement, int scalars)
    {
        using var fixture = await Fixture.Create(neighbors: true); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", replacement);
        await session.ReadOwnedOutputAsync(); var counts = session.VerifyOwnedGeometryInputs();
        Assert.Equal(3, counts.Roots); Assert.Equal(5, counts.Slots); Assert.Equal(scalars, counts.Scalars);
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        Assert.Equal(3, operation.GeometryExpectation!.Roots.Count);
        Assert.True(operation.GeometryExpectation.Roots[2].IdentityOnly);
        Assert.False(operation.GeometryExpectation.Roots[0].IdentityOnly);
        Assert.Equal(new[] { 0, 2, 0 }, operation.GeometryExpectation.Roots[1].Slots.Select(s => s.Scalars.Count));
    }

    [Fact]
    public async Task DeletionRetainsTheActualMeasuredQualificationAndFiniteEmptySlot()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "ABC", "");
        await session.ReadOwnedOutputAsync(); var counts = session.VerifyOwnedGeometryInputs();
        Assert.Equal(1, counts.Slots); Assert.Equal(0, counts.Scalars);
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var root = Assert.Single(operation.GeometryExpectation!.Roots);
        Assert.False(root.IdentityOnly); Assert.Empty(Assert.Single(root.Slots).Scalars);
    }

    [Fact]
    public async Task NonzeroPageOriginAndOriginalTranslationAreAppliedExactlyOnce()
    {
        using var fixture = await Fixture.Create(translated: true); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"); await session.ReadOwnedOutputAsync();
        Assert.Equal(4, session.VerifyOwnedGeometryInputs().Scalars);
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var root = Assert.Single(operation.GeometryExpectation!.Roots);
        Assert.Equal(5, root.PageX); Assert.Equal(7, root.PageY); Assert.Equal(.5, root.Tx); Assert.Equal(.75, root.Ty);
        Assert.Equal(2, root.Slots[0].Scalars[0].X); Assert.Equal(8, root.Slots[0].Scalars[0].Y);
    }

    [Fact]
    public async Task UnchangedNeighborRetainsOmittedAxisAndItsAuthoritativeInheritedStart()
    {
        using var fixture = await Fixture.Create(neighbors: true, inherited: true); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"); await session.ReadOwnedOutputAsync();
        Assert.Equal(6, session.VerifyOwnedGeometryInputs().Scalars);
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var inherited = operation.GeometryExpectation!.Roots[1].Slots[1];
        Assert.DoesNotContain(inherited.Attributes, a => a.Name == "X"); Assert.Equal(2, inherited.X);
        var actual = Assert.IsType<OfdDocumentPackage>(Field(session, "ownedReadback"));
        Assert.Equal(0, ((OfdTextElement)actual.Pages[0].Elements[1]).Runs[1].XMillimeters);
    }

    [Theory]
    [InlineData("source")]
    [InlineData("context")]
    public async Task PastAliasLoanPermanentlyRefusesPrivateQualification(string loan)
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        if (loan == "source") _ = session.BoundSource; else _ = session.BoundContext;
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"));
        Assert.Contains("AliasWasLoaned", error.Message);
        Assert.Null(Field(session, "ownedWriter")); Assert.Null(Field(session, "ownedSpool"));
        await Assert.ThrowsAsync<InvalidOperationException>(() => session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"));
    }

    [Theory]
    [InlineData("text")]
    [InlineData("raw-axes")]
    [InlineData("fill")]
    [InlineData("size")]
    [InlineData("font-metadata")]
    [InlineData("resource-base")]
    [InlineData("family-name")]
    [InlineData("charset")]
    [InlineData("fixed-width")]
    [InlineData("root-name")]
    [InlineData("root-visible")]
    [InlineData("transform-presence")]
    [InlineData("boundary-lexical")]
    [InlineData("comment")]
    [InlineData("cdata")]
    [InlineData("font-model")]
    [InlineData("font-entry")]
    [InlineData("font-both")]
    [InlineData("extra-root")]
    public async Task ActualReadbackMutantsCannotSupplyTheirOwnExpectedInputs(string mutation)
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--");
        await session.ReadOwnedOutputAsync();
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var readback = Assert.IsType<OfdDocumentPackage>(Field(session, "ownedReadback"));
        // Actual model/authoritative entries, never fabricated expected records.
        Mutate(readback, mutation);
        var error = Record.Exception(() => OfdTextGeometryInputVerifier.CompareIdentityBoundComponent(operation.GeometryExpectation!, operation, readback, default));
        Assert.True(error is InvalidDataException or NotSupportedException, error?.ToString());
        Assert.Throws<InvalidOperationException>(() => operation.Context.AssertActive());
    }

    [Fact]
    public async Task DistinctEqualFontArraysBothReceiveFreshHashes()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"); await session.ReadOwnedOutputAsync();
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var actual = Assert.IsType<OfdDocumentPackage>(Field(session, "ownedReadback"));
        actual.Fonts[0].Data = actual.Fonts[0].Data.ToArray();
        var counts = OfdTextGeometryInputVerifier.CompareIdentityBoundComponent(operation.GeometryExpectation!, operation, actual, default);
        Assert.Equal(2, counts.FreshFontHashOperations); Assert.Equal(6, counts.FreshFontHashBytes);
    }

    [Fact]
    public async Task ForeignOperationAndCancelledComparisonAreRejected()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"); await session.ReadOwnedOutputAsync();
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var actual = Assert.IsType<OfdDocumentPackage>(Field(session, "ownedReadback"));
        using var foreign = new OfdTextReplacementOperation(actual, new(), default);
        Assert.Throws<InvalidOperationException>(() => OfdTextGeometryInputVerifier.CompareIdentityBoundComponent(operation.GeometryExpectation!, foreign, actual, default));
        Assert.Throws<InvalidOperationException>(() => foreign.Context.AssertActive());
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdTextGeometryInputVerifier.CompareIdentityBoundComponent(operation.GeometryExpectation!, operation, actual, cancelled.Token));
        Assert.Throws<InvalidOperationException>(() => operation.Context.AssertActive());
    }

    [Fact]
    public async Task SupplementaryInsertionKeepsItsExactScalarSpanGlyphAndPreparedProvenance()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "\U00020000");
        await session.ReadOwnedOutputAsync(); var counts = session.VerifyOwnedGeometryInputs();
        Assert.Equal(3, counts.Scalars);
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var inserted = operation.GeometryExpectation!.Roots[0].Slots[0].Scalars[1];
        Assert.True(inserted.Inserted); Assert.Equal(0x20000, inserted.Metric.Scalar); Assert.Equal((uint)0x20000, inserted.Metric.GlyphId);
        Assert.Equal(2, inserted.RawLength); Assert.Equal(2, inserted.Metric.Utf16Length); Assert.Equal(-1, inserted.OriginalRawStart);
    }

    [Fact]
    public async Task FreshCaptureQuotaAdmissionFailsTerminallyWithoutAFontHashOrFallback()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        await session.PrepareAndWriteOwnedAsync(new TextReplacementTests.Metrics(), "B", "--"); await session.ReadOwnedOutputAsync();
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        var actual = Assert.IsType<OfdDocumentPackage>(Field(session, "ownedReadback")); var material = operation.Context.Material;
        material.ChargeCopy(536_870_912L - material.CloneCopyBytes - 3072L, default); var work = material.ValidationWorkBytes;
        var error = Assert.Throws<InvalidDataException>(() => OfdTextGeometryInputVerifier.CompareIdentityBoundComponent(operation.GeometryExpectation!, operation, actual, default));
        Assert.Contains("clone-copy", error.Message); Assert.Equal(work, material.ValidationWorkBytes);
        Assert.Throws<InvalidOperationException>(() => operation.Context.AssertActive());
    }

    [Fact]
    public async Task ProviderClosureFailureNeverSealsTheFrozenLiveDraft()
    {
        using var fixture = await Fixture.Create(); using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        var primary = new InvalidDataException("actual final provider disposal failed");
        var provider = new TextReplacementTests.Metrics { Callback = phase => { if (phase == "dispose") throw primary; } };
        Assert.Same(primary, await Record.ExceptionAsync(() => session.PrepareAndWriteOwnedAsync(provider, "B", "--")));
        var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
        Assert.Null(operation.GeometryExpectation); Assert.Null(Field(session, "ownedWriter")); Assert.Null(Field(session, "ownedSpool"));
        Assert.Equal(1, provider.Disposals);
    }

    [Fact]
    public async Task LongEqualDistinctFontMetadataPaysAgainForEveryRootProjection()
    {
        using var fixture = await Fixture.Create(longMetadata: true);
        using var input = File.OpenRead(fixture.Input); var package = await new OfdReader().ReadAsync(input);
        using var context = new OfdControlledEditContext(new());
        using var keyword = new OfdKeywordOperation(material: context.Material);
        var budget = new OfdKeywordLocatorBudget(new(), keyword);
        using var snapshot = new OfdKeywordSourceSnapshot(package, budget);
        var admission = new OfdKeywordSourceAdmission(package, snapshot, budget); admission.PrepareDocument();
        var text = TextReplacementTests.Text(package); var font = package.Fonts[0];
        var before = budget.ValidationCharacters;
        var first = admission.CaptureFontInput(font.Id, text); var firstCharacters = budget.ValidationCharacters - before;
        Assert.Equal(font.FontName, first.FontName); Assert.NotSame(font.FontName, first.FontName);
        Assert.Equal(font.FamilyName, first.FamilyName); Assert.NotSame(font.FamilyName, first.FamilyName);
        Assert.Equal(font.Charset, first.Charset); Assert.NotSame(font.Charset, first.Charset);
        var characterLowerBound = font.FontName.Length + font.FamilyName!.Length + font.Charset!.Length;
        Assert.True(firstCharacters >= characterLowerBound);
        before = budget.ValidationCharacters; var work = context.Material.ValidationWorkBytes;
        _ = admission.CaptureFontInput(font.Id, text);
        Assert.True(budget.ValidationCharacters - before >= characterLowerBound);
        Assert.True(context.Material.ValidationWorkBytes - work >= characterLowerBound);
    }

    [Fact]
    public async Task LongMetadataProjectionRefusesBeforeScanningWhenRemainingWorkIsInsufficient()
    {
        using var fixture = await Fixture.Create(longMetadata: true);
        using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        var context = Assert.IsType<OfdControlledEditContext>(Field(session, "context"));
        var package = Assert.IsType<OfdDocumentPackage>(Field(session, "source"));
        var setupReached = false; var projectionReturned = false;
        var metrics = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "measure" || setupReached) return;
            using var keyword = new OfdKeywordOperation(material: context.Material);
            var budget = new OfdKeywordLocatorBudget(new(), keyword);
            using var snapshot = new OfdKeywordSourceSnapshot(package, budget);
            var admission = new OfdKeywordSourceAdmission(package, snapshot, budget); admission.PrepareDocument();
            setupReached = true;
            var font = package.Fonts[0];
            // Allow less than even one complete long-name comparison after all
            // real admission setup. The enclosing owned route owns terminality.
            context.Material.ChargeWork(536_870_912L - context.Material.ValidationWorkBytes - font.FontName.Length / 2, default);
            _ = admission.CaptureFontInput(font.Id, TextReplacementTests.Text(package));
            projectionReturned = true;
        } };
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => session.PrepareAndWriteOwnedAsync(metrics, "B", "--"));
        Assert.True(setupReached, error.ToString()); Assert.False(projectionReturned); Assert.True(OfdTraversalBudget.IsLimit(error));
        Assert.Contains("validation work", error.Message); Assert.Contains("CaptureFontInput", error.StackTrace);
        Assert.True(Record.Exception(() => context.AssertActive()) is InvalidOperationException or ObjectDisposedException);
        Assert.Null(Field(session, "ownedWriter")); Assert.Null(Field(session, "ownedSpool"));
    }

    [Fact]
    public async Task FrozenLivePathAndSharedDigestReferencesAreAlreadyRetainedBeforeProviderClosure()
    {
        using var fixture = await Fixture.Create(neighbors: true);
        using var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(fixture.Input);
        var inspected = false;
        var metrics = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "statistics") return;
            var operation = Assert.IsType<OfdTextReplacementOperation>(Field(session, "ownedOperation"));
            Assert.Null(operation.GeometryExpectation);
            var builder = Field(operation, "geometryBuilder")!;
            var roots = Assert.IsType<OfdTextGeometryExpectation.Root[]>(Field(builder, "frozenRoots"));
            Assert.Equal(3, roots.Length); Assert.Same(roots[0].FontSha256, roots[1].FontSha256); Assert.Same(roots[1].FontSha256, roots[2].FontSha256);
            var storage = operation.Context.Material.StorageBytes;
            foreach (var root in roots)
            {
                operation.Context.Material.StoreString(root.Path); operation.Context.Material.StoreString(root.FontSha256);
                operation.Context.Material.StoreString(root.Path); operation.Context.Material.StoreString(root.FontSha256);
            }
            Assert.Equal(storage, operation.Context.Material.StorageBytes); inspected = true;
        } };
        await session.PrepareAndWriteOwnedAsync(metrics, "B", "--"); Assert.True(inspected);
    }

    [Fact]
    public async Task LowRemainingStorageRefusesAdmissionPathBeforeAnyCapsuleCanBeSealed()
    {
        using var fixture = await Fixture.Create();
        const long smallComponentCap = 16_000_000;
        using var calibrationInput = File.OpenRead(fixture.Input);
        var calibrationPackage = await new OfdReader().ReadAsync(calibrationInput);
        using var calibration = new OfdTextReplacementOperation(calibrationPackage,
            new OfdTextReplacementOptions(new(maxMaterializedBytes: smallComponentCap)), default);
        CaptureLiveComponent(calibration, new TextReplacementTests.Metrics());
        var allowance = checked(2 * (calibration.GeometryExpectation!.SourceVersion.Length + calibration.GeometryExpectation.Roots[0].Path.Length) - 1);
        using var input = File.OpenRead(fixture.Input); var package = await new OfdReader().ReadAsync(input);
        using var operation = new OfdTextReplacementOperation(package,
            new OfdTextReplacementOptions(new(maxMaterializedBytes: smallComponentCap)), default);
        var context = operation.Context; var filled = false;
        var metrics = new TextReplacementTests.Metrics { Callback = phase =>
        {
            if (phase != "measure" || filled) return;
            filled = true;
            // Fill a deliberately small component Storage cap, deriving the
            // remainder from real frozen strings. No production limit changes.
            context.Material.StoreString(new string('x', checked((int)((smallComponentCap - context.Material.StorageBytes - allowance) / 2))));
        } };
        var error = Assert.Throws<InvalidDataException>(() => CaptureLiveComponent(operation, metrics));
        operation.Fail(error);
        Assert.True(filled); Assert.True(OfdTraversalBudget.IsLimit(error)); Assert.Contains("input storage", error.Message);
        Assert.Contains("Builder.CaptureRoot", error.StackTrace); Assert.DoesNotContain("OfdStrictTextGeometry", error.StackTrace);
        Assert.Throws<InvalidOperationException>(() => context.AssertActive()); Assert.Null(operation.GeometryExpectation);
    }

    private static void CaptureLiveComponent(OfdTextReplacementOperation operation, TextReplacementTests.Metrics metrics)
    {
        // Friendly component invocation reuses the real analysis/fit/compiler
        // and its callback guards. It does not issue a qualified CLI lease.
        _ = OfdKeywordEditSource.Analyze(operation.Source!, new[] { 0 }, "B", metrics,
            new(maxNodes: 128, maxReferences: 256, maxOccurrences: 128), operation.Options.SearchOptions,
            operation.Context.Material, () => operation.Validate(default), capture =>
            {
                var requests = capture.CurrentHits.Select(hit => new OfdTextReplacement(hit, "--")).ToArray();
                var builder = operation.BeginGeometryExpectation(capture, 1);
                _ = OfdTextReplacementFitAdapter.Plan(capture, requests, operation, default, new() { 0 }, builder);
            }, default, operation.BeforeCallback);
        operation.CloseGeometryPreparation();
    }

    private static void Mutate(OfdDocumentPackage package, string mutation)
    {
        var text = TextReplacementTests.Text(package);
        if (mutation == "font-model") { package.Fonts[0].Data = package.Fonts[0].Data.ToArray(); package.Fonts[0].Data[0]++; return; }
        var fontEntry = package.PreservedEntries.First(p => p.Value.SequenceEqual(package.Fonts[0].Data)).Key;
        if (mutation == "font-entry") { package.PreservedEntries[fontEntry] = package.PreservedEntries[fontEntry].ToArray(); package.PreservedEntries[fontEntry][0]++; return; }
        if (mutation == "font-both") { package.Fonts[0].Data[0]++; return; }
        if (mutation is "font-metadata" or "resource-base" or "family-name" or "charset" or "fixed-width")
        {
            var resource = package.PreservedEntries.First(p => p.Key.EndsWith("Res.xml") && Encoding.UTF8.GetString(p.Value).Contains("FontFile"));
            var xml = XElement.Parse(Encoding.UTF8.GetString(resource.Value).TrimStart('\ufeff'));
            var font = xml.Descendants().First(n => n.Name.LocalName == "Font");
            if (mutation == "font-metadata") font.SetAttributeValue("Serif", "true");
            else if (mutation == "fixed-width") font.SetAttributeValue("FixedWidth", "false");
            else if (mutation == "family-name") { font.SetAttributeValue("FamilyName", "changed"); package.Fonts[0].FamilyName = "changed"; }
            else if (mutation == "charset") { font.SetAttributeValue("Charset", "changed"); package.Fonts[0].Charset = "changed"; }
            else xml.SetAttributeValue("BaseLoc", "./Res");
            package.PreservedEntries[resource.Key] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)); return;
        }
        if (mutation == "extra-root") { package.Pages[0].Elements.Add(text); return; }
        var root = XElement.Parse(text.SourceXml!); var code = root.Elements().First(n => n.Name.LocalName == "TextCode");
        if (mutation == "text") { code.Value = "A--D"; text.Text = "A--D"; text.Runs[0].Text = "A--D"; }
        if (mutation == "raw-axes") code.SetAttributeValue("X", "02");
        if (mutation == "root-name") root.SetAttributeValue("Name", "changed");
        if (mutation == "root-visible") root.SetAttributeValue("Visible", "true");
        if (mutation == "boundary-lexical") root.SetAttributeValue("Boundary", root.Attribute("Boundary")!.Value.Replace("20 ", "20.0 "));
        if (mutation == "transform-presence") { root.SetAttributeValue("CTM", "1 0 0 1 0 0"); text.Transform = new double[] { 1, 0, 0, 1, 0, 0 }; }
        if (mutation == "comment") code.Add(new XComment("unexpected"));
        if (mutation == "cdata") code.ReplaceNodes(new XCData(text.Text));
        if (mutation == "size") { root.SetAttributeValue("Size", "5"); text.FontSizeMillimeters = 5; }
        if (mutation == "fill")
        {
            var fill = root.Elements().FirstOrDefault(n => n.Name.LocalName == "FillColor");
            if (fill is null) { fill = new XElement(root.Name.Namespace + "FillColor"); root.AddFirst(fill); }
            fill.SetAttributeValue("Value", "1 0 0"); text.FillColor = new OfdColor(1, 0, 0);
        }
        text.SourceXml = root.ToString(SaveOptions.DisableFormatting);
        var entry = package.Pages[0].SourceEntryPath!; var page = XElement.Parse(Encoding.UTF8.GetString(package.PreservedEntries[entry]).TrimStart('\ufeff'));
        page.Descendants().First(n => n.Name.LocalName == "TextObject").ReplaceWith(new XElement(root));
        package.PreservedEntries[entry] = Encoding.UTF8.GetBytes(page.ToString(SaveOptions.DisableFormatting));
    }
    private static object? Field(object owner, string name) => owner.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(owner);
    private sealed class Fixture(string directory) : IDisposable
    {
        internal string Input => Path.Combine(directory, "input.ofd");
        internal static async Task<Fixture> Create(bool neighbors = false, bool translated = false, bool inherited = false, bool longMetadata = false)
        {
            var directory = Path.Combine(Path.GetTempPath(), "ofdrw-input-equivalence-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
            var fixture = new Fixture(directory); var package = await TextReplacementTests.Read();
            if (longMetadata)
            {
                package.Fonts[0].FontName = new string('n', 8_192) + "name";
                package.Fonts[0].FamilyName = new string('f', 8_192) + "family";
                package.Fonts[0].Charset = new string('c', 8_192) + "charset";
                TextReplacementTests.Text(package).FontName = package.Fonts[0].FontName;
            }
            if (translated)
            {
                package.Pages[0].XMillimeters = 5; package.Pages[0].YMillimeters = 7;
                var original = TextReplacementTests.Text(package); var root = XElement.Parse(original.SourceXml!);
                root.SetAttributeValue("CTM", "1 0 0 1 0.5 0.75"); original.SourceXml = root.ToString(SaveOptions.DisableFormatting);
                original.Transform = new double[] { 1, 0, 0, 1, .5, .75 };
            }
            if (neighbors)
            {
                var original = TextReplacementTests.Text(package);
                var neighbor = Neighbor(original, "41", "20 60 20 20", new[] { "", "WW", "" });
                if (inherited)
                {
                    var xml = XElement.Parse(neighbor.SourceXml!); xml.Elements(xml.Name.Namespace + "TextCode").ElementAt(1).Attribute("X")!.Remove();
                    neighbor.SourceXml = xml.ToString(SaveOptions.DisableFormatting);
                }
                package.Pages[0].Elements.Add(neighbor);
                package.Pages[0].Elements.Add(Neighbor(original, "42", "20 85 20 20", new[] { "" }));
            }
            using var file = File.Create(fixture.Input); await new OfdPackageWriter().WriteAsync(package, file); return fixture;
        }
        private static OfdTextElement Neighbor(OfdTextElement original, string id, string boundary, string[] texts)
        {
            var root = XElement.Parse(original.SourceXml!); root.SetAttributeValue("ID", id); root.SetAttributeValue("Boundary", boundary);
            root.Elements(root.Name.Namespace + "TextCode").Remove();
            foreach (var value in texts)
                root.Add(new XElement(root.Name.Namespace + "TextCode", new XAttribute("X", "2"), new XAttribute("Y", "8"),
                    value.Length > 1 ? new XAttribute("DeltaX", "3") : null, value));
            return new OfdTextElement { ObjectId = id, LayerId = original.LayerId, LayerType = original.LayerType, FontResourceId = original.FontResourceId,
                FontName = original.FontName, SourceXml = root.ToString(SaveOptions.DisableFormatting), Text = string.Concat(texts) };
        }
        public void Dispose() => Directory.Delete(directory, true);
    }
}
