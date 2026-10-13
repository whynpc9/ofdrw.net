using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TypedAnnotationTests
{
    [Fact]
    public async Task FreshLocalAppearanceRoundTripsAndMovesOnceWithIndependentClone()
    {
        var package = Fresh(); var initial = package.Pages[0].Annotations!.Items[0];
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, bytes);
        bytes.Position = 0; var loaded = await new OfdReader().ReadAsync(bytes);
        var annotation = Assert.Single(loaded.Pages[0].Annotations!.Items);
        Assert.Equal(new DateTime(2026, 10, 14), annotation.LastModifiedDate);
        Assert.Equal(2, annotation.Appearance.Elements[0].XMillimeters);
        var local = annotation.Appearance.Elements[0];
        var projection = OfdAnnotationProjection.Resolve(loaded.Pages[0], new(new OfdModelLimits()));
        Assert.Same(local, Assert.Single(projection).Definition);
        Assert.Equal(32, projection[0].RenderElement.XMillimeters);
        Assert.Equal("LOCAL", new OfdTextExtractor().Extract(loaded));
        using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        var clone = owned.GetMapped(annotation); clone.Appearance.XMillimeters += 10;
        Assert.Equal(30, annotation.Appearance.XMillimeters);
        using var output = new MemoryStream(); using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); await prepared.WriteAsync(output);
        output.Position = 0; var saved = await new OfdReader().ReadAsync(output);
        Assert.Equal(40, saved.Pages[0].Annotations!.Items[0].Appearance.XMillimeters);
        Assert.Equal(42, OfdAnnotationProjection.Resolve(saved.Pages[0], new(new OfdModelLimits()))[0].RenderElement.XMillimeters);
        Assert.Equal("LOCAL", new OfdTextExtractor().Extract(saved));
        Assert.Equal(2, saved.Pages[0].Annotations!.Items[0].Appearance.Elements[0].XMillimeters);
    }

    [Fact]
    public async Task EmptyClearsOwnedSlotsAndNullCannotWithdrawDirtyEdits()
    {
        var package = await RoundTrip(Fresh()); var page = package.Pages[0]; var collection = page.Annotations!;
        page.Annotations = null; page.Annotations = collection;
        collection.Items[0].Remark = "changed";
        Assert.Throws<InvalidOperationException>(() => page.Annotations = null);
        collection.Items.Clear(); Assert.Empty(OfdAnnotationProjection.Resolve(page, new(new OfdModelLimits())));
        var saved = await RoundTrip(package); Assert.Null(saved.Pages[0].Annotations); Assert.Empty(saved.Pages[0].AnnotationAppearances.OfType<OfdTextElement>());
    }

    [Fact]
    public async Task FrozenProjectionAndImportedPrimitiveMismatchRejectBeforeWriting()
    {
        var package = await RoundTrip(Fresh()); var page = package.Pages[0];
        Assert.IsType<OfdTextElement>(page.AnnotationAppearances[0]).Text = "forged";
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
        package = await RoundTrip(Fresh()); page = package.Pages[0]; Assert.IsType<OfdTextElement>(page.Annotations!.Items[0].Appearance.Elements[0]).Text = "new";
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
        Assert.Throws<NotSupportedException>(() => new OfdTextExtractor().Extract(package));
    }

    [Fact]
    public async Task FreshSplitRetainsLocalMetadataAndTypedMergeRefuses()
    {
        var source = Fresh(); source.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297 });
        var split = OfdDocumentSplitter.Split(source, [0]);
        Assert.NotSame(source.Pages[0].Annotations, split.Pages[0].Annotations);
        Assert.Equal("author", split.Pages[0].Annotations!.Items[0].Creator);
        Assert.Empty(split.Pages[0].Elements);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
        var saved = await RoundTrip(split); Assert.Equal("LOCAL", new OfdTextExtractor().Extract(saved));
    }

    [Fact]
    public async Task ImportedUnknownSlotsAndMetadataSurviveAppearanceEditAndKnownClear()
    {
        var source = await RoundTrip(Fresh()); var page = source.Pages[0]; var annotation = page.Annotations!.Items[0]; var path = annotation.Source!.Entry;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[path])); var ns = xml.Root!.Name.Namespace; var vendor = XNamespace.Get("urn:vendor");
        xml.Root.SetAttributeValue(XNamespace.Xmlns + "v", vendor.NamespaceName); xml.Root.AddFirst(new XProcessingInstruction("keep", "yes"));
        var annot = xml.Root.Element(ns + "Annot")!; annot.SetAttributeValue(vendor + "flag", "retained");
        annot.Add(new XElement(ns + "Parameters", new XElement(ns + "Parameter", new XAttribute("Name", "duplicate"), "one"), new XElement(ns + "Parameter", new XAttribute("Name", "duplicate"), "two")));
        xml.Root.Add(new XElement(vendor + "Annot", new XAttribute("ID", "900"), new XElement(vendor + "Appearance", "opaque")));
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)); source = await ReadEntries(source);
        Assert.NotNull(source.Pages[0].Annotations); source.Pages[0].Annotations!.Items[0].Appearance.XMillimeters = 50;
        var edited = await RoundTrip(source); var preserved = XDocument.Parse(Encoding.UTF8.GetString(edited.PreservedEntries[path]));
        Assert.Equal("retained", preserved.Root!.Element(ns + "Annot")!.Attribute(vendor + "flag")!.Value);
        Assert.Equal(2, preserved.Descendants(ns + "Parameter").Count()); Assert.Single(preserved.Root.Elements(vendor + "Annot")); Assert.Single(preserved.Root.Nodes().OfType<XProcessingInstruction>());
        edited.Pages[0].Annotations!.Items.Clear(); var cleared = await RoundTrip(edited);
        preserved = XDocument.Parse(Encoding.UTF8.GetString(cleared.PreservedEntries[path])); Assert.Empty(preserved.Root!.Elements(ns + "Annot")); Assert.Single(preserved.Root.Elements(vendor + "Annot"));
    }

    [Theory]
    [InlineData("attribute")] [InlineData("pi")] [InlineData("comment")]
    public async Task PageRemovalCannotDestroyUnknownFileOwner(string kind)
    {
        var source = Fresh(); source.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297 }); source = await RoundTrip(source);
        var path = source.Pages[0].Annotations!.Items[0].Source!.Entry;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[path]));
        if (kind == "attribute") xml.Root!.SetAttributeValue(XName.Get("keep", "urn:vendor"), "yes");
        else if (kind == "pi") xml.Root!.AddFirst(new XProcessingInstruction("keep", "yes"));
        else xml.Root!.AddFirst(new XComment("keep"));
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)); source = await ReadEntries(source);
        Assert.Throws<NotSupportedException>(() => OfdDocumentEditor.RemovePages(source, [0])); Assert.Equal(2, source.Pages.Count);
        Assert.Throws<NotSupportedException>(() => OfdDocumentSplitter.Split(source, [1]));
        source.Pages.RemoveAt(0); using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task RevealedHiddenAnnotationPaintOrderMatchesReadbackAndInsertionRefuses()
    {
        var source = Fresh(); var first = source.Pages[0].Annotations!.Items[0]; first.Visible = false;
        var second = new OfdAnnotation { Creator = "author", LastModifiedDate = new DateTime(2026, 10, 14), Type = OfdAnnotationType.Stamp,
            Appearance = new() { WidthMillimeters = 80, HeightMillimeters = 20 } };
        second.Appearance.Elements.Add(new OfdTextElement { Text = "SECOND", WidthMillimeters = 30, HeightMillimeters = 10 }); source.Pages[0].Annotations!.Items.Add(second);
        source = await RoundTrip(source); source.Pages[0].Annotations!.Items[0].Visible = true;
        var before = OfdAnnotationProjection.Resolve(source.Pages[0], new(new OfdModelLimits())).Select(o => ((OfdTextElement)o.RenderElement).Text).ToArray();
        Assert.Equal(new[] { "LOCAL", "SECOND" }, before);
        var saved = await RoundTrip(source); Assert.Equal(before, OfdAnnotationProjection.Resolve(saved.Pages[0], new(new OfdModelLimits())).Select(o => ((OfdTextElement)o.RenderElement).Text));
        saved.Pages[0].Annotations!.Items.Insert(0, new OfdAnnotation { Creator = "author", LastModifiedDate = new DateTime(2026, 10, 14), Appearance = new() { WidthMillimeters = 1, HeightMillimeters = 1 } });
        Assert.Throws<NotSupportedException>(() => new OfdTextExtractor().Extract(saved));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task WithdrawnAndReplacementCollectionsHaveClosedOwnedCloneAndSplitMaps(bool replacement)
    {
        var source = await RoundTrip(Fresh());
        source.Pages[0].Annotations = replacement ? new OfdAnnotationCollection() : null;
        using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var candidate = owned.Package; owned.Dispose(); context.Dispose();
        var read = await RoundTrip(candidate); Assert.Equal(replacement ? "" : "LOCAL", new OfdTextExtractor().Extract(read));
        var split = OfdDocumentSplitter.Split(source, [0]); read = await RoundTrip(split); Assert.Equal(replacement ? "" : "LOCAL", new OfdTextExtractor().Extract(read));
    }

    [Fact]
    public async Task WithdrawnCollectionCannotBeChangedThroughOldReference()
    {
        var source = await RoundTrip(Fresh()); var old = source.Pages[0].Annotations!; source.Pages[0].Annotations = null; old.Items[0].Creator = "changed";
        Assert.Throws<InvalidDataException>(() => new OfdTextExtractor().Extract(source));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task StrictCallerObjectQuotaAppliesToTypedAndHiddenBeforeConstruction(bool hidden)
    {
        var source = Fresh(); source.Pages[0].Annotations!.Items[0].Visible = !hidden;
        source.Pages[0].Annotations!.Items[0].Appearance.Elements.Add(new OfdPathElement { WidthMillimeters = 10, HeightMillimeters = 10, AbbreviatedData = "M 0 0 L 10 10" });
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, bytes); bytes.Position = 0;
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(bytes, new OfdPackageLoadOptions { MaxAnnotationObjectCount = 1 }));
        bytes.Position = 0; using var context = new OfdControlledEditContext(new());
        var quota = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadControlledAsync(bytes, new() { MaxAnnotationObjectCount = 1 }, context));
        Assert.Contains("object budget", quota.Message); Assert.Equal(0, context.Material.ReadAnnotationFileAdmissions);
        Assert.Throws<InvalidOperationException>(context.AssertActive);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task IndexOrDeclarationChangesInvalidateTypedAuthorityBeforeOutput(bool declaration)
    {
        var source = await RoundTrip(Fresh());
        if (declaration)
        {
            var position = source.PreservedDocumentElements.FindIndex(xml => XElement.Parse(xml).Name.LocalName == "Annotations");
            source.PreservedDocumentElements[position] = source.PreservedDocumentElements[position].Replace("Annotations_0.xml", "other.xml");
        }
        else
        {
            var path = source.AnnotationIndexSource!.IndexDigests.Keys.First(p => p.Contains("Annotations_"));
            source.PreservedEntries[path] = Encoding.UTF8.GetBytes("<Annotations xmlns='http://www.ofdspec.org'/>");
        }
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
        Assert.Throws<InvalidDataException>(() => new OfdTextExtractor().Extract(source));
    }

    [Fact]
    public async Task LocalOccurrenceUsesDefinitionAndParentMatrixOnceWhileRemainingUneditable()
    {
        var source = await RoundTrip(Fresh()); var page = source.Pages[0]; var annotation = page.Annotations!.Items[0]; annotation.Appearance.Transform = [0, 1, -1, 0, 5, 6];
        var traversal = OfdModelTraversal.ReadGraphics(source, [page], new() { IncludeAnnotations = true, RequireSupportedSource = true });
        var occurrence = Assert.Single(traversal.Occurrences); Assert.Same(annotation.Appearance.Elements[0], occurrence.Element);
        Assert.Equal(new double[] { 0, 1, -1, 0, 35, 46 }, occurrence.ParentToPage); Assert.False(occurrence.IsEditableSource);
        Assert.Equal(OfdSourceAuthority.LocalAnnotationDefinition, occurrence.SourceAuthority); Assert.Equal(annotation.Source!.Entry, occurrence.SourceEntryPath);
    }

    [Theory]
    [InlineData("http://www.ofdspec.org")] [InlineData("http://www.ofdspec.org/2016")]
    public async Task AnnotationOnlyImageAndEmbeddedFontResourcesSurviveReadback(string documentNamespace)
    {
        var source = Fresh(); source.Options.Namespace = documentNamespace;
        source.Fonts.Add(new OfdFontResource { Id = "100", FontName = "test", Data = [1, 2, 3], FileName = "font.ttf" });
        ((OfdTextElement)source.Pages[0].Annotations!.Items[0].Appearance.Elements[0]).FontResourceId = "100";
        var image = new OfdImageElement { Data = [1, 2, 3, 4], WidthMillimeters = 15, HeightMillimeters = 15, FileName = "sample.png", MediaType = "image/png" };
        image.ClippingXml = $"<Clips xmlns='{documentNamespace}'><Clip><Area><Path><AbbreviatedData>M 0 0 L 15 0 L 15 15 C</AbbreviatedData></Path></Area></Clip></Clips>";
        source.Pages[0].Annotations!.Items[0].Appearance.Elements.Add(image);
        Assert.Equal(2, OfdAnnotationProjection.Resolve(source.Pages[0], new(new OfdModelLimits()), documentNamespace: documentNamespace).Count);
        var read = await RoundTrip(source); Assert.Equal(new byte[] { 1, 2, 3 }, Assert.Single(read.Fonts).Data);
        Assert.Equal(image.Data, Assert.IsType<OfdImageElement>(read.Pages[0].Annotations!.Items[0].Appearance.Elements[1]).Data);
    }

    [Fact]
    public async Task CandidateGuardCoversUnqueriedMetadataAndPreparedWriteRejectsMutationBeforeOutput()
    {
        var source = await RoundTrip(Fresh()); using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); owned.Package.Pages[0].Annotations!.Items[0].LastModifiedDate = new DateTime(2026, 10, 15);
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task RepeatedAnnotationOwnershipCannotIssueOwnedCertificate()
    {
        var source = Fresh(); var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297, Annotations = new() };
        page.Annotations.Items.Add(source.Pages[0].Annotations!.Items[0]); source.Pages.Add(page);
        using var context = new OfdControlledEditContext(new());
        Assert.Throws<InvalidDataException>(() => context.CaptureSource(source));
    }

    [Fact]
    public async Task MultipleIndexAndPageAnnotFilesKeepPathsOrderAndDuplicateRowsPaintOnce()
    {
        var source = await RoundTrip(Fresh()); var first = source.Pages[0].Annotations!.Items[0]; var firstPath = first.Source!.Entry;
        var document = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[source.DocumentEntryPath!])); var ns = document.Root!.Name.Namespace;
        var indexPath = source.AnnotationIndexSource!.IndexDigests.Keys.First(path => path.Contains("Annotations_"));
        var index = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[indexPath])); index.Root!.Add(new XElement(index.Root.Elements().Single()));
        source.PreservedEntries[indexPath] = Encoding.UTF8.GetBytes(index.ToString());
        var second = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[firstPath])); second.Root!.Element(ns + "Annot")!.SetAttributeValue("ID", "200");
        second.Descendants(ns + "TextObject").Single().SetAttributeValue("ID", "201"); second.Descendants(ns + "TextCode").Single().Value = "SECOND";
        var secondPath = "Doc_0/Annots/second.xml"; source.PreservedEntries[secondPath] = Encoding.UTF8.GetBytes(second.ToString());
        var secondIndexPath = "Doc_0/Annots/second-index.xml";
        source.PreservedEntries[secondIndexPath] = Encoding.UTF8.GetBytes(new XElement(ns + "Annotations", new XElement(ns + "Page", new XAttribute("PageID", source.Pages[0].Id!), new XElement(ns + "FileLoc", "second.xml"))).ToString());
        document.Root.Add(new XElement(ns + "Annotations", "Annots/second-index.xml")); source.PreservedEntries[source.DocumentEntryPath!] = Encoding.UTF8.GetBytes(document.ToString());
        source = await ReadEntries(source); Assert.Equal(2, source.Pages[0].Annotations!.Items.Count);
        Assert.Equal(new[] { "LOCAL", "SECOND" }, OfdAnnotationProjection.Resolve(source.Pages[0], new(new OfdModelLimits())).Select(o => ((OfdTextElement)o.RenderElement).Text));
        source.Pages[0].Annotations!.Items[1].Appearance.XMillimeters = 75;
        var saved = await RoundTrip(source); Assert.Equal(75, saved.Pages[0].Annotations!.Items[1].Appearance.XMillimeters);
        Assert.Equal(firstPath, saved.Pages[0].Annotations!.Items[0].Source!.Entry); Assert.Equal(secondPath, saved.Pages[0].Annotations!.Items[1].Source!.Entry);
        Assert.Equal(source.PreservedEntries[indexPath], saved.PreservedEntries[indexPath]); Assert.Equal(source.PreservedEntries[secondIndexPath], saved.PreservedEntries[secondIndexPath]);
    }

    [Fact]
    public async Task SharedPageAnnotStaysUnclaimedAndCannotBeClaimedOrDestructivelySplit()
    {
        var source = Fresh(); source.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297 }); source = await RoundTrip(source);
        var path = source.AnnotationIndexSource!.IndexDigests.Keys.First(p => p.Contains("Annotations_")); var index = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[path]));
        var row = new XElement(index.Root!.Elements().Single()); row.SetAttributeValue("PageID", source.Pages[1].Id); index.Root.Add(row);
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(index.ToString()); source = await ReadEntries(source);
        Assert.All(source.Pages, page => { Assert.Null(page.Annotations); Assert.Single(page.AnnotationAppearances.OfType<OfdTextElement>()); });
        Assert.Throws<NotSupportedException>(() => OfdDocumentSplitter.Split(source, [0]));
        source.Pages[0].Annotations = new(); using var output = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task InvalidCivilDateAndPreferredIdCollisionFailBeforeDestination()
    {
        var source = Fresh(); source.Pages[0].Annotations!.Items[0].LastModifiedDate = DateTime.SpecifyKind(new DateTime(2026, 10, 14), DateTimeKind.Utc);
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
        source = Fresh(); source.Pages[0].Annotations!.Items[0].Id = "100"; source.Fonts.Add(new OfdFontResource { Id = "100", FontName = "collision" });
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(source, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task SmallQuotaBoundsHiddenFieldsAndValidationWorkWithoutReset()
    {
        var source = Fresh(); source.Pages[0].Annotations!.Items[0].Visible = false; source.Pages[0].Annotations!.Items[0].Remark = new string('x', 2000);
        using var small = new OfdControlledEditContext(new() { MaxMaterializedBytes = 1024 });
        Assert.Throws<InvalidDataException>(() => small.CaptureSource(source)); Assert.Throws<InvalidOperationException>(() => small.CaptureSource(Fresh()));
        using var context = new OfdControlledEditContext(new()); var guard = context.CaptureSource(source); var work = context.Material.ValidationWorkBytes;
        guard.AssertUnchanged(source, context); Assert.True(context.Material.ValidationWorkBytes > work);
        source.Pages[0].Annotations!.Items[0].Remark += "changed"; Assert.Throws<InvalidDataException>(() => guard.AssertUnchanged(source, context));
    }

    [Fact]
    public async Task RetainedTextSlotSeamDoesNotGrantAnnotationEditingAuthority()
    {
        var source = await RoundTrip(Fresh()); using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var local = owned.GetMapped((OfdTextElement)source.Pages[0].Annotations!.Items[0].Appearance.Elements[0]);
        Assert.Throws<NotSupportedException>(() => owned.ApplyTextSlots(local, [new(0, "CHANGED", 0, 4)])); Assert.Equal("LOCAL", local.Text);
    }

    [Theory]
    [InlineData("http://www.ofdspec.org", "http://www.ofdspec.org/2016")]
    [InlineData("http://www.ofdspec.org/2016", "http://www.ofdspec.org")]
    public async Task CrossNamespaceAnnotRemainsOpaqueAndOrdinarySavePreservesIt(string rootNamespace, string foreignNamespace)
    {
        var source = Fresh(); source.Options.Namespace = rootNamespace; source = await RoundTrip(source);
        var path = source.Pages[0].Annotations!.Items[0].Source!.Entry;
        var xml = XDocument.Parse(Encoding.UTF8.GetString(source.PreservedEntries[path]));
        var annotation = xml.Root!.Elements().Single();
        foreach (var node in annotation.DescendantsAndSelf()) node.Name = XNamespace.Get(foreignNamespace) + node.Name.LocalName;
        source.PreservedEntries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        source = await ReadEntries(source);
        Assert.Null(source.Pages[0].Annotations); Assert.Empty(source.Pages[0].AnnotationAppearances.OfType<OfdTextElement>());
        var opaque = source.PreservedEntries[path]; var saved = await RoundTrip(source);
        Assert.Null(saved.Pages[0].Annotations); Assert.Equal(opaque, saved.PreservedEntries[path]);
    }

    [Theory]
    [InlineData("http://www.ofdspec.org")] [InlineData("http://www.ofdspec.org/2016")]
    public async Task ControlledTypedReadCloneAppearanceEditWriteAndReadbackShareOneContext(string documentNamespace)
    {
        var source = Fresh(); source.Options.Namespace = documentNamespace;
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, input); input.Position = 0;
        using var context = new OfdControlledEditContext(new()); var reader = new OfdReader();
        var loaded = await reader.ReadControlledAsync(input, new(), context);
        var original = Assert.Single(loaded.Pages[0].Annotations!.Items); var copy = context.Material.CloneCopyBytes;
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        var target = owned.GetMapped(original); target.Appearance.XMillimeters += 10; target.Remark = "controlled metadata";
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        output.Position = 0; var readback = await reader.ReadControlledAsync(output, new(), context);
        var annotation = Assert.Single(readback.Pages[0].Annotations!.Items);
        Assert.Equal(40, annotation.Appearance.XMillimeters); Assert.Equal("controlled metadata", annotation.Remark);
        Assert.Equal("LOCAL", new OfdTextExtractor().Extract(readback)); Assert.Equal(2, annotation.Appearance.Elements[0].XMillimeters);
        Assert.True(context.Material.CloneCopyBytes > copy); Assert.Equal(2, context.Material.ReadMetadataReservations);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SameTickCivilDateKindMutationIsDetectedBeforePreparedOutput(bool mutateSource)
    {
        var source = await RoundTrip(Fresh()); using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        var annotation = (mutateSource ? source : owned.Package).Pages[0].Annotations!.Items[0];
        annotation.LastModifiedDate = DateTime.SpecifyKind(annotation.LastModifiedDate!.Value, DateTimeKind.Utc);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task ResetNullCannotWithdrawSameTickDateKindMutation()
    {
        var source = await RoundTrip(Fresh()); var page = source.Pages[0]; var annotation = page.Annotations!.Items[0];
        annotation.LastModifiedDate = DateTime.SpecifyKind(annotation.LastModifiedDate!.Value, DateTimeKind.Utc);
        Assert.Throws<InvalidOperationException>(() => page.Annotations = null);
    }

    private static OfdDocumentPackage Fresh()
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297, Annotations = new() };
        var annotation = new OfdAnnotation { Type = OfdAnnotationType.Stamp, Creator = "author", LastModifiedDate = new DateTime(2026, 10, 14), Remark = "remark",
            Appearance = new() { XMillimeters = 30, YMillimeters = 40, WidthMillimeters = 80, HeightMillimeters = 20 } };
        annotation.Appearance.Elements.Add(new OfdTextElement { Text = "LOCAL", XMillimeters = 2, YMillimeters = 3, WidthMillimeters = 50, HeightMillimeters = 10 });
        page.Annotations.Items.Add(annotation); source.Pages.Add(page); return source;
    }
    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage source)
    { using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, bytes); bytes.Position = 0; return await new OfdReader().ReadAsync(bytes); }
    private static async Task<OfdDocumentPackage> ReadEntries(OfdDocumentPackage source)
    {
        using var bytes = new MemoryStream(); using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, true)) foreach (var entry in source.PreservedEntries)
        { using var stream = zip.CreateEntry(entry.Key).Open(); stream.Write(entry.Value); }
        bytes.Position = 0; return await new OfdReader().ReadAsync(bytes);
    }
}
