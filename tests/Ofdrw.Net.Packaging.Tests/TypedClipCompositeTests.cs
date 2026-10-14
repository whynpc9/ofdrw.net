using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TypedClipCompositeTests
{
    internal static OfdDocumentPackage Create()
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var resource = new OfdCompositeResource { WidthMillimeters = 40, HeightMillimeters = 20 };
        resource.Elements.Add(new OfdTextElement { Text = "shared text", WidthMillimeters = 35, HeightMillimeters = 10 });
        resource.Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 30 0 L 30 15 L 0 15 C", WidthMillimeters = 30, HeightMillimeters = 15, Fill = true });
        package.CompositeResources.Items.Add(resource);
        var page = new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = 10, YMillimeters = 20, WidthMillimeters = 40, HeightMillimeters = 20 });
        page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = 55, YMillimeters = 50, WidthMillimeters = 40, HeightMillimeters = 20, Transform = new double[] { 1, 0.1, 0, 1, 0, 0 } });
        var path = new OfdPathElement { AbbreviatedData = "M 0 0 L 40 0 L 40 20 L 0 20 C", WidthMillimeters = 40, HeightMillimeters = 20, Clips = new() };
        var clip = new OfdClip(); clip.Areas.Add(new OfdClipArea { Transform = new double[] { 1, 0, 0, 1, 2, 3 },
            Path = new OfdPathElement { Fill = true, Stroke = false, FillRule = OfdFillRule.EvenOdd, AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C" } });
        path.Clips.Items.Add(clip); page.Elements.Add(path); package.Pages.Add(page); return package;
    }

    [Fact]
    public async Task PublicConstructionAndEditPreserveSharedDefinitionsAndClips()
    {
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), bytes); bytes.Position = 0;
        var loaded = await new OfdReader().ReadAsync(bytes);
        var resource = Assert.Single(loaded.CompositeResources!.Items);
        Assert.All(loaded.Pages[0].Elements.OfType<OfdCompositeElement>(), placement => Assert.Same(resource, placement.Resource));
        var path = Assert.Single(loaded.Pages[0].Elements.OfType<OfdPathElement>());
        Assert.Equal<OfdFillRule?>(OfdFillRule.EvenOdd, path.Clips!.Items[0].Areas[0].Path.FillRule);
        path.Clips.Items[0].Areas[0].Transform![4] = 7;
        ((OfdCompositeElement)loaded.Pages[0].Elements[0]).Transform = new double[] { 1, 0, 0, 1, 3, 4 };
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, output); output.Position = 0;
        var readback = await new OfdReader().ReadAsync(output);
        Assert.Equal(7, Assert.Single(readback.Pages[0].Elements.OfType<OfdPathElement>()).Clips!.Items[0].Areas[0].Transform![4]);
        Assert.Single(readback.CompositeResources!.Items);
        Assert.Equal(3, ((OfdCompositeElement)readback.Pages[0].Elements[0]).Transform![4]);
    }

    [Fact]
    public async Task CyclesAndMixedClipAuthorityFailBeforeOutput()
    {
        var package = Create(); var resource = package.CompositeResources!.Items[0];
        resource.Elements.Add(new OfdCompositeElement { Resource = resource, WidthMillimeters = 1, HeightMillimeters = 1 });
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
        package = Create(); package.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml = "<Clips/>";
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task OwnedClonePreservesDagAndIndependentClipMatrices()
    {
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), bytes); bytes.Position = 0;
        var source = await new OfdReader().ReadAsync(bytes);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var original = source.CompositeResources!.Items[0]; var copied = candidate.Package.CompositeResources!.Items[0];
        Assert.NotSame(original, copied);
        Assert.All(candidate.Package.Pages[0].Elements.OfType<OfdCompositeElement>(), placement => Assert.Same(copied, placement.Resource));
        var path = candidate.Package.Pages[0].Elements.OfType<OfdPathElement>().Single();
        path.Clips!.Items[0].Areas[0].Transform![4] = 11;
        Assert.Equal(2, source.Pages[0].Elements.OfType<OfdPathElement>().Single().Clips!.Items[0].Areas[0].Transform![4]);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(candidate); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        output.Position = 0; Assert.Equal(11, (await new OfdReader().ReadAsync(output)).Pages[0].Elements.OfType<OfdPathElement>().Single().Clips!.Items[0].Areas[0].Transform![4]);
    }

    [Fact]
    public void StrictTraversalRefusesAnEmptyCompositeAndFreshTypedClip()
    {
        var package = Create(); package.CompositeResources!.Items[0].Elements.Clear();
        var result = OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Kind == OfdTraversalDiagnosticKind.UnsupportedSource);
        Assert.Throws<System.NotSupportedException>(() => OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions { RequireSupportedSource = true }));
        Assert.Throws<System.NotSupportedException>(() => new Ofdrw.Net.Reader.Keywords.OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(package, "missing"));
        package.Pages[0].Elements.RemoveAll(e => e is OfdCompositeElement);
        Assert.Throws<System.NotSupportedException>(() => OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions { RequireSupportedSource = true }));
    }

    [Fact]
    public void PublicStrictQueryRejectsFreshTypedTextClip()
    {
        var package = KeywordLocatorTests.Model("A");
        var text = package.Pages[0].Elements.OfType<OfdTextElement>().Single(); text.Clips = new();
        Assert.Null(text.SourceXml);
        Assert.Throws<System.NotSupportedException>(() => new Ofdrw.Net.Reader.Keywords.OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(package, "missing"));
    }

    [Fact]
    public async Task WithdrawingDirtyTypedOwnersAndLateMutationReject()
    {
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), input); input.Position = 0;
        var package = await new OfdReader().ReadAsync(input);
        package.CompositeResources!.Items[0].WidthMillimeters += 1;
        Assert.Throws<System.InvalidOperationException>(() => package.CompositeResources = null);
        var path = package.Pages[0].Elements.OfType<OfdPathElement>().Single();
        path.Clips!.Items[0].Areas[0].Transform![0] = 2;
        Assert.Throws<System.InvalidOperationException>(() => path.Clips = null);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(package, context.CaptureSource(package), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(candidate);
        package.CompositeResources.Items[0].Elements.OfType<OfdPathElement>().Single().FillRule = OfdFillRule.EvenOdd;
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task SplitKeepsSharedClosureAndLastDeletedUsePrunesDefinition()
    {
        var package = Create(); var page = new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { Text = "retained", WidthMillimeters = 40, HeightMillimeters = 10 }); package.Pages.Add(page);
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        package = await new OfdReader().ReadAsync(input);
        var first = Ofdrw.Net.Layout.Editing.OfdDocumentSplitter.Split(package, new[] { 0 });
        Assert.Single(first.CompositeResources!.Items);
        Assert.Same(first.CompositeResources.Items[0], ((OfdCompositeElement)first.Pages[0].Elements[0]).Resource);
        var second = Ofdrw.Net.Layout.Editing.OfdDocumentSplitter.Split(package, new[] { 1 });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(second, output); output.Position = 0;
        Assert.Null((await new OfdReader().ReadAsync(output)).CompositeResources);
    }

    [Fact]
    public async Task UnknownCompositeAndClipXmlSurvivesRoundtrip()
    {
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), input);
        var entries = new System.Collections.Generic.Dictionary<string, byte[]>();
        input.Position = 0;
        using (var zip = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, true))
            foreach (var entry in zip.Entries) { using var data = entry.Open(); using var bytes = new MemoryStream(); data.CopyTo(bytes); entries.Add(entry.FullName, bytes.ToArray()); }
        foreach (var path in entries.Keys.ToArray())
        {
            if (!path.EndsWith(".xml")) continue;
            var xml = System.Xml.Linq.XDocument.Load(new MemoryStream(entries[path]));
            foreach (var node in xml.Descendants().Where(n => n.Name.LocalName == "CompositeGraphicUnit" || n.Name.LocalName == "Area" && n.Parent?.Name.LocalName == "Clip"))
                node.Add(new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("Keep", "urn:vendor"), new System.Xml.Linq.XAttribute("ID", "opaque"), new System.Xml.Linq.XCData("preserved")));
            entries[path] = System.Text.Encoding.UTF8.GetBytes(xml.ToString());
        }
        using var modified = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(modified, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var target = zip.CreateEntry(entry.Key).Open(); target.Write(entry.Value); }
        modified.Position = 0; var loaded = await new OfdReader().ReadAsync(modified);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, saved); saved.Position = 0;
        var readback = await new OfdReader().ReadAsync(saved);
        Assert.Contains("preserved", readback.CompositeResources!.Items[0].SourceXml);
        Assert.Contains("preserved", readback.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml);
    }

    [Fact]
    public void HiddenUnusedDefinitionsStillConsumeDepthAndCancellationBudget()
    {
        var package = Create(); var child = package.CompositeResources!.Items[0];
        for (var i = 0; i < 6; i++)
        {
            var parent = new OfdCompositeResource { WidthMillimeters = 1, HeightMillimeters = 1 };
            parent.Elements.Add(new OfdCompositeElement { Resource = child, WidthMillimeters = 1, HeightMillimeters = 1 });
            package.CompositeResources.Items.Add(parent); child = parent;
        }
        Assert.Throws<InvalidDataException>(() => OfdCompositeGraph.Validate(package, new OfdModelMaterialization(new OfdModelLimits { MaxDepth = 3 }), default));
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<System.OperationCanceledException>(() => OfdCompositeGraph.Validate(package, new OfdModelMaterialization(new OfdModelLimits()), cancellation.Token));
    }

    [Fact]
    public void MixedCompositeAndLegacySharedBodyStillReturnsNoneditableOccurrences()
    {
        var package = Create(); var text = new OfdTextElement { Text = "shared" };
        package.Pages[0].Elements.Add(text); package.Pages[0].Elements.Add(text);
        var result = OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions());
        Assert.Equal(2, result.Occurrences.Count(o => ReferenceEquals(o.Element, text)));
        Assert.All(result.Occurrences.Where(o => ReferenceEquals(o.Element, text)), o => Assert.False(o.IsEditableSource));
    }

    [Theory]
    [InlineData("clear")]
    [InlineData("replace")]
    [InlineData("area")]
    public async Task OpaqueClipSlotsCannotBeDeleted(string operation)
    {
        var package = await ReadModified(Create(), xml =>
        {
            foreach (var area in xml.Descendants().Where(n => n.Name.LocalName == "Area" && n.Parent?.Name.LocalName == "Clip"))
                area.Add(new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("Keep", "urn:vendor"), "opaque"));
        });
        var path = package.Pages[0].Elements.OfType<OfdPathElement>().Single();
        if (operation == "clear") path.Clips!.Items.Clear();
        else if (operation == "replace") path.Clips = new();
        else
        {
            path.Clips!.Items[0].Areas.RemoveAt(0);
            path.Clips.Items[0].Areas.Add(new OfdClipArea { Path = new OfdPathElement { Stroke = false, Fill = true, AbbreviatedData = "M 0 0 L 10 10" } });
        }
        using var output = new MemoryStream(); await Assert.ThrowsAsync<System.NotSupportedException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task KnownClipClearPreservesOpaqueWrapper()
    {
        var package = await ReadModified(Create(), xml =>
        {
            foreach (var clips in xml.Descendants().Where(n => n.Name.LocalName == "Clips")) clips.SetAttributeValue(System.Xml.Linq.XName.Get("flag", "urn:vendor"), "keep");
        });
        package.Pages[0].Elements.OfType<OfdPathElement>().Single().Clips = new();
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); output.Position = 0;
        Assert.Contains("keep", (await new OfdReader().ReadAsync(output)).Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml);
        using var knownInput = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), knownInput); knownInput.Position = 0;
        var known = await new OfdReader().ReadAsync(knownInput); known.Pages[0].Elements.OfType<OfdPathElement>().Single().Clips!.Items.Clear();
        using var cleared = new MemoryStream(); await new OfdPackageWriter().WriteAsync(known, cleared); cleared.Position = 0;
        Assert.Null((await new OfdReader().ReadAsync(cleared)).Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml);
    }

    [Theory]
    [InlineData("attribute")]
    [InlineData("sameNamespace")]
    [InlineData("pi")]
    public async Task OpaqueReferenceGrammarRetainsOtherwiseUnusedComposite(string kind)
    {
        var package = Create(); package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 100 });
        package = await ReadModified(package, xml =>
        {
            if (xml.Root?.Name.LocalName != "Res" || !xml.Descendants().Any(n => n.Name.LocalName == "CompositeGraphicUnit")) return;
            var id = xml.Descendants().First(n => n.Name.LocalName == "CompositeGraphicUnit").Attribute("ID")!.Value;
            if (kind == "attribute") xml.Root.SetAttributeValue(System.Xml.Linq.XName.Get("Refs", "urn:vendor"), id + " 99999");
            if (kind == "sameNamespace") xml.Root.Add(new System.Xml.Linq.XElement(xml.Root.Name.Namespace + "UnknownReferences", id + " 99999"));
            if (kind == "pi") xml.Root.AddFirst(new System.Xml.Linq.XProcessingInstruction("refs", id + " 99999"));
        });
        var retained = Ofdrw.Net.Layout.Editing.OfdDocumentSplitter.Split(package, new[] { 1 });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(retained, output); output.Position = 0;
        Assert.Single((await new OfdReader().ReadAsync(output)).CompositeResources!.Items);
    }

    [Fact]
    public void LongNumericLiteralAndRepeatedAncestorClipsRespectCumulativeBudget()
    {
        var longMatrix = string.Join(" ", Enumerable.Repeat("1", 10000));
        Assert.Throws<InvalidDataException>(() => OfdClipXml.Matrix(longMatrix, new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = 1024 }), default));
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<System.OperationCanceledException>(() => OfdClipXml.Matrix(longMatrix, new OfdModelMaterialization(new OfdModelLimits()), cancellation.Token));
        var package = Create(); var placement = package.Pages[0].Elements.OfType<OfdCompositeElement>().First();
        placement.Clips = new(); var clip = new OfdClip();
        clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement { Stroke = false, Fill = true, AbbreviatedData = string.Concat(Enumerable.Repeat("M 0 0 L 10 0 L 10 10 C ", 500)) } }); placement.Clips.Items.Add(clip);
        var once = new OfdModelMaterialization(new OfdModelLimits()); OfdCompositeProjection.Expand(new[] { placement }, once, default);
        var cap = checked(once.CloneCopyBytes + 1024);
        Assert.Throws<InvalidDataException>(() => OfdCompositeProjection.Expand(Enumerable.Repeat<OfdElement>(placement, 8), new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = cap }), default));
    }

    private static async Task<OfdDocumentPackage> ReadModified(OfdDocumentPackage package, System.Action<System.Xml.Linq.XDocument> edit)
    {
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        using var output = new MemoryStream();
        using (var source = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, true))
        using (var target = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var entry in source.Entries)
            {
                using var data = entry.Open(); using var destination = target.CreateEntry(entry.FullName).Open();
                if (entry.FullName.EndsWith(".xml")) { var xml = System.Xml.Linq.XDocument.Load(data); edit(xml); xml.Save(destination); }
                else data.CopyTo(destination);
            }
        output.Position = 0; return await new OfdReader().ReadAsync(output);
    }
}
