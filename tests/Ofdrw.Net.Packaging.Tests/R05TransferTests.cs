using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class R05TransferTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MergeAndMixCopyClosedGraphForEachSourceOccurrence(bool persisted)
    {
        var source = TypedClipCompositeTests.Create();
        source.Pages[0].Annotations = AnnotationSource().Pages[0].Annotations;
        if (persisted)
        {
            using var bytes = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(source, bytes);
            bytes.Position = 0;
            source = await new OfdReader().ReadAsync(bytes);
        }
        var originalResource = Assert.Single(source.CompositeResources!.Items);
        var originalClip = Assert.Single(source.Pages[0].Elements.OfType<OfdPathElement>()).Clips!.Items[0];
        var merged = OfdDocumentMerger.Merge([source, source]);
        var mixed = OfdDocumentMixer.Mix([new OfdMixSource(source, 0), new OfdMixSource(source, 0)]);
        Assert.Equal(2, merged.Pages.Count);
        Assert.Single(mixed.Pages);
        foreach (var result in new[] { merged, mixed })
        {
            Assert.Equal(2, result.CompositeResources!.Items.Count);
            Assert.NotSame(result.CompositeResources.Items[0], result.CompositeResources.Items[1]);
            Assert.Equal(4, result.Pages.SelectMany(page => page.Elements).OfType<OfdCompositeElement>().Count());
            Assert.All(result.Pages.SelectMany(page => page.Elements).OfType<OfdCompositeElement>(), placement =>
                Assert.Contains(placement.Resource, result.CompositeResources.Items));
            Assert.Equal(2, result.Pages.SelectMany(page => page.Elements).OfType<OfdPathElement>()
                .Count(path => path.Clips is not null));
            Assert.All(result.Pages.SelectMany(page => page.Elements).OfType<OfdPathElement>()
                .Where(path => path.Clips is not null), path =>
                Assert.Equal(OfdFillRule.EvenOdd, path.Clips!.Items[0].Areas[0].Path.FillRule));
            Assert.All(result.CompositeResources.Items, resource => Assert.Null(resource.SourceXml));
            using var output = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(result, output);
            output.Position = 0;
            var readback = await new OfdReader().ReadAsync(output);
            Assert.Equal(2, readback.CompositeResources!.Items.Count);
            Assert.Equal(2, readback.CompositeResources.Items.Select(resource => resource.Id).Distinct().Count());
            Assert.Equal(ReferenceEquals(result, merged) ? 2 : 0,
                readback.Pages.Sum(page => page.Annotations?.Items.Count ?? 0));
            if (ReferenceEquals(result, merged))
                Assert.Equal(2, readback.Pages.SelectMany(page => page.Annotations!.Items)
                    .Select(annotation => annotation.Id).Distinct().Count());
            Assert.Equal(4, readback.Pages.SelectMany(page => page.Elements).OfType<OfdCompositeElement>().Count());
        }
        Assert.Equal(2, merged.Pages.Sum(page => page.Annotations!.Items.Count));
        Assert.Null(mixed.Pages[0].Annotations);
        Assert.Equal(4, mixed.Pages[0].Elements.OfType<OfdPathElement>().Count());
        Assert.IsType<OfdPathElement>(mixed.Pages[0].Elements[3]);
        Assert.IsType<OfdCompositeElement>(mixed.Pages[0].Elements[4]);
        Assert.Same(originalResource, source.CompositeResources.Items[0]);
        Assert.Same(originalClip, Assert.Single(source.Pages[0].Elements.OfType<OfdPathElement>()).Clips!.Items[0]);
        Assert.NotSame(originalResource, merged.CompositeResources!.Items[0]);
        var firstClip = Assert.Single(merged.Pages[0].Elements.OfType<OfdPathElement>()).Clips!.Items[0];
        firstClip.Areas[0].Transform![4] = 99;
        Assert.Equal(2, originalClip.Areas[0].Transform![4]);
        Assert.Equal(2, Assert.Single(merged.Pages[1].Elements.OfType<OfdPathElement>()).Clips!.Items[0].Areas[0].Transform![4]);
        merged.Pages[0].Annotations!.Items[0].Creator = "edited";
        Assert.Equal("R05", source.Pages[0].Annotations!.Items[0].Creator);
    }

    [Fact]
    public void UnknownCompositeAndAnnotationActionsFailClosed()
    {
        var source = TypedClipCompositeTests.Create();
        var resource = source.CompositeResources!.Items[0];
        resource.SourceXml = "<CompositeGraphicUnit ID='7' Width='40' Height='20'><Vendor/></CompositeGraphicUnit>";
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source]));
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge([source],
            new OfdDocumentMergeOptions { SkipUnsupportedRawElements = true }));
        source = AnnotationSource();
        source.Pages[0].Annotations!.Items[0].AttachmentId = "42";
        Assert.ThrowsAny<Exception>(() => OfdDocumentMerger.Merge([source]));
    }

    [Fact]
    public void UnusedOpaqueDefinitionIsOutsideSelectedResourceClosure()
    {
        var source = TypedClipCompositeTests.Create();
        source.CompositeResources!.Items.Add(new OfdCompositeResource { WidthMillimeters = 1,
            HeightMillimeters = 1, SourceXml = "<CompositeGraphicUnit ID='99' Width='1' Height='1'><Vendor/></CompositeGraphicUnit>" });
        Assert.Single(OfdDocumentMerger.Merge([source]).CompositeResources!.Items);
        Assert.Single(OfdDocumentMixer.Mix([new OfdMixSource(source, 0)]).CompositeResources!.Items);
    }

    [Fact]
    public async Task NestedDefinitionsKeepSharingWithinAnOccurrenceAndSeparateAcrossOccurrences()
    {
        var source = TypedClipCompositeTests.Create();
        var child = source.CompositeResources!.Items[0];
        var clippedChild = Assert.Single(child.Elements.OfType<OfdPathElement>());
        clippedChild.Clips = new OfdClipCollection();
        var childClip = new OfdClip();
        childClip.Areas.Add(new OfdClipArea { Path = new OfdPathElement { Stroke = false, Fill = true,
            AbbreviatedData = "M 0 0 L 15 0 L 15 10 C" } });
        clippedChild.Clips.Items.Add(childClip);
        var parent = new OfdCompositeResource { WidthMillimeters = 40, HeightMillimeters = 20 };
        parent.Elements.Add(new OfdCompositeElement { Resource = child, WidthMillimeters = 40, HeightMillimeters = 20 });
        source.CompositeResources.Items.Add(parent);
        source.Pages[0].Elements.Add(new OfdCompositeElement { Resource = parent,
            XMillimeters = 5, YMillimeters = 5, WidthMillimeters = 40, HeightMillimeters = 20 });
        var merged = OfdDocumentMerger.Merge([source, source]);
        Assert.Equal(4, merged.CompositeResources!.Items.Count);
        for (var index = 0; index < 2; index++)
        {
            var leaf = merged.CompositeResources.Items[index * 2];
            var root = merged.CompositeResources.Items[index * 2 + 1];
            Assert.Single(Assert.Single(leaf.Elements.OfType<OfdPathElement>()).Clips!.Items);
            Assert.Same(leaf, Assert.IsType<OfdCompositeElement>(Assert.Single(root.Elements)).Resource);
            Assert.Same(root, Assert.IsType<OfdCompositeElement>(merged.Pages[index].Elements.Last()).Resource);
        }
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(merged, output);
        output.Position = 0;
        var readback = await new OfdReader().ReadAsync(output);
        Assert.Equal(4, readback.CompositeResources!.Items.Count);
        Assert.Single(Assert.Single(readback.CompositeResources.Items[0].Elements.OfType<OfdPathElement>()).Clips!.Items);
    }

    [Fact]
    public async Task ImportedTextAnnotationRebuildsLocalPaintOnce()
    {
        var source = AnnotationSource();
        var annotation = source.Pages[0].Annotations!.Items[0];
        annotation.Appearance.Elements.Clear();
        const string logicalText = "批注 \\ R05";
        annotation.Appearance.Elements.Add(new OfdTextElement { Text = logicalText, FontSizeMillimeters = 4,
            XMillimeters = 2, YMillimeters = 2, WidthMillimeters = 32, HeightMillimeters = 8 });
        using var input = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(source, input);
        input.Position = 0;
        var imported = await new OfdReader().ReadAsync(input);
        Assert.Contains("\\005C", Assert.IsType<OfdTextElement>(
            imported.Pages[0].Annotations!.Items[0].Appearance.Elements[0]).Text);
        Assert.Equal(logicalText, new OfdTextExtractor().Extract(imported).Trim());
        var merged = OfdDocumentMerger.Merge([imported]);
        var mixed = OfdDocumentMixer.Mix([new OfdMixSource(imported, 0)]);
        Assert.Empty(merged.Pages[0].AnnotationAppearances);
        var transferred = Assert.IsType<OfdTextElement>(merged.Pages[0].Annotations!.Items[0].Appearance.Elements[0]);
        Assert.Null(transferred.SourceXml);
        Assert.Equal(logicalText, transferred.Text);
        Assert.Equal(logicalText, Assert.IsType<OfdTextElement>(Assert.Single(mixed.Pages[0].Elements)).Text);
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(merged, output);
        output.Position = 0;
        var readback = await new OfdReader().ReadAsync(output);
        Assert.Equal(logicalText, new OfdTextExtractor().Extract(readback).Trim());
        using var mixedOutput = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(mixed, mixedOutput);
        mixedOutput.Position = 0;
        Assert.Equal(logicalText, new OfdTextExtractor().Extract(await new OfdReader().ReadAsync(mixedOutput)).Trim());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DeletedImportedAnnotationDoesNotBlockTransferOrReappear(bool deleteAll)
    {
        var source = AnnotationSource();
        var retained = new OfdAnnotation
        {
            Type = OfdAnnotationType.Stamp,
            Creator = "retained",
            LastModifiedDate = new DateTime(2026, 10, 26),
            Appearance = new() { XMillimeters = 60, YMillimeters = 30,
                WidthMillimeters = 30, HeightMillimeters = 20 }
        };
        retained.Appearance.Elements.Add(new OfdPathElement
        {
            AbbreviatedData = "M 0 0 L 20 0 L 20 10 C",
            WidthMillimeters = 20, HeightMillimeters = 10, Fill = true, Stroke = false
        });
        source.Pages[0].Annotations!.Items.Add(retained);
        using var input = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(source, input);
        input.Position = 0;
        var imported = await new OfdReader().ReadAsync(input);
        Assert.Equal(2, imported.Pages[0].AnnotationAppearances.Count);
        imported.Pages[0].Annotations!.Items.RemoveAt(0);
        if (deleteAll) imported.Pages[0].Annotations.Items.Clear();

        var merged = OfdDocumentMerger.Merge([imported]);
        var mixed = OfdDocumentMixer.Mix([new OfdMixSource(imported, 0)]);
        var expected = deleteAll ? 0 : 1;
        Assert.Equal(expected, merged.Pages[0].Annotations!.Items.Count);
        Assert.Empty(merged.Pages[0].AnnotationAppearances);
        if (!deleteAll) Assert.Equal("retained", merged.Pages[0].Annotations.Items[0].Creator);
        Assert.Null(mixed.Pages[0].Annotations);
        Assert.Equal(expected, mixed.Pages[0].Elements.OfType<OfdPathElement>().Count());
        foreach (var result in new[] { merged, mixed })
        {
            using var output = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(result, output);
            output.Position = 0;
            var readback = await new OfdReader().ReadAsync(output);
            Assert.Equal(expected, result == merged
                ? readback.Pages[0].Annotations?.Items.Count ?? 0
                : readback.Pages[0].Elements.OfType<OfdPathElement>().Count());
        }
    }

    private static OfdDocumentPackage AnnotationSource()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100, Annotations = new() };
        var annotation = new OfdAnnotation
        {
            Type = OfdAnnotationType.Stamp,
            Creator = "R05",
            LastModifiedDate = new DateTime(2026, 10, 26),
            Appearance = new() { XMillimeters = 20, YMillimeters = 30, WidthMillimeters = 40, HeightMillimeters = 20 }
        };
        annotation.Appearance.Elements.Add(new OfdPathElement
        {
            AbbreviatedData = "M 0 0 L 30 0 L 30 10 C",
            WidthMillimeters = 30,
            HeightMillimeters = 10,
            Fill = true,
            Stroke = false
        });
        page.Annotations.Items.Add(annotation);
        package.Pages.Add(page);
        return package;
    }
}
