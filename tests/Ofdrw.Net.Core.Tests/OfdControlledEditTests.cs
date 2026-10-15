using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdControlledEditTests
{
    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)] [InlineData(5)] [InlineData(6)] [InlineData(7)]
    public void GuardCoversUnqueriedScalarsArraysDatesAndOrder(int mutation)
    {
        var source = Document(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source);
        switch (mutation)
        {
            case 0: source.Options.EnableDeflateCompression = false; break;
            case 1: source.Options.Metadata.CreationDate = DateTimeOffset.UtcNow; break;
            case 2: source.Attachments[0].Id = "changed"; break;
            case 3: source.Attachments[0].Data[0]++; break;
            case 4: source.PreservedEntries["opaque.bin"][0]++; break;
            case 5: Assert.IsType<OfdImageElement>(source.Pages[1].Elements[0]).Alpha--; break;
            case 6: Assert.IsType<OfdImageElement>(source.Pages[1].Elements[0]).Transform![0]++; break;
            case 7: source.Pages.Reverse(); break;
        }
        Assert.Throws<InvalidDataException>(() => guard.AssertUnchanged(source, context));
    }

    [Fact]
    public void GuardRehashesSameArrayAndDoesNotAcceptAnotherOperation()
    {
        var source = Document(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source); var before = context.Material.ValidationWorkBytes;
        guard.AssertUnchanged(source, context); Assert.True(context.Material.ValidationWorkBytes > before);
        using var foreign = new OfdControlledEditContext(new OfdModelLimits());
        Assert.Throws<InvalidOperationException>(() => guard.AssertUnchanged(source, foreign));
        Assert.Throws<InvalidOperationException>(() => guard.AssertUnchanged(Document(), context));
        source.Attachments[0].Data = source.Attachments[0].Data.ToArray();
        Assert.Throws<InvalidDataException>(() => guard.AssertUnchanged(source, context));
    }

    [Fact]
    public void WholeClonePreservesOrderIndicesAndAllAttachmentFieldsWithoutAliases()
    {
        var source = Document(); using var context = new OfdControlledEditContext(new OfdModelLimits());
        var guard = context.CaptureSource(source); using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context);
        var copy = owned.Package;
        Assert.Equal(new[] { 8, 3 }, copy.Pages.Select(page => page.Index)); Assert.NotSame(source.Options, copy.Options);
        Assert.NotSame(source.Options.Metadata, copy.Options.Metadata); Assert.Equal(source.Options.Metadata.CreationDate, copy.Options.Metadata.CreationDate);
        var attachment = copy.Attachments[0]; Assert.Equal("att-9", attachment.Id);
        Assert.Equal(source.Attachments[0].CreationDate, attachment.CreationDate); Assert.Equal(source.Attachments[0].ModificationDate, attachment.ModificationDate);
        Assert.Equal(1.5, attachment.SizeKilobytes); Assert.Equal(false, attachment.Visible); Assert.Equal("usage", attachment.Usage);
        Assert.NotSame(source.Attachments[0].Data, attachment.Data); Assert.NotSame(source.PreservedEntries["opaque.bin"], copy.PreservedEntries["opaque.bin"]);
        Assert.Same(copy.Pages[0], owned.GetMapped(source.Pages[0]));
        var image = Assert.IsType<OfdImageElement>(copy.Pages[1].Elements[0]);
        Assert.NotSame(Assert.IsType<OfdImageElement>(source.Pages[1].Elements[0]).Transform, image.Transform);
        attachment.Data[0] = 99; image.Data[0] = 99; image.Transform![0] = 99;
        guard.AssertUnchanged(source, context);
        context.Dispose(); Assert.Equal(99, image.Data[0]);
        Assert.Throws<ObjectDisposedException>(() => owned.GetMapped(source.Pages[0]));
    }

    [Fact]
    public void RepeatedOwningGraphicsAndRunsFailBeforeClone()
    {
        var source = Document(); source.Pages[1].Elements.Add(source.Pages[1].Elements[0]);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        Assert.Throws<InvalidDataException>(() => context.CaptureSource(source));
        source = Document(); var text = Assert.IsType<OfdTextElement>(source.Pages[0].Elements[0]);
        var run = new OfdTextRun { Text = "x" }; text.Runs.Add(run); text.Runs.Add(run);
        using var second = new OfdControlledEditContext(new OfdModelLimits());
        Assert.Throws<InvalidDataException>(() => second.CaptureSource(source));
    }

    [Fact]
    public void WorkQuotaAndCancellationArePrecharged()
    {
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = 100 });
        Assert.Throws<InvalidDataException>(() => material.HashBytes(new byte[101], default));
        Assert.Equal(0, material.ValidationWorkBytes);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => material.HashBytes(new byte[1], cancelled.Token));
        Assert.Equal(0, material.ValidationWorkBytes);
    }

    private static OfdDocumentPackage Document()
    {
        var source = new OfdDocumentPackage(); source.Options.Metadata.CreationDate = DateTimeOffset.Parse("2020-01-02T03:04:05+08:00");
        var first = new OfdPage { Index = 8, WidthMillimeters = 210, HeightMillimeters = 297 }; first.Elements.Add(new OfdTextElement { Text = "中英" });
        var second = new OfdPage { Index = 3, WidthMillimeters = 210, HeightMillimeters = 297 };
        second.Elements.Add(new OfdImageElement { Data = new byte[] { 1, 2 }, Transform = new double[] { 1, 0, 0, 1, 0, 0 } });
        source.Pages.Add(first); source.Pages.Add(second); source.PreservedEntries.Add("opaque.bin", new byte[] { 3 });
        source.Attachments.Add(new OfdAttachment { Id = "att-9", Data = new byte[] { 4 }, CreationDate = source.Options.Metadata.CreationDate,
            ModificationDate = source.Options.Metadata.CreationDate?.AddDays(1), SizeKilobytes = 1.5, Visible = false, Usage = "usage" });
        return source;
    }
}
