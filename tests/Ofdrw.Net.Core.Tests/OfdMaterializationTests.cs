using System.Text;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdMaterializationTests
{
    [Fact]
    public void SeparateLimitsSnapshotConfigurationAndDeduplicateInputIdentity()
    {
        var limits = new OfdModelLimits { MaxMaterializedBytes = 8 };
        var material = new OfdModelMaterialization(limits); limits.MaxMaterializedBytes = 100;
        var data = new byte[8]; material.StoreArray(data); material.StoreArray(data);
        material.ChargeCopy(8, default); material.ChargeOutput(8);
        Assert.Equal(8, material.StorageBytes); Assert.Equal(8, material.CloneCopyBytes); Assert.Equal(8, material.OutputUpperBoundBytes);
        Assert.Throws<InvalidDataException>(() => material.StoreArray(new byte[1]));
        Assert.Throws<InvalidDataException>(() => material.ChargeCopy(1, default));
        Assert.Throws<InvalidDataException>(() => material.ChargeOutput(1));
    }

    [Fact]
    public void MatricesCopyPerEdgeEvenWhenPayloadSharingIsRequested()
    {
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = 96 });
        var matrix = new double[] { 1, 0, 0, 1, 0, 0 };
        var source = new OfdTextElement { Transform = matrix };
        var first = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(source, clonePayloads: false, material: material));
        var second = Assert.IsType<OfdTextElement>(OfdModelCloner.CloneElement(source, clonePayloads: false, material: material));
        Assert.NotSame(matrix, first.Transform); Assert.NotSame(first.Transform, second.Transform); Assert.Equal(96, material.CloneCopyBytes);
        Assert.Throws<InvalidDataException>(() => OfdModelCloner.CloneElement(source, clonePayloads: false, material: material));
        Assert.Equal(new double[] { 1, 0, 0, 1, 0, 0 }, matrix);
    }

    [Fact]
    public void XmlAliasesShareStreamingInspectionButPayEveryDomMaterialization()
    {
        var data = Encoding.UTF8.GetBytes("<Root Value='one two'/>");
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = data.Length * 6L + 128 });
        material.GuardEntry("first.xml", data, default); material.GuardEntry("alias.dat", data, default);
        _ = material.ReadXml("first.xml", data, default); _ = material.ReadXml("alias.dat", data, default);
        Assert.Equal(data.Length * 6L + 128, material.CloneCopyBytes);
        Assert.Throws<InvalidDataException>(() => material.ReadXml("third.xml", data, default));
    }

    [Fact]
    public void LargeNamespaceDecoderIsRejectedBeforeDomAndCancellationIsNotOpaqueFallback()
    {
        var data = Encoding.UTF8.GetBytes("<Root xmlns:v='" + new string('x', 1000) + "'/>");
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = 100 });
        Assert.Throws<InvalidDataException>(() => material.GuardEntry("large.xml", data, default));
        Assert.DoesNotContain("large.xml", material.OpaqueEntries);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => material.GuardEntry("cancel.xml", data, cancelled.Token));
        Assert.DoesNotContain("cancel.xml", material.OpaqueEntries);
    }

    [Fact]
    public void EmptyArrayFloodStillConsumesReferenceBudget()
    {
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxReferenceCount = 3 });
        for (var i = 0; i < 3; i++) material.StoreArray(Array.Empty<byte>());
        Assert.Throws<InvalidDataException>(() => material.StoreArray(Array.Empty<byte>()));
    }

    [Fact]
    public void XmlCacheCannotAuthenticateChangedBytesOrAnOpaqueToXmlChange()
    {
        var data = Encoding.UTF8.GetBytes("<Root ID='10'/>"); var material = new OfdModelMaterialization(new OfdModelLimits());
        material.GuardEntry("x.xml", data, default); data[10] = (byte)'9';
        Assert.Throws<InvalidDataException>(() => material.ReadXml("x.xml", data, default));
        Assert.Throws<InvalidDataException>(() => material.Reservations(data, out _));
        var opaque = Encoding.UTF8.GetBytes("!Root ID='10'/>"); material.GuardEntry("opaque", opaque, default);
        opaque[0] = (byte)'<'; material.GuardEntry("opaque", opaque, default);
        Assert.DoesNotContain("opaque", material.OpaqueEntries); Assert.Contains(10, material.Reservations(opaque, out _));
    }

    [Theory]
    [InlineData("<Root><Broken></Root>", true)]
    [InlineData("<!DOCTYPE Root [<!ENTITY x 'value'>]><Root>&x;</Root>", false)]
    public void MalformedExtensionsMayBeOpaqueButDtdCannot(string xml, bool opaque)
    {
        var material = new OfdModelMaterialization(new OfdModelLimits()); var data = Encoding.UTF8.GetBytes(xml);
        if (opaque) { material.GuardEntry("extension.dat", data, default); Assert.Contains("extension.dat", material.OpaqueEntries); }
        else { Assert.Throws<InvalidDataException>(() => material.GuardEntry("extension.dat", data, default)); Assert.DoesNotContain("extension.dat", material.OpaqueEntries); }
    }
}
