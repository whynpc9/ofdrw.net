using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Xunit;

namespace Ofdrw.Net.Core.Tests;

/// <summary>
/// Package-level options are copied in one place. A new option must be carried by <see cref="OfdDocumentOptions"/>'s
/// copy and fingerprinted by the serialized inventory, or these tests fail; before, each copy site (controlled edit,
/// merger, splitter, DOCX raster adapter) had to be edited separately and they had already drifted.
/// </summary>
public sealed class OfdDocumentOptionsCopyTests
{
    private static IEnumerable<(object Owner, PropertyInfo Property)> Settable(OfdDocumentOptions options) =>
        new object[] { options, options.FontEmbedding, options.Metadata }.SelectMany(owner => owner.GetType()
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanWrite && p.Name is not ("FontEmbedding" or "Metadata" or "CustomDatas"))
            .Select(p => (owner, p)));

    private static object Distinct(Type type, int seed)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;
        if (type == typeof(string)) return "v" + seed;
        if (type == typeof(bool)) return true;
        if (type == typeof(int)) return 1000 + seed;
        if (type == typeof(long)) return 100_000L + seed;
        if (type == typeof(double)) return 123.5 + seed;
        if (type == typeof(DateTimeOffset)) return new DateTimeOffset(2001, 2, 3, 4, 5, seed % 60, TimeSpan.Zero);
        if (type.IsEnum) return Enum.GetValues(type).Cast<object>().Last();
        throw new NotSupportedException("Add a distinct sample for " + type);
    }

    private static OfdDocumentOptions Populated()
    {
        var options = new OfdDocumentOptions();
        var seed = 0;
        foreach (var (owner, property) in Settable(options)) property.SetValue(owner, Distinct(property.PropertyType, ++seed));
        return options;
    }

    [Fact]
    public void CopyCarriesEveryOptionIntoNewObjects()
    {
        var source = Populated();
        var copy = source.Copy();
        Assert.NotSame(source.FontEmbedding, copy.FontEmbedding);
        Assert.NotSame(source.Metadata, copy.Metadata);
        foreach (var ((owner, property), (copyOwner, _)) in Settable(source).Zip(Settable(copy), (a, b) => (a, b)))
            Assert.True(Equals(property.GetValue(owner), property.GetValue(copyOwner)), property.DeclaringType!.Name + "." + property.Name);
    }

    [Fact]
    public void ANewDocumentDoesNotInheritTheSourceIdentity()
    {
        var source = Populated();
        var copy = source.CopyForNewDocument();
        Assert.Equal(new OfdDocumentOptions().DocumentId, copy.DocumentId);
        Assert.Null(copy.Metadata.DocId);
        Assert.Equal(source.LinearizedEntryOrder, copy.LinearizedEntryOrder);
        Assert.Equal(source.FontEmbedding.Mode, copy.FontEmbedding.Mode);
    }

    [Fact]
    public void TheSerializedInventoryFingerprintsEveryOption()
    {
        foreach (var (_, property) in Settable(new OfdDocumentOptions()))
        {
            var baseline = Fingerprint(new OfdDocumentOptions());
            var changed = new OfdDocumentOptions();
            var owner = Settable(changed).Single(x => x.Property.Name == property.Name && x.Property.DeclaringType == property.DeclaringType).Owner;
            var value = Distinct(property.PropertyType, 7);
            if (Equals(value, property.GetValue(owner))) value = value is bool flag ? !flag : throw new InvalidOperationException(property.Name);
            property.SetValue(owner, value);
            Assert.False(baseline.SequenceEqual(Fingerprint(changed)), property.DeclaringType!.Name + "." + property.Name + " is not in the inventory");
        }
    }

    private static List<string> Fingerprint(OfdDocumentOptions options)
    {
        var sink = new Sink();
        OfdSerializedInventory.Walk(new OfdDocumentPackage { Options = options }, sink,
            new OfdModelMaterialization(new OfdModelLimits()), CancellationToken.None);
        return sink.Items;
    }

    private sealed class Sink : OfdSerializedInventory.ISink
    {
        internal List<string> Items { get; } = new List<string>();
        public void Item(string role, object? value, bool identity = false) { if (!identity) Items.Add(role + "=" + value); }
        public void Bytes(string role, byte[] value, object? owner = null) => Items.Add(role + "#" + value.Length);
        public void Matrix(double[]? value) => Items.Add("m");
        public void Xml(string? value) => Items.Add("x=" + value);
        public void Path(string? value) => Items.Add("p=" + value);
    }
}
