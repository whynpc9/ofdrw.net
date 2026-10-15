using System.Collections;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Packaging.Tests;

// Separately authorized narrow P0 immutable-response construction repair.
public sealed class GlyphGeometryConstructionTests
{
    [Fact]
    public void ActualPublicConstructorFreezesCountBeforeAnAdversarialListCanGrow()
    {
        var input = new ChangingCount(1, 16_000_001);
        var result = new OfdGlyphGeometry(input);
        Assert.Single(result.Clusters); Assert.Equal(1, input.CountReads); Assert.Equal(1, input.IndexReads);
    }
    [Theory]
    [InlineData(-1)]
    [InlineData(16_000_001)]
    public void InvalidInitialCountRejectsBeforeCopyOrIndexerAccess(int count)
    {
        var input = new ChangingCount(count, 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdGlyphGeometry(input));
        Assert.Equal(1, input.CountReads); Assert.Equal(0, input.IndexReads);
    }
    [Fact]
    public void NullClustersAndUnsupportedResponseValidationAreRetained()
    {
        Assert.Throws<ArgumentNullException>(() => new OfdGlyphGeometry(null!));
        Assert.Throws<ArgumentException>(() => new OfdGlyphGeometry(new OfdGlyphClusterGeometry[] { null! }));
        Assert.Throws<ArgumentException>(() => new OfdGlyphGeometry(new ChangingCount(1, 0), "reason"));
        Assert.False(new OfdGlyphGeometry(Array.Empty<OfdGlyphClusterGeometry>(), "reason").IsSupported);
    }
    private sealed class ChangingCount(int initial, int later) : IReadOnlyList<OfdGlyphClusterGeometry>
    {
        internal int CountReads, IndexReads;
        public int Count => ++CountReads == 1 ? initial : later;
        public OfdGlyphClusterGeometry this[int index]
        {
            get { IndexReads++; if (index != 0) throw new InvalidOperationException("The list grew after its initial limit check."); return new(0, 1, 1, null, 1); }
        }
        public IEnumerator<OfdGlyphClusterGeometry> GetEnumerator() => throw new NotSupportedException();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
