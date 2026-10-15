using System;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Reader.Keywords.Internal;

internal abstract class OfdOwnedGlyphGeometryHost : OfdOwnedHostIssuer, IDisposable
{
    protected OfdOwnedGlyphGeometryHost() { }
    internal abstract void Open(OfdGlyphGeometryLimits limits, CancellationToken token);
    internal abstract OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken token);
    internal abstract OfdGlyphGeometryStatistics ReadStatistics();
    internal abstract void Close();
    public abstract void Dispose();
}
