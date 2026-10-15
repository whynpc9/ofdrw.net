using System;
using System.Threading;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Converter.Pdf.Geometry;

internal static class OfdPrivateGeometryFactory
{
    internal static OfdPrivateGeometryHost Create(OfdPrivateSaveBootstrap bootstrap,
        OfdTrustedEmbeddedFontIdentity[] concreteCopiedIdentities)
    {
        if (bootstrap is null) throw new ArgumentNullException(nameof(bootstrap));
        if (concreteCopiedIdentities is null) throw new ArgumentNullException(nameof(concreteCopiedIdentities));
        var count = concreteCopiedIdentities.Length;
        if (count > OfdTrustedEmbeddedFontGeometryProvider.MaximumTrustedIdentities)
            throw new ArgumentOutOfRangeException(nameof(concreteCopiedIdentities));
        // Bootstrap enforces its original preparation token in addition to this
        // per-call token and reserves fixed host/facade/session metadata and the
        // bounded identity dictionary/array before any copies or provider objects.
        bootstrap.ReserveHostConstruction(count, default);
        NoHooksHost? host = null;
        try
        {
            var frozen = new OfdTrustedEmbeddedFontIdentity[count];
            for (var i = 0; i < count; i++)
                frozen[i] = concreteCopiedIdentities[i] ?? throw new ArgumentException("Null trusted identity.");
            host = new NoHooksHost(frozen);
            return new OfdPrivateGeometryHost(bootstrap, host);
        }
        catch
        {
            // Bootstrap retires the ticket/context first. Cleanup cannot replace
            // construction/registration failure; the worker sees the primary.
            try { bootstrap.Dispose(); } catch { }
            try { host?.Dispose(); } catch { }
            throw;
        }
    }

    private sealed class NoHooksHost : OfdOwnedGlyphGeometryHost
    {
        private enum State { Fresh, Opening, Open, Failed, Closed, Disposed }
        private readonly OfdTrustedEmbeddedFontGeometryProvider provider;
        private OfdTrustedEmbeddedFontGeometryProvider.Session? session;
        private State state;
        internal NoHooksHost(OfdTrustedEmbeddedFontIdentity[] frozen)
            => provider = new OfdTrustedEmbeddedFontGeometryProvider(frozen);

        internal override void Open(OfdGlyphGeometryLimits limits, CancellationToken token)
        {
            if (state != State.Fresh) throw new InvalidOperationException("Owned host cannot be reopened.");
            state = State.Opening;
            try
            {
                session = (OfdTrustedEmbeddedFontGeometryProvider.Session)provider.OpenSession(limits, token);
                state = State.Open;
            }
            catch { state = State.Failed; throw; }
        }
        internal override OfdGlyphGeometry Measure(OfdGlyphGeometryRequest request, CancellationToken token)
        {
            if (state != State.Open || session is null) throw new InvalidOperationException("Owned host is not open.");
            try { return session.Measure(request, token); }
            catch { state = State.Failed; throw; }
        }
        internal override OfdGlyphGeometryStatistics ReadStatistics()
        {
            if (state != State.Open || session is null) throw new InvalidOperationException("Owned host is not open.");
            try { return session.Statistics; }
            catch { state = State.Failed; throw; }
        }
        internal override void Close()
        {
            if (state is not (State.Open or State.Failed)) throw new InvalidOperationException("Owned host is already closed or unused.");
            // Retire the actual session before its Dispose. The Reader caller
            // issues closed-host authority only after this returns and final
            // Q2/P0 + snapshot-byte comparisons succeed.
            var closing = session; session = null; state = State.Closed;
            try { closing?.Dispose(); }
            catch { state = State.Failed; throw; }
        }
        public override void Dispose()
        {
            if (state == State.Disposed) return;
            var closing = session; session = null; state = State.Disposed;
            closing?.Dispose();
        }
    }
}
