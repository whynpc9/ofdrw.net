using System;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Layout.Editing.Internal;

// A private-constructed real compiler expectation, never a caller assertion.
internal sealed class OfdOwnedGeometryExpectation : OfdOwnedExpectationIssuer, IDisposable
{
    private readonly OfdTextReplacementOperation operation;
    private readonly OfdTextGeometryExpectation capsule;
    private OfdOwnedTextRootKey[]? keys;
    private int compared;
    private OfdOwnedGeometryExpectation(OfdTextReplacementOperation operation, OfdTextGeometryExpectation capsule)
    { this.operation = operation; this.capsule = capsule; }
    internal static OfdOwnedGeometryExpectation FromActualCompiler(OfdTextReplacementOperation operation,
        OfdTextGeometryExpectation capsule, CancellationToken token)
    {
        capsule.AssertBinding(operation, token); operation.Context.Material.ChargeCopy(512, token);
        return new OfdOwnedGeometryExpectation(operation, capsule);
    }
    internal OfdOwnedWriteScope SealActualBatch(OfdOwnedAnalysisClaim claim, OfdOwnedHostClosure closure,
        OfdRetainedTextEditBatch actualBatch, CancellationToken token)
    {
        capsule.AssertBinding(operation, token);
        operation.Context.Material.ChargeCopy(checked(1024L + 16L * capsule.Roots.Count), token);
        var owners = new OfdTextElement[capsule.Roots.Count];
        for (var i = 0; i < owners.Length; i++) capsule.Roots[i].CopyOriginalOwner(owners, i);
        var scope = OfdOwnedSaveAuthority.SealCompiledBatch(claim, closure, this, actualBatch, owners, out var issued, token);
        BindRootKeys(issued); return scope;
    }
    internal void BindRootKeys(OfdOwnedTextRootKey[] exactIssuedKeys)
    {
        if (keys is not null || exactIssuedKeys.Length != capsule.Roots.Count) throw new InvalidOperationException("ExpectationKeysAlreadyBoundOrIncomplete");
        operation.Context.Material.ChargeCopy(checked(128L + IntPtr.Size * (long)exactIssuedKeys.Length), default);
        keys = (OfdOwnedTextRootKey[])exactIssuedKeys.Clone();
    }
    internal OfdOwnedStageScope CompareActual(OfdOwnedCompareScope scope, OfdDocumentPackage actual,
        OfdReadbackRootAddress[] completeAddresses, CancellationToken token)
    {
        // Authentication failure belongs to the scope's true owner. No foreign
        // operation/context or actual model is touched before this succeeds.
        RequireActualReadback(scope, actual);
        try
        {
            if (Interlocked.Exchange(ref compared, 1) != 0 || keys is null) throw new InvalidOperationException("OwnedComparisonAlreadyConsumed");
            OfdTextGeometryInputVerifier.CompareReceiptBoundOwned(capsule, operation, this, scope, actual, keys, completeAddresses, token);
            return CompleteActualComparison(scope, token); // only after the actual full strict comparator tail
        }
        catch { operation.Context.Fail(); throw; }
    }
    internal OfdOwnedVerifiedFontSummary CurrentFont(OfdOwnedCompareScope scope, OfdDocumentPackage actual, OfdFontResource font)
        => RequireVerifiedFontSummary(scope, actual, font);
    public void Dispose() { if (keys is not null) Array.Clear(keys, 0, keys.Length); keys = null; }
}
