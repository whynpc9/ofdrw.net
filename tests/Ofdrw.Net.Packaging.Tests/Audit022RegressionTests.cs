using System.Collections;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Audit022RegressionTests
{
    [Fact]
    public void Audit022_PreCanceled()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var sources = new CancelingSources(cancellation, cancelBeforeSecond: false);
        var exception = Assert.Throws<OperationCanceledException>(() => Merge(sources, null, cancellation.Token));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(0, sources.MoveNextCalls);
        Assert.Equal(0, sources.CurrentReads);
    }

    [Fact]
    public void Audit022_CancelBetweenSources()
    {
        using var cancellation = new CancellationTokenSource();
        var sources = new CancelingSources(cancellation, cancelBeforeSecond: true);
        var exception = Assert.Throws<OperationCanceledException>(() => Merge(sources, null, cancellation.Token));
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(2, sources.MoveNextCalls);
        Assert.Equal(1, sources.CurrentReads);
        Assert.True(sources.Disposed);
    }

    [Fact]
    public void Audit022_DiagnosticsControl()
    {
        var source = Source();
        source.Pages[0].Elements.Add(new OfdRawElement { ObjectId = "raw", Xml = "<Unknown />" });
        var result = Merge([source], new OfdDocumentMergeOptions { SkipUnsupportedRawElements = true }, CancellationToken.None);
        Assert.Contains("object raw", Assert.Single(result.Diagnostics));
        Assert.Equal("CONTROL", Assert.IsType<OfdTextElement>(Assert.Single(Assert.Single(result.Package.Pages).Elements)).Text);
        Assert.Throws<NotSupportedException>(() => Merge([source], null, CancellationToken.None));
    }

    [Fact]
    public void Audit022_OldOverload()
    {
        Assert.Single<OfdPage>(OfdDocumentMerger.MergeWithResult([Source()]).Package.Pages);
        Assert.Single<OfdPage>(OfdDocumentMerger.MergeWithResult([Source()], null).Package.Pages);
        Assert.Single<OfdPage>(OfdDocumentMerger.MergeWithResult([Source()], new OfdDocumentMergeOptions()).Package.Pages);
    }

    private static OfdDocumentMergeResult Merge(IEnumerable<OfdDocumentPackage> sources, OfdDocumentMergeOptions? options, CancellationToken token)
        => OfdDocumentMerger.MergeWithResult(sources, options, token);

    private static OfdDocumentPackage Source()
    {
        var source = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { Text = "CONTROL" });
        source.Pages.Add(page);
        return source;
    }

    private sealed class CancelingSources(CancellationTokenSource cancellation, bool cancelBeforeSecond)
        : IEnumerable<OfdDocumentPackage>, IEnumerator<OfdDocumentPackage>
    {
        public int MoveNextCalls { get; private set; }
        public int CurrentReads { get; private set; }
        public bool Disposed { get; private set; }
        public OfdDocumentPackage Current
        {
            get
            {
                CurrentReads++;
                if (MoveNextCalls == 2) throw new InvalidOperationException("Canceled second source was accessed.");
                return Source();
            }
        }
        object IEnumerator.Current => Current;
        public IEnumerator<OfdDocumentPackage> GetEnumerator() => this;
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public bool MoveNext()
        {
            MoveNextCalls++;
            if (MoveNextCalls == 2 && cancelBeforeSecond) cancellation.Cancel();
            return MoveNextCalls <= (cancelBeforeSecond ? 2 : 1);
        }
        public void Reset() => throw new NotSupportedException();
        public void Dispose() => Disposed = true;
    }
}
