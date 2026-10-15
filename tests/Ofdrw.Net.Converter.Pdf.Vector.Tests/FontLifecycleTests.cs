using System.Reflection;
using Ofdrw.Net.Converter.Pdf.Vector;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Parser;
using UglyToad.PdfPig.Tokenization.Scanner;
using UglyToad.PdfPig.Tokens;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class FontLifecycleTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));
    [Fact]
    public void NumericAdjustmentsUseHorizontalScalingAndGraphicsRestoreWithoutLoadingAFont()
    {
        using var fixture = new PdfFixture(Font); using var document = PdfDocument.Open(fixture.Create(new[] { "" }));
        document.AddPageFactory<VectorPageContext, VectorContextFactory>(); var context = document.GetPage<VectorPageContext>(1);
        var parser = new PageContentParser(ReflectionGraphicsStateOperationFactory.Instance, new StackDepthGuard(64), false);
        var metadataReads = 0;
        using var processor = new VectorStreamProcessor(context, parser, 420, 595, new(), new(), CancellationToken.None, 1,
            _ => throw new Exception("Numeric TJ unexpectedly loaded a font program."), _ => metadataReads++);
        processor.BeginText(); processor.SetFontAndSize(NameToken.Create("F1"), 12); processor.SetHorizontalScaling(80);
        processor.TextMatrices.TextMatrix = TransformationMatrix.GetTranslationMatrix(60, 90);
        processor.TextMatrices.TextLineMatrix = processor.TextMatrices.TextMatrix;
        processor.PushState(); processor.SetFontAndSize(NameToken.Create("F2"), 24); processor.SetHorizontalScaling(50);
        processor.ShowPositionedText(new IToken[] { new NumericToken(100), new StringToken(""), new NumericToken(-50), new NumericToken(25) });
        Assert.Equal(59.1, processor.TextMatrices.TextMatrix.E, 9);
        processor.PopState(); processor.ShowPositionedText(new IToken[] { new NumericToken(100) });
        Assert.Equal(58.14, processor.TextMatrices.TextMatrix.E, 9);
        Assert.Equal(60, processor.TextMatrices.TextLineMatrix.E); Assert.Equal(90, processor.TextMatrices.TextLineMatrix.F);
        Assert.Equal(2, metadataReads); Assert.Empty(processor.OutputFonts); processor.EndText();
    }
    [Theory]
    [InlineData(1, false)] [InlineData(2, false)]
    [InlineData(1, true)] [InlineData(2, true)]
    public void ActualStockLoaderFailureOrCancellationBalancesAllResourceOverlays(int failAt, bool cancel)
    {
        using var fixture = new PdfFixture(Font); var a = fixture.Hex("A");
        var body = $"BT /F1 12 Tf 1 0 0 1 30 90 Tm <{a}> Tj /F2 12 Tf <{a}> Tj ET";
        var second = "<< /Type /Font /Subtype /Type0 /BaseFont /FixtureFace /Encoding /Identity-H /DescendantFonts [4 0 R] /ToUnicode 6 0 R >>";
        using var document = PdfDocument.Open(fixture.Create(new[] { body }, fontResources: "/F1 3 0 R /F2 11 0 R", trailingObjects: new[] { second }));
        document.AddPageFactory<VectorPageContext, VectorContextFactory>(); var context = document.GetPage<VectorPageContext>(1);
        var store = context.Resources; var field = store.GetType().GetField("fontFactory", BindingFlags.Instance | BindingFlags.NonPublic)!;
        var original = field.GetValue(store)!;
        var proxy = DispatchProxy.Create(field.FieldType, typeof(FaultingFontFactory));
        using var cancellation = new CancellationTokenSource();
        var fault = (FaultingFontFactory)(object)proxy; fault.Target = original; fault.FailAt = failAt;
        if (cancel) fault.CancelAfterGet = cancellation.Cancel;
        var stacks = new[] { "namedColorSpaces", "currentFontState", "currentXObjectState", "extendedGraphicsStates" };
        int Depth(string name)
        {
            var stack = store.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(store)!;
            var values = (System.Collections.ICollection)stack.GetType().GetField("values", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(stack)!;
            return values.Count;
        }
        var before = stacks.Select(Depth).ToArray(); field.SetValue(store, proxy);
        try
        {
            if (cancel) Assert.Throws<OperationCanceledException>(() => context.Convert(new(), cancellation.Token));
            else
            {
                var error = Assert.Throws<IOException>(() => context.Convert(new(), cancellation.Token));
                Assert.Contains("injected stock font loader failure", error.Message);
            }
            Assert.Equal(failAt, fault.Calls); Assert.Equal(before, stacks.Select(Depth));
        }
        finally { field.SetValue(store, original); }
    }
    public class FaultingFontFactory : DispatchProxy
    {
        public object Target = null!; public int FailAt; public int Calls; public Action? CancelAfterGet;
        protected override object? Invoke(MethodInfo? method, object?[]? args)
        {
            var selected = method!.Name == "Get" && ++Calls == FailAt;
            if (selected && CancelAfterGet is null) throw new IOException("injected stock font loader failure");
            try
            {
                var result = method.Invoke(Target, args);
                if (selected) CancelAfterGet!();
                return result;
            }
            catch (TargetInvocationException error) when (error.InnerException is not null) { throw error.InnerException; }
        }
    }
}
