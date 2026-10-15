using System.Text;
using Ofdrw.Net.Converter.Docx.Converters;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit003RegressionTests
{
    [Fact]
    public async Task Audit003_PartialOutputFailure()
    {
        var attempts = new List<DocxConversionEngine>();
        var candidates = new List<string>();
        var validPdf = CreatePdf();
        var converter = CreateConverter((engine, path, _) =>
        {
            attempts.Add(engine);
            candidates.Add(path);
            File.WriteAllBytes(path, validPdf);
            return Success();
        });
        using var input = new MemoryStream(new byte[] { 1 });
        using var output = new FailingOutput();

        var failure = await Assert.ThrowsAsync<IOException>(() => converter.ConvertAsync(input, output));

        Assert.Same(output.Failure, failure);
        Assert.Equal(7, output.Length);
        Assert.Equal(new[] { DocxConversionEngine.LibreOffice }, attempts);
        Assert.All(candidates, path => Assert.False(Directory.Exists(Path.GetDirectoryName(path))));
    }

    [Fact]
    public Task Audit003_PdfPrefixOnly() => AssertInvalidCandidateFallsBack(Encoding.ASCII.GetBytes("%PDF-"));

    [Theory]
    [InlineData("%PDF-1.7\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\ntrailer << /Root 1 0 R >>\n%%EOF")]
    [InlineData("%PDF-1.7\n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n2 0 obj << /Type /Pages /Kids [] /Count 0 >> endobj\ntrailer << /Root 1 0 R >>\n%%EOF")]
    public Task Audit003_BrokenStructure(string candidate) => AssertInvalidCandidateFallsBack(Encoding.ASCII.GetBytes(candidate));

    [Theory]
    [InlineData("/Kids [] /Count 0", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>")]
    [InlineData("/Kids [3 0 R] /Count 1", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 0 842] >>")]
    [InlineData("/Kids [99 0 R] /Count 1", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>")]
    public Task Audit003_BrokenPageTree(string pages, string page)
    {
        // Keep a complete xref/trailer so these exercise page validation after opening.
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int> { 0 };
        var objects = new[] { "<< /Type /Catalog /Pages 2 0 R >>", $"<< /Type /Pages {pages} >>", page };
        for (var index = 0; index < objects.Length; index++)
        {
            offsets.Add(pdf.Length);
            pdf.Append($"{index + 1} 0 obj\n{objects[index]}\nendobj\n");
        }
        var xref = pdf.Length;
        pdf.Append("xref\n0 4\n0000000000 65535 f \n");
        foreach (var offset in offsets.Skip(1)) pdf.Append($"{offset:D10} 00000 n \n");
        pdf.Append($"trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        return AssertInvalidCandidateFallsBack(Encoding.ASCII.GetBytes(pdf.ToString()));
    }

    [Fact]
    public async Task Audit003_ValidFallback()
    {
        var expected = CreatePdf();
        var candidates = new List<string>();
        var converter = CreateConverter((engine, path, _) =>
        {
            candidates.Add(path);
            if (engine == DocxConversionEngine.LibreOffice)
                throw new InvalidOperationException("The fixture renderer could not start.");
            File.WriteAllBytes(path, expected);
            return Success();
        });
        using var input = new MemoryStream(new byte[] { 1 });
        using var output = new MemoryStream();

        var result = await converter.ConvertWithResultAsync(input, output);

        Assert.Equal(expected, output.ToArray());
        Assert.Equal(DocxConversionEngine.BuiltIn, result.ActualEngine);
        Assert.Equal(new[] { DocxConversionEngine.LibreOffice, DocxConversionEngine.BuiltIn }, result.AttemptedEngines);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_ENGINE_FAILED");
        Assert.All(candidates, path => Assert.False(Directory.Exists(Path.GetDirectoryName(path))));
    }

    [Fact]
    public async Task Audit003_CanceledOutput()
    {
        var attempts = new List<DocxConversionEngine>();
        var candidates = new List<string>();
        using var cancellation = new CancellationTokenSource();
        var converter = CreateConverter((engine, path, _) =>
        {
            attempts.Add(engine);
            candidates.Add(path);
            File.WriteAllBytes(path, CreatePdf());
            return Success();
        });
        using var input = new MemoryStream(new byte[] { 1 });
        using var output = new CancelingOutput(cancellation);

        var failure = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            converter.ConvertAsync(input, output, cancellation.Token));

        Assert.Equal(cancellation.Token, failure.CancellationToken);
        Assert.Equal(new[] { DocxConversionEngine.LibreOffice }, attempts);
        Assert.All(candidates, path => Assert.False(Directory.Exists(Path.GetDirectoryName(path))));
    }

    [Fact]
    public async Task Audit003_InputSafetyFailureDoesNotFallback()
    {
        var attempts = new List<DocxConversionEngine>();
        var rejection = new InvalidDataException("Unsafe fixture input.");
        var converter = CreateConverter((engine, _, _) =>
        {
            attempts.Add(engine);
            throw rejection;
        });
        using var input = new MemoryStream(new byte[] { 1 });
        using var output = new MemoryStream();

        Assert.Same(rejection, await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, output)));
        Assert.Equal(new[] { DocxConversionEngine.LibreOffice }, attempts);
        Assert.Equal(0, output.Length);
    }

    private static async Task AssertInvalidCandidateFallsBack(byte[] candidate)
    {
        var validPdf = CreatePdf();
        var candidates = new List<string>();
        var converter = CreateConverter((engine, path, _) =>
        {
            candidates.Add(path);
            File.WriteAllBytes(path, engine == DocxConversionEngine.LibreOffice ? candidate : validPdf);
            return Success();
        });
        using var input = new MemoryStream(new byte[] { 1 });
        using var output = new MemoryStream();

        var result = await converter.ConvertWithResultAsync(input, output);

        Assert.Equal(DocxConversionEngine.BuiltIn, result.ActualEngine);
        Assert.Equal(new[] { DocxConversionEngine.LibreOffice, DocxConversionEngine.BuiltIn }, result.AttemptedEngines);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_ENGINE_FAILED");
        Assert.Equal(validPdf, output.ToArray());
        using var pdf = PdfPigDocument.Open(output.ToArray());
        Assert.Single(pdf.GetPages());
        Assert.All(candidates, path => Assert.False(Directory.Exists(Path.GetDirectoryName(path))));
    }

    private static DocxToPdfConverter CreateConverter(
        Func<DocxConversionEngine, string, CancellationToken, Task<IReadOnlyList<DocxConversionDiagnostic>>> render) =>
        new(new DocxConversionOptions { Engine = DocxConversionEngine.Auto }, () => false, () => true, render);

    private static Task<IReadOnlyList<DocxConversionDiagnostic>> Success() =>
        Task.FromResult<IReadOnlyList<DocxConversionDiagnostic>>(Array.Empty<DocxConversionDiagnostic>());

    private static byte[] CreatePdf()
    {
        using var document = new PdfSharpCore.Pdf.PdfDocument();
        var page = document.AddPage();
        page.Width = 595;
        page.Height = 842;
        using var output = new MemoryStream();
        document.Save(output, false);
        return output.ToArray();
    }

    private sealed class FailingOutput : MemoryStream
    {
        internal IOException Failure { get; } = new("Fixture output accepted a prefix and then failed.");
        private bool _failed;

        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) =>
            new(WriteAsync(buffer.ToArray(), 0, buffer.Length, cancellationToken));

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            if (!_failed)
            {
                _failed = true;
                base.Write(buffer, offset, Math.Min(7, count));
                throw Failure;
            }
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }
    }

    private sealed class CancelingOutput(CancellationTokenSource cancellation) : MemoryStream
    {
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken cancellationToken = default) =>
            new(WriteAsync(buffer.ToArray(), 0, buffer.Length, cancellationToken));

        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            cancellation.Cancel();
            cancellationToken.ThrowIfCancellationRequested();
            return base.WriteAsync(buffer, offset, count, cancellationToken);
        }
    }
}
