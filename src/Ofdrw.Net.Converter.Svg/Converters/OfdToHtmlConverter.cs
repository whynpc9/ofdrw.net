using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Ofdrw.Net.Converter.Svg.Internal;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Svg.Converters;

/// <summary>Exports the Reader's first document as bounded, static, self-contained HTML with inline SVG pages.</summary>
/// <remarks>Only supported drawing projections are admitted. No scripts, external resources or interactive actions are exported.</remarks>
public sealed class OfdToHtmlConverter
{
    private readonly OfdHtmlOptions _options;
    private readonly OfdPackageLoadOptions _load;
    private readonly Func<Stream, OfdPackageLoadOptions, CancellationToken, Task<OfdDocumentPackage>> _read;
    /// <summary>Creates an exporter with finite default limits.</summary>
    public OfdToHtmlConverter() : this(new OfdHtmlOptions()) { }
    /// <summary>Creates an exporter, copying and validating the supplied limits.</summary>
    /// <param name="options">Positive output/generation limits and a plain text title.</param>
    /// <param name="loadOptions">Independent package input limits, or defaults.</param>
    public OfdToHtmlConverter(OfdHtmlOptions options, OfdPackageLoadOptions? loadOptions = null)
        : this(options, loadOptions, (input, load, token) => new OfdReader().ReadAsync(input, load, token)) { }
    internal OfdToHtmlConverter(OfdHtmlOptions options, OfdPackageLoadOptions? loadOptions,
        Func<Stream, OfdPackageLoadOptions, CancellationToken, Task<OfdDocumentPackage>> read)
    { _options = Snapshot(options); _load = CopyLoad(loadOptions ?? new OfdPackageLoadOptions()); _read = read; }

    /// <summary>Reads one package once and writes selected page occurrences in the requested order.</summary>
    /// <param name="ofdInput">Readable input at its current position; left open.</param>
    /// <param name="htmlOutput">Writable output at its current position; left open.</param>
    /// <param name="pageIndexes">Zero-based positions in sorted Reader pages, preserving duplicates; null selects all. Empty selections fail.</param>
    /// <param name="cancellationToken">Cancellation for reading, generating, serializing and final copying.</param>
    /// <returns>A task that completes after the output has been flushed.</returns>
    /// <remarks>Generation fails before writing public output. Final copy/flush failures or cancellation can leave a prefix in arbitrary streams; discard that output. No general stream rollback is promised.</remarks>
    public async Task ConvertAsync(Stream ofdInput, Stream htmlOutput, IReadOnlyList<int>? pageIndexes = null, CancellationToken cancellationToken = default)
    {
        if (ofdInput is null) throw new ArgumentNullException(nameof(ofdInput));
        if (htmlOutput is null) throw new ArgumentNullException(nameof(htmlOutput));
        if (!ofdInput.CanRead || !htmlOutput.CanWrite || ReferenceEquals(ofdInput, htmlOutput))
            throw new ArgumentException("Distinct readable input and writable output streams are required.");
        cancellationToken.ThrowIfCancellationRequested();
        var options = Snapshot(_options);
        int[]? selected = null;
        if (pageIndexes is not null)
        {
            var count = pageIndexes.Count;
            if (count == 0 || count > options.MaxPages) throw new ArgumentException("Invalid HTML page occurrence count.", nameof(pageIndexes));
            selected = new int[count];
            for (var i = 0; i < selected.Length; i++) { cancellationToken.ThrowIfCancellationRequested(); selected[i] = pageIndexes[i]; }
        }
        var package = await _read(ofdInput, CopyLoad(_load), cancellationToken).ConfigureAwait(false);
        var pages = package.Pages.OrderBy(p => p.Index).ToArray();
        if (pages.Length == 0) throw new InvalidDataException("HTML preview requires at least one page.");
        if (selected is null)
        {
            if (pages.Length > options.MaxPages) throw new InvalidDataException("HTML preview capacity exceeded: pages.");
            selected = Enumerable.Range(0, pages.Length).ToArray();
        }
        foreach (var index in selected)
            if (index < 0 || index >= pages.Length) throw new ArgumentOutOfRangeException(nameof(pageIndexes));
        var material = new OfdModelMaterialization(new OfdModelLimits());
        OfdCompositeGraph.Validate(package, material, cancellationToken);
        OfdAnnotationContract.Validate(package, material, cancellationToken);
        var context = new HtmlRenderContext(options, cancellationToken);
        var fontBudget = new FontParseBudget(FontParseBudget.DefaultMaximumWork, cancellationToken);
        using var spool = new HtmlSpool(options, cancellationToken);
        using (var writer = XmlWriter.Create(spool, new XmlWriterSettings { Encoding = new UTF8Encoding(false, true), OmitXmlDeclaration = true, CloseOutput = false, CheckCharacters = true }))
        {
            writer.WriteDocType("html", null, null, null);
            writer.WriteStartElement("html"); writer.WriteAttributeString("lang", "zh-CN"); writer.WriteStartElement("head");
            Meta(writer, "charset", "utf-8");
            writer.WriteStartElement("meta"); writer.WriteAttributeString("http-equiv", "Content-Security-Policy");
            writer.WriteAttributeString("content", "default-src 'none'; script-src 'none'; connect-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'"); writer.WriteEndElement();
            writer.WriteStartElement("meta"); writer.WriteAttributeString("name", "viewport"); writer.WriteAttributeString("content", "width=device-width, initial-scale=1"); writer.WriteEndElement();
            writer.WriteElementString("title", options.Title);
            writer.WriteElementString("style", "body{margin:0;background:#eee;color:#222;font-family:sans-serif}main{padding:16px}section{margin:0 auto 24px;width:fit-content;max-width:100%}h2{font-size:16px;font-weight:normal}svg{display:block;background:white;max-width:100%;height:auto;font-family:initial} @media print{body{background:white}main{padding:0}section{margin:0;break-after:page;page-break-after:always}section:last-child{break-after:auto;page-break-after:auto}h2{display:none}}");
            writer.WriteEndElement(); writer.WriteStartElement("body"); writer.WriteStartElement("main"); writer.WriteAttributeString("id", "ofd-preview-root");
            writer.Flush();
            for (var occurrence = 0; occurrence < selected.Length; occurrence++)
            {
                cancellationToken.ThrowIfCancellationRequested(); context.Occurrence(occurrence);
                var svg = OfdToSvgConverter.BuildSvg(package, pages[selected[occurrence]], material, fontBudget, cancellationToken, context);
                SvgPreviewPolicy.Validate(svg, context.Prefix, context.FontPrefix, cancellationToken);
                spool.BeginPage();
                writer.WriteStartElement("section"); writer.WriteAttributeString("id", "ofd-preview-page-" + occurrence);
                writer.WriteAttributeString("data-source-page", (selected[occurrence] + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteElementString("h2", $"第 {selected[occurrence] + 1} 页（预览 {occurrence + 1}）");
                svg.WriteTo(writer); writer.WriteFullEndElement(); writer.Flush(); spool.EndPage();
            }
            writer.WriteFullEndElement(); writer.WriteFullEndElement(); writer.WriteFullEndElement(); writer.Flush();
        }
        cancellationToken.ThrowIfCancellationRequested();
        await spool.CopyResultAsync(htmlOutput, cancellationToken).ConfigureAwait(false);
    }
    private static void Meta(XmlWriter writer, string name, string value)
    { writer.WriteStartElement("meta"); writer.WriteAttributeString(name, value); writer.WriteEndElement(); }
    private static OfdHtmlOptions Snapshot(OfdHtmlOptions value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        var copy = new OfdHtmlOptions { Title = value.Title, MaxTitleCharacters = value.MaxTitleCharacters, MaxPages = value.MaxPages,
            MaxOutputBytes = value.MaxOutputBytes, MaxPageBytes = value.MaxPageBytes, MaxGeneratedNodes = value.MaxGeneratedNodes,
            MaxGeneratedCharacters = value.MaxGeneratedCharacters, MaxEmbeddedResourceBytes = value.MaxEmbeddedResourceBytes, MaxDecodedImagePixels = value.MaxDecodedImagePixels };
        if (copy.Title is null || copy.MaxTitleCharacters <= 0 || copy.Title.Length > copy.MaxTitleCharacters || copy.MaxPages <= 0 ||
            copy.MaxOutputBytes <= 0 || copy.MaxPageBytes <= 0 || copy.MaxPageBytes > copy.MaxOutputBytes || copy.MaxGeneratedNodes <= 0 ||
            copy.MaxGeneratedCharacters <= 0 || copy.MaxEmbeddedResourceBytes <= 0 || copy.MaxDecodedImagePixels <= 0)
            throw new ArgumentException("HTML limits must be positive; page bytes cannot exceed output bytes; title must fit its limit.", nameof(value));
        try { HtmlRenderContext.Characters(copy.Title); } catch (InvalidDataException ex) { throw new ArgumentException("Invalid HTML title.", nameof(value), ex); }
        return copy;
    }
    private static OfdPackageLoadOptions CopyLoad(OfdPackageLoadOptions value) => new()
    { MaxInputBytes = value.MaxInputBytes, MaxEntryCount = value.MaxEntryCount, MaxPageCount = value.MaxPageCount,
        MaxAnnotationObjectCount = value.MaxAnnotationObjectCount, MaxAnnotationXmlBytes = value.MaxAnnotationXmlBytes,
        MaxEntryUncompressedBytes = value.MaxEntryUncompressedBytes, MaxTotalUncompressedBytes = value.MaxTotalUncompressedBytes, MaxCompressionRatio = value.MaxCompressionRatio };
}
