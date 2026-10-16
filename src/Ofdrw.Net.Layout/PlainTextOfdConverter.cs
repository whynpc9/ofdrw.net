using System.Text;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout;

/// <summary>Converts logical plain text to native OFD using a single Flow paragraph.</summary>
/// <remarks>CRLF/CR/LF are line controls; form feed is a page control. Terminal controls do not add empty tail pages.
/// Other whitespace is preserved. Font-family names use Flow's portable metrics, without embedding or coverage guarantees.
/// Each call owns its layout state. The returned package is independent; saving it is the caller's responsibility.</remarks>
public sealed class PlainTextOfdConverter
{
    private readonly FlowDocumentOptions _layout = new();
    private readonly long _maxInputBytes;

    /// <summary>Creates a converter with A4 Flow defaults and finite budgets.</summary>
    public PlainTextOfdConverter() : this(new PlainTextOfdOptions()) { }

    /// <summary>Validates and snapshots all options; later mutations do not affect this converter.</summary>
    /// <param name="options">Page, font and resource settings.</param>
    /// <exception cref="ArgumentException">Options are invalid.</exception>
    public PlainTextOfdConverter(PlainTextOfdOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        if (options.MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxInputBytes));
        _maxInputBytes = options.MaxInputBytes;
        Copy(options.Layout, _layout);
        Positive(_layout.PageWidthMillimeters, nameof(_layout.PageWidthMillimeters));
        Positive(_layout.PageHeightMillimeters, nameof(_layout.PageHeightMillimeters));
        Positive(_layout.DefaultFontSizeMillimeters, nameof(_layout.DefaultFontSizeMillimeters));
        Nonnegative(_layout.MarginLeftMillimeters, nameof(_layout.MarginLeftMillimeters));
        Nonnegative(_layout.MarginRightMillimeters, nameof(_layout.MarginRightMillimeters));
        Nonnegative(_layout.MarginTopMillimeters, nameof(_layout.MarginTopMillimeters));
        Nonnegative(_layout.MarginBottomMillimeters, nameof(_layout.MarginBottomMillimeters));
        if (_layout.MarginLeftMillimeters + _layout.MarginRightMillimeters >= _layout.PageWidthMillimeters ||
            _layout.MarginTopMillimeters + _layout.MarginBottomMillimeters >= _layout.PageHeightMillimeters)
            throw new ArgumentException("Margins leave no usable page area.", nameof(options));
        if (string.IsNullOrWhiteSpace(_layout.DefaultFontFamily) || _layout.DefaultColor is null)
            throw new ArgumentException("A font family and color are required.", nameof(options));
        if (_layout.MaxPageCount < 1 || _layout.MaxCharacters < 1 || _layout.MaxTextElements < 1 || _layout.MaxTableCells < 1)
            throw new ArgumentOutOfRangeException(nameof(options), "Flow budgets must be positive.");
    }

    /// <summary>Renders Unicode text. A leading U+FEFF is content. Empty text produces one blank page.</summary>
    /// <param name="text">Well-formed UTF-16 containing XML 1.0 scalars, plus form feed as a layout control.</param>
    /// <param name="cancellationToken">Cooperative cancellation for validation and Flow rendering.</param>
    /// <returns>A new in-memory native OFD package.</returns>
    /// <exception cref="InvalidDataException">Invalid Unicode/XML text, a budget excess, or text which cannot fit.</exception>
    public OfdDocumentPackage Render(string text, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (text.Length > _layout.MaxCharacters) throw new InvalidDataException("Plain text exceeds MaxCharacters.");
        var pendingHigh = false;
        Validate(text, ref pendingHigh, cancellationToken);
        if (pendingHigh) throw new InvalidDataException("Plain text ends with an unpaired surrogate.");
        return RenderValidated(text, cancellationToken);
    }

    /// <summary>Reads strict UTF-8 from the stream's current position, stripping only one initial byte BOM.</summary>
    /// <remarks>Uses bounded chunks and at most one extra byte to detect a byte-budget excess. Does not seek or close the borrowed stream.
    /// Invalid UTF-8 is rejected without replacement. Failure/cancellation may advance the input position. Layout materializes the admitted text.</remarks>
    /// <param name="input">Readable, possibly non-seekable UTF-8 stream.</param>
    /// <param name="cancellationToken">Passed to each read and to Flow rendering.</param>
    /// <returns>A new in-memory native OFD package.</returns>
    /// <exception cref="InvalidDataException">Invalid UTF-8/Unicode/XML text, a budget excess, or text which cannot fit.</exception>
    public async Task<OfdDocumentPackage> RenderUtf8Async(Stream input, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (input is null) throw new ArgumentNullException(nameof(input));
        if (!input.CanRead) throw new ArgumentException("Input must be readable.", nameof(input));
        var encoding = new UTF8Encoding(false, true);
        var decoder = encoding.GetDecoder();
        var bytes = new byte[4096];
        var chars = new char[encoding.GetMaxCharCount(bytes.Length)];
        var text = new StringBuilder();
        long readTotal = 0;
        var firstCharacter = true;
        var pendingHigh = false;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remaining = _maxInputBytes - readTotal;
            var requested = remaining >= bytes.Length ? bytes.Length : (int)remaining + 1;
            var read = await input.ReadAsync(bytes, 0, requested, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (read > remaining) throw new InvalidDataException("Plain text exceeds MaxInputBytes.");
            readTotal = checked(readTotal + read);
            int count;
            try { count = decoder.GetChars(bytes, 0, read, chars, 0, flush: read == 0); }
            catch (DecoderFallbackException ex) { throw new InvalidDataException("Input is not strict UTF-8.", ex); }
            var start = 0;
            if (count > 0 && firstCharacter)
            {
                firstCharacter = false;
                if (chars[0] == '\uFEFF') start = 1;
            }
            count -= start;
            if (count > _layout.MaxCharacters - text.Length) throw new InvalidDataException("Plain text exceeds MaxCharacters.");
            var chunk = new string(chars, start, count);
            Validate(chunk, ref pendingHigh, cancellationToken);
            text.Append(chunk);
            if (read == 0) break;
        }
        if (pendingHigh) throw new InvalidDataException("Plain text ends with an unpaired surrogate.");
        cancellationToken.ThrowIfCancellationRequested();
        return RenderValidated(text.ToString(), cancellationToken);
    }

    private OfdDocumentPackage RenderValidated(string text, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        var flow = new FlowDocument();
        Copy(_layout, flow.Options);
        flow.Blocks.Add(new Paragraph(text));
        token.ThrowIfCancellationRequested();
        OfdDocumentPackage result;
        try { result = flow.Render(token); }
        catch (InvalidOperationException ex) { throw new InvalidDataException("Plain text exceeds a Flow budget or cannot fit the page.", ex); }
        token.ThrowIfCancellationRequested();
        return result;
    }

    private static void Validate(string text, ref bool pendingHigh, CancellationToken token)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if ((i & 1023) == 0) token.ThrowIfCancellationRequested();
            var c = text[i];
            if (pendingHigh)
            {
                if (!char.IsLowSurrogate(c)) throw new InvalidDataException("Unpaired UTF-16 surrogate.");
                pendingHigh = false;
            }
            else if (char.IsHighSurrogate(c)) pendingHigh = true;
            else if (char.IsLowSurrogate(c) || c is '\uFFFE' or '\uFFFF' || (c < 0x20 && c is not ('\t' or '\n' or '\r' or '\f')))
                throw new InvalidDataException("Plain text contains an invalid Unicode/XML scalar.");
        }
    }

    private static void Positive(double value, string name)
    { if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) throw new ArgumentOutOfRangeException(name); }
    private static void Nonnegative(double value, string name)
    { if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) throw new ArgumentOutOfRangeException(name); }

    private static void Copy(FlowDocumentOptions source, FlowDocumentOptions target)
    {
        target.PageWidthMillimeters = source.PageWidthMillimeters;
        target.PageHeightMillimeters = source.PageHeightMillimeters;
        target.MarginLeftMillimeters = source.MarginLeftMillimeters;
        target.MarginRightMillimeters = source.MarginRightMillimeters;
        target.MarginTopMillimeters = source.MarginTopMillimeters;
        target.MarginBottomMillimeters = source.MarginBottomMillimeters;
        target.DefaultFontFamily = source.DefaultFontFamily;
        target.DefaultFontSizeMillimeters = source.DefaultFontSizeMillimeters;
        target.DefaultColor = source.DefaultColor;
        target.MaxCharacters = source.MaxCharacters;
        target.MaxPageCount = source.MaxPageCount;
        target.MaxTextElements = source.MaxTextElements;
        target.MaxTableCells = source.MaxTableCells;
    }
}
