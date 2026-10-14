using System.Globalization;
using Ofdrw.Net.Converter.Pdf.Geometry;

internal sealed class TextReplacementArguments
{
    private const int MaxSelectedPages = 1_000;
    internal required string InputPath { get; init; }
    internal required string OutputPath { get; init; }
    internal required string Keyword { get; init; }
    internal required string Replacement { get; init; }
    internal IReadOnlyList<int>? Pages { get; init; }
    internal required IReadOnlyList<OfdTrustedEmbeddedFontIdentity> TrustedFonts { get; init; }

    internal static TextReplacementArguments Parse(string[] args)
    {
        if (args.Length < 2 || args.Length > 2 + 2 * (3 + OfdTrustedEmbeddedFontGeometryProvider.MaximumTrustedIdentities))
            throw new ArgumentException("Input/output paths and bounded replacement options are required.");
        if (string.IsNullOrWhiteSpace(args[0]) || string.IsNullOrWhiteSpace(args[1]))
            throw new ArgumentException("Input and output paths cannot be empty.");
        string? keyword = null, replacement = null, pages = null;
        var trusted = new List<OfdTrustedEmbeddedFontIdentity>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var i = 2; i < args.Length; i += 2)
        {
            if (i + 1 >= args.Length) throw new ArgumentException("Missing replacement option value.");
            var option = args[i]; var value = args[i + 1];
            if (option != "--trusted-font" && !seen.Add(option)) throw new ArgumentException("Duplicate replacement option.");
            switch (option)
            {
                case "--find": keyword = value; break;
                case "--replace": replacement = value; break;
                case "--pages": pages = value; break;
                case "--trusted-font":
                    if (trusted.Count == OfdTrustedEmbeddedFontGeometryProvider.MaximumTrustedIdentities)
                        throw new ArgumentException("Too many trusted original-font identities.");
                    var colon = value.IndexOf(':');
                    if (colon != 64 || value.Length > 76 || !int.TryParse(value.AsSpan(colon + 1), NumberStyles.None, CultureInfo.InvariantCulture, out var length))
                        throw new ArgumentException("Trusted font must be SHA256:BYTECOUNT.");
                    trusted.Add(new OfdTrustedEmbeddedFontIdentity(value[..colon], length)); break;
                default: throw new ArgumentException("Unknown replacement option: " + option);
            }
        }
        if (string.IsNullOrEmpty(keyword)) throw new ArgumentException("A nonempty --find value is required.");
        if (replacement is null) throw new ArgumentException("--replace is required; an empty value deletes matching text.");
        if (keyword.Length > 1_000_000 || replacement.Length > 1_000_000)
            throw new ArgumentException("Replacement text capacity exceeded.");
        if (trusted.Count == 0) throw new ArgumentException("At least one explicitly vetted --trusted-font identity is required.");
        // Validate conflicting duplicate digests before opening input/staging output.
        _ = new OfdTrustedEmbeddedFontGeometryProvider(trusted);
        var comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var input = Path.GetFullPath(args[0]); var output = Path.GetFullPath(args[1]);
        if (string.Equals(input, output, comparison)) throw new ArgumentException("Replacement input and output must use different paths.");
        return new() { InputPath = input, OutputPath = output, Keyword = keyword, Replacement = replacement,
            Pages = pages is null ? null : ParsePages(pages), TrustedFonts = Array.AsReadOnly(trusted.ToArray()) };
    }
    private static IReadOnlyList<int> ParsePages(string value)
    {
        if (value.Length == 0 || value.Length > 32 * MaxSelectedPages) throw new ArgumentException("Invalid replacement page selection.");
        var positions = new List<int>(); var seen = new HashSet<int>();
        foreach (var part in value.Split(','))
        {
            var range = part.Split('-');
            if (range.Length is < 1 or > 2 || !int.TryParse(range[0], NumberStyles.None, CultureInfo.InvariantCulture, out var first) || first <= 0)
                throw new ArgumentException("Page numbers must be positive 1-based integers.");
            var last = first;
            if (range.Length == 2 && (!int.TryParse(range[1], NumberStyles.None, CultureInfo.InvariantCulture, out last) || last < first))
                throw new ArgumentException("Invalid replacement page range.");
            var count = (long)last - first + 1;
            if (count > MaxSelectedPages - positions.Count) throw new ArgumentException("Replacement selected-page capacity exceeded.");
            for (var offset = 0; offset < count; offset++)
            {
                var index = first - 1 + offset;
                if (!seen.Add(index)) throw new ArgumentException("Replacement page positions must be distinct.");
                positions.Add(index);
            }
        }
        return Array.AsReadOnly(positions.ToArray());
    }
}
