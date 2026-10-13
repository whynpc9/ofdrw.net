using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Reader.Keywords.Internal;

internal sealed class OfdKeywordLocatorBudget(OfdKeywordLocatorLimits limits, OfdKeywordOperation operation,
    Action<string>? beforeAllocation = null)
{
    internal OfdKeywordLocatorLimits Limits => limits;
    internal long Nodes, References, XmlBytes, TextUnits, DeltaValues, DeltaTokens, Fragments, Diagnostics;
    internal long ValidationFields, ValidationCharacters, ValidationBytes;
    internal long ReservedGlyphs, ReturnedGlyphs;
    internal long CorrelationVisits;
    internal long ReusedGeometryClusters;
    internal long RetainedFontDigestBytes, RetainedFontDigestOperations;
    internal void Check() => operation.Check();
    internal void Copy(long bytes) => operation.Material?.ChargeCopy(bytes, operation.Token);
    internal void ReserveCopy(string phase, long bytes) { Copy(bytes); beforeAllocation?.Invoke(phase); }
    internal void Work(long bytes) => operation.Work(bytes);
    internal void Correlate(int keyCharacters = 0)
    { Charge(ref CorrelationVisits, 1, limits.MaxValidationFields, "correlation visits"); Work(32L + 2L * keyCharacters); }
    private void Charge(ref long counter, long count, long maximum, string kind)
    {
        Check();
        if (count < 0 || count > maximum - counter) throw new InvalidDataException("Keyword locator " + kind + " capacity exceeded.");
        counter += count;
    }
    internal void Text(int length) => Charge(ref TextUnits, length, operation.Options.Limits.MaxTextUtf16Units, "source text");
    internal void Delta(int count) => Charge(ref DeltaValues, count, limits.MaxDeltaValues, "expanded deltas");
    internal void DeltaToken() => Charge(ref DeltaTokens, 1, limits.MaxDeltaTokens, "delta tokens");
    internal void Fragment(int count) => Charge(ref Fragments, count, limits.MaxFragments, "fragments");
    internal void Diagnostic() => Charge(ref Diagnostics, 1, limits.MaxDiagnostics, "diagnostics");
    internal void Field(long count = 1) { Charge(ref ValidationFields, count, limits.MaxValidationFields, "source comparisons"); Work(checked(16 * count)); }
    internal void Characters(int count) { Charge(ref ValidationCharacters, count, limits.MaxValidationCharacters, "string comparisons"); Work(2L * count); }
    internal void Bytes(int count) { Charge(ref ValidationBytes, count, limits.MaxValidationBytes, "entry comparisons"); Work(count); }
    internal void Reference() => Charge(ref References, 1, limits.MaxReferences, "admission references");
    internal void ReserveReferences(int count) => Charge(ref References, count, limits.MaxReferences, "admission references");
    internal void ReserveGlyphs(int count) => Charge(ref ReservedGlyphs, count,
        Math.Min(operation.Options.Limits.MaxClusterVisits, operation.Options.GeometryLimits.MaxGlyphOperations), "reserved scalar glyphs");
    internal void Returned(int count) => Charge(ref ReturnedGlyphs, count,
        Math.Min(operation.Options.Limits.MaxClusterVisits, operation.Options.GeometryLimits.MaxGlyphOperations), "returned clusters");
    internal void Reused(int count) => Charge(ref ReusedGeometryClusters, count,
        Math.Min(operation.Options.Limits.MaxClusterVisits, operation.Options.GeometryLimits.MaxGlyphOperations), "reused geometry clusters");

    // First pass reserves XML bytes/nodes/references before any DOM construction.
    internal XElement Parse(string value)
    {
        Check();
        if (value.Length > limits.MaxSingleXmlBytes) throw new InvalidDataException("Control XML character capacity exceeded.");
        var byteCount = Utf8ByteCount(value);
        if (byteCount > limits.MaxSingleXmlBytes) throw new InvalidDataException("Control XML byte capacity exceeded.");
        Charge(ref XmlBytes, byteCount, limits.MaxXmlBytes, "control XML bytes");
        // Reserve streaming decoder/string work before XmlReader, then reserve
        // the measured node/attribute DOM bound before constructing the DOM.
        Copy(12L * value.Length + 8192); Work(2L * value.Length);
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = limits.MaxSingleXmlBytes, IgnoreComments = false };
        try
        {
            var beforeNodes = Nodes; var beforeReferences = References;
            using (var reader = XmlReader.Create(new StringReader(value), settings))
                while (reader.Read())
                {
                    Check();
                    if (reader.Depth >= limits.MaxDepth) throw new InvalidDataException("Control XML depth capacity exceeded.");
                    Charge(ref Nodes, 1, limits.MaxNodes, "admission nodes");
                    for (var a = 0; a < reader.AttributeCount; a++) Reference();
                }
            Check();
            Copy(8L * value.Length + 256L * (Nodes - beforeNodes) + 128L * (References - beforeReferences));
            using var domReader = XmlReader.Create(new StringReader(value), settings);
            return XElement.Load(domReader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException exception) { throw new InvalidDataException("Malformed control XML.", exception); }
    }
    private long Utf8ByteCount(string value)
    {
        // Count with one bounded reusable buffer and scalar-aligned chunks so a
        // supplementary scalar crossing a chunk remains intact. The payload
        // copy and scan are prepaid, with cancellation between chunks.
        Copy(16_896L + 2L * value.Length); Work(2L * value.Length);
        var characters = new char[8192]; var encoding = new UTF8Encoding(false, true); long bytes = 0;
        for (var start = 0; start < value.Length;)
        {
            Check(); var count = Math.Min(characters.Length, value.Length - start);
            if (start + count < value.Length && char.IsHighSurrogate(value[start + count - 1])) count--;
            value.CopyTo(start, characters, 0, count); start += count;
            bytes = checked(bytes + encoding.GetByteCount(characters, 0, count));
        }
        Check(); return bytes;
    }
}
