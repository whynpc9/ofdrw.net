using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Reader.Keywords.Internal;

internal sealed class OfdKeywordLocatorBudget(OfdKeywordLocatorLimits limits, OfdKeywordOperation operation)
{
    internal OfdKeywordLocatorLimits Limits => limits;
    internal long Nodes, References, XmlBytes, TextUnits, DeltaValues, DeltaTokens, Fragments, Diagnostics;
    internal long ValidationFields, ValidationCharacters, ValidationBytes;
    internal long ReservedGlyphs, ReturnedGlyphs;
    internal void Check() => operation.Check();
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
    internal void Field(long count = 1) => Charge(ref ValidationFields, count, limits.MaxValidationFields, "source comparisons");
    internal void Characters(int count) => Charge(ref ValidationCharacters, count, limits.MaxValidationCharacters, "string comparisons");
    internal void Bytes(int count) => Charge(ref ValidationBytes, count, limits.MaxValidationBytes, "entry comparisons");
    internal void Reference() => Charge(ref References, 1, limits.MaxReferences, "admission references");
    internal void ReserveReferences(int count) => Charge(ref References, count, limits.MaxReferences, "admission references");
    internal void ReserveGlyphs(int count) => Charge(ref ReservedGlyphs, count,
        Math.Min(operation.Options.Limits.MaxClusterVisits, operation.Options.GeometryLimits.MaxGlyphOperations), "reserved scalar glyphs");
    internal void Returned(int count) => Charge(ref ReturnedGlyphs, count,
        Math.Min(operation.Options.Limits.MaxClusterVisits, operation.Options.GeometryLimits.MaxGlyphOperations), "returned clusters");

    // First pass reserves XML bytes/nodes/references before any DOM construction.
    internal XElement Parse(string value)
    {
        Check();
        if (value.Length > limits.MaxSingleXmlBytes) throw new InvalidDataException("Control XML character capacity exceeded.");
        var byteCount = new UTF8Encoding(false, true).GetByteCount(value);
        if (byteCount > limits.MaxSingleXmlBytes) throw new InvalidDataException("Control XML byte capacity exceeded.");
        Charge(ref XmlBytes, byteCount, limits.MaxXmlBytes, "control XML bytes");
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = limits.MaxSingleXmlBytes, IgnoreComments = false };
        try
        {
            using (var reader = XmlReader.Create(new StringReader(value), settings))
                while (reader.Read())
                {
                    Check();
                    if (reader.Depth >= limits.MaxDepth) throw new InvalidDataException("Control XML depth capacity exceeded.");
                    Charge(ref Nodes, 1, limits.MaxNodes, "admission nodes");
                    for (var a = 0; a < reader.AttributeCount; a++) Reference();
                }
            Check();
            using var domReader = XmlReader.Create(new StringReader(value), settings);
            return XElement.Load(domReader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException exception) { throw new InvalidDataException("Malformed control XML.", exception); }
    }
}
