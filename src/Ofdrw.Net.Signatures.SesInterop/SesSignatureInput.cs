using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Signatures.SesInterop;

internal static class SesSignatureInput
{
    internal static DateTime Validate(byte[] xml, string property, bool allowJavaLocalTime = false)
    {
        if (xml is null) throw new ArgumentNullException(nameof(xml));
        if (property is null) throw new ArgumentNullException(nameof(property));
        if (xml.Length == 0 || xml.Length > 1024 * 1024) throw new InvalidDataException("Signature XML size limit.");
        if (property.Length == 0 || property.Length > 1024 || property[0] != '/' ||
            property.Any(c => c < 33 || c > 126 || c == '\\') ||
            property.Split('/').Skip(1).Any(s => s.Length == 0 || s == "." || s == "..") ||
            !property.EndsWith("/Signature.xml", StringComparison.Ordinal))
            throw new InvalidDataException("Expected a bounded absolute Signature.xml package path.");
        using var stream = new MemoryStream(xml, writable: false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = 1024 * 1024
        });
        var document = XDocument.Load(reader);
        XNamespace ns = "http://www.ofdspec.org/2016";
        if (document.Root?.Name != ns + "Signature") throw new InvalidDataException("Unexpected signature namespace/root.");
        XElement Unique(string name, XElement parent)
        {
            var all = document.Descendants().Where(e => e.Name.LocalName == name).ToArray();
            if (all.Length != 1 || all[0].Name != ns + name || all[0].Parent != parent || all[0].HasElements)
                throw new InvalidDataException("Ambiguous signature field: " + name);
            return all[0];
        }
        var infos = document.Descendants().Where(e => e.Name.LocalName == "SignedInfo").ToArray();
        if (infos.Length != 1 || infos[0].Name != ns + "SignedInfo" || infos[0].Parent != document.Root)
            throw new InvalidDataException("Ambiguous SignedInfo.");
        if (Unique("SignatureMethod", infos[0]).Value != SesTestCrypto.Algorithm)
            throw new InvalidDataException("Expected SM2-with-SM3 method.");
        var text = Unique("SignatureDateTime", infos[0]).Value;
        if (!DateTime.TryParseExact(text, "yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var time))
        {
            // Java OFDRW writes a timezone-less local wall clock in Signature.xml.
            // Its value is a signed assertion, not a UTC timestamp. Use the current UTC
            // instant when producing a new V5 TBS; never infer an offset from the text.
            if (!allowJavaLocalTime || !DateTime.TryParseExact(text, "yyyy-MM-dd HH:mm:ss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                throw new InvalidDataException("Expected a supported second-resolution signature time.");
            time = DateTime.UtcNow;
        }
        return time;
    }
}
